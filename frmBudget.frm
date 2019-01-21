VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmManpower 
   Caption         =   "Budgeted Manpower"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   11415
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmBudget.frx":0000
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "frmBudget.frx":0014
      TabIndex        =   39
      Top             =   120
      Width           =   11025
   End
   Begin VB.TextBox Updstats 
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      DataField       =   "BD_LUSER"
      Height          =   285
      Index           =   2
      Left            =   6840
      TabIndex        =   54
      Text            =   "LUSER"
      Top             =   10080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Updstats 
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      DataField       =   "BD_LTIME"
      Height          =   285
      Index           =   1
      Left            =   5760
      TabIndex        =   53
      Text            =   "LTIME"
      Top             =   10080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Updstats 
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      DataField       =   "BD_LDATE"
      Height          =   285
      Index           =   0
      Left            =   4680
      TabIndex        =   52
      Text            =   "LDATE"
      Top             =   10080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.VScrollBar scrControl 
      Height          =   5295
      LargeChange     =   350
      Left            =   10920
      Max             =   100
      SmallChange     =   350
      TabIndex        =   51
      Top             =   2520
      Width           =   300
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8040
      Top             =   10080
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   50
      Top             =   10290
      Width           =   11415
      _Version        =   65536
      _ExtentX        =   20135
      _ExtentY        =   1164
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Font3D          =   1
      Alignment       =   1
      Begin VB.CommandButton cmdRemoveAllActual 
         Caption         =   "Remove All Actuals"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   32
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmdFlip 
         Caption         =   "Show Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   31
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmdUnfreeze 
         Caption         =   "Unfreeze Month"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   30
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmdFreeze 
         Caption         =   "Freeze Month"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   29
         Top             =   0
         Width           =   1575
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   9420
         Top             =   30
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   1
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin Threed.SSPanel panDetails 
      Height          =   7335
      Left            =   120
      TabIndex        =   41
      Top             =   2520
      Width           =   11175
      _Version        =   65536
      _ExtentX        =   19711
      _ExtentY        =   12938
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin VB.CheckBox chkUseFTE 
         Alignment       =   1  'Right Justify
         Caption         =   "Use FTE"
         DataField       =   "BD_FTE"
         Height          =   255
         Left            =   100
         TabIndex        =   7
         Top             =   1680
         Width           =   975
      End
      Begin VB.Frame frActEmployees 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   4800
         TabIndex        =   82
         Top             =   2030
         Visible         =   0   'False
         Width           =   4455
         Begin VB.TextBox txtAFTAA 
            DataField       =   "ACTUAL_FT_A"
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox txtATMPAA 
            DataField       =   "ACTUAL_TMP_A"
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Actual Full-Time Employees"
            Height          =   195
            Left            =   0
            TabIndex        =   84
            Top             =   135
            Width           =   1935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Actual Other Employees"
            Height          =   195
            Left            =   0
            TabIndex        =   83
            Top             =   495
            Width           =   1695
         End
      End
      Begin VB.Frame frTSTechFields 
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   120
         TabIndex        =   73
         Top             =   2880
         Width           =   9135
         Begin VB.TextBox txtAbsentFT 
            DataField       =   "ABSENT_HOURS_FT"
            Height          =   285
            Left            =   2280
            TabIndex        =   14
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtAbsentTMP 
            DataField       =   "ABSENT_HOURS_TMP"
            Height          =   285
            Left            =   2280
            TabIndex        =   15
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtSales 
            DataField       =   "TOTAL_SALES"
            Height          =   285
            Left            =   7200
            TabIndex        =   16
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox txtMaterial 
            DataField       =   "TOTAL_MATERIAL_COST"
            Height          =   285
            Left            =   7200
            TabIndex        =   17
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtValue 
            DataField       =   "TOTAL_VALUE_ADDED"
            Height          =   285
            Left            =   7200
            TabIndex        =   18
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtValAssociate 
            DataField       =   "VALUE_ADDED_ASSOC"
            Height          =   285
            Left            =   7200
            TabIndex        =   19
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtSchedFT 
            DataField       =   "SCHED_HOURS_FT"
            Height          =   285
            Left            =   2280
            TabIndex        =   12
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox txtShedTMP 
            DataField       =   "SCHED_HOURS_TMP"
            Height          =   285
            Left            =   2280
            TabIndex        =   13
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblAbsentFt 
            Caption         =   "Absenteeism Full Time"
            Height          =   255
            Left            =   0
            TabIndex        =   81
            Top             =   855
            Width           =   2295
         End
         Begin VB.Label lblAbsentTMP 
            Caption         =   "Absenteeism Temporary"
            Height          =   255
            Left            =   0
            TabIndex        =   80
            Top             =   1215
            Width           =   2175
         End
         Begin VB.Label lblSales 
            Caption         =   "Total Sales ($000)"
            Height          =   255
            Left            =   4680
            TabIndex        =   79
            Top             =   135
            Width           =   1935
         End
         Begin VB.Label lblMaterial 
            Caption         =   "Total Material Cost ($000)"
            Height          =   255
            Left            =   4680
            TabIndex        =   78
            Top             =   495
            Width           =   2055
         End
         Begin VB.Label lblValAdded 
            Caption         =   "Total Value Added ($000)"
            Height          =   255
            Left            =   4680
            TabIndex        =   77
            Top             =   855
            Width           =   1935
         End
         Begin VB.Label lblValAssoc 
            Caption         =   "Value Added per Assoc ($)"
            Height          =   255
            Left            =   4680
            TabIndex        =   76
            Top             =   1215
            Width           =   2055
         End
         Begin VB.Label lblSchedFT 
            Caption         =   "Scheduled Hours Full Time"
            Height          =   255
            Left            =   0
            TabIndex        =   75
            Top             =   135
            Width           =   2175
         End
         Begin VB.Label lblShedTMP 
            Caption         =   "Scheduled Hours Temporary"
            Height          =   255
            Left            =   0
            TabIndex        =   74
            Top             =   495
            Width           =   2175
         End
      End
      Begin VB.Frame frmBudgetYearDelete 
         Height          =   680
         Left            =   5040
         TabIndex        =   71
         Top             =   4440
         Width           =   3615
         Begin VB.TextBox txtBudYearDel 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2640
            MaxLength       =   4
            TabIndex        =   24
            Tag             =   "30-seniority Flag - Y or N"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmdBudYearDel 
            Caption         =   "Budget Year Delete"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   200
            Width           =   1815
         End
         Begin VB.Label lblBudYDel 
            Caption         =   "Year"
            Height          =   255
            Left            =   2160
            TabIndex        =   72
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame frmBudgetYearCopy 
         Height          =   680
         Left            =   120
         TabIndex        =   68
         Top             =   4440
         Width           =   4815
         Begin VB.CommandButton cmdBudYearCopy 
            Caption         =   "Budget Year Copy"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   200
            Width           =   1815
         End
         Begin VB.TextBox txtBudYearFrom 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2640
            MaxLength       =   4
            TabIndex        =   21
            Tag             =   "30-seniority Flag - Y or N"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtBudYearTo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3840
            MaxLength       =   4
            TabIndex        =   22
            Tag             =   "30-seniority Flag - Y or N"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblBudYF 
            Caption         =   "From"
            Height          =   255
            Left            =   2160
            TabIndex        =   70
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblBudYT 
            Caption         =   "To"
            Height          =   255
            Left            =   3480
            TabIndex        =   69
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.TextBox txtAOTS 
         DataField       =   "ACTUAL_OTHER_S"
         Height          =   285
         Left            =   7080
         TabIndex        =   26
         Top             =   5160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtATMPS 
         DataField       =   "ACTUAL_TMP_S"
         Height          =   285
         Left            =   7320
         TabIndex        =   67
         Top             =   2505
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtAFTS 
         DataField       =   "ACTUAL_FT_S"
         Height          =   285
         Left            =   7320
         TabIndex        =   66
         Top             =   2145
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtAOTA 
         DataField       =   "ACTUAL_OTHER_A"
         Height          =   285
         Left            =   2520
         TabIndex        =   25
         Top             =   5160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtATMPA 
         DataField       =   "ACTUAL_TMP_A"
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   2505
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtAFTA 
         DataField       =   "ACTUAL_FT_A"
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   2145
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkFreeze 
         Alignment       =   1  'Right Justify
         Caption         =   "Frozen"
         DataField       =   "BD_FREEZE"
         Enabled         =   0   'False
         Height          =   255
         Left            =   9120
         TabIndex        =   65
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtBDFTS 
         DataField       =   "BUDGET_FT_S"
         Height          =   285
         Left            =   7320
         TabIndex        =   36
         Top             =   2160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtBTMPS 
         DataField       =   "BUDGET_TMP_S"
         Height          =   285
         Left            =   7320
         TabIndex        =   37
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtBOTS 
         DataField       =   "BUDGET_OTHER_S"
         Height          =   285
         Left            =   7080
         TabIndex        =   38
         Top             =   5160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtAbsentOther 
         DataField       =   "ABSENT_HOURS_OT"
         Height          =   285
         Left            =   7080
         TabIndex        =   27
         Top             =   5520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtSchedOT 
         DataField       =   "SCHED_HOURS_OT"
         Height          =   285
         Left            =   7080
         TabIndex        =   28
         Top             =   5880
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         ItemData        =   "frmBudget.frx":E044
         Left            =   4440
         List            =   "frmBudget.frx":E046
         TabIndex        =   1
         Top             =   450
         Width           =   1575
      End
      Begin VB.TextBox txtBMonth 
         BackColor       =   &H80000000&
         DataField       =   "BUDGET_MONTH"
         Height          =   285
         Left            =   4680
         TabIndex        =   40
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtBYear 
         DataField       =   "BUDGET_YEAR"
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   465
         Width           =   1575
      End
      Begin VB.TextBox txtMonthSeq 
         DataField       =   "MONTH_SEQ"
         Height          =   285
         Left            =   7920
         TabIndex        =   2
         Top             =   465
         Width           =   975
      End
      Begin VB.TextBox txtBOA 
         DataField       =   "BUDGET_OTHER_A"
         Height          =   285
         Left            =   2520
         TabIndex        =   35
         Top             =   5160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtBTMPA 
         DataField       =   "BUDGET_TMP_A"
         Height          =   285
         Left            =   2400
         TabIndex        =   34
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox txtBFTA 
         DataField       =   "BUDGET_FT_A"
         Height          =   285
         Left            =   2400
         TabIndex        =   33
         Top             =   2160
         Width           =   1575
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "BD_DEPT"
         Height          =   285
         Index           =   0
         Left            =   7680
         TabIndex        =   5
         Top             =   945
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   7
         LookupType      =   2
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "GL_NUMBER"
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   3
         Tag             =   "00-General Ledger - Code"
         Top             =   945
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   3
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         DataField       =   "BD_DIV"
         Height          =   285
         Left            =   7680
         TabIndex        =   6
         Top             =   1305
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "BD_ADMINBY"
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   4
         Tag             =   "00-Administered By"
         Top             =   1305
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "BD_LOCATION"
         Height          =   285
         Index           =   3
         Left            =   1680
         TabIndex        =   63
         Tag             =   "00-Location - Code"
         Top             =   5520
         Visible         =   0   'False
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin VB.Label lblLegend 
         Caption         =   "NOTE: Rows highlighted in red are Frozen and will not recalculate on the manpower report"
         Height          =   255
         Left            =   0
         TabIndex        =   64
         Top             =   0
         Width           =   9015
      End
      Begin VB.Label lblDiv 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5640
         TabIndex        =   62
         Top             =   1350
         Width           =   555
      End
      Begin VB.Label lblLocation 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   600
         TabIndex        =   61
         Top             =   5520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblAdminBy 
         Caption         =   "Administered By"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblBFTS 
         Caption         =   "Budget Full Time Supervisors"
         Height          =   255
         Left            =   4800
         TabIndex        =   59
         Top             =   2160
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblBDTMS 
         Caption         =   "Budget Temporary Supervisors"
         Height          =   255
         Left            =   4800
         TabIndex        =   58
         Top             =   2520
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblBOTS 
         Caption         =   "Budget Other Supervisors"
         Height          =   255
         Left            =   4560
         TabIndex        =   57
         Top             =   5160
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblAbsentOT 
         Caption         =   "Absenteeism Other"
         Height          =   255
         Left            =   4560
         TabIndex        =   56
         Top             =   5520
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblSchedTmp 
         Caption         =   "Scheduled Hours Other"
         Height          =   255
         Left            =   4560
         TabIndex        =   55
         Top             =   6000
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblGL 
         Caption         =   "G/L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblMonthSeq 
         Caption         =   "Month Sequence"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   48
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblDept 
         Caption         =   "Department"
         Height          =   255
         Left            =   5640
         TabIndex        =   47
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblBOtherA 
         Caption         =   "Budget Other Associates"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   5160
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblBDTMPA 
         Caption         =   "Budget Temporary Associates"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label lblBDFTA 
         Caption         =   "Budget Full Time Associates"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label lblBMonth 
         Caption         =   "Budget Month"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   43
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblBYear 
         Caption         =   "Budget Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   42
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmManpower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************
'*                                                     *
'*      Form: frmManpower                              *
'*                                                     *
'*             Created: 12/Jul/05    By: Bryan         *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: To enter budgeted Manpower data*
'*                                                     *
'*******************************************************

Option Explicit
Dim fglbNew
Dim rsDATA As New ADODB.Recordset
Dim FRS As ADODB.Recordset
Dim fglbActual As Boolean 'True show actual/false show budget

Private Sub cboMonth_Change()
    If cboMonth.ListIndex > -1 Then
        txtBMonth = cboMonth.ItemData(cboMonth.ListIndex)
    End If
End Sub

Private Sub cboMonth_Click()
    If cboMonth.ListIndex > -1 Then
        txtBMonth = cboMonth.ItemData(cboMonth.ListIndex)
    End If
End Sub

Private Sub cmdBudYearCopy_Click() 'Ticket #13005
Dim rsBudYearFrom As New ADODB.Recordset
Dim rsBudYearTo As New ADODB.Recordset
Dim SQLQ As String
Dim a As Integer, Msg As String, x%

    If Len(txtBudYearFrom.Text) = 0 Then
        MsgBox "From Year is blank."
        txtBudYearFrom.SetFocus
        Exit Sub
    End If
    If Len(txtBudYearTo.Text) = 0 Then
        MsgBox "To Year is blank."
        txtBudYearTo.SetFocus
        Exit Sub
    End If
    If Not Len(txtBudYearFrom.Text) = 4 Then
        MsgBox "Invalid From Year."
        txtBudYearFrom.SetFocus
        Exit Sub
    End If
    If Not Len(txtBudYearTo.Text) = 4 Then
        MsgBox "Invalid To Year."
        txtBudYearTo.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtBudYearFrom.Text) Then
        MsgBox "Invalid From Year."
        txtBudYearFrom.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtBudYearTo.Text) Then
        MsgBox "Invalid To Year."
        txtBudYearTo.SetFocus
        Exit Sub
    End If
    
    'Check if the From Year data exists
    SQLQ = "SELECT * FROM HRBUDGET WHERE BUDGET_YEAR = " & txtBudYearFrom.Text
    rsBudYearFrom.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsBudYearFrom.EOF Then
        MsgBox "No " & txtBudYearFrom.Text & " Budgeted Manpower Setup. "
        Exit Sub
    End If
    
    
    'Check if the To Year data exists
    SQLQ = "SELECT * FROM HRBUDGET WHERE BUDGET_YEAR = " & txtBudYearTo.Text
    rsBudYearTo.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsBudYearTo.EOF Then
        MsgBox txtBudYearTo.Text & " Budgeted Manpower exists. "
        Exit Sub
    End If
    
    Msg = "Are You Sure You Want To Copy Budgeted Manpower From Year " & txtBudYearFrom.Text & " To Year " & txtBudYearTo.Text & "? "
    a% = MsgBox(Msg, 36, "Confirm Copy")
    If a% <> 6 Then Exit Sub

    Screen.MousePointer = HOURGLASS
    Do While Not rsBudYearFrom.EOF
        rsBudYearTo.AddNew
        rsBudYearTo("BUDGET_YEAR") = txtBudYearTo.Text
        rsBudYearTo("BUDGET_MONTH") = rsBudYearFrom("BUDGET_MONTH")
        rsBudYearTo("BUDGET_FT_A") = rsBudYearFrom("BUDGET_FT_A")
        rsBudYearTo("BUDGET_TMP_A") = rsBudYearFrom("BUDGET_TMP_A")
        rsBudYearTo("BUDGET_OTHER_A") = rsBudYearFrom("BUDGET_OTHER_A")
        rsBudYearTo("ACTUAL_FT_A") = 0
        rsBudYearTo("ACTUAL_TMP_A") = 0
        rsBudYearTo("ACTUAL_OTHER_A") = 0
        rsBudYearTo("BUDGET_FT_S") = rsBudYearFrom("BUDGET_FT_S")
        rsBudYearTo("BUDGET_TMP_S") = rsBudYearFrom("BUDGET_TMP_S")
        rsBudYearTo("BUDGET_OTHER_S") = rsBudYearFrom("BUDGET_OTHER_S")
        rsBudYearTo("ACTUAL_FT_S") = 0
        rsBudYearTo("ACTUAL_TMP_S") = 0
        rsBudYearTo("ACTUAL_OTHER_S") = 0
        rsBudYearTo("MONTH_SEQ") = rsBudYearFrom("MONTH_SEQ")
        rsBudYearTo("BD_DEPT") = rsBudYearFrom("BD_DEPT")
        rsBudYearTo("GL_NUMBER") = rsBudYearFrom("GL_NUMBER")
        rsBudYearTo("BD_ADMINBY_TABL") = rsBudYearFrom("BD_ADMINBY_TABL")
        rsBudYearTo("BD_ADMINBY") = rsBudYearFrom("BD_ADMINBY")
        rsBudYearTo("BD_Location") = rsBudYearFrom("BD_Location")
        rsBudYearTo("BD_Div") = rsBudYearFrom("BD_Div")
        rsBudYearTo("ABSENT_HOURS_FT") = rsBudYearFrom("ABSENT_HOURS_FT")
        rsBudYearTo("ABSENT_HOURS_TMP") = rsBudYearFrom("ABSENT_HOURS_TMP")
        rsBudYearTo("ABSENT_HOURS_OT") = rsBudYearFrom("ABSENT_HOURS_OT")
        rsBudYearTo("SCHED_HOURS_FT") = rsBudYearFrom("SCHED_HOURS_FT")
        rsBudYearTo("SCHED_HOURS_TMP") = rsBudYearFrom("SCHED_HOURS_TMP")
        rsBudYearTo("SCHED_HOURS_OT") = rsBudYearFrom("SCHED_HOURS_OT")
        rsBudYearTo("TOTAL_SALES") = rsBudYearFrom("TOTAL_SALES")
        rsBudYearTo("TOTAL_MATERIAL_COST") = rsBudYearFrom("TOTAL_MATERIAL_COST")
        rsBudYearTo("TOTAL_VALUE_ADDED") = rsBudYearFrom("TOTAL_VALUE_ADDED")
        rsBudYearTo("VALUE_ADDED_ASSOC") = rsBudYearFrom("VALUE_ADDED_ASSOC")
        rsBudYearTo("BD_FREEZE") = 0
        rsBudYearTo("BD_LUSER") = glbUserID
        rsBudYearTo("BD_LTIME") = Time$
        rsBudYearTo("BD_LDATE") = Date
        rsBudYearTo("BD_MGR") = rsBudYearFrom("BD_MGR")
        rsBudYearTo("BD_COOR") = rsBudYearFrom("BD_COOR")
        rsBudYearTo.Update
        rsBudYearFrom.MoveNext
    Loop
    rsBudYearFrom.Close
    rsBudYearTo.Close
    Screen.MousePointer = DEFAULT
    
    MsgBox "    Update Completed!    "
    txtBudYearFrom.Text = ""
    txtBudYearTo.Text = ""
    Data1.Refresh
End Sub

Private Sub cmdBudYearDel_Click() 'Ticket #13005
Dim rsBudYearDel As New ADODB.Recordset
Dim SQLQ As String
Dim a As Integer, Msg As String, x%

    If Len(txtBudYearDel.Text) = 0 Then
        MsgBox "Year is blank."
        txtBudYearDel.SetFocus
        Exit Sub
    End If

    If Not Len(txtBudYearDel.Text) = 4 Then
        MsgBox "Invalid Year."
        txtBudYearDel.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtBudYearDel.Text) Then
        MsgBox "Invalid Year."
        txtBudYearDel.SetFocus
        Exit Sub
    End If
    
    'Check if the From Year data exists
    SQLQ = "SELECT * FROM HRBUDGET WHERE BUDGET_YEAR = " & txtBudYearDel.Text
    rsBudYearDel.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsBudYearDel.EOF Then
        MsgBox "No " & txtBudYearDel.Text & " Budgeted Manpower. "
        Exit Sub
    End If
    rsBudYearDel.Close
    
    Msg = "Are You Sure You Want To Delete Budgeted Manpower For Year " & txtBudYearDel.Text & "? "
    a% = MsgBox(Msg, 36, "Confirm Delete")
    If a% <> 6 Then Exit Sub

    Screen.MousePointer = HOURGLASS
    SQLQ = "DELETE FROM HRBUDGET WHERE BUDGET_YEAR = " & txtBudYearDel.Text
    gdbAdoIhr001.Execute SQLQ
    Screen.MousePointer = DEFAULT
    
    MsgBox "   Delete Completed!    "
    txtBudYearDel.Text = ""
    Data1.Refresh
End Sub

Private Sub cmdFlip_Click()

    fglbActual = Not fglbActual
    
    Call setGridCol
    
    txtBFTA.Visible = Not fglbActual
    txtBTMPA.Visible = Not fglbActual
    txtAFTA.Visible = fglbActual
    txtATMPA.Visible = fglbActual

    If glbCompSerial = "S/N - 2369W" Then
        txtBDFTS.Visible = Not fglbActual
        txtBTMPS.Visible = Not fglbActual
        txtAFTS.Visible = fglbActual
        txtATMPS.Visible = fglbActual
    End If
    
    If fglbActual = True Then
        cmdFlip.Caption = "Show Budgeted"
        If glbCompSerial = "S/N - 2369W" Then
            Me.lblBDFTA.Caption = "Actual Full-Time Associates"
            Me.lblBDTMPA.Caption = "Actual Temporary Associates"
            Me.lblBDTMS.Caption = "Actual Temporary Supervisors"
            Me.lblBFTS.Caption = "Actual Full-Time Supervisors"
        Else
            Me.lblBDFTA.Caption = "Actual Full-Time Employees"
            Me.lblBDTMPA.Caption = "Actual Other Employees"
        End If
        
    Else
        cmdFlip.Caption = "Show Actual"
        If glbCompSerial = "S/N- 2369W" Then
            Me.lblBDFTA.Caption = "Budget Full-Time Associates"
            Me.lblBDTMPA.Caption = "Budget Temporary Associates"
            Me.lblBDTMS.Caption = "Budget Temporary Supervisors"
            Me.lblBFTS.Caption = "Budget Full-Time Supervisors"
        Else
            Me.lblBDFTA.Caption = "Budget Full-Time Employees"
            Me.lblBDTMPA.Caption = "Budget Other Employees"

        End If
    End If
    
End Sub

Private Sub cmdFreeze_Click()
    On Error GoTo EH
    Dim Response As Integer
    Dim strSQL As String
    Dim xID As Long
    
    'Ticket #21100
    If glbCompSerial = "S/N - 2369W" Then
        Response = MsgBox("This will freeze " & cboMonth.Text & ", " & txtBYear.Text & ", Do you wish to continue?", vbQuestion + vbYesNo, "Confirm Freeze")
    Else
        Response = MsgBox("This function will freeze and recalculate the Actual for " & txtBYear.Text & ". Do you wish to continue?", vbQuestion + vbYesNo, "Confirm Freeze and Recalculate Actual")
    End If
    If Response = vbYes Then
        'Ticket #21100
        If glbCompSerial <> "S/N - 2369W" Then
            'Calculate the Actual values by reading the current demographics settings
            'Ticket #21876 - Recalculate based on the 'Use FTE' setting
            'Call ReCalculate_Manpower_Actual
            Call ReCalculate_Manpower_Actual_FTE
            
            'Freeze entire year
            strSQL = "UPDATE HRBUDGET SET BD_FREEZE = 1 WHERE BUDGET_YEAR=" & txtBYear.Text
        Else
            'Freeze the selected Month
            strSQL = "UPDATE HRBUDGET SET BD_FREEZE = 1 WHERE BUDGET_MONTH=" & txtBMonth.Text & " AND BUDGET_YEAR=" & txtBYear.Text
        End If
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute strSQL
        gdbAdoIhr001.CommitTrans
        rsDATA.Resync
        xID = rsDATA!BUDGET_ID
        Data1.Refresh
        Data1.Recordset.Find "BUDGET_ID=" & xID
        FRS.Requery
        
    End If

    Exit Sub

EH:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdFreeze", "HRBUDGET", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
        
End Sub

Private Sub cmdRemoveAllActual_Click()
    On Error GoTo Err_cmdRemoveAllActual
    Dim Response As Integer
    Dim strSQL As String
    Dim xID As Long
    
    
    Response = MsgBox("This function will unfreeze and remove the actual counts for the entire year '" & txtBYear.Text & "'. Do you wish to continue?", vbQuestion + vbYesNo, "Confirm Remove All Actuals")
    If Response = vbYes Then
        'Reset the Actuals to Null first
        'For all the records in the Budget Year selected
        strSQL = "UPDATE HRBUDGET SET ACTUAL_FT_A = NULL, ACTUAL_TMP_A=NULL, ACTUAL_OTHER_A=NULL, ACTUAL_FT_S = NULL, ACTUAL_TMP_S=NULL, ACTUAL_OTHER_S=NULL "
        strSQL = strSQL & " WHERE BD_FREEZE=1 and BUDGET_YEAR=" & txtBYear.Text
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute strSQL
        gdbAdoIhr001.CommitTrans
        
        'Unfreeze
        strSQL = "UPDATE HRBUDGET SET BD_FREEZE = 0 WHERE BUDGET_YEAR=" & txtBYear.Text
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute strSQL
        gdbAdoIhr001.CommitTrans
        rsDATA.Resync
        xID = rsDATA!BUDGET_ID
        Data1.Refresh
        Data1.Recordset.Find "BUDGET_ID=" & xID
        FRS.Requery
    End If

    Exit Sub

Err_cmdRemoveAllActual:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdRemoveAllActual", "HRBUDGET", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub cmdUnfreeze_Click()
    On Error GoTo EH
    Dim Response As Integer
    Dim strSQL As String
    Dim xID As Long
    
    If glbCompSerial <> "S/N - 2369W" Then
        Response = MsgBox("This function will unfreeze and remove the actual counts for " & cboMonth.Text & ", " & txtBYear.Text & ", Do you wish to continue?", vbQuestion + vbYesNo, "Confirm Remove Actual")
    Else
        Response = MsgBox("This will unfreeze " & cboMonth.Text & ", " & txtBYear.Text & ", Do you wish to continue?", vbQuestion + vbYesNo, "Confirm Unfreeze")
    End If
    If Response = vbYes Then
        'Ticket #21876
        If glbCompSerial <> "S/N - 2369W" Then
            'Reset the Actuals to Null first
            'For the selected Budget Year & Month
            strSQL = "UPDATE HRBUDGET SET ACTUAL_FT_A = NULL, ACTUAL_TMP_A=NULL, ACTUAL_OTHER_A=NULL, ACTUAL_FT_S = NULL, ACTUAL_TMP_S=NULL, ACTUAL_OTHER_S=NULL "
            strSQL = strSQL & " WHERE BD_FREEZE=1 and BUDGET_YEAR=" & txtBYear.Text & " AND BUDGET_MONTH=" & txtBMonth.Text
            gdbAdoIhr001.BeginTrans
            gdbAdoIhr001.Execute strSQL
            gdbAdoIhr001.CommitTrans
        End If
    
        strSQL = "UPDATE HRBUDGET SET BD_FREEZE = 0 WHERE BUDGET_MONTH=" & txtBMonth.Text & " AND BUDGET_YEAR=" & txtBYear.Text
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute strSQL
        gdbAdoIhr001.CommitTrans
        rsDATA.Resync
        xID = rsDATA!BUDGET_ID
        Data1.Refresh
        Data1.Recordset.Find "BUDGET_ID=" & xID
        FRS.Requery
    End If

    Exit Sub

EH:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdunFreeze", "HRBUDGET", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub Form_Activate()
    Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Dim SQLQ As String
Dim c%

glbOnTop = "FRMMANPOWER"
    
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Call setCaption(lblDiv)
Call setCaption(lblDept)
Call setCaption(lblGL)
Call setCaption(lblAdminBy)
Call setCaption(lblLocation)

Screen.MousePointer = HOURGLASS

fglbNew = False
fglbActual = False

Data1.ConnectionString = glbAdoIHRDB

SQLQ = "SELECT * FROM HRBUDGET "
SQLQ = SQLQ & "ORDER BY BUDGET_YEAR DESC, MONTH_SEQ DESC, GL_NUMBER"

Data1.RecordSource = SQLQ
Data1.Refresh
Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True
vbxTrueGrid.MarqueeStyle = 3

For c% = 1 To 12
    cboMonth.AddItem MonthName(c%), c% - 1
    cboMonth.ItemData(c% - 1) = c%
Next c%

Call setCaption(Me.vbxTrueGrid.Columns(3)) 'Department
Call setCaption(Me.vbxTrueGrid.Columns(4)) 'G/L
Call setCaption(Me.vbxTrueGrid.Columns(5))  'Admin By
Call setCaption(Me.vbxTrueGrid.Columns(6))  'Location

If glbCompSerial = "S/N - 2369W" Then 'TS Tech
    lblAdminBy.FontBold = True
    lblDept.FontBold = True
    lblDiv.FontBold = True
    Me.lblBDFTA.Caption = "Budget Full-Time Associates"
    Me.lblBDTMPA.Caption = "Budget Temporary Associates"
    Me.lblValAssoc.Caption = "Value Added per Assoc ($)"
    Me.lblBDTMS.Visible = True
    Me.lblBFTS.Visible = True
    Me.txtAFTS.Visible = False
    Me.txtATMPS.Visible = False
    Me.txtBDFTS.Visible = True
    Me.txtBTMPS.Visible = True
    
    'Ticket #21876
    Me.cmdRemoveAllActual.Visible = False
    Me.chkUseFTE.Visible = False
Else
    lblAdminBy.FontBold = False
    lblDept.FontBold = False
    lblDiv.FontBold = False
    Me.lblBDFTA.Caption = "Budget Full-Time Employees"
    Me.lblBDTMPA.Caption = "Budget Other Employees"
    Me.lblValAssoc.Caption = "Value Added per Emp ($)"
    Me.lblBDTMS.Visible = False
    Me.lblBFTS.Visible = False
    Me.txtAFTS.Visible = False
    Me.txtATMPS.Visible = False
    Me.txtBDFTS.Visible = False
    Me.txtBTMPS.Visible = False
    
    'Ticket #21100
    frTSTechFields.Visible = False
    frActEmployees.Visible = True
    cmdFlip.Visible = False     'Show Actual/Show Budgeted
    cmdFreeze.Caption = "Update Actual"
    cmdUnfreeze.Caption = "Remove Actual"
           
End If

Call setGridCol

Call INI_Controls(Me)


fglbNew = False

Screen.MousePointer = DEFAULT

End Sub

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum

If fglbNew Then
    UpdateState = NewRecord
    TF = True
ElseIf Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
panDetails.Enabled = TF

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Resize()
'Commented by Bryan
If Me.WindowState <> vbMinimized Then
    If Me.Height >= vbxTrueGrid.Height + panDetails.Height + panControls.Height + 230 Then
        scrControl.Value = 0
        panDetails.Top = vbxTrueGrid.Height + 500 '240
        scrControl.Visible = False
        Exit Sub
    End If
    If Me.Height < vbxTrueGrid.Height + scrControl.Top + panControls.Height Then Exit Sub
    scrControl.Visible = True
    scrControl.Max = vbxTrueGrid.Height + panDetails.Height + panControls.Height - Me.Height + 250
    scrControl.Left = Me.Width - scrControl.Width - 120
    scrControl.Height = Me.Height - scrControl.Top - panControls.Height - 400
'    panDetails.Width = Me.Width
'    vbxTrueGrid.Width = Me.Width
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
Set frmManpower = Nothing
End Sub

Private Sub scrControl_Change()
    panDetails.Top = 500 + vbxTrueGrid.Height - scrControl.Value * 2.5
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Basic
End Property

Public Property Get Addable() As Boolean
Addable = True
End Property

Public Property Get Updateble() As Boolean
Updateble = True
End Property

Public Property Get Deleteble() As Boolean
Deleteble = True
End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Sub cmdCancel_Click()
Dim x, bk
On Error GoTo Can_Err

Call Display_Value

fglbNew = False

Call SET_UP_MODE

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_BENFTS_GROUP", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, x%

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
  MsgBox "Nothing to Delete"
  Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh
Call Display_Value

fglbNew = False

Exit Sub
Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRBENFTGROUP", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Sub cmdView_Click()

Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = "Budgeted Manpower "
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Destination = 0

Me.vbxCrystal.Action = 1

End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = "Budgeted Manpower "
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Sub cmdNew_Click()
Dim SQLQ As String, Msg$, x%
Dim Title$, DgDef As Variant, Response%

On Error GoTo AddN_Err

fglbNew = True

Call SET_UP_MODE
'If Not gSec_Upd_Benefits Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If

Call Set_Control("B", Me)

Screen.MousePointer = DEFAULT

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "ATTEND", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Sub cmdOK_Click()
Dim xID As Long

On Error GoTo Add_Err
    
If chk_Manpower = False Then Exit Sub

rsDATA.Requery

If fglbNew Then rsDATA.AddNew

Call UpdUStats(Me)
        
gdbAdoIhr001.BeginTrans

Call Set_Control("U", Me, rsDATA)

rsDATA.Update
gdbAdoIhr001.CommitTrans
rsDATA.Resync

xID = rsDATA!BUDGET_ID

Data1.Refresh

Data1.Recordset.Find "BUDGET_ID=" & xID

fglbNew = False

Call SET_UP_MODE

Exit Sub

Add_Err:
If Err = 3022 Then
    Data1.Recordset.CancelUpdate    ' no dups
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRBENFT", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Sub Display_Value()
Dim SQLQ
    
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If glbCompSerial <> "S/N - 2369W" Then 'TS Tech
        cmdFreeze.Enabled = False
        cmdUnfreeze.Enabled = False
        cmdRemoveAllActual.Enabled = False
    End If
Else
    If glbCompSerial <> "S/N - 2369W" Then 'TS Tech
        cmdFreeze.Enabled = True
        cmdUnfreeze.Enabled = True
        cmdRemoveAllActual.Enabled = True
    End If
    
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    SQLQ = "SELECT * FROM HRBUDGET "
    SQLQ = SQLQ & " WHERE BUDGET_ID = " & Data1.Recordset!BUDGET_ID
    SQLQ = SQLQ & " ORDER BY BUDGET_YEAR DESC, MONTH_SEQ ASC"
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)

End If

Call SET_UP_MODE

End Sub

Private Sub txtBMonth_Change()
    If Len(txtBMonth.Text) > 0 Then
        cboMonth.Text = MonthName(txtBMonth)
    End If
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo EH
    'added by Bryan 18/Jan/06 Ticket#10222
    FRS.Requery
    FRS.Bookmark = Bookmark
    If FRS("BD_FREEZE") = True Then
        RowStyle.ForeColor = vbRed
    End If
    
EH:
    Exit Sub
End Sub

Private Sub vbxTrueGrid_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
    Dim SQLQ As String
    
    If vbxTrueGrid.Columns(ColIndex).DataField <> "BUDGET_YEAR" Then
        
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT * FROM HRBUDGET "
        SQLQ = SQLQ & "ORDER BY BUDGET_YEAR DESC, " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
        Set FRS = Data1.Recordset.Clone
        vbxTrueGrid.FetchRowStyle = True
    End If
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
End Sub

Private Function chk_Manpower() As Boolean
Dim retVal As Boolean
Dim rs As New ADODB.Recordset
Dim strSQL As String
chk_Manpower = False

'if the month combo box hasn't been clicked we need to know the month anyway
    If txtBMonth = "" Then
        txtBMonth = month(cboMonth.Text & " 1, 2005")
    End If

If Len(txtBMonth.Text) = 0 Then
    MsgBox "Budget Month is required."
    txtBMonth.SetFocus
    Exit Function
End If

If Len(txtBYear.Text) = 0 Then
    MsgBox "Budget Year is required."
    txtBYear.SetFocus
    Exit Function
End If

If Len(txtMonthSeq.Text) = 0 Then
    MsgBox "Month Sequence is required."
    txtMonthSeq.SetFocus
    Exit Function
End If

If glbCompSerial = "S/N - 2369W" Then
    If Len(clpCode(0).Text) = 0 Then
        MsgBox lStr("Department") & " is required."
        clpCode(0).SetFocus
        Exit Function
    End If
    If Len(clpDiv.Text) = 0 Then
        MsgBox lStr("Division") & " is required."
        clpDiv.SetFocus
        Exit Function
    End If
    
    If Len(clpCode(2).Text) = 0 Then
        MsgBox lStr("Administered By") & " is required."
        clpCode(2).SetFocus
        Exit Function
    End If
End If

If Len(clpCode(1).Text) = 0 Then
    MsgBox lStr("G/L") & " is required."
    clpCode(1).SetFocus
    Exit Function
End If

'Ticket #21100 - Modified the SELECT statement to look for duplicate record
strSQL = "SELECT BUDGET_ID FROM HRBUDGET WHERE "
strSQL = strSQL & "BUDGET_YEAR=" & txtBYear.Text & " AND MONTH_SEQ=" & txtMonthSeq.Text
If Len(clpCode(0).Text) > 0 Then
    strSQL = strSQL & " AND BD_DEPT='" & clpCode(0).Text & "'"
End If
If Len(clpCode(1).Text) > 0 Then
    strSQL = strSQL & " AND GL_NUMBER='" & clpCode(1).Text & "' "
End If
If Len(clpDiv.Text) > 0 Then
    strSQL = strSQL & " AND BD_DIV='" & clpDiv.Text & "'"
End If
If Len(clpCode(3).Text) > 0 Then
    strSQL = strSQL & " AND BD_Location='" & clpCode(3).Text & "' "
End If
If Len(clpCode(2).Text) > 0 Then
    strSQL = strSQL & " AND BD_ADMINBY = '" & clpCode(2).Text & "'"
End If
strSQL = strSQL & " AND BUDGET_MONTH = " & txtBMonth.Text
If Not fglbNew Then
    strSQL = strSQL & " AND BUDGET_ID <> " & Data1.Recordset!BUDGET_ID
End If
rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
If rs.EOF = False And rs.BOF = False Then
    MsgBox "Duplicate record found in database"
    txtBYear.SetFocus
    rs.Close
    Set rs = Nothing
    Exit Function
End If
rs.Close
Set rs = Nothing

'Ticket #21100
'strSQL = "SELECT BUDGET_ID FROM HRBUDGET WHERE "
'strSQL = strSQL & "BUDGET_YEAR=" & txtBYear.Text & " AND MONTH_SEQ=" & txtMonthSeq.Text
'strSQL = strSQL & " AND BUDGET_MONTH = " & txtBMonth.Text
'If Not fglbNew Then
'    strSQL = strSQL & " AND BUDGET_ID <> " & Data1.Recordset!BUDGET_ID
'End If
'rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
'If rs.EOF = False And rs.BOF = False Then
'    MsgBox "Duplicate record found in database"
'    txtBYear.SetFocus
'    rs.Close
'    Set rs = Nothing
'    Exit Function
'End If
'rs.Close
'Set rs = Nothing

chk_Manpower = True

End Function

Private Sub setGridCol()

'Ticket #21100
If glbCompSerial = "S/N - 2369W" Then
    vbxTrueGrid.Columns(8).Visible = Not fglbActual
    vbxTrueGrid.Columns(9).Visible = Not fglbActual
    vbxTrueGrid.Columns(10).Visible = Not fglbActual
    vbxTrueGrid.Columns(11).Visible = Not fglbActual
    vbxTrueGrid.Columns(12).Visible = Not fglbActual
    vbxTrueGrid.Columns(13).Visible = Not fglbActual
    vbxTrueGrid.Columns(14).Visible = fglbActual
    vbxTrueGrid.Columns(15).Visible = fglbActual
    vbxTrueGrid.Columns(16).Visible = fglbActual
    vbxTrueGrid.Columns(17).Visible = fglbActual
    vbxTrueGrid.Columns(18).Visible = fglbActual
    vbxTrueGrid.Columns(19).Visible = fglbActual
Else
    vbxTrueGrid.Columns(10).Visible = False
    vbxTrueGrid.Columns(11).Visible = False
    vbxTrueGrid.Columns(12).Visible = False
    vbxTrueGrid.Columns(13).Visible = False
    vbxTrueGrid.Columns(16).Visible = False
    vbxTrueGrid.Columns(17).Visible = False
    vbxTrueGrid.Columns(18).Visible = False
    vbxTrueGrid.Columns(19).Visible = False
    vbxTrueGrid.Columns(20).Visible = False
    vbxTrueGrid.Columns(21).Visible = False
    vbxTrueGrid.Columns(22).Visible = False
    vbxTrueGrid.Columns(23).Visible = False
    vbxTrueGrid.Columns(24).Visible = False
    vbxTrueGrid.Columns(25).Visible = False
    vbxTrueGrid.Columns(26).Visible = False
    vbxTrueGrid.Columns(27).Visible = False
    vbxTrueGrid.Columns(28).Visible = False
    
    vbxTrueGrid.Columns(8).Caption = "Budgeted Full-Time Employees"
    vbxTrueGrid.Columns(9).Caption = "Budgeted Other Employees"
    vbxTrueGrid.Columns(14).Caption = "Actual Full-Time Employees"
    vbxTrueGrid.Columns(15).Caption = "Actual Other Employees"
End If

End Sub

Private Sub ReCalculate_Manpower_Actual()
    Dim strSQL As String
    Dim rsDATA As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim FTA, FTS, TMA, TMS
    Dim DTE As String
    Dim c% 'Month
    Dim x% 'Sequence

    'Reset the Actuals to Null first
    'For all the records in the Budget Year selected
    strSQL = "UPDATE HRBUDGET SET ACTUAL_FT_A = NULL, ACTUAL_TMP_A=NULL, ACTUAL_OTHER_A=NULL, ACTUAL_FT_S = NULL, ACTUAL_TMP_S=NULL, ACTUAL_OTHER_S=NULL "
    strSQL = strSQL & " WHERE BD_FREEZE=0 and BUDGET_YEAR=" & txtBYear.Text
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute strSQL
    gdbAdoIhr001.CommitTrans
    
    'For each Month of the Budget Year
    For c% = 1 To 12
        'Initialise
        x% = 0
        
        'Retrieve the Month Sequence
        strSQL = "SELECT MONTH_SEQ FROM HRBUDGET WHERE BUDGET_MONTH=" & c%
        strSQL = strSQL & " AND BUDGET_YEAR=" & txtBYear.Text
        rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
        
        If rsDATA.EOF = False And rsDATA.BOF = False Then
            x% = rsDATA("MONTH_SEQ")
        End If
        rsDATA.Close
            
        If x% > 0 Then
            'Compute end of the month date
            If c% < x% Then 'If month is less than the sequence then it must be next year
                DTE = getEOM(c%) & "/" & MonthName(c%, True) & "/" & (txtBYear.Text + 1)
            Else
                DTE = getEOM(c%) & "/" & MonthName(c%, True) & "/" & txtBYear.Text
            End If
            
            'Check if end of the month date date is less than today's date. Cannot calculate actual if the
            'budget end date is greater than today's date.
            If CDate(DTE) < Date Then
                'Find Actual Associates for this month(c)
                strSQL = "SELECT * FROM HRBUDGET WHERE BUDGET_YEAR=" & txtBYear.Text & " And BUDGET_MONTH = " & c%
                strSQL = strSQL & " AND BD_FREEZE=0"
                rsTemp.Open strSQL, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
                
                If rsTemp.EOF = False And rsTemp.EOF = False Then
                    Do
                        'Initialise
                        FTA = Null: FTS = Null: TMA = Null: TMS = Null
                        
                        'Calculate Actual for FT and other Categories
                        strSQL = "SELECT Count(HREMP.ED_EMPNBR) AS EMPCNT, HREMP.ED_PT, HRJOB.JB_STATUS "
                        If glbOracle Then
                            strSQL = strSQL & "FROM HREMP, HR_JOB_HISTORY, HRJOB "
                            strSQL = strSQL & "WHERE HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE AND (HREMP.ED_DOH <= " & Date_SQL(DTE) & ") AND (HR_JOB_HISTORY.JH_CURRENT<>0)"
                        Else
                            strSQL = strSQL & "FROM  (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
                            strSQL = strSQL & "WHERE (HREMP.ED_DOH <= " & Date_SQL(DTE) & ")  AND (HR_JOB_HISTORY.JH_CURRENT<>0) "
                        End If
                        If Len(rsTemp("BD_DEPT")) > 0 Then
                            strSQL = strSQL & "AND ED_DEPTNO='" & rsTemp("BD_DEPT") & "' "
                        End If
                        If Len(rsTemp("GL_NUMBER")) > 0 Then
                            strSQL = strSQL & "AND ED_GLNO='" & rsTemp("GL_NUMBER") & "' "
                        End If
                        If Len(rsTemp("BD_ADMINBY")) > 0 Then
                            strSQL = strSQL & "AND ED_ADMINBY='" & rsTemp("BD_ADMINBY") & "' "
                        End If
                        If Len(rsTemp("BD_DIV")) > 0 Then
                            strSQL = strSQL & "AND ED_DIV='" & rsTemp("BD_DIV") & "' "
                        End If
                        strSQL = strSQL & "GROUP BY  HREMP.ED_PT, HRJOB.JB_STATUS "
                        strSQL = strSQL & "HAVING HRJOB.JB_STATUS <> 'NA' "
                        
                        'TS Tech - FT and TMP only
                        If glbCompSerial = "S/N - 2369W" Then
                            strSQL = strSQL & " AND (HREMP.ED_PT='FT' Or HREMP.ED_PT='TMP') "
                        End If
                        
                        rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
                        If rsDATA.EOF = False And rsDATA.BOF = False Then
                            Do
                                'TS Tech
                                If glbCompSerial = "S/N - 2369W" And rsDATA("JB_STATUS") = "L" Then
                                    If rsDATA("ED_PT") = "FT" Then
                                        If IsNull(FTS) Then FTS = 0
                                        FTS = FTS + rsDATA("EMPCNT")
                                    Else
                                        If IsNull(TMS) Then TMS = 0
                                        TMS = TMS + rsDATA("EMPCNT")
                                    End If
                                Else
                                    'All other clients
                                    If rsDATA("ED_PT") = "FT" Then
                                        If IsNull(FTA) Then FTA = 0
                                        FTA = FTA + rsDATA("EMPCNT")
                                    Else
                                        If IsNull(TMA) Then TMA = 0
                                        TMA = TMA + rsDATA("EMPCNT")
                                    End If
                                End If
                                rsDATA.MoveNext
                            Loop Until rsDATA.EOF
                        End If
                        rsDATA.Close
                    
   
                        'Changed by Bryan 07/Mar/2006 Ticket#10493
                        'Will select terminated employees terminated on the last day of the month.
                        '**********************
                        'Compute end of the month date
                        If c% < x% Then
                            DTE = getEOM(c%) & "/" & MonthName(c%, True) & "/" & (txtBYear.Text + 1)
                        Else
                            DTE = getEOM(c%) & "/" & MonthName(c%, True) & "/" & txtBYear
                        End If
                        
                        'Check if end of the month date date is less than today's date. Cannot calculate actual if the
                        'budget end date is greater than today's date.
                        If CDate(DTE) < Date Then
                            'Clear the Temp table to load the fresh set of Term employees
                            strSQL = "DELETE FROM HRMANTERM WHERE WRKEMP='" & glbUserID & "'"
                            gdbAdoIhr001W.BeginTrans
                            gdbAdoIhr001W.Execute strSQL
                            gdbAdoIhr001W.CommitTrans
                            
                            'Find Actual from the terminated for this month(c) - between Hire Date and Term Date
                            strSQL = "SELECT TERM_HREMP.ED_EMPNBR, TERM_HREMP.ED_DEPTNO, TERM_HREMP.ED_GLNO, TERM_HREMP.ED_PT, TERM_HREMP.ED_ADMINBY, Term_JOB_HISTORY.JH_JOB, TERM_HREMP.ED_LOC, TERM_HREMP.ED_DIV  "
                            If glbOracle Then
                                strSQL = strSQL & "FROM TERM_HREMP, TERM_HRTRMEMP, Term_JOB_HISTORY "
                                strSQL = strSQL & "WHERE TERM_HREMP.ED_EMPNBR = TERM_HRTRMEMP.TERM_SEQ AND TERM_HREMP.TERM_SEQ = Term_JOB_HISTORY.JH_EMPNBR AND (TERM_HREMP.ED_DOH <= " & Date_SQL(DTE) & ") AND (TERM_HRTRMEMP.TERM_DOT = " & Date_SQL(DTE) & ") "
                            Else
                                strSQL = strSQL & "FROM ((TERM_HREMP INNER JOIN TERM_HRTRMEMP ON TERM_HREMP.TERM_SEQ = TERM_HRTRMEMP.TERM_SEQ) INNER JOIN Term_JOB_HISTORY ON TERM_HREMP.TERM_SEQ = Term_JOB_HISTORY.TERM_SEQ)"
                                strSQL = strSQL & "WHERE (TERM_HREMP.ED_DOH <= " & Date_SQL(DTE) & ") AND (TERM_HRTRMEMP.TERM_DOT = " & Date_SQL(DTE) & ") "
                            End If
                            strSQL = strSQL & "AND Term_JOB_HISTORY.JH_CURRENT<>0"
                            If Len(rsTemp("BD_DEPT")) > 0 Then
                                strSQL = strSQL & "AND ED_DEPTNO='" & rsTemp("BD_DEPT") & "' "
                            End If
                            If Len(rsTemp("GL_NUMBER")) > 0 Then
                                strSQL = strSQL & "AND ED_GLNO='" & rsTemp("GL_NUMBER") & "' "
                            End If
                            If Len(rsTemp("BD_ADMINBY")) > 0 Then
                                strSQL = strSQL & "AND ED_ADMINBY='" & rsTemp("BD_ADMINBY") & "' "
                            End If
                            If Len(rsTemp("BD_DIV")) > 0 Then
                                strSQL = strSQL & "AND ED_DIV='" & rsTemp("BD_DIV") & "' "
                            End If
                            
                            'TS Tech - Actual for FT and TMP only
                            If glbCompSerial = "S/N - 2369W" Then
                                strSQL = strSQL & "AND (TERM_HREMP.ED_PT='FT' Or TERM_HREMP.ED_PT='TMP') "
                            End If
                            
                            rsDATA.Open strSQL, gdbAdoIhr001X, adOpenStatic, adLockOptimistic
                            If rsDATA.EOF = False And rsDATA.BOF = False Then
                                Do
                                    'Insert term records retrieved into temp. table
                                    strSQL = "INSERT INTO HRMANTERM (ED_EMPNBR, ED_DEPTNO, ED_GLNO, ED_PT, ED_ADMINBY, JH_JOB, ED_LOC, ED_DIV, WRKEMP) "
                                    strSQL = strSQL & "VALUES (" & rsDATA("ED_EMPNBR") & ", '" & rsDATA("ED_DEPTNO") & "', '" & rsDATA("ED_GLNO") & "', '" & rsDATA("ED_PT") & "', '" & rsDATA("ED_ADMINBY") & "', '" & rsDATA("JH_JOB") & "', '" & rsDATA("ED_LOC") & "', '" & rsDATA("ED_DIV") & "', '" & glbUserID & "')"
                                    gdbAdoIhr001W.BeginTrans
                                    gdbAdoIhr001W.Execute strSQL
                                    gdbAdoIhr001W.CommitTrans
                                    rsDATA.MoveNext
                                Loop Until rsDATA.EOF
                            End If
                            rsDATA.Close
                            
                            'Join the Terminated Employees to Job Status and count.
                            strSQL = "SELECT Count(HRMANTERM.ED_EMPNBR) AS EMPCNT, HRMANTERM.ED_PT,  HRJOB.JB_STATUS "
                            If glbOracle Then
                                strSQL = strSQL & "FROM  HRMANTERM, hrjob "
                                strSQL = strSQL & "WHERE HRMANTERM.JH_JOB = hrjob.JB_CODE AND HRMANTERM.WRKEMP = '" & glbUserID & "'"
                            Else
                                strSQL = strSQL & "FROM  HRMANTERM INNER JOIN hrjob ON HRMANTERM.JH_JOB = hrjob.JB_CODE WHERE HRMANTERM.WRKEMP = '" & glbUserID & "'"
                            End If
                            strSQL = strSQL & "GROUP BY HRMANTERM.ED_PT,  HRJOB.JB_STATUS "
                            strSQL = strSQL & "HAVING (hrjob.JB_STATUS<>'NA')"
                            rsDATA.Open strSQL, gdbAdoIhr001W, adOpenStatic, adLockOptimistic
                            If rsDATA.EOF = False And rsDATA.BOF = False Then
                                Do
                                    'TS Tech - compute actual for FT and TMP
                                    If glbCompSerial = "S/N - 2369W" And rsDATA("JB_STATUS") = "L" Then
                                        If rsDATA("ED_PT") = "FT" Then
                                            If IsNull(FTS) Then FTS = 0
                                            FTS = FTS + rsDATA("EMPCNT")
                                        Else
                                            If IsNull(TMS) Then TMS = 0
                                            TMS = TMS + rsDATA("EMPCNT")
                                        End If
                                     Else
                                        'All the other clients, compute FT and other categories total
                                        If rsDATA("ED_PT") = "FT" Then
                                            If IsNull(FTA) Then FTA = 0
                                            FTA = FTA + rsDATA("EMPCNT")
                                        Else
                                            If IsNull(TMA) Then TMA = 0
                                            TMA = TMA + rsDATA("EMPCNT")
                                        End If
                                    End If
                                    rsDATA.MoveNext
                                Loop Until rsDATA.EOF
                            End If
                            rsDATA.Close
                        End If
                        
                        rsTemp("ACTUAL_FT_A") = FTA
                        rsTemp("ACTUAL_FT_S") = FTS
                        rsTemp("ACTUAL_TMP_A") = TMA
                        rsTemp("ACTUAL_TMP_s") = TMS

                        rsTemp.Update
                        
                        rsTemp.MoveNext
                    Loop Until rsTemp.EOF
                End If
                rsTemp.Close
            End If
        End If
        
    Next c

End Sub

Private Sub ReCalculate_Manpower_Actual_FTE()
    Dim strSQL As String
    Dim rsDATA As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim FTA, FTS, TMA, TMS
    Dim DTE As String
    Dim c% 'Month
    Dim x% 'Sequence

    'Reset the Actuals to Null first
    'For all the records in the Budget Year selected
    strSQL = "UPDATE HRBUDGET SET ACTUAL_FT_A = NULL, ACTUAL_TMP_A=NULL, ACTUAL_OTHER_A=NULL, ACTUAL_FT_S = NULL, ACTUAL_TMP_S=NULL, ACTUAL_OTHER_S=NULL "
    strSQL = strSQL & " WHERE BD_FREEZE=0 and BUDGET_YEAR=" & txtBYear.Text
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute strSQL
    gdbAdoIhr001.CommitTrans
    
    'For each Month of the Budget Year
    For c% = 1 To 12
        'Initialise
        x% = 0
        
        'Retrieve the Month Sequence
        strSQL = "SELECT MONTH_SEQ FROM HRBUDGET WHERE BUDGET_MONTH=" & c%
        strSQL = strSQL & " AND BUDGET_YEAR=" & txtBYear.Text
        rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
        
        If rsDATA.EOF = False And rsDATA.BOF = False Then
            x% = rsDATA("MONTH_SEQ")
        End If
        rsDATA.Close
            
        If x% > 0 Then
            'Compute end of the month date
            If c% < x% Then 'If month is less than the sequence then it must be next year
                DTE = getEOM(c%) & "/" & MonthName(c%, True) & "/" & (txtBYear.Text + 1)
            Else
                DTE = getEOM(c%) & "/" & MonthName(c%, True) & "/" & txtBYear.Text
            End If
            
            'Check if end of the month date date is less than today's date. Cannot calculate actual if the
            'budget end date is greater than today's date.
            If CDate(DTE) < Date Then
                'Find Actual Associates for this month(c)
                strSQL = "SELECT * FROM HRBUDGET WHERE BUDGET_YEAR=" & txtBYear.Text & " And BUDGET_MONTH = " & c%
                strSQL = strSQL & " AND BD_FREEZE=0"
                rsTemp.Open strSQL, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
                
                If rsTemp.EOF = False And rsTemp.EOF = False Then
                    Do
                        'Initialise
                        FTA = Null: FTS = Null: TMA = Null: TMS = Null
                        
                        'Calculate Actual for FT and other Categories
                        If rsTemp("BD_FTE") = 0 Or IsNull(rsTemp("BD_FTE")) Then
                            strSQL = "SELECT Count(HREMP.ED_EMPNBR) AS EMPCNT, HREMP.ED_PT, HRJOB.JB_STATUS "
                        Else
                            strSQL = "SELECT Sum(HR_JOB_HISTORY.JH_FTENUM) AS EMPCNT, HREMP.ED_PT, HRJOB.JB_STATUS "
                        End If
                        If glbOracle Then
                            strSQL = strSQL & "FROM HREMP, HR_JOB_HISTORY, HRJOB "
                            strSQL = strSQL & "WHERE HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE AND (HREMP.ED_DOH <= " & Date_SQL(DTE) & ") AND (HR_JOB_HISTORY.JH_CURRENT<>0)"
                        Else
                            strSQL = strSQL & "FROM  (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
                            strSQL = strSQL & "WHERE (HREMP.ED_DOH <= " & Date_SQL(DTE) & ")  AND (HR_JOB_HISTORY.JH_CURRENT<>0) "
                        End If
                        If Len(rsTemp("BD_DEPT")) > 0 Then
                            strSQL = strSQL & "AND ED_DEPTNO='" & rsTemp("BD_DEPT") & "' "
                        End If
                        If Len(rsTemp("GL_NUMBER")) > 0 Then
                            strSQL = strSQL & "AND ED_GLNO='" & rsTemp("GL_NUMBER") & "' "
                        End If
                        If Len(rsTemp("BD_ADMINBY")) > 0 Then
                            strSQL = strSQL & "AND ED_ADMINBY='" & rsTemp("BD_ADMINBY") & "' "
                        End If
                        If Len(rsTemp("BD_DIV")) > 0 Then
                            strSQL = strSQL & "AND ED_DIV='" & rsTemp("BD_DIV") & "' "
                        End If
                        strSQL = strSQL & "GROUP BY  HREMP.ED_PT, HRJOB.JB_STATUS "
                        strSQL = strSQL & "HAVING HRJOB.JB_STATUS <> 'NA' "
                        
                        'TS Tech - FT and TMP only
                        If glbCompSerial = "S/N - 2369W" Then
                            strSQL = strSQL & " AND (HREMP.ED_PT='FT' Or HREMP.ED_PT='TMP') "
                        End If
                        
                        rsDATA.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic
                        If rsDATA.EOF = False And rsDATA.BOF = False Then
                            Do
                                'TS Tech
                                If glbCompSerial = "S/N - 2369W" And rsDATA("JB_STATUS") = "L" Then
                                    If rsDATA("ED_PT") = "FT" Then
                                        If IsNull(FTS) Then FTS = 0
                                        FTS = FTS + rsDATA("EMPCNT")
                                    Else
                                        If IsNull(TMS) Then TMS = 0
                                        TMS = TMS + rsDATA("EMPCNT")
                                    End If
                                Else
                                    'All other clients
                                    If rsDATA("ED_PT") = "FT" Then
                                        If IsNull(FTA) Then FTA = 0
                                        FTA = FTA + IIf(IsNull(rsDATA("EMPCNT")), 0, rsDATA("EMPCNT"))
                                    Else
                                        If IsNull(TMA) Then TMA = 0
                                        TMA = TMA + IIf(IsNull(rsDATA("EMPCNT")), 0, rsDATA("EMPCNT"))
                                    End If
                                End If
                                rsDATA.MoveNext
                            Loop Until rsDATA.EOF
                        End If
                        rsDATA.Close
                    
   
                        'Changed by Bryan 07/Mar/2006 Ticket#10493
                        'Will select terminated employees terminated on the last day of the month.
                        '**********************
                        'Compute end of the month date
                        If c% < x% Then
                            DTE = getEOM(c%) & "/" & MonthName(c%, True) & "/" & (txtBYear.Text + 1)
                        Else
                            DTE = getEOM(c%) & "/" & MonthName(c%, True) & "/" & txtBYear
                        End If
                        
                        'Check if end of the month date date is less than today's date. Cannot calculate actual if the
                        'budget end date is greater than today's date.
                        If CDate(DTE) < Date Then
                            'Clear the Temp table to load the fresh set of Term employees
                            strSQL = "DELETE FROM HRMANTERM WHERE WRKEMP='" & glbUserID & "'"
                            gdbAdoIhr001W.BeginTrans
                            gdbAdoIhr001W.Execute strSQL
                            gdbAdoIhr001W.CommitTrans
                            
                            'Find Actual from the terminated for this month(c) - between Hire Date and Term Date
                            strSQL = "SELECT TERM_HREMP.ED_EMPNBR, TERM_HREMP.ED_DEPTNO, TERM_HREMP.ED_GLNO, TERM_HREMP.ED_PT, TERM_HREMP.ED_ADMINBY, Term_JOB_HISTORY.JH_JOB, TERM_HREMP.ED_LOC, TERM_HREMP.ED_DIV, Term_JOB_HISTORY.JH_FTENUM  "
                            If glbOracle Then
                                strSQL = strSQL & "FROM TERM_HREMP, TERM_HRTRMEMP, Term_JOB_HISTORY "
                                strSQL = strSQL & "WHERE TERM_HREMP.ED_EMPNBR = TERM_HRTRMEMP.TERM_SEQ AND TERM_HREMP.TERM_SEQ = Term_JOB_HISTORY.JH_EMPNBR AND (TERM_HREMP.ED_DOH <= " & Date_SQL(DTE) & ") AND (TERM_HRTRMEMP.TERM_DOT = " & Date_SQL(DTE) & ") "
                            Else
                                strSQL = strSQL & "FROM ((TERM_HREMP INNER JOIN TERM_HRTRMEMP ON TERM_HREMP.TERM_SEQ = TERM_HRTRMEMP.TERM_SEQ) INNER JOIN Term_JOB_HISTORY ON TERM_HREMP.TERM_SEQ = Term_JOB_HISTORY.TERM_SEQ)"
                                strSQL = strSQL & "WHERE (TERM_HREMP.ED_DOH <= " & Date_SQL(DTE) & ") AND (TERM_HRTRMEMP.TERM_DOT = " & Date_SQL(DTE) & ") "
                            End If
                            strSQL = strSQL & "AND Term_JOB_HISTORY.JH_CURRENT<>0"
                            If Len(rsTemp("BD_DEPT")) > 0 Then
                                strSQL = strSQL & "AND ED_DEPTNO='" & rsTemp("BD_DEPT") & "' "
                            End If
                            If Len(rsTemp("GL_NUMBER")) > 0 Then
                                strSQL = strSQL & "AND ED_GLNO='" & rsTemp("GL_NUMBER") & "' "
                            End If
                            If Len(rsTemp("BD_ADMINBY")) > 0 Then
                                strSQL = strSQL & "AND ED_ADMINBY='" & rsTemp("BD_ADMINBY") & "' "
                            End If
                            If Len(rsTemp("BD_DIV")) > 0 Then
                                strSQL = strSQL & "AND ED_DIV='" & rsTemp("BD_DIV") & "' "
                            End If
                            
                            'TS Tech - Actual for FT and TMP only
                            If glbCompSerial = "S/N - 2369W" Then
                                strSQL = strSQL & "AND (TERM_HREMP.ED_PT='FT' Or TERM_HREMP.ED_PT='TMP') "
                            End If
                            
                            rsDATA.Open strSQL, gdbAdoIhr001X, adOpenStatic, adLockOptimistic
                            If rsDATA.EOF = False And rsDATA.BOF = False Then
                                Do
                                    'Insert term records retrieved into temp. table
                                    strSQL = "INSERT INTO HRMANTERM (ED_EMPNBR, ED_DEPTNO, ED_GLNO, ED_PT, ED_ADMINBY, JH_JOB, ED_LOC, ED_DIV, JH_FTENUM, WRKEMP) "
                                    strSQL = strSQL & "VALUES (" & rsDATA("ED_EMPNBR") & ", '" & rsDATA("ED_DEPTNO") & "', '" & rsDATA("ED_GLNO") & "', '" & rsDATA("ED_PT") & "', '" & rsDATA("ED_ADMINBY") & "', '" & rsDATA("JH_JOB") & "', '" & rsDATA("ED_LOC") & "', '" & rsDATA("ED_DIV") & "', " & rsDATA("JH_FTENUM") & ", '" & glbUserID & "')"
                                    gdbAdoIhr001W.BeginTrans
                                    gdbAdoIhr001W.Execute strSQL
                                    gdbAdoIhr001W.CommitTrans
                                    rsDATA.MoveNext
                                Loop Until rsDATA.EOF
                            End If
                            rsDATA.Close
                            
                            'Join the Terminated Employees to Job Status and count.
                            If rsTemp("BD_FTE") = 0 Or IsNull(rsTemp("BD_FTE")) Then
                                strSQL = "SELECT Count(HRMANTERM.ED_EMPNBR) AS EMPCNT, HRMANTERM.ED_PT,  HRJOB.JB_STATUS "
                            Else
                                strSQL = "SELECT Sum(HRMANTERM.JH_FTENUM) AS EMPCNT, HRMANTERM.ED_PT,  HRJOB.JB_STATUS "
                            End If
                            If glbOracle Then
                                strSQL = strSQL & "FROM  HRMANTERM, hrjob "
                                strSQL = strSQL & "WHERE HRMANTERM.JH_JOB = hrjob.JB_CODE AND HRMANTERM.WRKEMP = '" & glbUserID & "'"
                            Else
                                strSQL = strSQL & "FROM  HRMANTERM INNER JOIN hrjob ON HRMANTERM.JH_JOB = hrjob.JB_CODE WHERE HRMANTERM.WRKEMP = '" & glbUserID & "'"
                            End If
                            strSQL = strSQL & "GROUP BY HRMANTERM.ED_PT,  HRJOB.JB_STATUS "
                            strSQL = strSQL & "HAVING (hrjob.JB_STATUS<>'NA')"
                            rsDATA.Open strSQL, gdbAdoIhr001W, adOpenStatic, adLockOptimistic
                            If rsDATA.EOF = False And rsDATA.BOF = False Then
                                Do
                                    'TS Tech - compute actual for FT and TMP
                                    If glbCompSerial = "S/N - 2369W" And rsDATA("JB_STATUS") = "L" Then
                                        If rsDATA("ED_PT") = "FT" Then
                                            If IsNull(FTS) Then FTS = 0
                                            FTS = FTS + rsDATA("EMPCNT")
                                        Else
                                            If IsNull(TMS) Then TMS = 0
                                            TMS = TMS + rsDATA("EMPCNT")
                                        End If
                                     Else
                                        'All the other clients, compute FT and other categories total
                                        If rsDATA("ED_PT") = "FT" Then
                                            If IsNull(FTA) Then FTA = 0
                                            FTA = FTA + IIf(IsNull(rsDATA("EMPCNT")), 0, rsDATA("EMPCNT"))
                                        Else
                                            If IsNull(TMA) Then TMA = 0
                                            TMA = TMA + IIf(IsNull(rsDATA("EMPCNT")), 0, rsDATA("EMPCNT"))
                                        End If
                                    End If
                                    rsDATA.MoveNext
                                Loop Until rsDATA.EOF
                            End If
                            rsDATA.Close
                        End If
                        
                        rsTemp("ACTUAL_FT_A") = FTA
                        rsTemp("ACTUAL_FT_S") = FTS
                        rsTemp("ACTUAL_TMP_A") = TMA
                        rsTemp("ACTUAL_TMP_s") = TMS

                        rsTemp.Update
                        
                        rsTemp.MoveNext
                    Loop Until rsTemp.EOF
                End If
                rsTemp.Close
            End If
        End If
        
    Next c

End Sub

Public Function getEOM(Mnt As Variant) As Integer
   Dim myDate As Date
   Dim myMonth As String
   Dim NextMonth As Date, EndOfMonth As Date
   
   'Get End of the Month date(Day)
   If IsNumeric(Mnt) Then
        myMonth = MonthName(Mnt, True)
    Else
        myMonth = Mnt
    End If
   
   myDate = Format("1/" & myMonth & "/2005", "dd/mmm/yyyy")
   NextMonth = DateAdd("m", 1, myDate)
   EndOfMonth = NextMonth - DatePart("d", NextMonth)
   getEOM = Day(EndOfMonth)

End Function

