VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmSWorkSchRule 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Work Schedule Rules"
   ClientHeight    =   9420
   ClientLeft      =   90
   ClientTop       =   1005
   ClientWidth     =   13920
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
   ScaleHeight     =   12930
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   6375
      LargeChange     =   300
      Left            =   13440
      Max             =   3000
      SmallChange     =   300
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmSWorkSchRule.frx":0000
      Height          =   2535
      Left            =   240
      OleObjectBlob   =   "frmSWorkSchRule.frx":0014
      TabIndex        =   12
      Top             =   240
      Width           =   13215
   End
   Begin VB.Frame frWorkScheduleRules 
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   240
      TabIndex        =   14
      Top             =   3000
      Width           =   13095
      Begin VB.TextBox txtTIgnoreHolidays 
         Appearance      =   0  'Flat
         DataField       =   "WR_T_IGNOR_HOLIDAYS"
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
         Left            =   6750
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   5790
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkTIgnoreHolidays 
         Caption         =   "Always Ignore Satutory Holidays in Holiday Master (Overrides Application Settings)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   510
         TabIndex        =   11
         Top             =   5760
         Width           =   6135
      End
      Begin VB.TextBox txtTIgnoreWKEnds 
         Appearance      =   0  'Flat
         DataField       =   "WR_T_IGNOR_WKEND"
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
         Left            =   6750
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   5400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkTIgnoreWKEnds 
         Caption         =   "Always Ignore Weekends (Overrides Application Settings)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   510
         TabIndex        =   10
         Tag             =   " "
         Top             =   5370
         Width           =   5535
      End
      Begin VB.TextBox txtTOvrideWrkSch 
         Appearance      =   0  'Flat
         DataField       =   "WR_T_OVRIDE_HRS"
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
         Left            =   6750
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   4995
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkTOvrideWrkSch 
         Caption         =   "Override Work Schedule 'Hours'"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   510
         TabIndex        =   9
         Top             =   4965
         Width           =   5535
      End
      Begin VB.TextBox txtTDefaultHrs 
         Appearance      =   0  'Flat
         DataField       =   "WR_T_DFLT_HRS"
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
         Left            =   6750
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4605
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkTDefaultHrs 
         Caption         =   "Default 'Hours' from Work Schedule"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   510
         TabIndex        =   8
         Top             =   4575
         Width           =   5535
      End
      Begin VB.TextBox txtTSingleDay 
         Appearance      =   0  'Flat
         DataField       =   "WR_T_SNGL_REQ"
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
         Left            =   6750
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   4215
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkTSingleDay 
         Caption         =   "Single Day Requests Ignores Work Schedule"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   510
         TabIndex        =   7
         Top             =   4185
         Width           =   5535
      End
      Begin VB.TextBox txtVSingleDay 
         Appearance      =   0  'Flat
         DataField       =   "WR_V_SNGL_REQ"
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
         Left            =   6750
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1770
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkVDefaultHrs 
         Caption         =   "Default 'Hours' from Work Schedule"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   510
         TabIndex        =   3
         Top             =   2130
         Width           =   5535
      End
      Begin VB.TextBox txtVDefaultHrs 
         Appearance      =   0  'Flat
         DataField       =   "WR_V_DFLT_HRS"
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
         Left            =   6750
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkVOvrideWrkSch 
         Caption         =   "Override Work Schedule 'Hours'"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   510
         TabIndex        =   4
         Top             =   2520
         Width           =   5535
      End
      Begin VB.TextBox txtVOvrideWrkSch 
         Appearance      =   0  'Flat
         DataField       =   "WR_V_OVRIDE_HRS"
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
         Left            =   6750
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2550
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkVIgnoreWKEnds 
         Caption         =   "Always Ignore Weekends (Overrides Application Settings)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   510
         TabIndex        =   5
         Tag             =   " "
         Top             =   2925
         Width           =   5535
      End
      Begin VB.TextBox txtVIgnoreWKEnds 
         Appearance      =   0  'Flat
         DataField       =   "WR_V_IGNOR_WKEND"
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
         Left            =   6750
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2955
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkVIgnoreHolidays 
         Caption         =   "Always Ignore Statutory Holidays in Holiday Master (Overrides Application Settings)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   510
         TabIndex        =   6
         Top             =   3315
         Width           =   6135
      End
      Begin VB.TextBox txtVIgnoreHolidays 
         Appearance      =   0  'Flat
         DataField       =   "WR_V_IGNOR_HOLIDAYS"
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
         Left            =   6750
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3345
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "WR_LUSER"
         Enabled         =   0   'False
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
         Left            =   11310
         MaxLength       =   10
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3375
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "WR_LDATE"
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   9870
         MaxLength       =   12
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3375
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "WR_LTIME"
         Enabled         =   0   'False
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
         Left            =   10620
         MaxLength       =   8
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3375
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CheckBox chkVSingleDay 
         Caption         =   "Single Day Requests Ignores Work Schedule"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   510
         TabIndex        =   2
         Top             =   1740
         Width           =   5535
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   9870
         Top             =   2775
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   9390
         Top             =   3615
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
         GridSource      =   "vbxTrueGrid"
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "WR_ADMINBY"
         Height          =   285
         Index           =   2
         Left            =   9240
         TabIndex        =   28
         Tag             =   "00-Administered By"
         Top             =   4920
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "WR_ORG"
         Height          =   285
         Index           =   1
         Left            =   9390
         TabIndex        =   29
         Tag             =   "00-Union"
         Top             =   6000
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOR"
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         DataField       =   "WR_DEPT"
         Height          =   285
         Left            =   9225
         TabIndex        =   30
         Tag             =   "00-Department"
         Top             =   4500
         Visible         =   0   'False
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   7
         LookupType      =   2
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         DataField       =   "WR_DIV"
         Height          =   285
         Left            =   9225
         TabIndex        =   31
         Tag             =   "00-Division"
         Top             =   4080
         Visible         =   0   'False
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "WR_EMP"
         Height          =   285
         Index           =   3
         Left            =   1590
         TabIndex        =   0
         Tag             =   "00-Employment Status"
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDEM"
      End
      Begin INFOHR_Controls.CodeLookup clpPT 
         DataField       =   "WR_PT"
         Height          =   285
         Left            =   1590
         TabIndex        =   1
         Tag             =   "EDPT-Category"
         Top             =   660
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDPT"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "WR_LOC"
         Height          =   285
         Index           =   4
         Left            =   9390
         TabIndex        =   32
         Tag             =   "00-Location"
         Top             =   6420
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "WR_SECTION"
         Height          =   285
         Index           =   0
         Left            =   9240
         TabIndex        =   33
         Tag             =   "00-Section"
         Top             =   5340
         Visible         =   0   'False
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOR"
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Time Requests:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   44
         Top             =   3855
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vacation Requests:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   43
         Top             =   1455
         Width           =   1680
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
         Left            =   7920
         TabIndex        =   42
         Top             =   6465
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblSection 
         Appearance      =   0  'Flat
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
         Left            =   7950
         TabIndex        =   41
         Top             =   5385
         Visible         =   0   'False
         Width           =   1260
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
         TabIndex        =   40
         Top             =   705
         Width           =   630
      End
      Begin VB.Label lblAdmin 
         Appearance      =   0  'Flat
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
         Left            =   7950
         TabIndex        =   39
         Top             =   4965
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lblEmpStatys 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employment Status"
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
         Top             =   285
         Width           =   1350
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
         Left            =   7920
         TabIndex        =   37
         Top             =   6045
         Visible         =   0   'False
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
         Left            =   7950
         TabIndex        =   36
         Top             =   4545
         Visible         =   0   'False
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
         Left            =   7950
         TabIndex        =   35
         Top             =   4125
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblCNum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comp"
         DataField       =   "WR_COMPNO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   10020
         TabIndex        =   34
         Top             =   3825
         Visible         =   0   'False
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmSWorkSchRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim rsDATA As New ADODB.Recordset
Dim UpdateState As UpdateStateEnum
Dim fglbVSQLQ As String
Dim fglbESQLQ As String

Private Function chkWorkScheduleRule()
Dim Msg As String
Dim x%, xchk
Dim SQLQ As String
Dim rsWR As New ADODB.Recordset
Dim xID
Dim a As Integer
Dim xWorkSch

chkWorkScheduleRule = False

For x% = 3 To 3     '0 To 4
    If Len(clpCode(x%).Text) > 0 And clpCode(x%).Caption = "Unassigned" Then
        MsgBox "If Code entered it must be valid."
        clpCode(x%).SetFocus
        Exit Function
    End If
Next x%

'Make sure if Employment Status entered is of Work Schedule type.
If Len(clpCode(3).Text) > 0 Then
    xWorkSch = GetCode_Data("EDEM", clpCode(3), "TB_WORKSCHED", False)
    If Not xWorkSch Or IsNull(xWorkSch) Then
        MsgBox "The Employment Status selected is not of Work Schedule type code.", vbExclamation, "Work Schedule Type Code"
        clpCode(3).SetFocus
        Exit Function
    End If
End If

'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
'    MsgBox lStr("If Department Entered - it must be valid.")
'    clpDept.SetFocus
'    Exit Function
'End If

'If clpDiv.Caption = "Unassigned" Then
'    MsgBox lStr("If Division Entered - it must be valid.")
'    clpDiv.SetFocus
'    Exit Function
'End If

If clpPT.Caption = "Unassigned" Then
    MsgBox "If " & lblPT.Caption & " Entered - it must be valid."
    clpPT.SetFocus
    Exit Function
End If

Call getWSQLQ
SQLQ = "SELECT * FROM HRWORKSCHDRULE WHERE " & fglbVSQLQ
rsWR.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsWR.EOF Then
    MsgBox "You can not add duplicate Work Schedule Rule.", vbExclamation
    clpCode(3).SetFocus
    Exit Function
End If
rsWR.Close
Set rsWR = Nothing

'Check if rule with not selection criteria found
If Len(clpCode(3).Text) > 0 Or Len(clpPT.Text) > 0 Then
    SQLQ = "SELECT * FROM HRWORKSCHDRULE WHERE (WR_EMP IS NULL OR WR_EMP='') AND (WR_PT IS NULL OR WR_PT='')"
    If fglbNew Then
        xID = 0
    Else
        If Not rsDATA.EOF Then
            xID = rsDATA("WR_ID")
        Else
            xID = 0
        End If
    End If
    If xID > 0 Then
        SQLQ = SQLQ & " AND NOT WR_ID = " & xID & " "
    End If
    rsWR.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsWR.EOF Then
        MsgBox "A Work Schedule Rule which applies to everyone already exists.", vbExclamation
        clpCode(3).SetFocus
        Exit Function
    End If
    rsWR.Close
    Set rsWR = Nothing
End If

'Check if rule with Employment Status only already exists
If Len(clpCode(3).Text) > 0 Then
    SQLQ = "SELECT * FROM HRWORKSCHDRULE WHERE WR_EMP = '" & clpCode(3).Text & "' AND (WR_PT IS NULL OR WR_PT='')"
    If fglbNew Then
        xID = 0
    Else
        If Not rsDATA.EOF Then
            xID = rsDATA("WR_ID")
        Else
            xID = 0
        End If
    End If
    If xID > 0 Then
        SQLQ = SQLQ & " AND NOT WR_ID = " & xID & " "
    End If
    rsWR.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsWR.EOF Then
        MsgBox "A Work Schedule Rule for this Employment Status which applies to everyone already exists.", vbExclamation
        clpCode(3).SetFocus
        Exit Function
    End If
    rsWR.Close
    Set rsWR = Nothing
End If

'Check if rule with Category only already exists
If Len(clpPT.Text) > 0 Then
    SQLQ = "SELECT * FROM HRWORKSCHDRULE WHERE (WR_EMP IS NULL OR WR_EMP='') AND (WR_PT = '" & clpPT.Text & "')"
    If fglbNew Then
        xID = 0
    Else
        If Not rsDATA.EOF Then
            xID = rsDATA("WR_ID")
        Else
            xID = 0
        End If
    End If
    If xID > 0 Then
        SQLQ = SQLQ & " AND NOT WR_ID = " & xID & " "
    End If
    rsWR.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsWR.EOF Then
        MsgBox "A Work Schedule Rule for this " & lStr("Category") & " which applies to everyone already exists.", vbExclamation
        clpPT.SetFocus
        Exit Function
    End If
    rsWR.Close
    Set rsWR = Nothing
End If

'Check if rule for a specific group already exists when adding a Employment Status rule for everyone
If Len(clpCode(3).Text) > 0 And Len(clpPT.Text) = 0 Then
    SQLQ = "SELECT * FROM HRWORKSCHDRULE WHERE WR_EMP = '" & clpCode(3).Text & "' AND (WR_PT IS NOT NULL)"
    If fglbNew Then
        xID = 0
    Else
        If Not rsDATA.EOF Then
            xID = rsDATA("WR_ID")
        Else
            xID = 0
        End If
    End If
    If xID > 0 Then
        SQLQ = SQLQ & " AND NOT WR_ID = " & xID & " "
    End If
    rsWR.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsWR.EOF Then
        MsgBox "A Work Schedule Rule which applies to this Employment Status and specific " & lStr("Category") & " already exists.", vbExclamation
        clpCode(3).SetFocus
        Exit Function
    End If
    rsWR.Close
    Set rsWR = Nothing
End If

'Check if rule for a specific group already exists when adding a Category rule for everyone
If Len(clpCode(3).Text) = 0 And Len(clpPT.Text) > 0 Then
    SQLQ = "SELECT * FROM HRWORKSCHDRULE WHERE (WR_EMP IS NOT NULL) AND (WR_PT = '" & clpPT.Text & "')"
    If fglbNew Then
        xID = 0
    Else
        If Not rsDATA.EOF Then
            xID = rsDATA("WR_ID")
        Else
            xID = 0
        End If
    End If
    If xID > 0 Then
        SQLQ = SQLQ & " AND NOT WR_ID = " & xID & " "
    End If
    rsWR.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsWR.EOF Then
        MsgBox "A Work Schedule Rule which applies to a specific Employment Status and this " & lStr("Category") & " already exists.", vbExclamation
        clpPT.SetFocus
        Exit Function
    End If
    rsWR.Close
    Set rsWR = Nothing
End If

'Additional rules by Jerry
If (chkVSingleDay And chkVDefaultHrs) Or (chkTSingleDay And chkTDefaultHrs) Then
    MsgBox "If 'Single Day Requests Ignores Work Schedule' and 'Default 'Hours' from Work Schedule' are both checked, the default hours will be taken from the Employee's Current Position and can be overwritten by the employee.", vbInformation, "info:HR - Default Hours"
End If

If (chkVSingleDay And chkVDefaultHrs And chkVOvrideWrkSch) Or (chkTSingleDay And chkTDefaultHrs And chkTOvrideWrkSch) Then
    a% = MsgBox("Conditions checked above can have unusual consequences in ESS." & vbCrLf & vbCrLf & "Are you sure you want to save this?", vbExclamation + vbYesNo, "info:HR - Confirm Conditions Checked")
    If a% <> 6 Then
        Exit Function
    End If
End If

chkWorkScheduleRule = True

End Function

Sub cmdCancel_Click()

On Error GoTo Can_Err

fglbNew = False

If fglbEmptyNew Then
    Me.vbxTrueGrid.Enabled = True
    Me.vbxTrueGrid.Refresh
End If

rsDATA.CancelUpdate

Call Display_Value

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HRWORKSCHDRULE", "Cancel")
Call RollBack '09June99 js

End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String

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

Call SET_UP_MODE
'Call ST_UPD_MODE(False)


Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRWORKSCHDRULE", "Delete")
Call RollBack '09June99 js

End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Call ST_UPD_MODE(True)

'clpDiv.SetFocus
clpCode(3).SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HRWORKSCHDRULE", "Modify")
Call RollBack '09June99 js

End Sub

Sub cmdNew_Click()

On Error GoTo AddN_Err

Call Set_Control("B", Me)

rsDATA.AddNew

lblCNum.Caption = "001"

fglbNew = True

Call SET_UP_MODE

'Call ST_UPD_MODE(True)
clpCode(3).SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRWORKSCHDRULE", "Add")
Call RollBack '09June99 js

End Sub

Sub cmdOK_Click()
Dim x%
Dim bmk As Variant

On Error GoTo cmdOK_Err

If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    bmk = 0
Else
    bmk = Data1.Recordset.Bookmark
End If

If Not chkWorkScheduleRule() Then Exit Sub


Call UpdUStats(Me) ' update user's stats (who did it and when)

Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans

Data1.Refresh
If Not bmk = 0 Then
    Data1.Recordset.Bookmark = bmk
End If

fglbNew = False

Call Display_Value

Me.vbxTrueGrid.Enabled = True
Me.vbxTrueGrid.SetFocus
Screen.MousePointer = DEFAULT

Exit Sub

cmdOK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRWORKSCHDRULE", "Update")
Call RollBack '09June99 js

End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = "Work Schedule Rules"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = "Work Schedule Rules"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub chkVSingleDay_Click()
    If chkVSingleDay.Value = 1 Then
        txtVSingleDay.Text = "1"
    Else
        txtVSingleDay.Text = "0"
    End If
End Sub

Private Sub chkVDefaultHrs_Click()
    If chkVDefaultHrs.Value = 1 Then
        txtVDefaultHrs.Text = "1"
    Else
        txtVDefaultHrs.Text = "0"
    End If
End Sub

Private Sub chkVOvrideWrkSch_Click()
    If chkVOvrideWrkSch.Value = 1 Then
        txtVOvrideWrkSch.Text = "1"
    Else
        txtVOvrideWrkSch.Text = "0"
    End If
End Sub

Private Sub chkVIgnoreWKEnds_Click()
    If chkVIgnoreWKEnds.Value = 1 Then
        txtVIgnoreWKEnds.Text = "1"
    Else
        txtVIgnoreWKEnds.Text = "0"
    End If
End Sub

Private Sub chkVIgnoreHolidays_Click()
    If chkVIgnoreHolidays.Value = 1 Then
        txtVIgnoreHolidays.Text = "1"
    Else
        txtVIgnoreHolidays.Text = "0"
    End If
End Sub

Private Sub chkTSingleDay_Click()
    If chkTSingleDay.Value = 1 Then
        txtTSingleDay.Text = "1"
    Else
        txtTSingleDay.Text = "0"
    End If
End Sub

Private Sub chkTDefaultHrs_Click()
    If chkTDefaultHrs.Value = 1 Then
        txtTDefaultHrs.Text = "1"
    Else
        txtTDefaultHrs.Text = "0"
    End If
End Sub

Private Sub chkTOvrideWrkSch_Click()
    If chkTOvrideWrkSch.Value = 1 Then
        txtTOvrideWrkSch.Text = "1"
    Else
        txtTOvrideWrkSch.Text = "0"
    End If
End Sub

Private Sub chkTIgnoreWKEnds_Click()
    If chkTIgnoreWKEnds.Value = 1 Then
        txtTIgnoreWKEnds.Text = "1"
    Else
        txtTIgnoreWKEnds.Text = "0"
    End If
End Sub

Private Sub chkTIgnoreHolidays_Click()
    If chkTIgnoreHolidays.Value = 1 Then
        txtTIgnoreHolidays.Text = "1"
    Else
        txtTIgnoreHolidays.Text = "0"
    End If
End Sub

Private Sub clpCode_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRWORKSCHDRULE", "SELECT")

End Sub

Private Sub Form_Activate()

Call SET_UP_MODE

Me.cmdModify_Click

End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim SQLQ

'Me.Show

glbOnTop = "FRMSWORKSCHRULE"

Screen.MousePointer = HOURGLASS

Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "SELECT * FROM HRWORKSCHDRULE ORDER BY WR_ID "
Data1.Refresh

Screen.MousePointer = DEFAULT

'Call setRptCaption(Me)
Call setCaption(lblPT)

'Call setCaption(lblDiv)
'Call setCaption(lblDept)
'Call setCaption(lblLocation)
'Call setCaption(lblRegion)
'Call setCaption(lblAdmin)
'Call setCaption(lblSection)
'Call setCaption(lblUnion)
'Call setCaption(lblPT)

vbxTrueGrid.Columns(0).Caption = lStr("Division")
vbxTrueGrid.Columns(1).Caption = lStr("Department")
vbxTrueGrid.Columns(2).Caption = lStr("Administered By")
vbxTrueGrid.Columns(3).Caption = lStr("Section")
vbxTrueGrid.Columns(5).Caption = lStr("Category")
vbxTrueGrid.Columns(6).Caption = lStr("Union")
vbxTrueGrid.Columns(7).Caption = lStr("Location")

vbxTrueGrid.Columns(0).Visible = False  '.lStr("Division")
vbxTrueGrid.Columns(1).Visible = False  'lStr("Department")
vbxTrueGrid.Columns(2).Visible = False  'lStr("Administered By")
vbxTrueGrid.Columns(3).Visible = False  'lStr("Section")
vbxTrueGrid.Columns(6).Visible = False  'lStr("Union")
vbxTrueGrid.Columns(7).Visible = False  'lStr("Location")


'Call Display_Value

Call ST_UPD_MODE(False)

'vbxTrueGrid.Columns(0).Caption = lStr("Section")
'lblTitle(0).Caption = lStr(lblTitle(0).Caption)
'lblTitle(1).Caption = lStr(lblTitle(1).Caption)
'lblTitle(2).Caption = lStr(lblTitle(2).Caption)
Call INI_Controls(Me)

Screen.MousePointer = DEFAULT                           '

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
On Error GoTo Err_WorkScheduleRule_Scroll

If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    If Me.Height >= vbxTrueGrid.Height + frWorkScheduleRules.Height + 1000 Then
        scrControl.Value = 0
        frWorkScheduleRules.Top = vbxTrueGrid.Height + 520
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        scrControl.Left = Me.ScaleWidth - scrControl.Width
        scrControl.Height = Me.Height - vbxTrueGrid.Height - 1000
    End If

End If

Cont:
Exit Sub

Err_WorkScheduleRule_Scroll:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form_Resize", "Work Schedule Rules", "Form Resize")
    Resume Cont
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
Dim I As Integer
End Sub

Private Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

fUPMode = TF
'vbxTrueGrid.Enabled = FT
'cmdOK.Enabled = TF              '
'cmdCancel.Enabled = TF          '
'cmdClose.Enabled = FT           '
'cmdModify.Enabled = FT          '
'cmdNew.Enabled = FT             '
'cmdDelete.Enabled = FT          '
'cmdPrint.Enabled = FT           '

'clpDiv.Enabled = TF
'clpDept.Enabled = TF
clpPT.Enabled = TF
'clpCode(0).Enabled = TF
'clpCode(1).Enabled = TF
'clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
'clpCode(4).Enabled = TF
chkVSingleDay.Enabled = TF      '
chkVDefaultHrs.Enabled = TF      '
chkVOvrideWrkSch.Enabled = TF      '
chkVIgnoreWKEnds.Enabled = TF      '
chkVIgnoreHolidays.Enabled = TF
chkTSingleDay.Enabled = TF      '
chkTDefaultHrs.Enabled = TF      '
chkTOvrideWrkSch.Enabled = TF      '
chkTIgnoreWKEnds.Enabled = TF      '
chkTIgnoreHolidays.Enabled = TF

End Sub

Private Sub scrControl_Change()
    frWorkScheduleRules.Top = 3000 - scrControl.Value
End Sub

Private Sub txtVSingleDay_Change()
    If txtVSingleDay = "-1" Or txtVSingleDay = "1" Then
        chkVSingleDay.Value = 1
    Else
        chkVSingleDay.Value = 0
    End If
End Sub

Private Sub txtVDefaultHrs_Change()
    If txtVDefaultHrs = "-1" Or txtVDefaultHrs = "1" Then
        chkVDefaultHrs.Value = 1
    Else
        chkVDefaultHrs.Value = 0
    End If
End Sub

Private Sub txtVOvrideWrkSch_Change()
    If txtVOvrideWrkSch = "-1" Or txtVOvrideWrkSch = "1" Then
        chkVOvrideWrkSch.Value = 1
    Else
        chkVOvrideWrkSch.Value = 0
    End If
End Sub

Private Sub txtVIgnoreWKEnds_Change()
    If txtVIgnoreWKEnds = "-1" Or txtVIgnoreWKEnds = "1" Then
        chkVIgnoreWKEnds.Value = 1
    Else
        chkVIgnoreWKEnds.Value = 0
    End If
End Sub

Private Sub txtVIgnoreHolidays_Change()
    If txtVIgnoreHolidays = "-1" Or txtVIgnoreHolidays = "1" Then
        chkVIgnoreHolidays.Value = 1
    Else
        chkVIgnoreHolidays.Value = 0
    End If
End Sub

Private Sub txtTSingleDay_Change()
    If txtTSingleDay = "-1" Or txtTSingleDay = "1" Then
        chkTSingleDay.Value = 1
    Else
        chkTSingleDay.Value = 0
    End If
End Sub

Private Sub txtTDefaultHrs_Change()
    If txtTDefaultHrs = "-1" Or txtTDefaultHrs = "1" Then
        chkTDefaultHrs.Value = 1
    Else
        chkTDefaultHrs.Value = 0
    End If
End Sub

Private Sub txtTOvrideWrkSch_Change()
    If txtTOvrideWrkSch = "-1" Or txtTOvrideWrkSch = "1" Then
        chkTOvrideWrkSch.Value = 1
    Else
        chkTOvrideWrkSch.Value = 0
    End If
End Sub

Private Sub txtTIgnoreWKEnds_Change()
    If txtTIgnoreWKEnds = "-1" Or txtTIgnoreWKEnds = "1" Then
        chkTIgnoreWKEnds.Value = 1
    Else
        chkTIgnoreWKEnds.Value = 0
    End If
End Sub

Private Sub txtTIgnoreHolidays_Change()
    If txtTIgnoreHolidays = "-1" Or txtTIgnoreHolidays = "1" Then
        chkTIgnoreHolidays.Value = 1
    Else
        chkTIgnoreHolidays.Value = 0
    End If
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
           
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    SQLQ = "SELECT * FROM HRWORKSCHDRULE "
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I%

On Error GoTo vbxTrueGrid_Err

Call Display_Value

If Data1.Recordset.EOF Or Data1.Recordset.BOF = 0 Then
    Exit Sub
End If

Exit Sub

vbxTrueGrid_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HRWORKSCHDRULE", "Select")
Call RollBack

End Sub

Private Function RollBack()
On Error GoTo rr
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
rr:
End Function

Private Sub Display_Value()
    Dim SQLQ
    
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
        Call SET_UP_MODE
        
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HRWORKSCHDRULE where WR_ID= " & Data1.Recordset!WR_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    
    Call Set_Control("R", Me, rsDATA)
    
    Call SET_UP_MODE
    
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fglbNew = True
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_WorkSchRule
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

Call ST_UPD_MODE(TF)

Call set_Buttons(UpdateState)

If Not UpdateRight Then TF = False

End Sub

Private Sub getWSQLQ() 'xType)
Dim xDiv, xDept, xORG, xAsOf, xEMP, xEmpMode, xGRPCE
Dim xLoc, xSection
Dim xFromDate
Dim xToDate
Dim xID
Dim SQLQ As String

fglbESQLQ = "" 'glbSeleDeptUn
fglbVSQLQ = " (1=1) "

'If Len(clpDiv.Text) = 0 Then
'    fglbVSQLQ = fglbVSQLQ & "AND (WR_DIV IS NULL OR WR_DIV='') "
'Else
'    fglbVSQLQ = fglbVSQLQ & "AND WR_DIV = '" & clpDiv.Text & "' "
'End If

'If Len(clpDept.Text) = 0 Then
'    fglbVSQLQ = fglbVSQLQ & " AND (WR_DEPT IS NULL OR WR_DEPT='') "
'Else
'    fglbVSQLQ = fglbVSQLQ & " AND WR_DEPT = '" & clpDept.Text & "' "
'End If

'If Len(clpCode(1).Text) = 0 Then
'    fglbVSQLQ = fglbVSQLQ & " AND (WR_ORG IS NULL OR WR_ORG='') "
'Else
'    fglbVSQLQ = fglbVSQLQ & " AND WR_ORG = '" & clpCode(1).Text & "' "
'End If

If Len(clpCode(3).Text) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (WR_EMP IS NULL OR WR_EMP='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND WR_EMP = '" & clpCode(3).Text & "' "
End If

If Len(clpPT.Text) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (WR_PT IS NULL OR WR_PT='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND WR_PT = '" & clpPT.Text & "' "
End If

'If Len(clpCode(2).Text) = 0 Then
'    fglbVSQLQ = fglbVSQLQ & " AND (WR_ADMINBY IS NULL OR WR_ADMINBY='') "
'Else
'    fglbVSQLQ = fglbVSQLQ & " AND WR_ADMINBY = '" & clpCode(2).Text & "' "
'End If

'If Len(clpCode(0).Text) = 0 Then
'    fglbVSQLQ = fglbVSQLQ & " AND (WR_SECTION IS NULL OR WR_SECTION='') "
'Else
'    fglbVSQLQ = fglbVSQLQ & " AND WR_SECTION = '" & clpCode(0).Text & "' "
'End If

'If Len(clpCode(4).Text) = 0 Then
'    fglbVSQLQ = fglbVSQLQ & " AND (WR_LOC IS NULL OR WR_LOC='') "
'Else
'    fglbVSQLQ = fglbVSQLQ & " AND WR_LOC = '" & clpCode(4).Text & "' "
'End If

If fglbNew Then
    xID = 0
Else
    If Not rsDATA.EOF Then
        xID = rsDATA("WR_ID")
    Else
        xID = 0
    End If
End If
If xID > 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND NOT WR_ID = " & xID & " "
End If
'getWSQLQ = fglbVSQLQ

End Sub

