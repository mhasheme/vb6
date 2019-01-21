VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEHSINCIDENT 
   AutoRedraw      =   -1  'True
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   900
   ClientWidth     =   12165
   DrawWidth       =   2
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10950
   ScaleWidth      =   12165
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDemo 
      Appearance      =   0  'Flat
      DataField       =   "EC_JOBDESC"
      Height          =   315
      Index           =   14
      Left            =   10305
      MaxLength       =   25
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtDemo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   12
      Left            =   8880
      MaxLength       =   25
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDemo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   13
      Left            =   9600
      MaxLength       =   25
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.VScrollBar scrControl 
      Height          =   5535
      LargeChange     =   315
      Left            =   11775
      Max             =   200
      SmallChange     =   315
      TabIndex        =   65
      Top             =   3240
      Width           =   300
   End
   Begin VB.TextBox txtDemo 
      Appearance      =   0  'Flat
      DataField       =   "EC_WORKCOUNTRY"
      Height          =   315
      Index           =   11
      Left            =   8040
      MaxLength       =   25
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDemo 
      Appearance      =   0  'Flat
      DataField       =   "EC_ADMINBY"
      Height          =   315
      Index           =   10
      Left            =   7320
      MaxLength       =   25
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDemo 
      Appearance      =   0  'Flat
      DataField       =   "EC_SECTION"
      Height          =   315
      Index           =   9
      Left            =   6600
      MaxLength       =   25
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDemo 
      Appearance      =   0  'Flat
      DataField       =   "EC_REGION"
      Height          =   315
      Index           =   8
      Left            =   5880
      MaxLength       =   25
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDemo 
      Appearance      =   0  'Flat
      DataField       =   "EC_PT"
      Height          =   315
      Index           =   7
      Left            =   5160
      MaxLength       =   25
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDemo 
      Appearance      =   0  'Flat
      DataField       =   "EC_EMP"
      Height          =   315
      Index           =   6
      Left            =   4440
      MaxLength       =   25
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDemo 
      Appearance      =   0  'Flat
      DataField       =   "EC_ORG"
      Height          =   315
      Index           =   5
      Left            =   3720
      MaxLength       =   25
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDemo 
      Appearance      =   0  'Flat
      DataField       =   "EC_LOC"
      Height          =   315
      Index           =   4
      Left            =   3000
      MaxLength       =   25
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDemo 
      Appearance      =   0  'Flat
      DataField       =   "EC_DIV"
      Height          =   315
      Index           =   3
      Left            =   2280
      MaxLength       =   25
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDemo 
      Appearance      =   0  'Flat
      DataField       =   "EC_DEPTNO"
      Height          =   315
      Index           =   2
      Left            =   1560
      MaxLength       =   25
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   615
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fehsinc.frx":0000
      Height          =   2325
      Left            =   120
      OleObjectBlob   =   "fehsinc.frx":0014
      TabIndex        =   45
      Top             =   600
      Width           =   11895
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   10440
      Top             =   10380
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Ado3"
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8520
      Top             =   10380
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Ado2"
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EC_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   2430
      MaxLength       =   25
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   10320
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EC_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   4170
      MaxLength       =   25
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   10320
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EC_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   5910
      MaxLength       =   25
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   10320
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   12165
      _Version        =   65536
      _ExtentX        =   21458
      _ExtentY        =   952
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BevelInner      =   2
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin VB.Label lblEEProdLine 
         AutoSize        =   -1  'True
         Caption         =   "Product Line"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   7560
         TabIndex        =   103
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblEENumber 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   160
         Width           =   1005
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1320
         TabIndex        =   51
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2880
         TabIndex        =   50
         Top             =   135
         Width           =   720
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   600
      Top             =   120
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
   Begin VB.Frame fraDetail 
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   120
      TabIndex        =   66
      Top             =   3000
      Width           =   11535
      Begin VB.Frame frmOHSStatus 
         Height          =   735
         Left            =   9360
         TabIndex        =   118
         Top             =   6600
         Visible         =   0   'False
         Width           =   6135
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   6
            Left            =   1560
            TabIndex        =   18
            Tag             =   "01-OH&&S Incidnet Status"
            Top             =   0
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ECST"
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            Height          =   285
            Index           =   8
            Left            =   1560
            TabIndex        =   19
            Tag             =   "41-Status Effective Date"
            Top             =   315
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Status Effective Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   28
            Left            =   0
            TabIndex        =   120
            Top             =   360
            Width           =   1515
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   27
            Left            =   0
            TabIndex        =   119
            Top             =   45
            Width           =   450
         End
      End
      Begin VB.TextBox txtAMPMNotified 
         Appearance      =   0  'Flat
         DataField       =   "EC_TIMNOT_FORMAT"
         Height          =   285
         Left            =   6390
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "01-AM or PM or leave blank"
         Top             =   810
         Width           =   330
      End
      Begin VB.TextBox txtAMPMIncident 
         Appearance      =   0  'Flat
         DataField       =   "EC_OCCTM_FORMAT"
         Height          =   285
         Left            =   3050
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "01-AM or PM or leave blank"
         Top             =   815
         Width           =   330
      End
      Begin VB.TextBox txtJobStartDate 
         Appearance      =   0  'Flat
         DataField       =   "EC_JBSDATE"
         Height          =   285
         Left            =   5880
         TabIndex        =   117
         Tag             =   "Job Start Date"
         Top             =   4680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbJBStartDate 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Tag             =   "Position Start Date"
         Top             =   4658
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.CommandButton cmdPostion 
         Caption         =   "P&ositions"
         Height          =   270
         Left            =   270
         TabIndex        =   35
         Tag             =   "Postions"
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdPageRight 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   10680
         Picture         =   "fehsinc.frx":5E2C
         Style           =   1  'Graphical
         TabIndex        =   115
         Tag             =   "Grant All Basic"
         Top             =   10
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtClaimNo 
         Appearance      =   0  'Flat
         DataField       =   "EC_REOCCUR_CLAIM_NUM"
         Enabled         =   0   'False
         Height          =   285
         Left            =   7560
         TabIndex        =   114
         Top             =   5400
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtClaimDate 
         Appearance      =   0  'Flat
         DataField       =   "EC_REOCCUR_DATE"
         Enabled         =   0   'False
         Height          =   285
         Left            =   8040
         TabIndex        =   113
         Top             =   5400
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.ComboBox comDateClaimNo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "fehsinc.frx":626E
         Left            =   6915
         List            =   "fehsinc.frx":6270
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Tag             =   "10-Type of Employee "
         Top             =   5040
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.CheckBox chkReoccurence 
         Caption         =   "Reoccurrence"
         DataField       =   "EC_REOCCURENCE"
         Height          =   225
         Left            =   3960
         TabIndex        =   41
         Tag             =   "00-Reoccurence?"
         Top             =   5085
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.CheckBox chkGenerateWSIB7 
         Caption         =   "Generate Form 7 for this Incident"
         DataField       =   "EC_FORM7"
         Height          =   225
         Left            =   270
         TabIndex        =   40
         Tag             =   "00-Will this incident generate a Form 7?"
         Top             =   5085
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.TextBox txtReptAuthorityFName 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   3600
         TabIndex        =   109
         Tag             =   "00-Employee First Name of individual's supervisor"
         Top             =   2185
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtReptAuthorityFName 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   3600
         TabIndex        =   108
         Tag             =   "00-Employee First Name of individual's supervisor"
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtReptAuthoritySName 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   107
         Tag             =   "00-Employee Surname Name of individual's supervisor"
         Top             =   2185
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtReptAuthoritySName 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   106
         Tag             =   "00-Employee Surname Name of individual's supervisor"
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox comRelIncident 
         Height          =   315
         Left            =   2160
         TabIndex        =   22
         Tag             =   "01-Related Incident"
         Top             =   4200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtRelIncident 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         MaxLength       =   8
         TabIndex        =   101
         Tag             =   "11-Incident Number"
         Top             =   4215
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdDemo 
         Appearance      =   0  'Flat
         Caption         =   "Demographics"
         Height          =   330
         Left            =   600
         TabIndex        =   44
         Top             =   6960
         Width           =   1860
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         DataField       =   "EC_SHIFT"
         Height          =   285
         Left            =   2130
         MaxLength       =   1
         TabIndex        =   9
         Tag             =   "00-Shift incident occurred on"
         Top             =   1165
         Width           =   330
      End
      Begin VB.ComboBox cmbShift 
         Height          =   315
         Left            =   2130
         TabIndex        =   8
         Top             =   1150
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CheckBox chkOvertime 
         Alignment       =   1  'Right Justify
         Caption         =   "Overtime Work"
         DataField       =   "EC_OTWORK"
         Height          =   225
         Left            =   6990
         TabIndex        =   25
         Tag             =   "00-Overtime Work"
         Top             =   845
         Width           =   2240
      End
      Begin VB.Frame frmWorkType 
         Caption         =   "Work Type"
         Height          =   1335
         Left            =   6990
         TabIndex        =   68
         Top             =   1470
         Width           =   3075
         Begin VB.CheckBox chkWType 
            Alignment       =   1  'Right Justify
            Caption         =   "Temporary Transfer"
            DataField       =   "EC_WT_TMPTRN"
            Height          =   225
            Index           =   2
            Left            =   150
            TabIndex        =   29
            Tag             =   "00-Work Type - Temporary Transfer"
            Top             =   930
            Width           =   2565
         End
         Begin VB.CheckBox chkWType 
            Alignment       =   1  'Right Justify
            Caption         =   "Training"
            DataField       =   "EC_WT_TRAIN"
            Height          =   225
            Index           =   1
            Left            =   150
            TabIndex        =   28
            Tag             =   "00-Work Type - Training"
            Top             =   630
            Width           =   2565
         End
         Begin VB.CheckBox chkWType 
            Alignment       =   1  'Right Justify
            Caption         =   "Regular Schedule Work Shift"
            DataField       =   "EC_WT_REG"
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   27
            Tag             =   "00-Work Type - Regular Schedule Work Shift"
            Top             =   300
            Width           =   2565
         End
      End
      Begin VB.TextBox txtHRat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         DataField       =   "EC_HAZARD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9720
         TabIndex        =   67
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox cmbHRat 
         Height          =   315
         Left            =   9030
         TabIndex        =   24
         Tag             =   "00-Hazard Rating - Choose A,B, or C"
         Top             =   465
         Width           =   645
      End
      Begin VB.CheckBox chkModDuties 
         Alignment       =   1  'Right Justify
         Caption         =   "Modified Duties Required"
         DataField       =   "EC_MODDUTIES"
         Height          =   225
         Left            =   6990
         TabIndex        =   26
         Tag             =   "00-Overtime Work"
         Top             =   1195
         Width           =   2240
      End
      Begin VB.TextBox txtReptAuthority 
         Appearance      =   0  'Flat
         DataField       =   "EC_SUPERVISOR"
         Height          =   285
         Index           =   1
         Left            =   2250
         MaxLength       =   12
         TabIndex        =   15
         Tag             =   "00-Enter Employee Number"
         Top             =   2520
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.TextBox txtReptAuthority 
         Appearance      =   0  'Flat
         DataField       =   "EC_EMPNOT"
         Height          =   285
         Index           =   0
         Left            =   2250
         MaxLength       =   12
         TabIndex        =   13
         Tag             =   "01-Enter Employee Number"
         Top             =   2185
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.TextBox memComments 
         Appearance      =   0  'Flat
         DataField       =   "EC_COMMENTS_INC"
         Height          =   1125
         Left            =   540
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Tag             =   "60-Comments "
         Top             =   5760
         Width           =   8565
      End
      Begin VB.CheckBox chkFollowed 
         DataField       =   "EC_POLICY_FLAG"
         Height          =   195
         Left            =   9030
         TabIndex        =   23
         Tag             =   "00-Policy/Procedure Followed"
         Top             =   180
         Width           =   285
      End
      Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
         Height          =   285
         Index           =   1
         Left            =   1830
         TabIndex        =   14
         Tag             =   "01-Enter Employee Number"
         Top             =   2520
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
         Height          =   285
         Index           =   0
         Left            =   1830
         TabIndex        =   12
         Tag             =   "01-Enter Employee Number"
         Top             =   2185
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "EC_RETURN"
         Height          =   285
         Index           =   3
         Left            =   8850
         TabIndex        =   31
         Tag             =   "41-Date returned to work"
         Top             =   3195
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "EC_LDAY"
         Height          =   285
         Index           =   2
         Left            =   8850
         TabIndex        =   30
         Tag             =   "41-Last Day Worked"
         Top             =   2850
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "EC_DATENOT"
         Height          =   285
         Index           =   1
         Left            =   5175
         TabIndex        =   3
         Tag             =   "41-Date notified of incident"
         Top             =   480
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "EC_OCCDATE"
         Height          =   285
         Index           =   0
         Left            =   1830
         TabIndex        =   2
         Tag             =   "41-Date incident occurred"
         Top             =   480
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_PROVIDEDBY"
         Height          =   285
         Index           =   4
         Left            =   1830
         TabIndex        =   17
         Tag             =   "01-Provided By"
         Top             =   3190
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECPB"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_FAPROVIDED"
         Height          =   285
         Index           =   3
         Left            =   1830
         TabIndex        =   16
         Tag             =   "01-First aid provided"
         Top             =   2855
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECFF"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_CLASS"
         Height          =   285
         Index           =   2
         Left            =   1830
         TabIndex        =   11
         Tag             =   "01-Classification of Incident- Code"
         Top             =   1850
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECCL"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_TYPE"
         Height          =   285
         Index           =   1
         Left            =   1830
         TabIndex        =   10
         Tag             =   "01-Type of Incident- Code"
         Top             =   1515
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECTY"
      End
      Begin MSMask.MaskEdBox medIncidentTime 
         DataField       =   "EC_OCCTM"
         Height          =   285
         Left            =   2145
         TabIndex        =   4
         Tag             =   "10-Time Incident Occurred"
         Top             =   815
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medNotifyTime 
         DataField       =   "EC_TIMNOT"
         Height          =   285
         Left            =   5490
         TabIndex        =   6
         Tag             =   "10-Time of Notification"
         Top             =   810
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medShiftsLost 
         Height          =   285
         Left            =   2145
         TabIndex        =   20
         Tag             =   "21-Enter salary"
         Top             =   3525
         Visible         =   0   'False
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         Height          =   285
         Index           =   4
         Left            =   8850
         TabIndex        =   34
         Tag             =   "41-Date Approved"
         Top             =   4185
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "EC_RETURN_REG"
         Height          =   285
         Index           =   5
         Left            =   8850
         TabIndex        =   32
         Tag             =   "41-Date returned to regular work"
         Top             =   3525
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "EC_RETURN_SUITABLE"
         Height          =   285
         Index           =   6
         Left            =   8850
         TabIndex        =   33
         Tag             =   "41-Date returned to suitable work"
         Top             =   3855
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   5
         Left            =   1830
         TabIndex        =   21
         Tag             =   "00-Impact"
         Top             =   3860
         Visible         =   0   'False
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECIM"
      End
      Begin VB.Frame frmWFC 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   91
         Top             =   120
         Visible         =   0   'False
         Width           =   6735
         Begin VB.TextBox txtIncidentNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5385
            MaxLength       =   4
            TabIndex        =   1
            Tag             =   "01-Type of Incident- Code"
            Top             =   0
            Width           =   870
         End
         Begin VB.TextBox txtYear 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2030
            MaxLength       =   4
            TabIndex        =   0
            Tag             =   "01-Yeat of Incident"
            Top             =   0
            Width           =   930
         End
         Begin VB.Label lblTitle 
            Caption         =   "Incident#"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   11
            Left            =   3670
            TabIndex        =   93
            Tag             =   "01-Incident Number"
            Top             =   60
            Width           =   1125
         End
         Begin VB.Label lblTitle 
            Caption         =   "Year"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   10
            Left            =   165
            TabIndex        =   92
            Top             =   60
            Width           =   1485
         End
      End
      Begin INFOHR_Controls.CodeLookup clpJob 
         DataField       =   "EC_JBCODE"
         Height          =   285
         Left            =   1830
         TabIndex        =   36
         Tag             =   "01-Job Code"
         Top             =   4673
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "n/a"
         MaxLength       =   6
         LookupType      =   5
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "EC_JBSDATE"
         Height          =   285
         Index           =   7
         Left            =   6720
         TabIndex        =   39
         Tag             =   "41-Position Start Date"
         Top             =   4673
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   26
         Left            =   3480
         TabIndex        =   116
         Top             =   4718
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date && Claim #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   25
         Left            =   5640
         TabIndex        =   112
         Top             =   5100
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Approved"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   6990
         TabIndex        =   111
         Top             =   4230
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Return to Suitable Work"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   20
         Left            =   6990
         TabIndex        =   110
         Top             =   3900
         Width           =   1710
      End
      Begin VB.Label lblReptAuthority 
         Caption         =   "Reported to/by"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   9360
         TabIndex        =   104
         Top             =   6000
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label lblReptAuthority 
         Caption         =   "Supervisor"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   9360
         TabIndex        =   105
         Top             =   6240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Related Incident"
         Height          =   195
         Index           =   24
         Left            =   270
         TabIndex        =   102
         Top             =   4260
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Return Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   6990
         TabIndex        =   80
         Top             =   3240
         Width           =   870
      End
      Begin VB.Image imgEmail 
         Height          =   320
         Left            =   0
         Picture         =   "fehsinc.frx":6272
         Stretch         =   -1  'True
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblUpdateDate 
         Caption         =   "Updated Date"
         Height          =   255
         Left            =   6120
         TabIndex        =   100
         Top             =   7020
         Width           =   1095
      End
      Begin VB.Label lblUpdDateDesc 
         Height          =   255
         Left            =   7200
         TabIndex        =   99
         Top             =   6180
         Width           =   1935
      End
      Begin VB.Label lblUpdateBy 
         Caption         =   "Updated By"
         Height          =   255
         Left            =   2760
         TabIndex        =   98
         Top             =   7020
         Width           =   975
      End
      Begin VB.Label lblUserDesc 
         Height          =   255
         Left            =   3720
         TabIndex        =   97
         Top             =   6180
         Width           =   2415
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Impact"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   23
         Left            =   270
         TabIndex        =   96
         Top             =   3900
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Time of Incident"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   90
         Top             =   855
         Width           =   1395
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Shift Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   89
         Top             =   1215
         Width           =   900
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   270
         TabIndex        =   88
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Classification"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   270
         TabIndex        =   87
         Top             =   1890
         Width           =   915
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reported to/by"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   270
         TabIndex        =   86
         Top             =   2235
         Width           =   1305
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Time Notified"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3750
         TabIndex        =   85
         Top             =   855
         Width           =   1140
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hazard Rating"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   6990
         TabIndex        =   84
         Tag             =   "00-Hazard Rationg"
         Top             =   495
         Width           =   1620
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Supervisor"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   270
         TabIndex        =   83
         Top             =   2565
         Width           =   750
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Provided By"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   270
         TabIndex        =   82
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "First Aid Provided"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   270
         TabIndex        =   81
         Top             =   2895
         Width           =   1230
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Day Worked"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   6990
         TabIndex        =   79
         Top             =   2895
         Width           =   1245
      End
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   3645
         TabIndex        =   78
         Top             =   1830
         Width           =   4005
      End
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   3645
         TabIndex        =   77
         Top             =   1500
         Width           =   4005
      End
      Begin VB.Label lblIncidentNo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         DataField       =   "EC_CASE"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2130
         TabIndex        =   76
         Top             =   210
         Width           =   90
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Incident Number"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   270
         TabIndex        =   75
         Top             =   180
         Width           =   1410
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Notified"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3780
         TabIndex        =   74
         Top             =   525
         Width           =   1140
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Incident"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   73
         Top             =   525
         Width           =   1395
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Shifts Lost"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   17
         Left            =   270
         TabIndex        =   72
         Top             =   3570
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Return to Regular Work"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   19
         Left            =   6990
         TabIndex        =   71
         Top             =   3570
         Width           =   1695
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   21
         Left            =   270
         TabIndex        =   70
         Top             =   5460
         Width           =   735
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Policy/Procedure Followed"
         Height          =   195
         Index           =   22
         Left            =   6990
         TabIndex        =   69
         Top             =   180
         Width           =   2145
      End
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EC_EMPNBR"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1560
      TabIndex        =   53
      Top             =   10440
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EC_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   30
      TabIndex        =   54
      Top             =   10440
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEHSINCIDENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean
Dim glbNew
Dim fsnapEENames As New ADODB.Recordset
Dim fglbNewCode
Dim savAuth
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control

Dim rsDATA3 As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim rsTWFC As New ADODB.Recordset
Dim xPlantCode, SQLC
Dim fglbJobList As String
Dim xJobSelected As String

Private Sub WFCPlantCode()
    If Not glbtermopen Then
        SQLC = "SELECT ED_SECTION FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
    Else
        SQLC = "SELECT ED_SECTION FROM Term_HREMP WHERE TERM_SEQ = " & glbTERM_Seq
    End If
    rsTWFC.Open SQLC, gdbAdoIhr001, adOpenStatic
    xPlantCode = ""
    If Not rsTWFC.EOF Then
        If Not IsNull(rsTWFC("ED_SECTION")) Then
            xPlantCode = rsTWFC("ED_SECTION")
        End If
    End If
    rsTWFC.Close
End Sub

Private Sub WFCStatusScreen() 'Ticket #27576 Franks 10/26/2015
    frmOHSStatus.BorderStyle = 0
    frmOHSStatus.Left = 270
    frmOHSStatus.Top = 3510
    frmOHSStatus.Visible = True
    clpCode(6).DataField = "EC_STATUS"
    dlpDate(8).DataField = "EC_STDATE"
End Sub

Private Function chkHSIncident()

Dim SQLQ As String, Msg As String, dd#
Dim RsEHST As New ADODB.Recordset
Dim DupFlag As Boolean, xMsg, xTempInc, xIncNum

chkHSIncident = False

On Error GoTo chkHSIncident_Err

If glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2172W" Then
'Ticket #16096 D. Muskoka
'Ticket #17112 County of Lanark
    If Len(txtYear) > 0 Then
        If (Not IsNumeric(txtYear)) Or Val(txtYear) < 1900 Or Val(txtYear) >= 2080 Then
            MsgBox "Incident Year is not a valid value."
            dlpDate(0).SetFocus
            Exit Function
        End If
    Else
        MsgBox "Incident Year is required."
        dlpDate(0).SetFocus
        Exit Function
    End If
    If Len(txtIncidentNo) >= 1 Then
        If (Not IsNumeric(txtIncidentNo)) Or Val(txtIncidentNo) <= 0 Or Val(txtIncidentNo) > 9999 Then
            MsgBox "Incident Number is not a valid value."
            dlpDate(0).SetFocus
            Exit Function
        End If
    Else
        MsgBox "Incident Number is required."
        dlpDate(0).SetFocus
        Exit Function
    End If

    xIncNum = Val(txtYear & Format(txtIncidentNo, "0000"))
    SQLQ = "SELECT * FROM HR_OCC_HEALTH_SAFETY WHERE EC_CASE = " & xIncNum
    DupFlag = False
    If Not fglbNew Then
        If Not IsNull(Data1.Recordset("EC_CASE")) Then
            SQLQ = SQLQ & " AND EC_CASE <> " & Data1.Recordset("EC_CASE")
            DupFlag = True
        End If
    End If

    RsEHST.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not RsEHST.EOF Then
        If DupFlag Then
            xTempInc = Val(Right(Str(Data1.Recordset("EC_CASE")), 4))
            xMsg = "Duplicate Incident Number - " & txtIncidentNo & Chr(10)
            xMsg = xMsg & "Keep the previous Incident Number - " & Str(xTempInc)
            MsgBox xMsg
            txtIncidentNo = xTempInc
            Exit Function
        Else
            xTempInc = Val(Right(Str(fglbNewCode), 4))
            xMsg = "Duplicate Incident Number - " & txtIncidentNo & Chr(10)
            If Not glbWFC Then
            xMsg = xMsg & "Next Available Incident Number is - " & Str(xTempInc)
            End If
            MsgBox xMsg
            txtIncidentNo = xTempInc
            Exit Function
        End If
    End If
    RsEHST.Close
End If

If glbWFC Then
    If Len(txtYear) > 0 Then
        If (Not IsNumeric(txtYear)) Or Val(txtYear) < 1900 Or Val(txtYear) >= 2080 Then
            MsgBox "Incident Year is not a valid value."
            dlpDate(0).SetFocus
            Exit Function
        End If
    Else
        MsgBox "Incident Year is required."
        dlpDate(0).SetFocus
        Exit Function
    End If
    If Len(txtIncidentNo) >= 1 Then
        If (Not IsNumeric(txtIncidentNo)) Or Val(txtIncidentNo) <= 0 Or Val(txtIncidentNo) > 9999 Then
            MsgBox "Incident Number is not a valid value."
            dlpDate(0).SetFocus
            Exit Function
        End If
    Else
        MsgBox "Incident Number is required."
        dlpDate(0).SetFocus
        Exit Function
    End If
    
    'Franks Jan 25,2002
    'To fix the following problem for WFC:
    'When click "Edit" or "New", if user enter a duplicate incident number, the system will crash
    'If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then

        xIncNum = Val(txtYear & Format(txtIncidentNo, "0000"))
        SQLQ = "SELECT * FROM HR_OCC_HEALTH_SAFETY WHERE EC_CASE = " & xIncNum
        DupFlag = False
        If Not fglbNew Then
            If Not IsNull(Data1.Recordset("EC_CASE")) Then
                SQLQ = SQLQ & " AND EC_CASE <> " & Data1.Recordset("EC_CASE")
                DupFlag = True
            End If
        End If
        'If glbWFC And Not glbtermopen Then
            If Len(xPlantCode) > 0 Then
                If Not glbtermopen Then
                    SQLQ = SQLQ & " AND EC_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & xPlantCode & "')"
                Else
                    SQLQ = SQLQ & " AND EC_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_SECTION = '" & xPlantCode & "')"
                End If
            End If
        'End If
        RsEHST.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not RsEHST.EOF Then
            If DupFlag Then
                xTempInc = Val(Right(Str(Data1.Recordset("EC_CASE")), 4))
                xMsg = "Duplicate Incident Number - " & txtIncidentNo & Chr(10)
                xMsg = xMsg & "Keep the previous Incident Number - " & Str(xTempInc)
                MsgBox xMsg
                txtIncidentNo = xTempInc
                Exit Function
            Else
                xTempInc = Val(Right(Str(fglbNewCode), 4))
                xMsg = "Duplicate Incident Number - " & txtIncidentNo & Chr(10)
                If Not glbWFC Then
                xMsg = xMsg & "Next Available Incident Number is - " & Str(xTempInc)
                End If
                MsgBox xMsg
                txtIncidentNo = xTempInc
            End If
        End If
        RsEHST.Close
    'End If
    'Franks Jan 25,2002
    
    'Ticket #15396 - Begin
    If Len(medIncidentTime.Text) = 0 Then
        MsgBox "Time of Incident is required."
        medIncidentTime.SetFocus
        Exit Function
    End If
    If Len(cmbShift.Text) = 0 Then
        MsgBox "Shift Code is required."
        cmbShift.SetFocus
        Exit Function
    End If
    If Len(clpCode(2).Text) = 0 Then
        MsgBox "Classification is required."
        clpCode(2).SetFocus
        Exit Function
    End If
    If Len(cmbHRat.Text) = 0 Then
        MsgBox "Hazard Rating is required."
        cmbHRat.SetFocus
        Exit Function
    End If
    'Ticket #15396 - End
End If

If glbWFC Then 'Ticket #27576 Franks 10/26/2015
    If clpCode(6).Caption = "Unassigned" Then
        MsgBox "Status code must be valid"
        If clpCode(6).Enabled Then clpCode(6).SetFocus
        Exit Function
    End If
    If Len(clpCode(6).Text) > 0 Then
        If Not IsDate(dlpDate(8).Text) Then
            MsgBox "Status Effective Date cannot be blank if Status is entered"
            If dlpDate(8).Enabled Then dlpDate(8).SetFocus
            Exit Function
        End If
    End If
    If IsDate(dlpDate(8).Text) Then
        If Len(clpCode(6).Text) = 0 Then
            MsgBox "Status Effective Date must be blank if Status is not entered"
            If clpCode(6).Enabled Then clpCode(6).SetFocus
            Exit Function
        End If
    End If
End If

If Len(dlpDate(0).Text) >= 1 Then
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "Incident Date is not a valid date."
        dlpDate(0).SetFocus
        Exit Function
    End If
Else
    MsgBox "Incident Date is required."
    dlpDate(0).SetFocus
    Exit Function
End If

Dim tTime As Variant
Dim Part1$, Part2$

'~~

If Len(dlpDate(1).Text) >= 1 Then
    If Not IsDate(dlpDate(1).Text) Then
        MsgBox "Notification Date is not a valid date."
        dlpDate(1).SetFocus
        Exit Function
    End If
Else
    MsgBox "Notification Date is required."
    dlpDate(1).SetFocus
    Exit Function
End If

dd# = DateDiff("d", CVDate(dlpDate(0).Text), CVDate(dlpDate(1).Text))

If dd# < 0 Then
    MsgBox "Notification date must be later than Incident."
    dlpDate(1).SetFocus
    Exit Function
End If

'~~

'Ticket #22682 - Release 8.0: Allow user to enter Time format - mainly used in Form 7
If Len(medIncidentTime) = 5 Then
    If txtAMPMIncident.Text <> "AM" And txtAMPMIncident.Text <> "PM" And txtAMPMIncident.Text <> "" Then
        MsgBox "'Time of Incident' format can only be 'AM' or 'PM' or blank."
        txtAMPMIncident.SetFocus
        Exit Function
    End If
End If

If Len(medIncidentTime) = 5 And txtAMPMIncident.Text <> "" Then
    Part1$ = Left$(medIncidentTime, 2)
    Part2$ = Right$(medIncidentTime, 2)
    If Not Left$(Part1$, 2) = "__" And Not Right$(Part2$, 2) = "__" Then
        If Not IsNumeric(Part1$) Or Not IsNumeric(Part2$) Then
            MsgBox "Not a valid time"
            medIncidentTime.SetFocus
            Exit Function
        End If
        If CInt(Part1$) > 24 Or CInt(Part2$) > 59 Then
            MsgBox "Not a valid time"
            medIncidentTime.SetFocus
            Exit Function
        End If
    End If
Else
    If Len(medIncidentTime) > 1 And Len(medIncidentTime) < 5 Then
            MsgBox "Not a valid time"
            medIncidentTime.SetFocus
            Exit Function
    End If
   'MsgBox "Invalid time."
   'medIncidentTime.SetFocus
   'Exit Function
End If

'Ticket #22682 - Release 8.0: Allow user to enter Time format - mainly used in Form 7
If Len(medNotifyTime) = 5 Then
    If txtAMPMNotified.Text <> "AM" And txtAMPMNotified.Text <> "PM" And txtAMPMNotified.Text <> "" Then
        MsgBox "'Time Notified' format can only be 'AM' or 'PM' or blank."
        txtAMPMNotified.SetFocus
        Exit Function
    End If
End If

If Len(medNotifyTime) = 5 And txtAMPMNotified.Text <> "" Then
    Part1$ = Left$(medNotifyTime, 2)
    Part2$ = Right$(medNotifyTime, 2)
    If Not Left$(Part1$, 2) = "__" And Not Right$(Part2$, 2) = "__" Then
        If Not IsNumeric(Part1$) Or Not IsNumeric(Part2$) Then
            MsgBox "Not a valid time"
            medNotifyTime.SetFocus
            Exit Function
        End If
        If CInt(Part1$) > 24 Or CInt(Part2$) > 59 Then
            MsgBox "Not a valid time"
            medNotifyTime.SetFocus
            Exit Function
        End If
    End If
Else
    If Len(medNotifyTime) > 1 And Len(medNotifyTime) < 5 Then
            MsgBox "Not a valid time"
            medNotifyTime.SetFocus
            Exit Function
    End If
   'MsgBox "Invalid time."
   'medNotifyTime.SetFocus
   'Exit Function
End If

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Incident Type code is a required field"
    clpCode(1).SetFocus
    Exit Function
End If

If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Incident Type code must be valid"
    clpCode(1).SetFocus
    Exit Function
End If

'If Len(clpCode(2).Text) < 1 Then   'As per Next Release Documentation
'    MsgBox "Classification code is a required field"
'    clpCode(2).SetFocus
'    Exit Function
'End If

If glbLinamar Then
    If Len(Trim(clpCode(2))) = 0 Then
        MsgBox lblTitle(6).Caption & " is a required field"
        clpCode(2).SetFocus
        Exit Function
    End If
End If

If clpCode(2).Caption = "Unassigned" Then
    If glbLinamar Then
        MsgBox lblTitle(6).Caption & " must be valid"
    Else
        MsgBox "Classification code must be valid"
    End If
    clpCode(2).SetFocus
    Exit Function
End If

If elpReptAuthShow(0).Text = "0" Then elpReptAuthShow(0).Text = ""
If elpReptAuthShow(1).Text = "0" Then elpReptAuthShow(1).Text = ""

'Ticket #28283 - Jerry said to make it non mandatory for all
'Release 8.1 - City of Sarnia do not want 'Reported to/by' to be mandatory
'If glbCompSerial <> "S/N - 2362W" Then
'    If Len(elpReptAuthShow(0).Text) < 1 Then
'        MsgBox lblTitle(7).Caption & " is a required field"
'        elpReptAuthShow(0).SetFocus
'        Exit Function
'    End If
'End If

If glbLinamar Then
    If Len(elpReptAuthShow(1).Text) < 1 Then
        MsgBox "Supervisor is a required field"
        elpReptAuthShow(1).SetFocus
        Exit Function
    End If
End If

If glbLinamar And Not glbtermopen Then
    If lblReptAuthority(0).Caption = "Unassigned" Then
        MsgBox "Reported by is not a valid entry"
        elpReptAuthShow(0).SetFocus
        Exit Function
    End If
    If lblReptAuthority(1).Caption = "Unassigned" And Len(elpReptAuthShow(1).Text) > 0 Then
        MsgBox "Supervisor is not a valid entry"
        elpReptAuthShow(1).SetFocus
        Exit Function
    End If
Else
    If elpReptAuthShow(0).Caption = "Unassigned" Then
        MsgBox lblTitle(7).Caption & " is not a valid entry"
        elpReptAuthShow(0).SetFocus
        Exit Function
    End If
    If elpReptAuthShow(1).Caption = "Unassigned" And Len(elpReptAuthShow(1).Text) > 0 Then
        MsgBox "Supervisor is not a valid entry"
        elpReptAuthShow(1).SetFocus
        Exit Function
    End If
End If


If glbLinamar Then
    If Len(Trim(clpCode(3).Text)) = 0 Then    'First Aid Provided By
        MsgBox lblTitle(14).Caption & " is a required field"
        clpCode(3).SetFocus
        Exit Function
    End If
    If Not clpCode(3).ListChecker Then Exit Function
    
    If Len(Trim(clpCode(4).Text)) = 0 Then    'Provided By
        MsgBox lblTitle(13).Caption & " is a required field"
        clpCode(4).SetFocus
        Exit Function
    End If
    If Not clpCode(4).ListChecker Then Exit Function
Else
    If Not clpCode(3).ListChecker Then Exit Function
End If

'If glbLinamar Then
'    If Len(clpCode(2)) = 0 Then
'        MsgBox "Classification is a required field"
'        clpCode(2).SetFocus
'        Exit Function
'    End If
'End If

If glbLinamar Then  'Ticket #14703
    If Len(Trim(comRelIncident.Text)) > 0 Then
        If Not IsNumeric(comRelIncident.Text) Then
            MsgBox "Invalid Related Incident"
            comRelIncident.SetFocus
            Exit Function
        Else
            If Not IfIncidentNo(Val(comRelIncident.Text)) Then
                MsgBox "Related Incident not a valid number"
                comRelIncident.SetFocus
                Exit Function
            End If
        End If
    End If
End If

If gSec_Inq_HSW7CmpMst And gSec_Inq_HSW7Injury And glbWSIBModule Then
    If Not clpJob.ListChecker Then Exit Function
    If Len(clpJob.Text) > 0 Then
        If dlpDate(7).Visible = True Then
            If Len(Trim(dlpDate(7).Text)) = 0 Then
                MsgBox "Job Start Date cannot be blank"
                dlpDate(7).SetFocus
                Exit Function
            ElseIf Not IsDate(dlpDate(7).Text) Then
                MsgBox "Invalid Job Start Date"
                dlpDate(7).SetFocus
                Exit Function
            End If
        ElseIf cmbJBStartDate.Text = "" Then
            MsgBox "Job Start Date cannot be blank"
            cmbJBStartDate.SetFocus
            Exit Function
        End If
    End If
End If

chkHSIncident = True

Exit Function

chkHSIncident_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HR_OCC_HEALTH_SAFETY", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub chkGenerateWSIB7_Click()
    If isWSIBModule Then
        If chkGenerateWSIB7.Value = 1 Then
            cmdPageRight(0).Visible = True
        Else
            cmdPageRight(0).Visible = False
        End If
    Else
        cmdPageRight(0).Visible = False
    End If
End Sub

Private Sub chkModDuties_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkOvertime_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkReoccurence_Click()
Dim X As Integer
If chkReoccurence.Value = 1 Then
    lblTitle(25).Visible = True
    comDateClaimNo.Visible = True

    If txtClaimDate.Text = "" And txtClaimNo.Text = "" Then
        comDateClaimNo.Clear
    End If

    'Retrieve all Claims with Dates for this employee
    Call Populate_Employee_ClaimDate

    If txtClaimDate.Text = "" And txtClaimNo.Text = "" And Not Data1.Recordset.EOF Then
        If Not IsNull(Data1.Recordset("EC_REOCCUR_CLAIM_NUM")) And Not IsNull(Data1.Recordset("EC_REOCCUR_DATE")) Then
            txtClaimNo.Text = Data1.Recordset("EC_REOCCUR_CLAIM_NUM")
            txtClaimDate.Text = Data1.Recordset("EC_REOCCUR_DATE")
        End If
    End If

    For X = 0 To comDateClaimNo.ListCount - 1
        If comDateClaimNo.List(X) = txtClaimNo.Text & " - " & txtClaimDate.Text Then
            comDateClaimNo.ListIndex = X
            Exit For
        Else
            'If txtClaimNo.Text = "" Then
            '    comDateClaimNo.ListIndex = -1
            'End If
        End If
    Next
Else
    lblTitle(25).Visible = False
    comDateClaimNo.Visible = False
    comDateClaimNo.Clear
    txtClaimDate.Text = ""
    txtClaimNo.Text = ""
End If
End Sub

Private Sub chkWType_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpJob_Change()
If clpJob.Text <> "" And clpJob.Visible Then
    txtDemo(14).Text = GetJobDesc(clpJob.Text)
End If
End Sub

Private Sub clpJob_GotFocus()
txtDemo(14).Text = GetJobDesc(clpJob.Text)
End Sub

Private Sub clpJob_LostFocus()
    If xJobSelected <> clpJob Then
        'Populate Position Start Date combo box
        Call Populate_Job_Start_Date
        
        'Current job selected
        xJobSelected = clpJob
    End If
End Sub

Private Sub cmbHRat_Change()
'Ticket # 6831 - For Burlington Tech.
If glbCompSerial = "S/N - 2351W" Or glbLinamar Then
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
ElseIf glbLinamar Then
    Select Case cmbHRat
        Case "Gradual": txtHRat = "G"
        Case "Sudden": txtHRat = "S"
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

Private Sub cmbJBStartDate_Change()
    txtJobStartDate.Text = cmbJBStartDate.Text
    dlpDate(7).Text = cmbJBStartDate.Text
End Sub

Private Sub cmbJBStartDate_Click()
    txtJobStartDate.Text = cmbJBStartDate.Text
    dlpDate(7).Text = cmbJBStartDate.Text
End Sub

Private Sub cmbJBStartDate_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbJBStartDate_LostFocus()
    txtJobStartDate.Text = cmbJBStartDate.Text
    dlpDate(7).Text = cmbJBStartDate.Text
End Sub

Private Sub cmbShift_Click()
If Not glbWFC Then Exit Sub
txtShift = Left(cmbShift, 1)
End Sub

Private Sub cmbShift_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

'Private Sub cmdCAction_Click()
'frmEHSCorrective.Show
'Unload Me
'End Sub

'Private Sub cmdCAction_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdCancel_Click()
Dim X As Integer
On Error GoTo Can_Err

'Data1.Recordset.CancelUpdate '.CancelBatch
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh

''' Sam add July 2002 * Remove Binding Control

Call Display_Value

For X = 0 To 1
    Call txtReptAuthority_Change(X)
Next
'Call ST_UPD_MODE(True)
fglbNew = False
Call SET_UP_MODE
Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_OCC_HEALTH_SAFETY", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If



End Sub

'Private Sub cmdCancel_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEHSINCIDENT" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdContact_Click()
'frmEHSContact.Show
'Unload Me
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, INo&, X


If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If


On Error GoTo Del_Err


Msg = "Are You Sure You Want To Delete This Record?  "
Msg = Msg & Chr(10) & Chr(10) & "All the related Incident and Cost records will also be deleted!"
INo& = CLng(lblIncidentNo.Caption)

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub


gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

If Not glbtermopen Then Call NukeHSCosts(INo&)

'Ticket #22682 - Delete the rest of the incident related records
If Not glbtermopen Then Call NukeHSRootCauses(INo&, glbLEE_ID)
If Not glbtermopen Then Call NukeHSContacts(INo&, glbLEE_ID)
If Not glbtermopen Then Call NukeHSCorrective(INo&, glbLEE_ID)
If Not glbtermopen Then Call NukeHSForm9(INo&, glbLEE_ID)
If Not glbtermopen Then Call NukeHSAttachment(INo&, glbLEE_ID)


fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_OCC_HEALTH_SAFETY", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

'Private Sub cmdDelete_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdInjLoc_Click()
'frmEHSINJURY.Show
'Unload Me
'End Sub

'Private Sub cmdInjLoc_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()

'If Not gSec_Upd_Health_Safety Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If

On Error GoTo Mod_Err

'If Not Fnd_Match_Data1() Then MsgBox "No Records Found"

'Call ST_UPD_MODE(True)
'Call SET_UP_MODE
'If glbWFC Then
'    txtYear.SetFocus
'Else
'    dlpDate(0).SetFocus
'End If

xJobSelected = clpJob

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_OCC_HEALTH_SAFETY", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

'Private Sub cmdModify_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
Dim SQLQ As String
Dim TX As New ADODB.Recordset
Dim rsDemo As New ADODB.Recordset
Call ST_UPD_MODE(True)

fglbNew = True
Call SET_UP_MODE
On Error GoTo AddN_Err


Data3.Refresh
fglbNewCode = 1
'If glbWFC Then
If glbWFC Or glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2172W" Then
'Ticket #16096 D. Muskoka
'Ticket #17112 County of Lanark
    If glbWFC Then
        If Not glbtermopen Then
            SQLQ = "SELECT DISTINCT EC_CASE FROM HR_OCC_HEALTH_SAFETY WHERE "
            SQLQ = SQLQ & "EC_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & xPlantCode & "') "
            SQLQ = SQLQ & "ORDER BY EC_CASE DESC "
        Else
            SQLQ = "SELECT DISTINCT EC_CASE FROM Term_HR_OCC_HEALTH_SAFETY WHERE "
            SQLQ = SQLQ & "EC_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_SECTION = '" & xPlantCode & "') "
            SQLQ = SQLQ & "ORDER BY EC_CASE DESC "
        End If
        rsTWFC.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTWFC.EOF Then
            If Not IsNull(rsTWFC("EC_CASE")) Then
                fglbNewCode = rsTWFC("EC_CASE") + 1
            End If
        End If
        rsTWFC.Close
    End If
    If glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2172W" Then
    'Ticket #16096 D. Muskoka
    'Ticket #17112 County of Lanark
        If Not glbtermopen Then
            SQLQ = "SELECT DISTINCT EC_CASE FROM HR_OCC_HEALTH_SAFETY WHERE (1=1) "
            'SQLQ = SQLQ & "EC_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & xPlantCode & "') "
            SQLQ = SQLQ & "ORDER BY EC_CASE DESC "
        Else
            SQLQ = "SELECT DISTINCT EC_CASE FROM Term_HR_OCC_HEALTH_SAFETY WHERE (1=1) "
            'SQLQ = SQLQ & "EC_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_SECTION = '" & xPlantCode & "') "
            SQLQ = SQLQ & "ORDER BY EC_CASE DESC "
        End If
        rsTWFC.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTWFC.EOF Then
            If Not IsNull(rsTWFC("EC_CASE")) Then
                fglbNewCode = rsTWFC("EC_CASE") + 1
            End If
        End If
        rsTWFC.Close
    End If
Else
    If Not (Data3.Recordset.EOF And Data3.Recordset.BOF) Then
        Data3.Recordset.MoveFirst
        If Not IsNull(Data3.Recordset("OHSNBR")) Then fglbNewCode = Data3.Recordset("OHSNBR")
    End If
    If fglbNewCode = 0 Then fglbNewCode = 1
    TX.Open "SELECT EC_CASE FROM HR_OCC_HEALTH_SAFETY", gdbAdoIhr001, adOpenDynamic, adLockReadOnly
    If Not TX.EOF Then
        Do While True
            TX.MoveFirst
            TX.Find "EC_CASE =" & fglbNewCode
            If TX.EOF Then Exit Do Else fglbNewCode = fglbNewCode + 1
        Loop
    End If
    TX.Close
    If glbtermopen Then
        TX.Open "SELECT EC_CASE FROM Term_HR_OCC_HEALTH_SAFETY ", gdbAdoIhr001X, adOpenDynamic, adLockReadOnly
        If Not TX.EOF Then
            Do While True
                TX.MoveFirst
                TX.Find "EC_CASE =" & fglbNewCode
                If TX.EOF Then Exit Do Else fglbNewCode = fglbNewCode + 1
            Loop
        End If
    End If
End If
Call Set_Control("B", Me)
'rsDATA.AddNew

chkModDuties = 0
chkOvertime = 0
chkWType(0) = 0
chkWType(1) = 0
chkWType(2) = 0
lblIncidentNo.Caption = fglbNewCode
If Not glbWFC And Not glbtermopen Then
    Data3.Refresh
    If Data3.Recordset.EOF Then
        Data3.Recordset.AddNew
        Data3.Recordset("OHSNBR") = Val(lblIncidentNo) + 1
        Data3.Recordset.Update
    End If
    If IsNull(Data3.Recordset("OHSNBR")) Then
        Data3.Recordset("OHSNBR") = Val(lblIncidentNo) + 1
        Data3.Recordset.Update
    End If
    If Val(lblIncidentNo) >= Data3.Recordset("OHSNBR") Then
        If glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2172W" Then
        'Ticket #16096 D. Muskoka
        'Ticket #17112 County of Lanark
            Data3.Recordset("OHSNBR") = Val(lblIncidentNo)
        Else
            Data3.Recordset("OHSNBR") = Val(lblIncidentNo) + 1
        End If
        Data3.Recordset.Update
    End If
Else
    If glbtermopen Then
        If Not Data1.Recordset.EOF Then
            Data1.Recordset("TERM_SEQ") = glbTERM_Seq
        End If
    End If
End If

If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"

If glbLinHS Then
    txtDemo(3) = glbDiv
Else
    If glbtermopen Then
        rsDemo.Open "SELECT * FROM Term_HREMP where TERM_SEQ = " & lblEEID, gdbAdoIhr001X, adOpenDynamic, adLockReadOnly
    Else
        rsDemo.Open "SELECT * FROM HREMP where ED_EMPNBR = " & lblEEID, gdbAdoIhr001, adOpenDynamic, adLockReadOnly
    End If
    If Not (rsDemo.BOF And rsDemo.EOF) Then
        If Not IsNull(rsDemo("ED_DEPTNO")) Then txtDemo(2) = rsDemo("ED_DEPTNO")
        If Not IsNull(rsDemo("ED_DIV")) Then txtDemo(3) = rsDemo("ED_DIV")
        If Not IsNull(rsDemo("ED_LOC")) Then txtDemo(4) = rsDemo("ED_LOC")
        If Not IsNull(rsDemo("ED_ORG")) Then txtDemo(5) = rsDemo("ED_ORG")
        If Not IsNull(rsDemo("ED_EMP")) Then txtDemo(6) = rsDemo("ED_EMP")
        If Not IsNull(rsDemo("ED_PT")) Then txtDemo(7) = rsDemo("ED_PT")
        If Not IsNull(rsDemo("ED_REGION")) Then txtDemo(8) = rsDemo("ED_REGION")
        If Not IsNull(rsDemo("ED_SECTION")) Then txtDemo(9) = rsDemo("ED_SECTION")
        If Not IsNull(rsDemo("ED_ADMINBY")) Then txtDemo(10) = rsDemo("ED_ADMINBY")
        If Not IsNull(rsDemo("ED_WORKCOUNTRY")) Then txtDemo(11) = rsDemo("ED_WORKCOUNTRY")
        If glbLinamar Then
            If Not IsNull(rsDemo("ED_HOMEOPRTNBR")) Then txtDemo(12) = Mid(rsDemo("ED_HOMEOPRTNBR"), 4) 'rsDemo("ED_HOMEOPRTNBR")
            If Not IsNull(rsDemo("ED_HOMELINE")) Then txtDemo(13) = Mid(rsDemo("ED_HOMELINE"), 4) 'rsDemo("ED_HOMELINE")
        End If
    End If
    If Not glbLinHS Then
        If Not Set_Cur_Position() Then Exit Sub
    End If
    elpReptAuthShow(1).Text = ShowEmpnbr(savAuth)
End If

If glbWFC Then
    txtYear.SetFocus
Else
    dlpDate(0).SetFocus
End If


Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_OCC_HEALTH_SAFETY", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub CmdNew_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Function cmdOK_Click()
Dim X, xReptAuthority, xFld
Dim xLocNewRec As Boolean

On Error GoTo Add_Err

cmdOK_Click = False

If Not chkHSIncident() Then Exit Function

rsDATA.Requery

If fglbNew Then rsDATA.AddNew
xLocNewRec = fglbNew

'If glbWFC Then 'And Not glbtermopen Then
If glbWFC Or glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2172W" Then
'Ticket #16096 D. Muskoka
'Ticket #17112 County of Lanark
    lblIncidentNo = txtYear & Format(txtIncidentNo, "0000")
    Data3.Refresh
    If Not glbtermopen Then
        If Data3.Recordset.EOF Then
            Data3.Recordset.AddNew
            Data3.Recordset("OHSNBR") = Val(lblIncidentNo) + 1
            Data3.Recordset.Update
        Else
            If Left(Format(Data3.Recordset("OHSNBR"), "00000000"), 4) = "0000" Then
                Data3.Recordset("OHSNBR") = txtYear & Right(Format(Data3.Recordset("OHSNBR"), "00000000"), 4)
                Data3.Recordset.Update
            End If
        End If
    End If
End If
If glbWFC And Not glbtermopen Then
    Data3.Refresh
    If Data3.Recordset.EOF Then
        Data3.Recordset.AddNew
        Data3.Recordset("OHSNBR") = Val(lblIncidentNo) + 1
        Data3.Recordset.Update
    End If
    If IsNull(Data3.Recordset("OHSNBR")) Then
        Data3.Recordset("OHSNBR") = Val(lblIncidentNo) + 1
        Data3.Recordset.Update
    End If
    If Val(lblIncidentNo) >= Data3.Recordset("OHSNBR") Then
        Data3.Recordset("OHSNBR") = Val(lblIncidentNo) + 1
        Data3.Recordset.Update
    End If
End If


Call UpdUStats(Me)
Call Set_Control("U", Me, rsDATA)

For X = 0 To 1
    xReptAuthority = getEmpnbr(elpReptAuthShow(X).Text)
    xFld = "EC_" & IIf(X = 0, "EMPNOT", "SUPERVISOR")
    If Val(xReptAuthority) = 0 Then
        rsDATA(xFld) = Null
    Else
        rsDATA(xFld) = xReptAuthority
    End If
Next

'WSIB Form 7 - set the Result to PEND as default value as per the requirement documentation.
If gSec_Inq_HSW7CmpMst And gSec_Inq_HSW7Injury And glbWSIBModule Then
    If fglbNew And chkGenerateWSIB7.Value = 1 Then
        rsDATA("EC_WCBRES") = "PEND"
    End If
    
    'Ticket #21550
    If cmbJBStartDate.Visible = True Then
        txtJobStartDate.Text = cmbJBStartDate.Text
    End If
End If

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
End If
Data1.Refresh


fglbNew = False
cmdOK_Click = True

'Call ST_UPD_MODE(True)
Call SET_UP_MODE

Me.vbxTrueGrid.SetFocus

If gsEMAIL_ONHSINCIDENT Then 'Ticket #28664 Franks 05/31/2016 (WFC needs it)
    If glbWFC Then
        If xLocNewRec Then 'new record only
            Call HS_INCIDENT_Email
        End If
    Else
        'Ticket #28815 - Opened for all so copied the WFC routine and making this general
        If xLocNewRec Then 'new record only
            Call All_HS_INCIDENT_Email
        End If
    End If
End If

If NextFormIF("Incident") Then
    Call cmdNew_Click
End If

Exit Function

Add_Err:
If Err = 3022 Then
    Data1.Recordset.CancelUpdate
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Function
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OCC_HEALTH_SAFETY", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Sub HS_INCIDENT_Email() 'Ticket #28664 Franks 05/31/2016
Dim rsCodeMatrix As New ADODB.Recordset
Dim strFields As String
Dim SQLQ As String
Dim xUptFlag As Boolean
Dim xDate1, xDate2
Dim xLDate
Dim a As Integer, Msg As String
Dim Title$, DgDef, Response%
Dim xYear, xMonth
Dim xEmail
Dim xToEmail As String
Dim IsSendEmail As Boolean
Dim MailBody

        'check if Classification code is setup for Sending Email
        IsSendEmail = False
        If Len(clpCode(2).Text) = 0 Then
            Exit Sub
        End If
        'check Code Matrix on ECTY ----------- begin
        If Len(xPlantCode) > 0 Then
            SQLQ = "SELECT * FROM CODEMATRIX WHERE CM_NAME = 'ECTY' AND CM_KEY = '" & clpCode(1).Text & "' "
            SQLQ = SQLQ & "AND CM_SECTION = '" & xPlantCode & "' "
            If rsCodeMatrix.State <> 0 Then rsCodeMatrix.Close
            rsCodeMatrix.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsCodeMatrix.EOF Then
                If Not IsNull(rsCodeMatrix("CM_PRKEY")) Then
                    If UCase(Trim(rsCodeMatrix("CM_PRKEY"))) = "Y" Then
                        IsSendEmail = True
                    End If
                End If
            End If
        End If
        If Not IsSendEmail Then 'Not found in this plant then check this code wihtout plant code
            SQLQ = "SELECT * FROM CODEMATRIX WHERE CM_NAME = 'ECTY' AND CM_KEY = '" & clpCode(1).Text & "' "
            'SQLQ = SQLQ & "AND CM_SECTION = '" & xPlantCode & "' "
            If rsCodeMatrix.State <> 0 Then rsCodeMatrix.Close
            rsCodeMatrix.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsCodeMatrix.EOF Then
                If Not IsNull(rsCodeMatrix("CM_PRKEY")) Then
                    If UCase(Trim(rsCodeMatrix("CM_PRKEY"))) = "Y" Then
                        IsSendEmail = True
                    End If
                End If
            End If
        End If
        'check Code Matrix on ECTY ----------- end
        
        'check Code Matrix on ECCL
        If Not IsSendEmail Then
            If Len(xPlantCode) > 0 Then
                SQLQ = "SELECT * FROM CODEMATRIX WHERE CM_NAME = 'ECCL' AND CM_KEY = '" & clpCode(2).Text & "' "
                SQLQ = SQLQ & "AND CM_SECTION = '" & xPlantCode & "' "
                If rsCodeMatrix.State <> 0 Then rsCodeMatrix.Close
                rsCodeMatrix.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsCodeMatrix.EOF Then
                    If Not IsNull(rsCodeMatrix("CM_PRKEY")) Then
                        If UCase(Trim(rsCodeMatrix("CM_PRKEY"))) = "Y" Then
                            IsSendEmail = True
                        End If
                    End If
                End If
            End If
            If Not IsSendEmail Then 'Not found in this plant then check this code wihtout plant code
                SQLQ = "SELECT * FROM CODEMATRIX WHERE CM_NAME = 'ECCL' AND CM_KEY = '" & clpCode(2).Text & "' "
                'SQLQ = SQLQ & "AND CM_SECTION = '" & xPlantCode & "' "
                If rsCodeMatrix.State <> 0 Then rsCodeMatrix.Close
                rsCodeMatrix.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsCodeMatrix.EOF Then
                    If Not IsNull(rsCodeMatrix("CM_PRKEY")) Then
                        If UCase(Trim(rsCodeMatrix("CM_PRKEY"))) = "Y" Then
                            IsSendEmail = True
                        End If
                    End If
                End If
            End If
        End If
        If Not IsSendEmail Then
            Exit Sub 'No Sending Email setup as Y then Do Not send email
        End If
        
        xToEmail = GetComPreferEmail("EMAIL_ONHSINCIDENT", glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONHSINCIDENT")
        End If
        If Len(xToEmail) = 0 Then 'cannot find email
             Exit Sub
        End If
        
        frmSendEmail.txtTo.Text = xToEmail
        frmSendEmail.txtCC.Text = GetCurUserEmail
        frmSendEmail.txtSubject.Text = "New Incident Email - " & lblEEName.Caption
        
        MailBody = "Employee #: " & glbLEE_ID & " - " & "Name: " & lblEEName.Caption & vbCrLf & vbCrLf
        If IsDate(dlpDate(0).Text) Then MailBody = MailBody & "A new incident occurred on: " & CVDate(dlpDate(0).Text) & vbCrLf & vbCrLf
        'MailBody = MailBody & "Incident Number: " & lblIncidentNo & vbCrLf & vbCrLf
        If Len(clpCode(1).Text) > 0 Then MailBody = MailBody & "Incident Type: " & GetTABLDesc("ECTY", clpCode(1).Text) & vbCrLf & vbCrLf
        If Len(clpCode(1).Text) > 0 Then MailBody = MailBody & "Classification Type: " & GetTABLDesc("ECCL", clpCode(2).Text) & vbCrLf & vbCrLf
        MailBody = MailBody & "Plant: " & GetTABLDesc("EDSE", xPlantCode) & vbCrLf & vbCrLf 'Ticket #28960 Franks 07/22/2016
        If Len(txtDemo(14).Text) > 0 Then
            MailBody = MailBody & "Position: " & txtDemo(14).Text & vbCrLf & vbCrLf 'Ticket #28960 Franks 07/22/2016
        End If
        frmSendEmail.txtBody.Text = MailBody
        
        MDIMain.panHelp(0).FloodType = 0 '
        MDIMain.panHelp(0).Caption = "Sending email..."
        frmSendEmail.Tag = ""
        frmSendEmail.cmdSend_Click
        Do
            DoEvents
        Loop Until frmSendEmail.Tag <> ""   ' MC - dkostka - 05/03/01 - Changed from = "DONE" to <> ""
        ' AC - dkostka - 05/03/01 - Added checking to make sure the email went through,
        '   otherwise refuse to terminate the employee.
        If frmSendEmail.Tag = "DONE" Then
            Unload frmSendEmail
            'AbortTerm = False
        Else
            Unload frmSendEmail
            'AbortTerm = True
        End If
        MDIMain.panHelp(0).Caption = ""
        MDIMain.panHelp(0).FloodType = 1
            
End Sub

'Ticket #28815 - Opened for all so copied the WFC routine and making this general
Private Sub All_HS_INCIDENT_Email() 'Ticket #28664 Franks 05/31/2016
Dim rsCodeMatrix As New ADODB.Recordset
Dim strFields As String
Dim SQLQ As String
Dim xUptFlag As Boolean
Dim xDate1, xDate2
Dim xLDate
Dim a As Integer, Msg As String
Dim Title$, DgDef, Response%
Dim xYear, xMonth
Dim xEmail
Dim xToEmail As String
Dim IsSendEmail As Boolean
Dim MailBody

        'At least Classification or Type code required for Sending Email
        IsSendEmail = False
        If Len(clpCode(1).Text) = 0 And Len(clpCode(2).Text) = 0 Then
            Exit Sub
        End If
        
        'check Code Matrix on ECTY ----------- begin
        If Not IsSendEmail Then
            SQLQ = "SELECT * FROM CODEMATRIX WHERE CM_NAME = 'ECTY' AND CM_KEY = '" & clpCode(1).Text & "' "
            If rsCodeMatrix.State <> 0 Then rsCodeMatrix.Close
            rsCodeMatrix.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsCodeMatrix.EOF Then
                If Not IsNull(rsCodeMatrix("CM_PRKEY")) Then
                    If UCase(Trim(rsCodeMatrix("CM_PRKEY"))) = "Y" Then
                        IsSendEmail = True
                    End If
                End If
            End If
        End If
        'check Code Matrix on ECTY ----------- end
        
        'check Code Matrix on ECCL
        If Not IsSendEmail Then
            SQLQ = "SELECT * FROM CODEMATRIX WHERE CM_NAME = 'ECCL' AND CM_KEY = '" & clpCode(2).Text & "' "
            If rsCodeMatrix.State <> 0 Then rsCodeMatrix.Close
            rsCodeMatrix.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsCodeMatrix.EOF Then
                If Not IsNull(rsCodeMatrix("CM_PRKEY")) Then
                    If UCase(Trim(rsCodeMatrix("CM_PRKEY"))) = "Y" Then
                        IsSendEmail = True
                    End If
                End If
            End If
        End If
        
        If Not IsSendEmail Then
            Exit Sub 'No Sending Email setup as Y then Do Not send email
        End If
        
        xToEmail = GetComPreferEmail("EMAIL_ONHSINCIDENT", glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONHSINCIDENT")
        End If
        
        'Cannot find email
        If Len(xToEmail) = 0 Then
             Exit Sub
        End If
        
        frmSendEmail.txtTo.Text = xToEmail
        frmSendEmail.txtCC.Text = GetCurUserEmail
        frmSendEmail.txtSubject.Text = "New Incident Email - " & lblEEName.Caption
        
        MailBody = "Employee #: " & glbLEE_ID & " - " & "Name: " & lblEEName.Caption & vbCrLf & vbCrLf
        If IsDate(dlpDate(0).Text) Then MailBody = MailBody & "A new incident occurred on: " & CVDate(dlpDate(0).Text) & vbCrLf & vbCrLf
        'MailBody = MailBody & "Incident Number: " & lblIncidentNo & vbCrLf & vbCrLf
        If Len(clpCode(1).Text) > 0 Then MailBody = MailBody & "Incident Type: " & GetTABLDesc("ECTY", clpCode(1).Text) & vbCrLf '& vbCrLf
        If Len(clpCode(2).Text) > 0 Then MailBody = MailBody & "Classification Type: " & GetTABLDesc("ECCL", clpCode(2).Text) & vbCrLf & vbCrLf
        
        'Ticket #28815 - No Plant Code concept for rest of the clients
        'MailBody = MailBody & "Plant: " & GetTABLDesc("EDSE", xPlantCode) & vbCrLf & vbCrLf 'Ticket #28960 Franks 07/22/2016
        
        If Len(txtDemo(14).Text) > 0 Then
            MailBody = MailBody & "Position: " & txtDemo(14).Text & vbCrLf & vbCrLf 'Ticket #28960 Franks 07/22/2016
        End If
        frmSendEmail.txtBody.Text = MailBody
        
        MDIMain.panHelp(0).FloodType = 0 '
        MDIMain.panHelp(0).Caption = "Sending email..."
        
        'frmSendEmail.Tag = ""
        
        frmSendEmail.Show 1
        'frmSendEmail.cmdSend_Click
        
        'Do
        '    DoEvents
        'Loop Until frmSendEmail.Tag <> ""   ' MC - dkostka - 05/03/01 - Changed from = "DONE" to <> ""
        
        ' AC - dkostka - 05/03/01 - Added checking to make sure the email went through,
        '   otherwise refuse to terminate the employee.
        'If frmSendEmail.Tag = "DONE" Then
        '    Unload frmSendEmail
        '    'AbortTerm = False
        'Else
        '    Unload frmSendEmail
        '    'AbortTerm = True
        'End If
        
        MDIMain.panHelp(0).Caption = ""
        MDIMain.panHelp(0).FloodType = 1
            
End Sub

'Private Sub cmdOK_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Incident"
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

RHeading = lblEEName & "'s Incident"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub


'Private Sub cmdPrint_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdTCause_Click()
'frmEHSCause.Show
'Unload Me
'End Sub

'Private Sub cmdTCause_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdWCBMed_Click()
'frmEHSWCB.Show
'Unload Me
'End Sub

'Private Sub cmdWCBMed_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdWSIB_Click()
'frmEHSWCBC.Show
'Unload Me
'End Sub


Function EERetrieve()

Dim X, xFld
Dim SQLQ As String

EERetrieve = False

Screen.MousePointer = HOURGLASS
On Error GoTo EERError
SQLQ = "SELECT " & FldList & ", "
For X = 0 To 1
    xFld = IIf(X = 0, "EMPNOT", "SUPERVISOR")
    If glbLinamar Then
        SQLQ = SQLQ & " CASE WHEN EC_" & xFld & " IS NOT NULL AND LEN(EC_" & xFld & ")>2 "
        SQLQ = SQLQ & " THEN RIGHT(EC_" & xFld & ",3)+'-'+"
        SQLQ = SQLQ & " LEFT(EC_" & xFld & ",LEN(EC_" & xFld & ")-3) "
        SQLQ = SQLQ & " ELSE STR(EC_" & xFld & ") END "
        SQLQ = SQLQ & " AS " & xFld & IIf(X = 1, "", ",")
    Else
        If glbOracle Then
                        SQLQ = SQLQ & "EC_" & xFld & " AS " & xFld & IIf(X = 1, "", ",")
        Else
                SQLQ = SQLQ & "STR(EC_" & xFld & ") AS " & xFld & IIf(X = 1, "", ",")
        End If
        
    End If
Next
If glbtermopen Then
    SQLQ = SQLQ & " FROM Term_HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = SQLQ & " FROM HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " WHERE EC_EMPNBR = " & glbLEE_ID
End If
'SQLQ = SQLQ & " ORDER BY EC_OCCDATE DESC"
SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
Data1.RecordSource = SQLQ
Data1.Refresh


'Ticket #14703
If glbLinamar Then
    Call Retrieve_Incidents 'Populate the Related Incident # dropdown list
End If


EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function


EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "OCH Retrieve", "HR_OCC_HEALTH_SAFETY", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

Exit Function

End Function

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


Private Sub cmdDemo_Click()
    frmEIncidentDemo.txtCountryOfEmp = txtDemo(11)
    frmEIncidentDemo.txtIncidentNo = Me.txtIncidentNo
    frmEIncidentDemo.clpCode(2) = txtDemo(2)
    frmEIncidentDemo.clpCode(3) = txtDemo(3)
    frmEIncidentDemo.clpCode(4) = txtDemo(4)
    frmEIncidentDemo.clpCode(5) = txtDemo(5)
    frmEIncidentDemo.clpCode(6) = txtDemo(6)
    frmEIncidentDemo.clpCode(7) = txtDemo(7)
    frmEIncidentDemo.clpCode(8) = txtDemo(8)
    frmEIncidentDemo.clpCode(9) = txtDemo(9)
    frmEIncidentDemo.clpCode(10) = txtDemo(10)
    If glbLinamar Then 'Mid(rsDATA("EC_HOMEOPRTNBR"), 4)
        If Len(txtDemo(12)) > 3 Then
            frmEIncidentDemo.clpHOME(1) = Mid(txtDemo(12), 4)
        End If
        If Len(txtDemo(13)) > 3 Then
            frmEIncidentDemo.clpHOME(2) = Mid(txtDemo(13), 4)
        End If
    End If
    frmEIncidentDemo.txtJobDesc = txtDemo(14)
    
    If glbLinamar Then
        'If Not IsNull(Data1.Recordset("ED_HOMEOPRTNBR")) Then
        '    clpHOME(1) = Mid(Data1.Recordset("ED_HOMEOPRTNBR"), 4)
        'Else
        '    clpHOME(1) = ""
        'End If
        'If Not IsNull(Data1.Recordset("ED_HOMELINE")) Then
        '    clpHOME(2) = Mid(Data1.Recordset("ED_HOMELINE"), 4)
        'Else
        '    clpHOME(2) = ""
        'End If
        If Not IsNull(frmEIncidentDemo.clpCode(8)) Then
            frmEIncidentDemo.clpCode(8) = Mid(frmEIncidentDemo.clpCode(8), 4)
        Else
            frmEIncidentDemo.clpCode(8) = ""
        End If
        If Not IsNull(frmEIncidentDemo.clpCode(9)) Then
            frmEIncidentDemo.clpCode(9) = Mid(frmEIncidentDemo.clpCode(9), 4)
        Else
            frmEIncidentDemo.clpCode(9) = ""
        End If
    End If
    
    frmEIncidentDemo.Show 1
End Sub

Private Sub cmdPageRight_Click(Index As Integer)
    'Save the data
    If Not cmdOK_Click() Then Exit Sub
    
    'Unload the current form and load the next one
    Unload Me
    
    'Next form
    Screen.MousePointer = HOURGLASS
    Load frmEHSINJURYWF7
    frmEHSINJURYWF7.ZOrder 0
    Screen.MousePointer = DEFAULT
    
End Sub

Private Sub cmdPostion_Click()
Dim OJOB As String, OJobD As String

OJOB = clpJob.Text
OJobD = txtDemo(14).Text

Load frmJOBS
frmJOBS.Show 1

'If Len(glbJob) < 1 Then
If Len(glbPos) < 1 Then
    clpJob.Text = OJOB
    txtDemo(14).Text = OJobD
Else
    clpJob.Text = glbPos
    txtDemo(14).Text = glbPosDesc
    Call Populate_Job_Start_Date
End If

End Sub

Private Sub comDateClaimNo_Click()
    If comDateClaimNo.Text <> "" Then
        txtClaimNo.Text = Trim(Left(comDateClaimNo.Text, InStr(1, comDateClaimNo.Text, " - ") - 1))
        txtClaimDate.Text = Trim(Mid(comDateClaimNo.Text, Len(txtClaimNo.Text) + 3))
    End If
End Sub

Private Sub comRelIncident_Change()
  If Not (Val(comRelIncident) = 0) Then
    txtRelIncident = comRelIncident
  Else
    txtRelIncident = ""
  End If
End Sub

Private Sub comRelIncident_Click()
  If Not (Val(comRelIncident) = 0) Then
    txtRelIncident = comRelIncident
  Else
    txtRelIncident = ""
  End If
End Sub

Private Sub dlpDate_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub elpReptAuthShow_Change(Index As Integer)
txtReptAuthority(Index).Text = getEmpnbr(elpReptAuthShow(Index).Text)

'Ticket #15172 Show employee name including term employee
If glbLinamar Then
    If Not glbtermopen Then
        lblReptAuthority(Index).Caption = GetSuperEmpName(elpReptAuthShow(Index).Text)
    End If
End If
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
'Me.cmdModify_Click
glbOnTop = "FRMEHSINCIDENT"
End Sub

Private Sub Form_GotFocus()
glbOnTop = "FRMEHSINCIDENT"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim X%
Dim SQLQ As String

glbOnTop = "FRMEHSINCIDENT"

elpReptAuthShow(0).Caption = ""
elpReptAuthShow(1).Caption = ""
fglbJobList = 0

If glbWFC Then 'Ticket #27576 Franks 10/26/2015
    Call WFCStatusScreen
End If

If glbLinamar Then
    txtDemo(12).DataField = "EC_HOMEOPRTNBR"
    txtDemo(13).DataField = "EC_HOMELINE"
    'lblTitle(6).FontBold = True
    txtReptAuthorityFName(0).DataField = "EC_EMPNOT_FNAME"
    txtReptAuthorityFName(1).DataField = "EC_SUPER_FNAME"
    txtReptAuthoritySName(0).DataField = "EC_EMPNOT_SURNAME"
    txtReptAuthoritySName(1).DataField = "EC_SUPER_SURNAME"
End If

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If
Data3.ConnectionString = glbAdoIHRDB
Data3.RecordSource = "HROHSNBR"


Screen.MousePointer = DEFAULT

If glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2172W" Then
'Ticket #16096 D. Muskoka
'Ticket #17112 County of Lanark
    frmWFC.Visible = True
End If

'WSIB Form 7 fields available only when the user has access to the security
If gSec_Inq_HSW7CmpMst And gSec_Inq_HSW7Injury And glbWSIBModule Then
    chkGenerateWSIB7.Visible = True
    chkReoccurence.Visible = True
    cmdPostion.Visible = True
    clpJob.Visible = True
    'Ticket #21550
    cmbJBStartDate.Visible = True
    lblTitle(26).Visible = True
    dlpDate(7).Visible = False
    
    'Ticket #21737
    lblTitle(7).Caption = "Reported To"
    lblTitle(14).Caption = "Health Care Provided"
    clpCode(3).Tag = "01-Health Care provided"
Else
    chkGenerateWSIB7.Visible = False
    chkReoccurence.Visible = False
    cmdPostion.Visible = False
    clpJob.Visible = False
    'Ticket #21550
    cmbJBStartDate.Visible = False
    lblTitle(26).Visible = False
    dlpDate(7).Visible = False
End If

If glbWFC Then
    frmWFC.Visible = True
    cmbShift.Visible = True
    txtShift.Visible = False
    cmbShift.AddItem "Days"
    cmbShift.AddItem "Afternoon"
    cmbShift.AddItem "Night"
    
    lblTitle(15).Enabled = False
    lblTitle(16).Enabled = False
    lblTitle(18).Enabled = False
    lblTitle(19).Enabled = False
    lblTitle(20).Enabled = False
    
    'Ticket #15396 - Begin
    lblTitle(2).FontBold = True
    lblTitle(4).FontBold = True
    lblTitle(6).FontBold = True
    lblTitle(9).FontBold = True
    'Ticket #15396 - End

End If
'Ticket # 6831 - For Burlington Tech.
If glbCompSerial = "S/N - 2351W" Then
    lblTitle(9).Caption = "Frequency"
    cmbHRat.Width = 1500
    cmbHRat.Left = 7560
    cmbHRat.AddItem "Frequent"
    cmbHRat.AddItem "Occassional"
    cmbHRat.AddItem "Rare"
    lblTitle(7).Caption = "Reported To"
    lblTitle(14).Caption = "First Aid Provided By"
    clpCode(3).Tag = "01-First aid provided By"
    lblTitle(13).Visible = False
    clpCode(4).Visible = False
ElseIf glbLinamar Then 'Ticket #15172
    lblTitle(9).Caption = "Onset"
    cmbHRat.Width = 1500
    cmbHRat.AddItem "Gradual"
    cmbHRat.AddItem "Sudden"
Else
    cmbHRat.AddItem "A"
    cmbHRat.AddItem "B"
    cmbHRat.AddItem "C"
End If
'Ticket# 7963 for CITY OF SARNIA
If glbCompSerial = "S/N - 2362W" Then
    lblTitle(17).Visible = True
    lblTitle(18).Visible = True
    medShiftsLost.Visible = True
    dlpDate(4).Visible = True
    medShiftsLost.DataField = "EC_SHIFTSLOST"
    dlpDate(4).DataField = "EC_APPRDATE"
    
    'Release 8.1 - Do not want 'Reported to/by' to be mandatory
    lblTitle(7).FontBold = False
Else
    'Ticket #28283 - Jerry said to make it non mandatory for all
    lblTitle(7).FontBold = False
End If

'Ticket #14573 - Linamar
If glbLinamar Then
    lblTitle(5).Caption = "Incident Type"
    lblTitle(6).Caption = "Type of Event"
    lblTitle(6).FontBold = True
    lblTitle(12).FontBold = True
    lblTitle(14).FontBold = True
    lblTitle(13).FontBold = True
    
    'Ticket #14703
    lblTitle(24).Visible = True
    comRelIncident.Visible = True
    'comRelIncident.DataField = "EC_RELINCIDENT"
    txtRelIncident.DataField = "EC_RELINCIDENT" 'Ticket #15827
    
    'Ticket #15172 - Begin
    lblTitle(22).Visible = False: chkFollowed.Visible = False 'Policy/Procedure Followed
    'lblTitle(9).Visible = False: cmbHRat.Visible = False 'Hazard Rating -> change to "Onset"
    chkOvertime.Caption = "OSHA Recordable"
    chkWType(0).Visible = False 'Regular Schedule Work Shift
    chkWType(2).Top = chkWType(1).Top
    chkWType(1).Top = chkWType(0).Top
    frmWorkType.Height = 1000
    lblTitle(15).Caption = "Expected Return to Modified Work"
    dlpDate(3).Tag = "41-Date expected to return to Modified Work"
    lblTitle(7).Caption = "Reported To"
    lblTitle(12).Caption = "Direct Supervisor"
    
    If Not glbtermopen Then 'Show employee name including term employee
        elpReptAuthShow(0).Width = 1600
        elpReptAuthShow(1).Width = 1600
        lblReptAuthority(0).Top = elpReptAuthShow(0).Top
        lblReptAuthority(1).Top = elpReptAuthShow(1).Top
        lblReptAuthority(0).Left = 3750
        lblReptAuthority(1).Left = 3750
        lblReptAuthority(0).Caption = ""
        lblReptAuthority(1).Caption = ""
        lblReptAuthority(0).Visible = True
        lblReptAuthority(1).Visible = True
    End If
    lblTitle(18).Caption = "Expected Return to Regular Work"
    lblTitle(18).Visible = True
    dlpDate(4).Visible = True
    dlpDate(4).DataField = "EC_APPRDATE"
    'Ticket #15172 - End
    
    'Hemu - Ticket #15172
    lblTitle(19).Caption = "Return to Modified Work"
    dlpDate(5).Tag = "41-Date returned to Modified Work"
    lblTitle(20).Caption = "Expected Return to Regular Work"
    dlpDate(6).Tag = "41-Date expected return to Regular Work"
    lblTitle(18).Caption = "Return to Regular Work"
    dlpDate(4).Tag = "41-Date return to Regular Work"
    'Hemu - End - Ticket #15172
End If

'Ticket# 13413 for Bird Packaging Limited
If glbCompSerial = "S/N - 2387W" Then
    imgEmail.Visible = True
    lblTitle(12).Caption = "Assigned To"
    lblTitle(14).Caption = "Step"
    lblTitle(13).Caption = "Status"
    vbxTrueGrid.Columns(9).Caption = "Assigned To"
    SQLQ = "UPDATE HRTABDES SET TD_DESC = 'INCIDENT STEP CODE' WHERE TD_NAME = 'ECFF'"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRTABDES SET TD_DESC = 'INCIDENT STATUS CODE' WHERE TD_NAME = 'ECPB'"
    gdbAdoIhr001.Execute SQLQ
    
End If

If glbLinHS Then 'Ticket #12401
    glbLinEmpNo = glbLEE_ID
    If Not glbtermopen Then
        If Len(glbDiv) = 0 Then Call Get_Div(False)  'True) 'frmDIVISIONS.Show 1
        If Len(glbDiv) = 0 Then Unload Me: Exit Sub
    Else
        If Len(glbDiv) = 0 Then Call Get_Div(False) 'frmDIVISIONS.Show 1
        If Len(glbDiv) = 0 Then Unload Me: Exit Sub
    End If
    glbLinHSDivNo = Val("999999" & glbDiv)
    glbLEE_ID = glbLinHSDivNo
    glbLEE_SName = glbDivDesc
    lblTitle(23).Top = lblTitle(17).Top
    lblTitle(23).Visible = True
    clpCode(5).Top = medShiftsLost.Top
    clpCode(5).Visible = True
    clpCode(5).DataField = "EC_IMPACT"
Else
    If glbLinamar Then
        If glbLEE_ID <> 0 Then
            If Left(Trim(Str(glbLEE_ID)), 6) = "999999" Then
                glbLEE_ID = 0
            End If
        End If
    End If
    
    If Not glbtermopen Then
        If glbLEE_ID = 0 Then frmEEFIND.Show 1
        If glbLEE_ID = 0 Then Unload Me: Exit Sub
    Else
        If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
        If glbTERM_ID = 0 Then Unload Me: Exit Sub
    End If
End If
If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

If Len(glbLEE_SName) < 1 Then Exit Sub
Screen.MousePointer = HOURGLASS

Me.vbxTrueGrid.SetFocus
If glbLinHS Then
    If Len(glbDivDesc) > 0 Then   ' dont do on add new until in
        Me.Caption = "Incident Data - " & glbDivDesc
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv
    lblEENumber.Caption = lStr("Division")
Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
        Me.Caption = "Incident Data - " & Left$(glbLEE_SName, 5)
        Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    lblEENum.Caption = ShowEmpnbr(lblEEID)
End If

Call ST_UPD_MODE(False)
Call Display_Value

'Get the list of curent jobs of the employee
Call CR_JobHis_Snap

'Change the Code Description for Linamar - Ticket #14573
If glbLinamar Then
    'OH&S CLASSIFICATION CODES to Type of Event
    Call Change_Code_Table_Description
End If

'If Not gSec_Upd_Health_Safety Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
'End If
Call INI_Controls(Me)
xPlantCode = ""
If glbWFC Then
    Call WFCPlantCode
End If

'Get the lookup for Position to show only current jobs of the employee
clpJob.seleEMPCode = fglbJobList

Screen.MousePointer = DEFAULT
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
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
fraDetail.Height = 7350 '6855 '6615 '6260 '5000
If Me.Height >= vbxTrueGrid.Height + panEEDESC.Height + fraDetail.Height + 530 Then
    scrControl.Value = 0
    'fraDetail.Top = vbxTrueGrid.Height + panEEDESC.Height + 240
    fraDetail.Top = 3000
    scrControl.Visible = False
    Exit Sub
End If
'If Me.Height < vbxTrueGrid.Height + panEEDESC.Height + scrControl.Top + 400 Then Exit Sub
If Me.Height < vbxTrueGrid.Height + panEEDESC.Height + scrControl.Top Then Exit Sub
scrControl.Visible = True

scrControl.Max = vbxTrueGrid.Height + panEEDESC.Height + fraDetail.Height + 1100 - Me.Height
'scrControl.Max = vbxTrueGrid.Height + panEEDESC.Height + fraDetail.Height '- Me.Height
scrControl.Left = Me.Width - scrControl.Width - 220
scrControl.Height = Me.Height - scrControl.Top - 700
End Sub

Private Sub Form_Unload(Cancel As Integer)

MDIMain.panHelp(0).Caption = "Select function from the menu."
Set frmEHSINCIDENT = Nothing 'carmen may 00
Call NextForm
End Sub

Private Sub imgEmail_Click()
Call EmailShow
End Sub

Private Sub lblIncidentNo_Change()
Dim xIncidentNo
xIncidentNo = Format(lblIncidentNo, "00000000")
txtYear = Left(xIncidentNo, 4)

If Val(xIncidentNo) = 0 Then txtIncidentNo = "" Else txtIncidentNo = Val(Right(xIncidentNo, 4))

End Sub


Private Sub medIncidentTime_GotFocus()
medIncidentTime.Mask = "##:##"
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medIncidentTime_LostFocus()
medIncidentTime.Mask = ""
End Sub

Private Sub medNotifyTime_GotFocus()
medNotifyTime.Mask = "##:##"
Call SetPanHelp(ActiveControl)
End Sub


Private Sub medNotifyTime_LostFocus()
medNotifyTime.Mask = ""
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

glbOHSEdit% = TF

'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF


'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
'cmdWCBMed.Enabled = FT
'cmdInjLoc.Enabled = FT
'cmdCAction.Enabled = FT
'cmdContact.Enabled = FT
'cmdTCause.Enabled = FT
'cmdWSIB.Enabled = FT

medIncidentTime.Enabled = TF
medNotifyTime.Enabled = TF
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
clpCode(4).Enabled = TF
dlpDate(0).Enabled = TF
dlpDate(1).Enabled = TF

elpReptAuthShow(0).Enabled = TF
elpReptAuthShow(1).Enabled = TF
txtShift.Enabled = TF
cmbHRat.Enabled = TF
chkOvertime.Enabled = TF
chkGenerateWSIB7.Enabled = TF
chkReoccurence.Enabled = TF
comDateClaimNo.Enabled = TF

If glbWFC Then
    dlpDate(2).Enabled = False
    dlpDate(3).Enabled = False
    dlpDate(4).Enabled = False
    dlpDate(5).Enabled = False
    dlpDate(6).Enabled = False
    chkWType(0).Enabled = False
    chkWType(1).Enabled = False
    chkWType(2).Enabled = False
    chkModDuties.Enabled = False
Else
    dlpDate(2).Enabled = TF
    dlpDate(3).Enabled = TF
    chkWType(0).Enabled = TF
    chkWType(1).Enabled = TF
    chkWType(2).Enabled = TF
    chkModDuties.Enabled = TF
    dlpDate(5).Enabled = TF
    dlpDate(6).Enabled = TF
End If

frmWFC.Enabled = TF
elpReptAuthShow(0).Enabled = TF
elpReptAuthShow(1).Enabled = TF
If Data1.Recordset.BOF Or Data1.Recordset.EOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If
If glbtermopen Then
'    cmdNew.Enabled = False
End If
If glbCompSerial = "S/N - 2362W" Then  'CITY OF SARNIA
    dlpDate(4).Enabled = TF
    medShiftsLost.Enabled = TF
End If
'vbxTrueGrid.Enabled = FT

'Ticket #14703
If glbLinamar Then
    comRelIncident.Enabled = TF
    dlpDate(4).Enabled = TF
End If

clpCode(6).Enabled = TF 'Ticket #27576 Franks 10/26/2015
dlpDate(8).Enabled = TF 'Ticket #27576 Franks 10/26/2015

End Sub

Private Sub scrControl_Change()
fraDetail.Top = 240 + vbxTrueGrid.Height + panEEDESC.Height - scrControl.Value * (((scrControl.Max + 100) / scrControl.Max) / scrControl.Max) * ScaleHeight
End Sub

Private Sub txtAMPMIncident_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtAMPMIncident_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtAMPMNotified_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtAMPMNotified_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtClaimDate_Change()
    Dim X As Integer
    For X = 0 To comDateClaimNo.ListCount - 1
        If comDateClaimNo.List(X) = txtClaimNo.Text & " - " & txtClaimDate.Text Then
            comDateClaimNo.ListIndex = X
            Exit For
        Else
            'If txtClaimNo.Text = "" Then
            '    comDateClaimNo.ListIndex = -1
            'End If
        End If
    Next
End Sub

Private Sub txtClaimNo_Change()
    Dim X As Integer
    For X = 0 To comDateClaimNo.ListCount - 1
        If comDateClaimNo.List(X) = txtClaimNo.Text & " - " & txtClaimDate.Text Then
            comDateClaimNo.ListIndex = X
            Exit For
        Else
            'If txtClaimNo.Text = "" Then
            '    comDateClaimNo.ListIndex = -1
            'End If
        End If
    Next
End Sub

Private Sub txtHRat_Change()
'Ticket # 6831 - For Burlington Tech.
If glbCompSerial = "S/N - 2351W" Then
    Select Case txtHRat
        Case "F": cmbHRat.ListIndex = 0
        Case "O": cmbHRat.ListIndex = 1
        Case "R": cmbHRat.ListIndex = 2
        Case Else
            cmbHRat.ListIndex = -1
    End Select
ElseIf glbLinamar Then
    Select Case txtHRat
        Case "G": cmbHRat.ListIndex = 0
        Case "S": cmbHRat.ListIndex = 1
        Case Else
            cmbHRat.ListIndex = -1
    End Select
Else
    cmbHRat = txtHRat
End If
End Sub

Private Sub txtHRat_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtIncidentNo_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtJobStartDate_Change()
    Dim X As Integer
    
    If cmbJBStartDate.Visible = True Then
        For X = 0 To cmbJBStartDate.ListCount - 1
            If cmbJBStartDate.List(X) = txtJobStartDate.Text Then
                cmbJBStartDate.ListIndex = X
                Exit For
            End If
        Next X
        If txtJobStartDate = "" Then
            cmbJBStartDate.ListIndex = -1
        End If
    End If
End Sub

Private Sub txtRelIncident_Change()
  If Not (Val(txtRelIncident) = 0) Then
    comRelIncident = txtRelIncident
  Else
    comRelIncident = ""
  End If
End Sub

Private Sub txtReptAuthority_Change(Index As Integer)
    elpReptAuthShow(Index).Text = ShowEmpnbr(txtReptAuthority(Index).Text)
    
    If IsNumeric(txtReptAuthority(Index).Text) Then
        txtReptAuthorityFName(Index).Text = GetEmpData(txtReptAuthority(Index).Text, "ED_FNAME")
        txtReptAuthoritySName(Index).Text = GetEmpData(txtReptAuthority(Index).Text, "ED_SURNAME")
    End If
End Sub

Private Sub txtShift_Change()
If Not glbWFC Then Exit Sub
Select Case txtShift
Case "D": cmbShift.ListIndex = 0
Case "A": cmbShift.ListIndex = 1
Case "N": cmbShift.ListIndex = 2
End Select
End Sub

Private Sub txtShift_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtYear_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Updstats_Change(Index As Integer)
    If Index = 0 Then
        'If IsDate(Updstats(Index).Text) Then
        lblUpdDateDesc.Caption = Updstats(Index).Text
        'End If
    End If
    If Index = 2 Then
        lblUserDesc.Caption = GetUserDesc(Updstats(Index))
    End If
End Sub

Private Sub vbxTrueGrid_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
 Dim SQLQ As String
 Dim X As Integer
 Dim xFld As String
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT " & FldList & ", "
        For X = 0 To 1
            xFld = IIf(X = 0, "EMPNOT", "SUPERVISOR")
            If glbLinamar Then
                SQLQ = SQLQ & " CASE WHEN EC_" & xFld & " IS NOT NULL AND LEN(EC_" & xFld & ")>2 "
                SQLQ = SQLQ & " THEN RIGHT(EC_" & xFld & ",3)+'-'+"
                SQLQ = SQLQ & " LEFT(EC_" & xFld & ",LEN(EC_" & xFld & ")-3) "
                SQLQ = SQLQ & " ELSE STR(EC_" & xFld & ") END "
                SQLQ = SQLQ & " AS " & xFld & IIf(X = 1, "", ",")
            Else
                If glbOracle Then
                                SQLQ = SQLQ & "EC_" & xFld & " AS " & xFld & IIf(X = 1, "", ",")
                Else
                        SQLQ = SQLQ & "STR(EC_" & xFld & ") AS " & xFld & IIf(X = 1, "", ",")
                End If
                
            End If
        Next
        If glbtermopen Then
            SQLQ = SQLQ & " FROM Term_HR_OCC_HEALTH_SAFETY "
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = SQLQ & " FROM HR_OCC_HEALTH_SAFETY "
            SQLQ = SQLQ & " WHERE EC_EMPNBR = " & glbLEE_ID
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
'    End If
End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$, X%
Dim SQLQ As String

On Error GoTo Tab1_Err

'Ticket #14703
If glbLinamar Then
    Call Retrieve_Incidents 'Populate the Related Incident # dropdown list
End If


Call Display_Value

Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_OCC_HEALTH_SAFETY", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Function Set_Cur_Position()
Dim SQLQ As String, Msg$
Dim Snap_Job_His As New ADODB.Recordset

Set_Cur_Position = False

On Error GoTo SCError
Screen.MousePointer = HOURGLASS

SQLQ = "Select HR_JOB_HISTORY.* from HR_JOB_HISTORY"
SQLQ = SQLQ & " where JH_EMPNBR = " & glbLEE_ID & " "

SQLQ = SQLQ & " and JH_CURRENT <>0 "

Snap_Job_His.Open SQLQ, gdbAdoIhr001, adOpenKeyset
If glbtermopen Then
    glbStopPerform% = False
    savAuth = ""
Else
    If Snap_Job_His.BOF And Snap_Job_His.EOF Then
        Msg$ = "No current position found "
        Msg$ = Msg$ & Chr(10) & "Please review position prior. "
        MsgBox Msg$
        glbStopPerform% = True
        Screen.MousePointer = DEFAULT
        Exit Function
    Else
        glbStopPerform% = False
    End If
    txtDemo(14).Text = GetJobDesc(Snap_Job_His("JH_JOB"))
    If IsNull(Snap_Job_His("JH_REPTAU")) Then savAuth = "" Else savAuth = Snap_Job_His("JH_REPTAU")
End If
Snap_Job_His.Close

Screen.MousePointer = DEFAULT
Set_Cur_Position = True

Exit Function

SCError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HR_JOB_HISTORY", "SELECT")
Call RollBack '28July99 js
Resume Next
End Function


Private Function FldList()
Dim SQLQ
SQLQ = ""
SQLQ = SQLQ & "EC_COMPNO, EC_EMPNBR, EC_CASE, EC_OCCDATE, "
SQLQ = SQLQ & "EC_OCCTM, EC_DATENOT, EC_TIMNOT, EC_EMPNOT,"
SQLQ = SQLQ & "EC_SHIFT, EC_CLASS_TABL, EC_CLASS, EC_TYPE_TABL,"
SQLQ = SQLQ & "EC_TYPE, EC_HAZARD, EC_OTWORK, EC_WT_REG,"
SQLQ = SQLQ & "EC_WT_TRAIN, EC_WT_TMPTRN, EC_SUPERVISOR, "
SQLQ = SQLQ & "EC_LDAY, EC_RETURN, EC_MODDUTIES, EC_FAPROVIDED, EC_PROVIDEDBY, "
SQLQ = SQLQ & "EC_RETURN_REG, EC_RETURN_SUITABLE, EC_COMMENTS_INC, EC_POLICY_FLAG, "
SQLQ = SQLQ & "EC_DEPTNO, EC_DIV, EC_LOC, EC_ORG, "
SQLQ = SQLQ & "EC_EMP, EC_PT, EC_REGION, EC_SECTION, "
SQLQ = SQLQ & "EC_ADMINBY, EC_WORKCOUNTRY, EC_JOBDESC,"
SQLQ = SQLQ & "EC_LDATE, EC_LTIME,EC_LUSER, "
SQLQ = SQLQ & "EC_EMPNOT_FNAME, EC_EMPNOT_SURNAME,EC_SUPER_FNAME,EC_SUPER_SURNAME "
SQLQ = SQLQ & ",EC_JBCODE,EC_FORM7, EC_WCBRES,EC_JBSDATE"
SQLQ = SQLQ & ",EC_REOCCURENCE, EC_REOCCUR_DATE,EC_REOCCUR_CLAIM_NUM,EC_WCBNBR,EC_WCBFDTE,EC_OCCTM_FORMAT,EC_TIMNOT_FORMAT "

If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"
If glbCompSerial = "S/N - 2362W" Then 'CITY OF SARNIA
    SQLQ = SQLQ & ",EC_SHIFTSLOST"
End If
If glbCompSerial = "S/N - 2362W" Or glbLinamar Then 'CITY OF SARNIA and Linamar
    SQLQ = SQLQ & ",EC_APPRDATE"
End If
If glbLinamar Then
    SQLQ = SQLQ & ",EC_HOMEOPRTNBR ,EC_HOMELINE,EC_IMPACT, EC_RELINCIDENT"
End If
If glbWFC Then 'Ticket #27576 Franks 10/26/2015
    SQLQ = SQLQ & ",EC_STATUS ,EC_STDATE "
End If
FldList = SQLQ
End Function

''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
    Dim SQLQ
    
Dim X, xFld
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbtermopen Then
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        Call SET_UP_MODE
Me.cmdModify_Click
        Exit Sub
    End If
    
    
    SQLQ = "SELECT " & FldList & ", "
For X = 0 To 1
    xFld = IIf(X = 0, "EMPNOT", "SUPERVISOR")
    If glbLinamar Then
        SQLQ = SQLQ & " CASE WHEN EC_" & xFld & " IS NOT NULL AND LEN(EC_" & xFld & ")>2 "
        SQLQ = SQLQ & " THEN RIGHT(EC_" & xFld & ",3)+'-'+"
        SQLQ = SQLQ & " LEFT(EC_" & xFld & ",LEN(EC_" & xFld & ")-3) "
        If glbOracle Then
            SQLQ = SQLQ & " ELSE EC_" & xFld & " END "
        Else
            SQLQ = SQLQ & " ELSE STR(EC_" & xFld & ") END "
        End If
        
        SQLQ = SQLQ & " AS " & xFld & IIf(X = 1, "", ",")
    Else
        If glbOracle Then
            SQLQ = SQLQ & "EC_" & xFld & " AS " & xFld & IIf(X = 1, "", ",")
        Else
            SQLQ = SQLQ & "STR(EC_" & xFld & ") AS " & xFld & IIf(X = 1, "", ",")
        End If
        
    End If
Next
If glbtermopen Then
    SQLQ = SQLQ & " FROM Term_HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " WHERE EC_CASE=" & Data1.Recordset!EC_CASE
    'If glbWFC Then
    If glbWFC Or glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2172W" Then
    'Ticket #16096 D. Muskoka
    'Ticket #17112 County of Lanark
        SQLQ = SQLQ & " AND TERM_SEQ = " & glbTERM_Seq
    End If
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = SQLQ & " FROM HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " WHERE EC_CASE=" & Data1.Recordset!EC_CASE
    'If glbWFC Then
    If glbWFC Or glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2172W" Then
    'Ticket #16096 D. Muskoka
    'Ticket #17112 County of Lanark
        SQLQ = SQLQ & " AND EC_EMPNBR =" & glbLEE_ID
    End If
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If
'SQLQ = SQLQ & " ORDER BY EC_OCCDATE DESC"
SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
    
If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    
Call Set_Control("R", Me, rsDATA)

Call SET_UP_MODE

'Get the list of curent jobs of the employee
Call CR_JobHis_Snap

'Get the lookup for Position to show only current jobs of the employee
If gSec_Inq_HSW7CmpMst And gSec_Inq_HSW7Injury And glbWSIBModule Then
    clpJob.seleEMPCode = fglbJobList
    If clpJob.Text = "" Then
        If cmbJBStartDate.Visible = True Then
            cmbJBStartDate.Clear
            dlpDate(7).Visible = False
        Else
            cmbJBStartDate.Visible = True
            dlpDate(7).Visible = False
        End If
    Else
        Call Populate_Job_Start_Date
    End If
End If

Me.cmdModify_Click

End Sub
'Private Sub Set_Control2(Act As String, Optional rsTA As ADODB.Recordset)
'  If Act = "U" Then
'            If Len(medNotifyTime) = 0 Then
'                  rsTA!EC_TIMNOT = Null
'            Else
'                  rsTA!EC_TIMNOT = medNotifyTime.Mask
'            End If
'  ElseIf Act = "B" Then
'            medNotifyTime.Mask = ""
'  ElseIf Act = "R" Then
'            medNotifyTime.Mask = ""
'            If rsTA.EOF Or rsTA.BOF Then Exit Sub
'                   If IsNull(rsTA!EC_TIMNOT) Then
'                        medNotifyTime.Mask = ""
'                    Else
'                        medNotifyTime = rsTA!EC_TIMNOT
'                    End If
'   End If
'End Sub


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
RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Health_Safety
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
ElseIf rsDATA.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
Call ST_UPD_MODE(TF)
End Sub
Private Sub lblEEID_Change()

If glbLinHS Then
    If Len(glbDivDesc) > 0 Then   ' dont do on add new until in
        frmEHSINCIDENT.Caption = "Incident Data - " & glbDivDesc
        frmEHSINCIDENT.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv
    
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = ""
    End If
Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
        frmEHSINCIDENT.Caption = "Incident Data - " & Left$(glbLEE_SName, 5)
        frmEHSINCIDENT.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
        
        If glbLinamar Then  'Ticket #14775
            lblEEProdLine = glbLEE_ProdLine
        Else
            lblEEProdLine = ""
        End If
    End If
End If

If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
'lblEEID = glbLEE_ID
If glbLinHS Then
    lblEENum.Caption = glbDiv
Else
    lblEENum = ShowEmpnbr(lblEEID)
End If

End Sub

Function GetJobDesc(xCode)
Dim SQLQ As String
Dim xDesc As String
Dim dynaJobHIS As New ADODB.Recordset
    xDesc = ""
    If Len(xCode) > 0 Then
        SQLQ = "SELECT JB_CODE,JB_DESCR FROM HRJOB WHERE JB_CODE = '" & xCode & "' "
        If dynaJobHIS.State <> 0 Then dynaJobHIS.Close
        dynaJobHIS.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not dynaJobHIS.EOF Then
            xDesc = dynaJobHIS("JB_DESCR")
        End If
        dynaJobHIS.Close
    End If
    GetJobDesc = xDesc
End Function

Private Sub EmailShow()
Dim rsTMail As New ADODB.Recordset
Dim xReportByEmail As String
Dim SQLQ As String

On Error GoTo Email_Err
    If gsEMAIL_SENDING Then
        xReportByEmail = ""
        SQLQ = "SELECT ED_EMPNBR, ED_EMAIL FROM HREMP WHERE ED_EMPNBR = " & elpReptAuthShow(0).Text & " "
        rsTMail.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTMail.EOF Then
            If Not IsNull(rsTMail("ED_EMAIL")) Then
                If Len((rsTMail("ED_EMAIL"))) Then
                    xReportByEmail = rsTMail("ED_EMAIL")
                End If
            End If
        End If
        rsTMail.Close
        If Len(xReportByEmail) > 0 Then
            frmSendEmail.txtTo.Text = xReportByEmail
            frmSendEmail.Tag = ""
            frmSendEmail.Show 1
        Else
            MsgBox "Reported to/by Email Address is blank."
        End If
    End If
    
    Exit Sub
    
Email_Err:
    If Err.Number = 364 Then
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail", "SMTP", "SENDEMAIL")
    Resume Next
    
End Sub

Private Sub Change_Code_Table_Description()
Dim rsHRTabDesc As New ADODB.Recordset
Dim SQLQ As String

    SQLQ = "SELECT * FROM HRTABDES WHERE (TD_NAME = 'ECCL')"
    rsHRTabDesc.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    If Not rsHRTabDesc.EOF Then
        If rsHRTabDesc("TD_DESC") = "OH&S CLASSIFICATION CODES" Then
            rsHRTabDesc("TD_DESC") = "OH&S TYPE OF EVENT CODES"
            rsHRTabDesc.Update
        End If
    End If
    rsHRTabDesc.Close
End Sub

Private Sub Retrieve_Incidents()
    Dim rsHROccHealth As New ADODB.Recordset
    Dim SQLQ As String
    
    If glbtermopen Then
        SQLQ = "SELECT EC_EMPNBR, EC_CASE from Term_HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        If Not fglbNew And Not Data1.Recordset.EOF Then
            SQLQ = SQLQ & " AND EC_CASE NOT IN (" & Data1.Recordset!EC_CASE & ")"
        End If
        SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
        rsHROccHealth.Open SQLQ, gdbAdoIhr001X, adOpenDynamic, adLockOptimistic
    Else
        SQLQ = "SELECT EC_EMPNBR, EC_CASE from HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & " WHERE EC_EMPNBR = " & glbLEE_ID
        If Not fglbNew And Not Data1.Recordset.EOF Then
            SQLQ = SQLQ & " AND EC_CASE NOT IN (" & Data1.Recordset!EC_CASE & ")"
        End If
        SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
        rsHROccHealth.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    End If

    comRelIncident.Clear
    Do While Not rsHROccHealth.EOF
      comRelIncident.AddItem rsHROccHealth("EC_CASE")
      rsHROccHealth.MoveNext
    Loop
    rsHROccHealth.Close
    
End Sub

Function IfIncidentNo(InciNo As Double)
    Dim rsHROccHealth As New ADODB.Recordset
    Dim SQLQ As String
    
    If glbtermopen Then
        SQLQ = "SELECT EC_EMPNBR, EC_CASE, EC_OCCDATE from Term_HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " AND EC_CASE = " & comRelIncident.Text
        If Not fglbNew Then
            SQLQ = SQLQ & " AND EC_CASE NOT IN (" & txtIncidentNo.Text & ")"
        End If
        SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
        rsHROccHealth.Open SQLQ, gdbAdoIhr001X, adOpenDynamic, adLockOptimistic
    Else
        SQLQ = "SELECT EC_EMPNBR, EC_CASE, EC_OCCDATE from HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & " WHERE EC_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND EC_CASE = " & comRelIncident.Text
        If Not fglbNew Then
            SQLQ = SQLQ & " AND EC_CASE NOT IN (" & txtIncidentNo.Text & ")"
        End If
        SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
        rsHROccHealth.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    End If
    
    IfIncidentNo = False
    
    If rsHROccHealth.EOF Then
        rsHROccHealth.Close
        Exit Function
    End If
    rsHROccHealth.Close
    
    IfIncidentNo = True

End Function

Private Sub Populate_Employee_ClaimDate()
    Dim rsHS As New ADODB.Recordset
    Dim SQLQ As String
    
    'comDateClaimNo.Clear
    
    SQLQ = "SELECT EC_WCBNBR, EC_WCBFDTE FROM HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & "WHERE EC_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND EC_WCBNBR <> '' "
    rsHS.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsHS.EOF
        comDateClaimNo.AddItem rsHS("EC_WCBNBR") & " - " & rsHS("EC_WCBFDTE")
        rsHS.MoveNext
    Loop
    rsHS.Close
    Set rsHS = Nothing
End Sub

Private Sub CR_JobHis_Snap()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String
Dim dynaJobHIS As New ADODB.Recordset

On Error GoTo JobHis_Err

fglbJobList = ""
Screen.MousePointer = HOURGLASS

If glbtermopen Then
    SQLQ = "SELECT * FROM Term_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
    'Ticket #23630 - Showing all the positions the employee ever had
    'SQLQ = SQLQ & " AND JH_CURRENT <> 0"
    SQLQ = SQLQ & " ORDER BY JH_CURRENT " & IIf(glbSQL, "DESC", "") & ",JH_SDATE DESC"

    If dynaJobHIS.State <> 0 Then dynaJobHIS.Close
    dynaJobHIS.Open SQLQ, gdbAdoIhr001X, adOpenStatic
Else
    SQLQ = "SELECT * FROM HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE JH_EMPNBR=" & glbLEE_ID
    'Ticket #23630 - Showing all the positions the employee ever had
    'SQLQ = SQLQ & " AND JH_CURRENT <> 0"
    SQLQ = SQLQ & " ORDER BY JH_CURRENT " & IIf(glbSQL, "DESC", "") & ",JH_SDATE DESC"

    If dynaJobHIS.State <> 0 Then dynaJobHIS.Close
    dynaJobHIS.Open SQLQ, gdbAdoIhr001, adOpenStatic
End If
If Not dynaJobHIS.EOF Then
    Do Until dynaJobHIS.EOF
        If Not IsNull(dynaJobHIS!JH_JOB) Then
            fglbJobList = fglbJobList & dynaJobHIS!JH_JOB & ","
        End If
        dynaJobHIS.MoveNext
    Loop
    If Right(fglbJobList, 1) = "," Then
        fglbJobList = Left(fglbJobList, Len(fglbJobList) - 1)
    End If
    dynaJobHIS.MoveFirst
End If
Screen.MousePointer = DEFAULT

Exit Sub

JobHis_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Current Jobs", "HR_JOB_HISTORY", "SELECT")
Screen.MousePointer = DEFAULT
Resume Next

End Sub

Private Sub Populate_Job_Start_Date()
    Dim SQLQ As String
    Dim rsEmpJob As New ADODB.Recordset
    Dim X As Integer
    
    If glbtermopen Then
        SQLQ = "SELECT JH_JOB,JH_SDATE FROM Term_JOB_HISTORY "
        SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
        SQLQ = SQLQ & " AND JH_EMPNBR=" & glbLEE_ID
        SQLQ = SQLQ & " AND JH_JOB= '" & clpJob.Text & "'"
        'Ticket #23630 - Show all the position's Start Date the employee ever had
        'SQLQ = SQLQ & " AND JH_CURRENT <> 0"
        SQLQ = SQLQ & " ORDER BY JH_CURRENT " & IIf(glbSQL, "DESC", "") & ",JH_SDATE DESC"
    
        If rsEmpJob.State <> 0 Then rsEmpJob.Close
        rsEmpJob.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Else
        SQLQ = "SELECT JH_JOB,JH_SDATE FROM HR_JOB_HISTORY "
        SQLQ = SQLQ & " WHERE JH_EMPNBR=" & glbLEE_ID
        SQLQ = SQLQ & " AND JH_JOB= '" & clpJob.Text & "'"
        'Ticket #23630 - Show all the position's Start Date the employee ever had
        'SQLQ = SQLQ & " AND JH_CURRENT <> 0"
        SQLQ = SQLQ & " ORDER BY JH_CURRENT " & IIf(glbSQL, "DESC", "") & ",JH_SDATE DESC"
    
        If rsEmpJob.State <> 0 Then rsEmpJob.Close
        rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
    End If
    
    'Populate the Position Start Date combo box
    If Not rsEmpJob.EOF Then
        cmbJBStartDate.Clear
        cmbJBStartDate.Visible = True
        dlpDate(7).Visible = False
        
        Do While Not rsEmpJob.EOF
            cmbJBStartDate.AddItem rsEmpJob("JH_SDATE")
            
            rsEmpJob.MoveNext
        Loop
    Else
        cmbJBStartDate.Clear
        cmbJBStartDate.Visible = False
        dlpDate(7).Visible = True
        dlpDate(7).Left = cmbJBStartDate.Left
    End If
    rsEmpJob.Close
    Set rsEmpJob = Nothing
    
    If txtJobStartDate.Text <> "" Then
        If cmbJBStartDate.Visible = True Then
            For X = 0 To cmbJBStartDate.ListCount - 1
                If cmbJBStartDate.List(X) = txtJobStartDate.Text Then
                    cmbJBStartDate.ListIndex = X
                    Exit For
                End If
            Next X
            If txtJobStartDate = "" Then
                cmbJBStartDate.ListIndex = -1
            End If
        End If
        dlpDate(7).Text = txtJobStartDate.Text
    End If
End Sub
