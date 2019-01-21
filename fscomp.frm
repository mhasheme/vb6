VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmComp 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Company Master"
   ClientHeight    =   8040
   ClientLeft      =   4395
   ClientTop       =   4080
   ClientWidth     =   11400
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
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8040
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtOptMaxEE 
      Appearance      =   0  'Flat
      DataField       =   "PC_OPT"
      DataSource      =   "Data1"
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
      Left            =   8640
      MaxLength       =   5
      TabIndex        =   42
      Tag             =   "11-Enter Max. Number of Employees for License"
      Top             =   4182
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.TextBox txtNextAvPos 
      Appearance      =   0  'Flat
      DataSource      =   "Data1"
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
      Left            =   4155
      MaxLength       =   9
      TabIndex        =   17
      Tag             =   "10-Next Available Position #"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpEDates 
      DataField       =   "PC_TDATE"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   5385
      TabIndex        =   3
      Tag             =   "41-To Date"
      Top             =   1290
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpEDates 
      DataField       =   "PC_FDATE"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   3840
      TabIndex        =   2
      Tag             =   "41-From Date"
      Top             =   1303
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin VB.TextBox txtFEDTAX 
      Appearance      =   0  'Flat
      DataField       =   "PC_FEDTAX"
      DataSource      =   "Data1"
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
      Left            =   4155
      MaxLength       =   10
      TabIndex        =   10
      Tag             =   "Federal Tax Exemption"
      Top             =   3824
      Width           =   1215
   End
   Begin VB.TextBox txtPROVTAX 
      Appearance      =   0  'Flat
      DataField       =   "PC_PROVTAX"
      DataSource      =   "Data1"
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
      Left            =   7800
      MaxLength       =   10
      TabIndex        =   11
      Tag             =   "Provincial Tax Exemption"
      Top             =   3824
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   5520
      Top             =   7560
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   61
      Top             =   7380
      Width           =   11400
      _Version        =   65536
      _ExtentX        =   20108
      _ExtentY        =   1164
      _StockProps     =   15
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
      Begin VB.CommandButton cmdESSToTracker 
         Appearance      =   0  'Flat
         Caption         =   "ESS Data To Tracker"
         Height          =   375
         Left            =   8400
         TabIndex        =   66
         Tag             =   "Recalculate for all employees"
         Top             =   120
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton CmdPointRecalc 
         Appearance      =   0  'Flat
         Caption         =   "&Points Recalculate"
         Height          =   375
         Left            =   1920
         TabIndex        =   65
         Tag             =   "Recalculate for all employees"
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton CmdRecalc 
         Appearance      =   0  'Flat
         Caption         =   "&Recalculate"
         Height          =   375
         Left            =   240
         TabIndex        =   62
         Tag             =   "Recalculate for all employees"
         Top             =   120
         Width           =   1335
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   4800
         Top             =   120
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
   End
   Begin VB.Frame fraEntDate 
      Caption         =   "Outstanding Based Upon"
      ForeColor       =   &H00FF0000&
      Height          =   1725
      Left            =   7800
      TabIndex        =   54
      Top             =   840
      Width           =   2325
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1 - Entitlements Date"
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
         Index           =   2
         Left            =   150
         TabIndex        =   60
         Top             =   270
         Width           =   1800
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "2 - Original Date of Hire"
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
         Index           =   3
         Left            =   150
         TabIndex        =   59
         Top             =   480
         Width           =   2040
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "3 - Seniority Date"
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
         Index           =   4
         Left            =   150
         TabIndex        =   58
         Top             =   690
         Width           =   1500
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "4 - Last Hire Date"
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
         Index           =   5
         Left            =   150
         TabIndex        =   57
         Top             =   900
         Width           =   1530
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "5 - User Defined Date"
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
         Left            =   150
         TabIndex        =   56
         Top             =   1110
         Width           =   1875
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "6 - Union Date"
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
         Left            =   150
         TabIndex        =   55
         Top             =   1320
         Width           =   1260
      End
   End
   Begin VB.ComboBox cmbCountry 
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
      Left            =   4155
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Tag             =   "40-Country Code"
      Top             =   4167
      Width           =   1455
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PC_LUSER"
      DataSource      =   "Data1"
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
      Left            =   7860
      MaxLength       =   25
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1590
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PC_LTIME"
      DataSource      =   "Data1"
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
      Left            =   7860
      MaxLength       =   25
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1305
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PC_LDATE"
      DataSource      =   "Data1"
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
      Index           =   0
      Left            =   7860
      MaxLength       =   25
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   990
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox txtCompany 
      Appearance      =   0  'Flat
      DataField       =   "PC_CO"
      DataSource      =   "Data1"
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
      Left            =   7800
      MaxLength       =   3
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txtMaxEE 
      Appearance      =   0  'Flat
      DataSource      =   "Data1"
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
      Left            =   7800
      MaxLength       =   5
      TabIndex        =   41
      Tag             =   "11-Enter Max. Number of Employees for License"
      Top             =   4182
      Width           =   750
   End
   Begin VB.TextBox txtSysGenEmpl 
      Appearance      =   0  'Flat
      DataField       =   "PC_SYSTEM_EMPLOYEE"
      DataSource      =   "Data1"
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
      Left            =   5055
      MaxLength       =   3
      TabIndex        =   18
      Top             =   4898
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtNextAvEmpl 
      Appearance      =   0  'Flat
      DataField       =   "PC_NEXT_AVAILABLE_NBR"
      DataSource      =   "Data1"
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
      Left            =   7800
      MaxLength       =   9
      TabIndex        =   15
      Tag             =   "10-Next Available Employee #"
      Top             =   4898
      Width           =   1215
   End
   Begin VB.ComboBox cmbSysGenEmpl 
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
      Left            =   4140
      TabIndex        =   14
      Tag             =   "System Generated Employee #  Yes/No"
      Text            =   "cmbSysGenEmpl"
      Top             =   4883
      Width           =   975
   End
   Begin VB.ComboBox cmbPrecision 
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
      Left            =   4155
      TabIndex        =   9
      Tag             =   "Enter Number of Decimal Places for Salary"
      Text            =   "cmbPrecision"
      Top             =   3451
      Width           =   870
   End
   Begin VB.ComboBox cmbEntBase 
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
      Left            =   4140
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Tag             =   "Base for Entitlements Mass Update"
      Top             =   3078
      Width           =   2280
   End
   Begin VB.ComboBox cmbMonAnn 
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
      Left            =   4155
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "Sick Time Entitlements Earned A/M"
      Top             =   2705
      Width           =   2055
   End
   Begin VB.TextBox txtDateUseS 
      Appearance      =   0  'Flat
      DataField       =   "PC_ENTOUTS"
      DataSource      =   "Data1"
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
      Left            =   4170
      MaxLength       =   1
      TabIndex        =   6
      Tag             =   "11-Enter number 1-6"
      Top             =   2362
      Width           =   330
   End
   Begin VB.ComboBox cmbMonAnn 
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
      Left            =   4170
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "Vacation Entitlements Earned A/M"
      Top             =   1989
      Width           =   1995
   End
   Begin VB.TextBox txtDateUsed 
      Appearance      =   0  'Flat
      DataField       =   "PC_ENTOUT"
      DataSource      =   "Data1"
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
      Left            =   4170
      MaxLength       =   1
      TabIndex        =   4
      Tag             =   "11-Enter number 1-6"
      Top             =   1646
      Width           =   330
   End
   Begin VB.TextBox txtSerialNo 
      Appearance      =   0  'Flat
      DataField       =   "PC_SERIAL"
      DataSource      =   "Data1"
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
      Left            =   4155
      MaxLength       =   11
      TabIndex        =   1
      Tag             =   "01-Enter Serial Number"
      Top             =   960
      Width           =   1215
   End
   Begin Threed.SSPanel panCompName 
      Height          =   750
      Left            =   1530
      TabIndex        =   19
      Top             =   120
      Width           =   9720
      _Version        =   65536
      _ExtentX        =   17145
      _ExtentY        =   1323
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
      BevelWidth      =   2
      BevelOuter      =   0
      Font3D          =   1
      Alignment       =   1
      Begin VB.TextBox txtComp 
         Appearance      =   0  'Flat
         DataField       =   "PC_NAME"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   0
         MaxLength       =   100
         TabIndex        =   0
         Tag             =   "01-Enter Company Name"
         Top             =   240
         Width           =   9465
      End
   End
   Begin MSMask.MaskEdBox medEntBase 
      DataField       =   "PC_WDATE"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   6420
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3093
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
      BackColor       =   12632256
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox SavWDate 
      Height          =   285
      Left            =   6900
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3093
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12632256
      Enabled         =   0   'False
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
   Begin VB.Label lblNextAvPos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Next Available Position #"
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
      Left            =   0
      TabIndex        =   68
      Top             =   5685
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label lblMessage 
      Height          =   615
      Left            =   9120
      TabIndex        =   67
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Provincial Tax Exemption"
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
      Index           =   13
      Left            =   5820
      TabIndex        =   64
      Top             =   3869
      Width           =   1905
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Federal Tax Exemption"
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
      Index           =   12
      Left            =   0
      TabIndex        =   63
      Top             =   3869
      Width           =   1620
   End
   Begin VB.Label lblCountry 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      DataField       =   "PC_COUNTRY"
      DataSource      =   "Data1"
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
      Left            =   3000
      TabIndex        =   53
      Top             =   4227
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblCntryCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Country Code"
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
      Left            =   0
      TabIndex        =   52
      Top             =   4227
      Width           =   960
   End
   Begin VB.Label lblCompName 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      Height          =   195
      Left            =   0
      TabIndex        =   51
      Top             =   435
      Width           =   1320
   End
   Begin VB.Label SavDateUses 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   4920
      TabIndex        =   50
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label SavDateUsed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   5040
      TabIndex        =   49
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMonAnn 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "PC_VACENT"
      DataSource      =   "Data1"
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
      Left            =   8640
      TabIndex        =   48
      Top             =   1800
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblMonAnn 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "PC_SICKENT"
      DataSource      =   "Data1"
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
      Left            =   8640
      TabIndex        =   47
      Top             =   2280
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employees for License"
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
      Index           =   8
      Left            =   6000
      TabIndex        =   38
      Top             =   4227
      Width           =   1695
   End
   Begin VB.Label lblPrecision 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      DataField       =   "PC_DECHR"
      DataSource      =   "Data1"
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
      Left            =   5640
      TabIndex        =   37
      Top             =   3511
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label SAVTDATES 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TDateS"
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
      Left            =   5655
      TabIndex        =   36
      Top             =   2407
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label SavTdate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TDate"
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
      Left            =   5655
      TabIndex        =   35
      Top             =   1691
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label SAVFDATES 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FDateS"
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
      Left            =   4575
      TabIndex        =   34
      Top             =   2407
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label SavFdate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FDate"
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
      Left            =   4575
      TabIndex        =   33
      Top             =   1691
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblLvlNbr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "PC_LVLNBR"
      DataSource      =   "Data1"
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
      Left            =   3000
      TabIndex        =   32
      Top             =   4555
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblCountDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "PC_WHEN_COUNTED"
      DataSource      =   "Data1"
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4155
      TabIndex        =   16
      Top             =   5265
      Width           =   2640
   End
   Begin VB.Label lblNumberEE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      DataField       =   "PC_NUMBER_EMPLOYEES"
      DataSource      =   "Data1"
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4155
      TabIndex        =   13
      Tag             =   "40-Number of Employees on file"
      Top             =   4540
      Width           =   600
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Employee Created on"
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
      Index           =   9
      Left            =   0
      TabIndex        =   31
      Top             =   5310
      Width           =   2235
   End
   Begin VB.Label lblNextAvEmpl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Next Available Employee #"
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
      Left            =   5820
      TabIndex        =   30
      Top             =   4943
      Width           =   1905
   End
   Begin VB.Label lblGenEmpl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "System Generated Employee #"
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
      Left            =   0
      TabIndex        =   29
      Top             =   4943
      Width           =   2895
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employees on File"
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
      Index           =   7
      Left            =   0
      TabIndex        =   28
      Top             =   4585
      Width           =   3075
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Salary Decimal Precision"
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
      Index           =   6
      Left            =   0
      TabIndex        =   27
      Top             =   3511
      Width           =   2115
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vacation / Sick Mass Update Based Upon"
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
      Left            =   0
      TabIndex        =   26
      Top             =   3153
      Width           =   3225
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sick Time Earned"
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
      Index           =   4
      Left            =   0
      TabIndex        =   25
      Top             =   2795
      Width           =   2520
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sick Time Outstanding Based Upon"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   0
      TabIndex        =   24
      Top             =   2430
      Width           =   3030
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vacation Earned"
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
      Left            =   0
      TabIndex        =   23
      Top             =   2079
      Width           =   2430
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vacation Outstanding Based Upon"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   22
      Top             =   1721
      Width           =   4035
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fiscal Year Date Range"
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
      Left            =   0
      TabIndex        =   21
      Top             =   1363
      Width           =   3045
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Serial Number"
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
      Left            =   0
      TabIndex        =   20
      Top             =   1005
      Width           =   1920
   End
End
Attribute VB_Name = "frmComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LProvs() As Variant
Dim fglbEmptyNew As Integer
Dim EntDteChange As Integer
Dim fglbNew As Boolean
Dim OldNextAvEmpl 'George for Granite Club System Generated Employee's number
Dim SavVacEntEarned

Private Function chkComp()
Dim x%, Response As Integer, Msg As String

lblCountry = cmbCountry '07June99 js
chkComp = False
If Len(txtSerialNo) < 1 Then
    MsgBox "Serial # is a required field"
    txtSerialNo.SetFocus
    Exit Function
End If

For x% = 0 To 1
    If Len(dlpEDates(x%).Text) < 1 Then
        MsgBox "Fiscal Year Date Range is a required field"
        dlpEDates(x%).SetFocus
        Exit Function
    End If
    If Not IsDate(dlpEDates(x%).Text) Then
        MsgBox "Invalid Fiscal Year Date"
        dlpEDates(x%).SetFocus
        Exit Function
    End If
Next x%

If CVDate(dlpEDates(0).Text) >= CVDate(dlpEDates(1).Text) Then
    MsgBox "Fiscal Year From Date Can Not Be Greater Than To Date"
    dlpEDates(0).SetFocus
    Exit Function
End If

If Not IsNumeric(txtDateUsed.Text) Then
    MsgBox "You Must Enter 1 - 6"
    txtDateUsed.SetFocus
    Exit Function
End If

If txtDateUsed < 1 Or txtDateUsed > 6 Then
    MsgBox "You Must Enter 1 - 6"
    txtDateUsed.SetFocus
    Exit Function
End If

'Ticket #29230 - Daily Vacation Entitlement - glbCompEntVacDaily is only used temporary setting during the development
'If glbCompEntVacDaily And cmbMonAnn(0).ListIndex = 3 And txtDateUsed <> "1" Then
If cmbMonAnn(0).ListIndex = 3 And txtDateUsed <> "1" Then
    MsgBox "Vacation Outstanding must be based on 1 - Entitlements Date"
    txtDateUsed.SetFocus
    Exit Function
End If

'For X% = 0 To 1
'    If Len(dlpEDatesS(X%).Text) < 1 Then
'        MsgBox "Sick Time Entitlement Date Range is a required field"
'        dlpEDatesS(X%).SetFocus
'        Exit Function
'    End If
'    If Not IsDate(dlpEDatesS(X%).Text) Then
'        MsgBox "Invalid Sick Time Entitlement Date"
'        dlpEDatesS(X%).SetFocus
'        Exit Function
'    End If
'Next X%
'
'If CVDate(dlpEDatesS(0).Text) >= CVDate(dlpEDatesS(1).Text) Then
'    MsgBox "Sick Time Entitlement From Date Can Not Be Greater Than To Date"
'    dlpEDatesS(0).SetFocus
'    Exit Function
'End If

If Not IsNumeric(txtDateUseS.Text) Then
    MsgBox "You Must Enter 1 - 6"
    txtDateUseS.SetFocus
    Exit Function
End If
If txtDateUseS < 1 Or txtDateUseS > 6 Then
    MsgBox "You Must Enter 1 - 6"
    txtDateUseS.SetFocus
    Exit Function
End If
If Len(txtFedTax) > 0 And Not IsNumeric(txtFedTax.Text) Then
    MsgBox "Federal Tax Exemption Must be Numeric"
    txtFedTax.SetFocus
    Exit Function
End If
If Len(txtPROVTAX) > 0 And Not IsNumeric(txtPROVTAX.Text) Then
    MsgBox "Provincial Tax Exemption Must be Numeric"
    txtPROVTAX.SetFocus
    Exit Function
End If

If txtMaxEE < 100 Then
    lblLvlNbr.Caption = "1"
Else
    If txtMaxEE < 250 Then
        lblLvlNbr.Caption = "2"
    Else
        If txtMaxEE < 800 Then
            lblLvlNbr.Caption = "3"
        Else
            lblLvlNbr.Caption = "4"
        End If
    End If
End If

chkComp = True

End Function

Private Sub cmbCountry_Change()
    Call SetPanHelp(ActiveControl) '07June99 js
End Sub

Private Sub cmbCountry_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbEntBase_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbEntBase_LostFocus()

Select Case cmbEntBase.ListIndex
    Case 0: medEntBase = "O"
    Case 1: medEntBase = "S"
    Case 2: medEntBase = "L"
    Case 3: medEntBase = "U"
    Case Else: medEntBase = "D"
End Select

End Sub

Private Sub cmbMonAnn_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbMonAnn_LostFocus(Index As Integer)
    If cmbMonAnn(Index).ListIndex = 0 Then
        lblMonAnn(Index).Caption = "M"
    ElseIf cmbMonAnn(Index).ListIndex = 1 Then
        lblMonAnn(Index).Caption = "A"
    Else
        'Ticket #29230 - Daily Vacation Entitlement - glbCompEntVacDaily is only used temporary setting during the development
        'If glbCompEntVacDaily And Index = 0 And cmbMonAnn(Index).ListIndex = 3 Then
        If Index = 0 And cmbMonAnn(Index).ListIndex = 3 Then
            lblMonAnn(Index).Caption = "D"
        Else
            lblMonAnn(Index).Caption = "N"
        End If
    End If
End Sub

Private Sub cmbPrecision_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbPrecision_LostFocus()
    lblPrecision.Caption = cmbPrecision.Text
End Sub

Private Sub cmbSysGenEmpl_Click()

    txtNextAvEmpl.Visible = cmbSysGenEmpl = "Yes"   'jaddy 10/14/99
    'Granite Club and City of Timmins (Ticket #10340) and Ticket #14968 - Listowel Tech.
    If glbCompSerial = "S/N - 2241W" Or glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2350W" Then
        txtNextAvEmpl.Enabled = True    'George for Granite Club System Generated Employee's number
    End If
    lblNextAvEmpl.Visible = cmbSysGenEmpl = "Yes"   'jaddy 10/14/99

End Sub

Private Sub cmbSysGenEmpl_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub cmdCancel_Click()

On Error GoTo Can_Err

Data1.Recordset.CancelUpdate
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

If txtSysGenEmpl = 0 Then cmbSysGenEmpl = "No"
If txtSysGenEmpl = -1 Then cmbSysGenEmpl = "Yes"

Call ST_UPD_MODE(True)
CmdRecalc.Enabled = True

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRPARCO", "Cancel")
Resume Next

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
Private Sub GetCaption()
SavDateUsed.Caption = Data1.Recordset.Fields("pc_entout")
SavFdate.Caption = Data1.Recordset.Fields("pc_fdate")
SavTdate.Caption = Data1.Recordset.Fields("pc_tdate")
SavDateUses.Caption = Data1.Recordset.Fields("pc_entoutS")
SAVFDATES.Caption = Data1.Recordset.Fields("pc_fdateS")
SAVTDATES.Caption = Data1.Recordset.Fields("pc_tdateS")
SavWDate = Data1.Recordset.Fields("pc_wdate")

'Ticket #29230 - Daily Vacation Entitlement - glbCompEntVacDaily is only used temporary setting during the development
'If glbCompEntVacDaily And lblMonAnn(0) = "D" Then
If lblMonAnn(0) = "D" Then
    SavVacEntEarned = Data1.Recordset.Fields("PC_VACENT")
End If

End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Call SET_UP_MODE
'Call ST_UPD_MODE(True)

txtComp.SetFocus
SavDateUsed.Caption = Data1.Recordset.Fields("pc_entout")
SavFdate.Caption = Data1.Recordset.Fields("pc_fdate")
SavTdate.Caption = Data1.Recordset.Fields("pc_tdate")
SavDateUses.Caption = Data1.Recordset.Fields("pc_entoutS")
SAVFDATES.Caption = Data1.Recordset.Fields("pc_fdateS")
SAVTDATES.Caption = Data1.Recordset.Fields("pc_tdateS")
SavWDate = Data1.Recordset.Fields("pc_wdate")

'Ticket #29230 - Daily Vacation Entitlement - glbCompEntVacDaily is only used temporary setting during the development
'If glbCompEntVacDaily And lblMonAnn(0) = "D" Then
If lblMonAnn(0) = "D" Then
    SavVacEntEarned = Data1.Recordset.Fields("PC_VACENT")
End If

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '08June99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()

Dim DtTm As Variant, rc As Integer, x%, Msg, Response
DtTm = Now

Call GetCaption

DoEvents

If cmbSysGenEmpl = "Yes" Then txtSysGenEmpl = -1    '28 Nov, 1997 laura
If cmbSysGenEmpl = "No" Then txtSysGenEmpl = 0

If Not chkComp() Then Exit Sub

EntDteChange = 0

If txtDateUsed.Text <> SavDateUsed Or txtDateUseS.Text <> SavDateUses Then
    EntDteChange = 2
End If

'If dlpEDates(0).Text <> SavFdate Or dlpEDates(1).Text <> SavTdate Then
'    EntDteChange = EntDteChange + 1
'End If

If EntDteChange > 0 Then
    Msg = "Change in " & Chr(34)
    If EntDteChange > 1 Then
        Msg = Msg & "Outstanding Based Upon "
        If EntDteChange > 2 Then Msg = Msg & "& "
    End If
    'If EntDteChange = 1 Or EntDteChange = 3 Then
    '    Msg = Msg & "Fiscal Year Date Range "
    'End If
    
    Msg = Msg & Chr(34) & " will affect employees outstanding "
    Msg = Msg & "sick and vacation."
    Msg = Msg & Chr(10) & "Do you wish to proceed and recalculate the "
    Msg = Msg & "Employee's Vacation / Sick Outstanding ?"
    
    Response = MsgBox(Msg, 52, "Warning")
    If Response = IDNO Then GoTo ExitCmdOK
    
    glbEntOutStanding$ = txtDateUsed.Text
    glbCompEdFrom = dlpEDates(0).Text
    glbCompEdTo = dlpEDates(1).Text
    glbEntOutStandingS$ = txtDateUseS.Text
    'glbCompEdFromS = dlpEDatesS(0).Text
    'glbCompEdToS = dlpEDatesS(1).Text
    glbCountry = lblCountry  '(VB3 js -2/9/99)
'    glbMultiGrid = chkMultiGrid
    Screen.MousePointer = HOURGLASS
    
    Call EntReCalc("")
End If

If medEntBase <> SavWDate Then
    Msg = "Change in " & Chr(34)
    Msg = Msg & "Vacation / Sick Mass Update Based Upon "
    Msg = Msg & Chr(34) & " will affect employees outstanding "
    Msg = Msg & "Hourly entitlements."
    Msg = Msg & Chr(10) & "Do you wish to proceed and recalculate the "
    Msg = Msg & "Employee's outstanding Hourly entitlement ?"
    Response = MsgBox(Msg, 52, "Warning")
    If Response = IDNO Then
        medEntBase = SavWDate
        GoTo ExitCmdOK
    End If
    glbCompWDate = medEntBase.Text
    Screen.MousePointer = HOURGLASS
    
    Call EntReCalcHr
End If

Call UpdUStats(Me) ' update user's stats (who did it and when)

'call updDataFld(txtFedTax, "N", Data1, "PC_FEDTAX")
'call updDataFld(txtPROVTAX, "N", Data1, "PC_PROVTAX")

If IsNull(txtFedTax) Or Len(txtFedTax) = 0 Then txtFedTax = 0
If IsNull(txtPROVTAX) Or Len(txtPROVTAX) = 0 Then txtPROVTAX = 0

'Data1.Recordset("PC_MULTIGRID") = chkMultiGrid

'Release 8.1 - The Max Employee License is encrypted now.
'Encrypt the Max Employee and save into the database
If txtMaxEE.Enabled Then
    If Not Data1.Recordset.EOF Then
        'Encrypt it first and then assign to table field
        'Encrypted value is stored in the different field.
        If gsMultiLang = "Y" Then 'For Listowel only
            Data1.Recordset.Fields("PC_OPT") = EncryptPasswordMultiLang_First(txtMaxEE.Text)
        ElseIf UCase(gsMultiLang) = "YES" Then 'whscc
            Data1.Recordset.Fields("PC_OPT") = EncryptPasswordMultiLang(txtMaxEE.Text)
        Else
            Data1.Recordset.Fields("PC_OPT") = EncryptPassword(txtMaxEE.Text)
        End If
        Data1.Recordset.Fields("PC_OPTUPD") = True
    End If
End If

Data1.Recordset("PC_NUMBER_EMPLOYEES") = Val(lblNumberEE)

Data1.Recordset.UpdateBatch

If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
If Not glbSQL And Not glbOracle Then Call Pause(0.5)

x% = setCompInfo(glbCompNo)

ExitCmdOK:
Screen.MousePointer = DEFAULT

fglbNew = False

Call SET_UP_MODE

'Call ST_UPD_MODE(False)

CmdRecalc.Enabled = True   'laura jan 12, 1998

Data1.Refresh

Select Case medEntBase
    Case "O": x% = 0
    Case "S": x% = 1
    Case "L": x% = 2
    Case "U": x% = 3
    Case Else: x% = 4
End Select

cmbEntBase.ListIndex = x%

If gSec_Upd_Company Then           'May99 js
 '   cmdModify.Enabled = True        '
End If                              '

    
    ' dkostka - 01/15/00 - Was losing these globals somehow, easiest way to fix is just to
    '   reload them from the fields here.
    glbEntOutStanding$ = txtDateUsed.Text
    glbCompEdFrom = dlpEDates(0).Text
    glbCompEdTo = dlpEDates(1).Text
    glbEntOutStandingS$ = txtDateUseS.Text
    'glbCompEdFromS = dlpEDatesS(0).Text
    'glbCompEdToS = dlpEDatesS(1).Text
    glbCountry = lblCountry  '(VB3 js -2/9/99)
    
    'Hemu - Vacation Earned bases also was loosing it's value, reloaded it.
    glbCompEntVac$ = lblMonAnn(0).Caption
    glbCompEntSick$ = lblMonAnn(1).Caption
    
    
'Ticket #29230 - Daily Vacation Entitlement - glbCompEntVacDaily is only used temporary setting during the development
'If glbCompEntVacDaily And cmbMonAnn(0).ListIndex = 3 And txtDateUsed = "1" And SavVacEntEarned <> "D" Then
If cmbMonAnn(0).ListIndex = 3 And txtDateUsed = "1" And SavVacEntEarned <> "D" Then
    'MsgBox "Daily Vacation Accrual Tables must be setup to use this option. To setup the table, use the Daily Vacation Mass Update function.", vbInformation, "info:HR - Daily Vacation Accrual"
    MsgBox "You have selected the 'Daily Vacation Accrual' option. " & vbCrLf & vbCrLf & "Please call the Support line at HR Systems Strategies Inc. to set up the Daily Vacation Accrual Tables under Mass Updates menu \ Entitlements \ Daily Vacation Accrual Master screen.", vbInformation, "info:HR - Daily Vacation Accrual"
End If

Exit Sub

Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRPARCO", "Update")
Resume Next
Unload Me

End Sub

Private Sub cmdESSToTracker_Click()
Dim xEmpCnt As Integer
Dim xAttCnt As Integer
Dim Msg, a%
Dim xMsgEmp As String
Dim xMsgAtt As String


    Msg = "This function transfers Employee and Attendance ESS changes to Tracker" & Chr(10) & Chr(10)
    Msg = Msg & "Are You Sure You Want To Do This? "
    a% = MsgBox(Msg, 36, "Confirm ")
    If a% <> 6 Then Exit Sub
    Screen.MousePointer = HOURGLASS
    
    cmdESSToTracker.Enabled = False
    
    xEmpCnt = ESS_To_Tracker_EMP
    xAttCnt = ESS_To_Tracker_ATT
    
    If xEmpCnt = 0 Then
        xMsgEmp = "No employee data updated."
    Else
        xMsgEmp = xEmpCnt & " employee(s) data updated."
    End If
    If xAttCnt = 0 Then
        xMsgAtt = "No attendance updated."
    Else
        xMsgAtt = xAttCnt & " attendance(s) updated."
    End If
    
    Screen.MousePointer = DEFAULT
    
    MsgBox xMsgEmp & Chr(10) & xMsgAtt
    
    cmdESSToTracker.Enabled = True
End Sub

Private Sub CmdPointRecalc_Click()
Dim Msg, a%
If glbBurlTech Then 'BTI Points Recalculate
    Msg = "This function will check Unexcused, Emergency Leave Flag," & Chr(10)
    Msg = Msg & "also reset the Absence Point" & Chr(10)
    Msg = Msg & "Are You Sure You Want To Do This? "
    a% = MsgBox(Msg, 36, "Confirm ")
    If a% <> 6 Then Exit Sub
    Call BTIPoint("ALL", True)
End If

'Ticket #20076 Frank 04/01/2011
'Small program
If glbCompSerial = "S/N - 2259W" Then
    Msg = "This function will do these tasks:" & Chr(10)
    Msg = Msg & "   1. Fix the blank employee salary step (from previous salary record) " & Chr(10)
    Msg = Msg & "   2. update employee salary from info:HR to GP" & Chr(10) & Chr(10)
    Msg = Msg & "Are You Sure You Want To Do This? "
    a% = MsgBox(Msg, 36, "Confirm ")
    If a% <> 6 Then Exit Sub
    Call OxfordGPSalary
End If
End Sub

Private Sub OxfordGPSalary()
Dim rsHRSal As New ADODB.Recordset
Dim rsHRPos As New ADODB.Recordset
Dim rsHRTmp As New ADODB.Recordset
Dim SQLQ As String
Dim xEmpNo, xID
Dim xTot, I
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).FloodPercent = 0
      
    Screen.MousePointer = HOURGLASS
    CmdPointRecalc.Enabled = False
    
    SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT (SH_CURRENT = 0) ORDER BY SH_EMPNBR "
    rsHRSal.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRSal.EOF Then
        xTot = rsHRSal.RecordCount
    End If
    I = 0
    Do While Not rsHRSal.EOF
        MDIMain.panHelp(0).FloodPercent = (I / xTot) * 100: I = I + 1
        DoEvents
        xEmpNo = rsHRSal("SH_EMPNBR")
        xID = rsHRSal("SH_ID")
        'update blank step
        If IsNull(rsHRSal("SH_GRADE")) Then
            'get previous step
            SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & xEmpNo & " "
            SQLQ = SQLQ & "AND SH_JOB = '" & rsHRSal("SH_JOB") & "' "
            SQLQ = SQLQ & "AND SH_EDATE < " & Date_SQL(rsHRSal("SH_EDATE")) & " "
            SQLQ = SQLQ & "ORDER BY SH_EDATE DESC, SH_ID DESC "
            If rsHRTmp.State <> 0 Then rsHRTmp.Close
            rsHRTmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsHRTmp.EOF Then
                rsHRSal("SH_GRADE") = rsHRTmp("SH_GRADE")
                rsHRSal.Update
            End If
            rsHRTmp.Close
        End If
        'update the GP salary for default position
        'call GP Salary integration
        SQLQ = "SELECT JH_EMPNBR, JH_POSITION_CONTROL FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & xEmpNo & " AND NOT (JH_CURRENT = 0) "
        SQLQ = SQLQ & "AND JH_JOB = '" & rsHRSal("SH_JOB") & "' "
        SQLQ = SQLQ & "AND JH_POSITION_CONTROL = 'YES' "
        If rsHRPos.State <> 0 Then rsHRPos.Close
        rsHRPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsHRPos.EOF Then
            If glbGP Then
                Call Salary_Integration(xEmpNo, , False, False, xID)
            End If
        End If
        rsHRSal.MoveNext
    Loop
    rsHRSal.Close
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).FloodPercent = 0
    CmdPointRecalc.Enabled = True
    Screen.MousePointer = DEFAULT
    
    MsgBox "   Done!   "
End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdRecalc_Click()
Dim Msg, Response, DgDef
Dim rsHr As New ADODB.Recordset
Dim CYear, CYear1
Dim SQLQ, ReasonCode
Msg = "Do you wish to proceed and recalculate the "
Msg = Msg & "Employee's " & Chr(10)
Msg = Msg & "outstanding entitlements, emergency leave taken and Hours Per Day ? "
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2

Response = MsgBox(Msg, DgDef, "ReCalculate")

If Response = IDNO Then Exit Sub

Screen.MousePointer = HOURGLASS
If glbCompSerial = "S/N - 2350W" Then 'listowel
    glbEntExcept = 0
Else
    glbEntExcept = 1 'Don't calculate the Vacation Entitlement
End If
Call EntReCalc("")
Call EntReCalcHr    'laura jan 12, 1998
If glbCompSerial = "S/N - 2433W" Then 'Kerry's Place Ticket #23417 Franks 05/31/2013
    Call EntReCalcHrTerm
End If

'for Vacation/Sick Entitlement FollowUp'Zahoor(Sam) Butt 02/10/2006
Screen.MousePointer = HOURGLASS
Call VacSickHourlyFollowUp("(1 = 1)")

glbflgFU = False

If Not (glbCompSerial = "S/N - 2241W") Then ' Not Granite Club
    Call CR_EMPLOYEE_SNAP 'laura 02/24/98
End If

If glbCompSerial = "S/N - 2373W" Then 'District Muskoka
    Call SalaryTotal
End If

If glbCompSerial = "S/N - 2173W" Then 'Town of Ajax 'Ticket #30402 Franks 08/02/2017
    Call Recalculate_OTBANK_Ajax_AllEmployees
End If

Screen.MousePointer = DEFAULT

glbENTScreen = True

End Sub

Private Sub SalaryTotal()
Dim SQLQ As String, I, xTot
Dim snapSal As New ADODB.Recordset
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).FloodPercent = 0
    SQLQ = "SELECT * FROM HR_SALARY_HISTORY"
    snapSal.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    If Not snapSal.EOF Then
        xTot = snapSal.RecordCount
    End If
    I = 0
    Do While Not snapSal.EOF
        MDIMain.panHelp(0).FloodPercent = (I / xTot) * 100: I = I + 1
        If I > xTot Then I = xTot
        snapSal("SH_TOTAL") = snapSal("SH_SALARY") + snapSal("SH_PREMIUM")
        snapSal.Update
        snapSal.MoveNext
    Loop
    snapSal.Close
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).FloodPercent = 0
End Sub

Private Sub CR_EMPLOYEE_SNAP()
Dim SQLQ As String, countr As Integer, SQLQ2 As String
Dim Desc As String
Dim DMaxNum, DMaxNumX
Dim snapEmp As New ADODB.Recordset
Dim snapEmpX As New ADODB.Recordset

On Error GoTo Emp_Err

DMaxNum = 0
SQLQ = "Select MAX(ED_EMPNBR) AS MAXEMPNBR from HREMP"
If glbCompSerial = "S/N - 2351W" Then ' Burlington Technologies
    SQLQ = SQLQ & " WHERE ED_EMPNBR< 90000000"
End If
snapEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not IsNull(snapEmp("MAXEMPNBR")) Then DMaxNum = snapEmp("MAXEMPNBR")

DMaxNumX = 0
If glbCompSerial <> "S/N - 2151W" Then ' IAPA - Ticket #10958
    SQLQ = "SELECT MAX(ED_EMPNBR) AS MAXEMPNBR FROM Term_HREMP"
    If glbCompSerial = "S/N - 2351W" Then ' Burlington Technologies
        SQLQ = SQLQ & " WHERE ED_EMPNBR< 90000000"
    End If
    snapEmpX.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    If Not IsNull(snapEmpX("MAXEMPNBR")) Then DMaxNumX = snapEmpX("MAXEMPNBR")
End If

If DMaxNum < DMaxNumX Then DMaxNum = DMaxNumX Else DMaxNum = DMaxNum

'Data1.DatabaseName = glbIHRDB   'laura nov 28, 1997
Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "HRPARCO"
 
SQLQ2 = "Select HRPARCO.* from HRPARCO"

Data1.RecordSource = SQLQ2
Data1.Refresh

If Data1.Recordset("PC_NEXT_AVAILABLE_NBR") <= DMaxNum Then
    'Data1.Recordset.Edit
    Data1.Recordset("PC_NEXT_AVAILABLE_NBR") = DMaxNum + 1
    Data1.Recordset.UpdateBatch
    Data1.Refresh
ElseIf Data1.Recordset("PC_NEXT_AVAILABLE_NBR") > DMaxNum + 1 Then  'laura 03/03/98
    'Data1.Recordset.Edit
    Data1.Recordset("PC_NEXT_AVAILABLE_NBR") = DMaxNum + 1
    Data1.Recordset.UpdateBatch
    Data1.Refresh
End If

Exit Sub

Emp_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Employees", "HREMP", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Function Check_EMPLOYEE_Number(mEmpNum)
Dim SQLQ As String, countr As Integer, SQLQ2 As String
Dim Desc As String
Dim blnFirst As Boolean
Dim DMaxNum, DMaxNumX
Dim snapEmp As New ADODB.Recordset

On Error GoTo Emp_Err

blnFirst = True
DMaxNum = 0
SQLQ = "Select ED_EMPNBR,ED_SurName,ED_FName from qry_HREMP "
SQLQ = SQLQ & "where ED_EMPNBR >=" & mEmpNum
SQLQ = SQLQ & " order by ED_EMPNBR "
snapEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic

If snapEmp.EOF Then
    DMaxNum = mEmpNum
Else
Do While Not snapEmp.EOF
    If snapEmp("ED_EMPNBR") <> mEmpNum Then
        If snapEmp("ED_EMPNBR") > mEmpNum Then
            DMaxNum = mEmpNum
            Exit Do
        End If
    Else
        If blnFirst Then
            blnFirst = False
            Check_EMPLOYEE_Number = mEmpNum & " assigned to " & snapEmp("ED_SURNAME") & "," & snapEmp("ED_FNAME")
        End If
        mEmpNum = mEmpNum + 1
    
    End If
    snapEmp.MoveNext
Loop
End If
If DMaxNum = 0 Then DMaxNum = mEmpNum

'Data1.DatabaseName = glbIHRDB   'laura nov 28, 1997
Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "HRPARCO"
 
SQLQ2 = "Select HRPARCO.* from HRPARCO"

Data1.RecordSource = SQLQ2
Data1.Refresh

'Data1.Recordset.Edit
Data1.Recordset("PC_NEXT_AVAILABLE_NBR") = DMaxNum
Data1.Recordset.UpdateBatch
Data1.Refresh

Exit Function

Emp_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Employees", "HREMP", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRSTATS", "SELECT")

End Sub

Private Sub Form_Activate()
    glbOnTop = "FRMCOMP"
    Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMCOMP"
End Sub

Private Sub Form_Load()
Dim ctylist, x

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = "FRMCOMP"

' part of question is two fields in this table
' relaying client's desire for 2 or 3 decimal salary
'retentions... only presently modifying one.

On Error GoTo Ld_Err

Screen.MousePointer = HOURGLASS

frmComp.Show
cmbPrecision.AddItem "2"
cmbPrecision.AddItem "3"
cmbPrecision.AddItem "4"

cmbSysGenEmpl.AddItem "Yes"  '28 Nov, 1997 laura
cmbSysGenEmpl.AddItem "No"   '28 Nov, 1997 laura

If glbWFC Then 'Ticket #27827 Franks 12/02/2015
    Call WFCSrceenSetup
End If

If glbLinamar Then 'Ticket #29759 Franks 02/22/2017
    Call LinamarSrceenSetup
End If

'Call function to populate the dropdown list with Countries from MTF file
ctylist = CountryList
x = 1
Do While x > 0
    x = InStr(ctylist, "&")
    If x > 0 Then
        cmbCountry.AddItem Left(ctylist, x - 1)
        'cmbCountryOfEmp.AddItem Left(ctylist, x - 1)
        ctylist = Mid(ctylist, x + 1)
    Else
        cmbCountry.AddItem ctylist
        'cmbCountryOfEmp.AddItem ctylist
    End If
Loop

cmbCountry.ListIndex = 0

'cmbCountry.AddItem "CANADA"     '07June99 js
'cmbCountry.AddItem "U.S.A."     '
'cmbCountry.AddItem "BAHAMAS"    '
'cmbCountry.AddItem "GERMANY"    '
'cmbCountry.AddItem "SINGAPORE"  '
'cmbCountry.AddItem "ENGLAND"    '

'Data1.DatabaseName = glbIHRDB
Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "HRPARCO"
Data1.Refresh

If IsNull(glbCountry) Then                          '07June99 js
    glbCountry = "CANADA"
        'commented by Bryan Ticket#12579 Feb 7, 2007
    'glbcountry can be the users country, since this is a bound field
    'there's no need to overwrite teh company country'
'Else                                                '
'    lblCountry = glbCountry '(VB3 js-2/9/99)        '
End If                                              '

'''Select Case lblCountry                              '
'''        Case "CANADA": cmbCountry.ListIndex = 0     '
'''        Case "U.S.A.": cmbCountry.ListIndex = 1     '
'''        Case "BAHAMAS": cmbCountry.ListIndex = 2    '
'''        Case "GERMANY": cmbCountry.ListIndex = 3    '
'''        Case "SINGAPORE": cmbCountry.ListIndex = 4  '
'''        Case "ENGLAND": cmbCountry.ListIndex = 5    '
'''        Case Else: cmbCountry.ListIndex = 0         '
'''End Select
'Ticket #18739 06/28/2010 Frank, wrong Country showed up
cmbCountry.Text = lblCountry

If txtSysGenEmpl = 0 Then cmbSysGenEmpl = "No"    '28 Nov, 1997 laura
If txtSysGenEmpl = -1 Then cmbSysGenEmpl = "Yes"   '28 Nov, 1997 laura

Call ST_UPD_MODE(False)

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    frmComp.Show
    Data1.RecordSource = "HRPARCO"
    Data1.Recordset.AddNew
    Call ST_UPD_MODE(True)
'    txtComp.SetFocus
End If

If DemoSystem Then txtMaxEE = DemoMaxEmp%

Call modLoadLists

cmbPrecision.Text = lblPrecision.Caption

frmComp.Show

Call cmbSysGenEmpl_Click

If txtComp.Text = "" Then
    txtCompany = "001"
    txtSerialNo.Enabled = True
    txtMaxEE.Enabled = True
    txtComp.Enabled = True
End If

If glbBurlTech Then
    CmdPointRecalc.Visible = True
End If

Screen.MousePointer = DEFAULT

If Not gSec_Upd_Company Then       'May99 js
'    cmdModify.Enabled = False   '
End If                          '

Call INI_Controls(Me)

OldNextAvEmpl = txtNextAvEmpl 'George for Granite Club System Generated Employee's number

''Ticket #20076 Frank 04/01/2011
''Small program
'If glbCompSerial = "S/N - 2259W" Then
'    CmdPointRecalc.Caption = "Small Program"
'    CmdPointRecalc.Visible = True
'End If

Exit Sub

Ld_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Company", "HRPARCO", "Select")
Resume Next

End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub modLoadLists()
Dim x%

cmbMonAnn(0).AddItem "Monthly"
cmbMonAnn(0).AddItem "Annual"
cmbMonAnn(0).AddItem "Annualized Monthly"

'Ticket #29230 - Daily Vacation Entitlement - glbCompEntVacDaily is only used temporary setting during the development
'If glbCompEntVacDaily Then
    cmbMonAnn(0).AddItem "Daily"
'End If

If Len(lblMonAnn(0)) > 0 Then
    If lblMonAnn(0) = "M" Then
        cmbMonAnn(0).ListIndex = 0
    ElseIf lblMonAnn(0) = "A" Then
        cmbMonAnn(0).ListIndex = 1
    Else
        'Ticket #29230 - Daily Vacation Entitlement - glbCompEntVacDaily is only used temporary setting during the development
        'If glbCompEntVacDaily And lblMonAnn(0) = "D" Then
        If lblMonAnn(0) = "D" Then
            cmbMonAnn(0).ListIndex = 3
        Else
            cmbMonAnn(0).ListIndex = 2
        End If
    End If
Else
    cmbMonAnn(0).ListIndex = -1
End If

cmbMonAnn(1).AddItem "Monthly"
cmbMonAnn(1).AddItem "Annual"
cmbMonAnn(1).AddItem "Annualized Monthly"

If Len(lblMonAnn(1)) > 0 Then
    If lblMonAnn(1) = "M" Then
        cmbMonAnn(1).ListIndex = 0
    ElseIf lblMonAnn(1) = "A" Then
        cmbMonAnn(1).ListIndex = 1
    Else
        cmbMonAnn(1).ListIndex = 2
    End If
Else
    cmbMonAnn(1).ListIndex = -1
End If

cmbEntBase.AddItem lStr("Original Hire Date")
cmbEntBase.AddItem lStr("Seniority Date")
cmbEntBase.AddItem lStr("Last Hire Date")
cmbEntBase.AddItem lStr("Union Date")
cmbEntBase.AddItem lStr("User Defined")

If Len(medEntBase) > 0 Then
    Select Case medEntBase
        Case "O": x% = 0
        Case "S": x% = 1
        Case "L": x% = 2
        Case "U": x% = 3
        Case Else: x% = 4
    End Select
Else
    x% = -1
End If
cmbEntBase.ListIndex = x%

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

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF

'CmdRecalc.Enabled = FT
'CmdRecalc.Enabled = FT

cmbEntBase.Enabled = TF
cmbMonAnn(0).Enabled = TF
cmbMonAnn(1).Enabled = TF
cmbPrecision.Enabled = TF
cmbSysGenEmpl.Enabled = TF And Not glbLinamar
cmbCountry.Enabled = TF

fraEntDate.Enabled = TF
panCompName.Enabled = TF
'txtComp.Enabled = TF
txtDateUsed.Enabled = TF
txtDateUseS.Enabled = TF
dlpEDates(0).Enabled = TF
dlpEDates(1).Enabled = TF
'dlpEDatesS(0).Enabled = TF
'dlpEDatesS(1).Enabled = TF
txtFedTax.Enabled = TF
txtPROVTAX.Enabled = TF
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub


Private Sub txtComp_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtDateUsed_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtDateUseS_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub
'Private Sub dlpEDates_Change(Index As Integer)
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub dlpEDates_DblClick(Index As Integer)
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub dlpEDates_GotFocus(Index As Integer)
'    Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub dlpEDates_KeyPress(Index As Integer, KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub
'Private Sub dlpEDatesS_Change(Index As Integer)
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub dlpEDatesS_DblClick(Index As Integer)
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub dlpEDatesS_GotFocus(Index As Integer)
'    Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub dlpEDatesS_KeyPress(Index As Integer, KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub

Private Sub txtFedTax_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtMaxEE_Change()
    'Release 8.1 - The Max Employee License is encrypted now.
    'Decrypt the Max Employee and put on the control visible to the users
    
    'Check if it's encrypted. If so, then decrypt it first and then assign to EMax.
    If Not Data1.Recordset.EOF Then
        If Not Data1.Recordset.Fields("PC_OPTUPD") Or IsNull(Data1.Recordset.Fields("PC_OPTUPD")) Then
            'Not Encrypted yet
        Else
            'Encrypted. Decrypt it first and then assign
            'Encrypted value is stored in the different field.
            If gsMultiLang = "Y" Then 'For Listowel only
                txtMaxEE.Text = DecryptPasswordMultiLang_First(Data1.Recordset.Fields("PC_OPT"))
            ElseIf UCase(gsMultiLang) = "YES" Then 'whscc
                txtMaxEE.Text = DecryptPasswordMultiLang(Data1.Recordset.Fields("PC_OPT"))
            Else
                txtMaxEE.Text = DecryptPassword(Data1.Recordset.Fields("PC_OPT"))
            End If
        End If
    End If

End Sub

Private Sub txtMaxEE_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtNextAvEmpl_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'George for Granite Club System Generated Employee's number
Private Sub txtNextAvEmpl_LostFocus()
Dim strMessage
'Granite Club and City of Timmins (Ticket #10340) and Ticket #14968 - Listowel Tech.
If glbCompSerial = "S/N - 2241W" Or glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2350W" Then

    If Not (OldNextAvEmpl = txtNextAvEmpl) Then
        strMessage = Check_EMPLOYEE_Number(Int(txtNextAvEmpl.Text))
        If strMessage <> "" Then  'George for Granite Club System Generated Employee's number
            lblMessage.Visible = True
            lblMessage.Caption = strMessage
        End If
    Else
        lblMessage.Visible = False
    End If
    OldNextAvEmpl = txtNextAvEmpl
End If
End Sub

Private Sub txtOptMaxEE_Change()
    'Release 8.1 - The Max Employee License is encrypted now.
    'Decrypt the Max Employee and put on the control visible to the users
    
    'Check if it's encrypted. If so, then decrypt it first and then assign to EMax.
    If Not Data1.Recordset.EOF Then
        If Not Data1.Recordset.Fields("PC_OPTUPD") Or IsNull(Data1.Recordset.Fields("PC_OPTUPD")) Then
            'Not Encrypted yet
        Else
            'Encrypted. Decrypt it first and then assign
            'Encrypted value is stored in the different field.
            If gsMultiLang = "Y" Then 'For Listowel only
                txtMaxEE.Text = DecryptPasswordMultiLang_First(Data1.Recordset.Fields("PC_OPT"))
            ElseIf UCase(gsMultiLang) = "YES" Then 'whscc
                txtMaxEE.Text = DecryptPasswordMultiLang(Data1.Recordset.Fields("PC_OPT"))
            Else
                txtMaxEE.Text = DecryptPassword(Data1.Recordset.Fields("PC_OPT"))
            End If
        End If
    End If
End Sub

Private Sub txtPROVTAX_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtSerialNo_GotFocus()
    Call SetPanHelp(ActiveControl)
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
UpdateRight = gSec_Upd_Company
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean

If gSec_Upd_Company Then           'May99 js
 Updateble = True
Else
Updateble = True
End If

End Property

Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = False
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
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
Call ST_UPD_MODE(TF)

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
If InStr(xCountryList, cmbCountry) = 0 And cmbCountry <> "" Then
    xCountryList = xCountryList & "&" & cmbCountry
    cmbCountry.AddItem cmbCountry
'    comCountryOfEmp.AddItem cmbCountry
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
 
Private Sub LinamarSrceenSetup() 'Ticket #29759 Franks 02/22/2017
    txtNextAvPos.DataField = "PC_NEXT_POS_NBR"
    lblNextAvPos.Caption = "Next Available Payroll ID"
    lblNextAvPos.Visible = True
    txtNextAvPos.Visible = True
End Sub

Private Sub WFCSrceenSetup() 'Ticket #27827 Franks 12/02/2015
    If glbUserID = "3142" Then 'Ticket #28373 Franks 03/29/2016
        cmdESSToTracker.Visible = True
    End If
    txtNextAvPos.DataField = "PC_NEXT_POS_NBR"
    lblNextAvPos.Visible = True
    txtNextAvPos.Visible = True
    Call WFCNextPosNoSetup("Reset")
End Sub
