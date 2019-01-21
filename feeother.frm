VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEmpOther 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Employee Other Information"
   ClientHeight    =   8355
   ClientLeft      =   135
   ClientTop       =   2535
   ClientWidth     =   12960
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8355
   ScaleWidth      =   12960
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtUserText 
      Appearance      =   0  'Flat
      DataSource      =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   6
      Left            =   8040
      MaxLength       =   15
      TabIndex        =   12
      Tag             =   "00-V.I.C"
      Top             =   4050
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox txtUserText 
      Appearance      =   0  'Flat
      DataSource      =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   5
      Left            =   8040
      MaxLength       =   15
      TabIndex        =   11
      Tag             =   "00-U.A.N "
      Top             =   3720
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   330
      Left            =   1680
      TabIndex        =   33
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtUserText 
      Appearance      =   0  'Flat
      DataField       =   "ER_TEXT4"
      DataSource      =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   4
      Left            =   3480
      MaxLength       =   20
      TabIndex        =   10
      Tag             =   "00-Other Text 4"
      Top             =   4065
      Width           =   1620
   End
   Begin VB.TextBox txtUserText 
      Appearance      =   0  'Flat
      DataField       =   "ER_TEXT3"
      DataSource      =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   3
      Left            =   3480
      MaxLength       =   20
      TabIndex        =   9
      Tag             =   "00-Other Text 3"
      Top             =   3735
      Width           =   1620
   End
   Begin VB.TextBox txtUserText 
      Appearance      =   0  'Flat
      DataField       =   "ER_TEXT2"
      DataSource      =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   2
      Left            =   3480
      MaxLength       =   20
      TabIndex        =   8
      Tag             =   "00-Other Text 2"
      Top             =   3390
      Width           =   1620
   End
   Begin VB.TextBox txtUserText 
      Appearance      =   0  'Flat
      DataField       =   "ER_TEXT1"
      DataSource      =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   1
      Left            =   3480
      MaxLength       =   20
      TabIndex        =   7
      Tag             =   "00-Other Text 1"
      Top             =   3060
      Width           =   1620
   End
   Begin VB.TextBox txtPassNo 
      Appearance      =   0  'Flat
      DataField       =   "ER_PASSPORTNO"
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
      Left            =   3480
      MaxLength       =   50
      TabIndex        =   4
      Tag             =   "00-Passport Number"
      Top             =   1815
      Width           =   3360
   End
   Begin VB.TextBox txtPassCountry 
      Appearance      =   0  'Flat
      DataField       =   "ER_PASSCOUNTRY"
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
      Left            =   3480
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "00-Passport Country"
      Top             =   1140
      Width           =   3360
   End
   Begin VB.TextBox txtPermit 
      Appearance      =   0  'Flat
      DataField       =   "ER_VISAPERMITNO"
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
      Left            =   3480
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "00-Visa/Work permit #"
      Top             =   2265
      Width           =   3360
   End
   Begin VB.TextBox txtCitizen 
      Appearance      =   0  'Flat
      DataField       =   "ER_CITIZENSHIP"
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
      Left            =   3480
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "00-Citizenship"
      Top             =   675
      Width           =   3360
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   2160
      Top             =   7800
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Height          =   540
      Left            =   0
      TabIndex        =   19
      Top             =   7815
      Visible         =   0   'False
      Width           =   12960
      _Version        =   65536
      _ExtentX        =   22860
      _ExtentY        =   952
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.74
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   7140
         Top             =   0
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
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "ER_LUSER"
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
      Left            =   7320
      MaxLength       =   25
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "LUser"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "ER_LTIME"
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
      Left            =   5760
      MaxLength       =   25
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "LTime"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "ER_LDATE"
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
      Left            =   4200
      MaxLength       =   25
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "Ldate"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   12960
      _Version        =   65536
      _ExtentX        =   22860
      _ExtentY        =   873
      _StockProps     =   15
      ForeColor       =   0
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
         Left            =   7320
         TabIndex        =   28
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label lblEEnum 
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
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   135
         Width           =   1005
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         DataField       =   "ER_EMPNBR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5040
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         Left            =   3150
         TabIndex        =   14
         Top             =   120
         Width           =   1740
      End
   End
   Begin INFOHR_Controls.DateLookup dlpPassDate 
      DataField       =   "ER_PASSCOUNTRYDATE"
      Height          =   285
      Left            =   3165
      TabIndex        =   3
      Tag             =   "40-Passport Expiration Date"
      Top             =   1470
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpPermitDate 
      DataField       =   "ER_VISAPERMITDATE"
      Height          =   285
      Left            =   3165
      TabIndex        =   6
      Tag             =   "40-Visa/Work Permit Expiration Date"
      Top             =   2610
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "WFC Other_Text 6"
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
      Left            =   6000
      TabIndex        =   36
      Top             =   4095
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "WFC Other_Text 5"
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
      Index           =   11
      Left            =   6000
      TabIndex        =   35
      Top             =   3750
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblImport 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Other Info."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   210
      TabIndex        =   34
      Top             =   4845
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Image imgNoSec 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1350
      Picture         =   "feeother.frx":0000
      Top             =   4845
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Text 4"
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
      Index           =   10
      Left            =   240
      TabIndex        =   32
      Top             =   4110
      Width           =   2565
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Text 3"
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
      Left            =   240
      TabIndex        =   31
      Top             =   3765
      Width           =   2565
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Text 2"
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
      Left            =   240
      TabIndex        =   30
      Top             =   3435
      Width           =   2565
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Text 1"
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
      Left            =   240
      TabIndex        =   29
      Top             =   3090
      Width           =   2565
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Passport Number"
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
      Left            =   240
      TabIndex        =   27
      Top             =   1845
      Width           =   1215
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "ER_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3000
      TabIndex        =   26
      Top             =   7320
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Visa/Work Permit Expiration Date"
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
      Index           =   5
      Left            =   240
      TabIndex        =   25
      Top             =   2640
      Width           =   2370
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Citizenship"
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
      Index           =   1
      Left            =   240
      TabIndex        =   24
      Top             =   720
      Width           =   750
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Visa/Work Permit #"
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
      Left            =   240
      TabIndex        =   23
      Top             =   2310
      Width           =   1395
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Passport Expiration Date"
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
      Left            =   240
      TabIndex        =   22
      Top             =   1515
      Width           =   1740
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Passport Country"
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
      Left            =   240
      TabIndex        =   21
      Top             =   1170
      Width           =   1200
   End
   Begin VB.Image imgSec 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1350
      Picture         =   "feeother.frx":014A
      Top             =   4845
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmEmpOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim fglbNew As Integer
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim oContName, oContNbr1, oContNbr2, oEmail
Dim oCitizen, oPassCountry, oPassDate, oPassNo, oPermit, oPermitDate


Public Sub cmdCancel_Click()
Dim X
On Error GoTo Can_Err

rsDATA.CancelUpdate
Call Display_Value
fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(True)  ' reset screen's attributes

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
Call RollBack   '23June99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEMPOTHER" Then glbOnTop = ""

End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdModify_Click()

On Error GoTo Mod_Err

oCitizen = txtCitizen.Text
oPassCountry = txtPassCountry.Text
oPassDate = dlpPassDate.Text
oPassNo = txtPassNo.Text
oPermit = txtPermit.Text
oPermitDate = dlpPermitDate.Text

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HREMP", "Modify")
Call RollBack  '23June99 - js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdOK_Click()

On Error GoTo Add_Err

If Not chkEmpOther() Then Exit Sub '17Aug99 js

If Not glbtermopen Then If Not AUDITOTHER() Then MsgBox "ERROR : AUDIT 2 FILE"

'rsDATA.Requery
Call UpdUStats(Me) ' update user's stats (who did it and when)

Call Set_Control("U", Me, rsDATA)
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

'Ticket #24564 - Macaulay Child Development Centre - changes and enhancements
If Not updFollow("U", "WorkPermit") Then Exit Sub
If Not updFollow("U", "Passport") Then Exit Sub     'Ticket #22682: Release 8.0 - Jerry asked to add for Passport expiration as well


fglbNew = False

Call SET_UP_MODE
'Call ST_UPD_MODE(True)
Call EERetrieve

Call NextForm
Exit Sub

Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREMP", "Update")
Call RollBack  '23June99 - js

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String, xReport, X%

'cmdPrint.Enabled = False

RHeading = lblEEName & "'s Emegency Contact Information"
Me.vbxCrystal.WindowTitle = lblEEName & "'s Emegency Contact Report"
Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
'Call setRptLabel(Me, 1)
If Not glbtermopen Then
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 1
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next
    End If
    xReport = glbIHRREPORTS & "Rgcontct.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{HREMP.ED_EMPNBR}=" & glbLEE_ID & " "
Else
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
        Me.vbxCrystal.DataFiles(1) = glbIHRAUDIT
    End If
    xReport = glbIHRREPORTS & "Rgcontc2.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{Term_HREMP.TERM_SEQ}=" & glbTERM_Seq & " "
    
End If


Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True
End Sub

Public Sub cmdView_Click()
Dim RHeading As String, xReport, X%

'cmdPrint.Enabled = False

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = lblEEName & "'s Emegency Contact Information"
Me.vbxCrystal.WindowTitle = lblEEName & "'s Emegency Contact Report"
Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
'Call setRptLabel(Me, 1)
If Not glbtermopen Then
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 1
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next
    End If
    xReport = glbIHRREPORTS & "Rgcontct.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{HREMP.ED_EMPNBR}=" & glbLEE_ID & " "
Else
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
        Me.vbxCrystal.DataFiles(1) = glbIHRAUDIT
    End If
    xReport = glbIHRREPORTS & "Rgcontc2.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{Term_HREMP.TERM_SEQ}=" & glbTERM_Seq & " "
    
End If

Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True
End Sub

Function EERetrieve()

Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS

'Release 8.0 - Ticket #22682: Get Employee # of the User - View Own security
If Not glbtermopen Then
    If glbUserEmpNo = glbLEE_ID And Not gSec_OthInfo_ViewOwn Then
        MsgBox "You cannot view your own Other Information.", vbCritical, "info:HR - Security"
        'glbLEE_ID = 0      'Ticket #25208
        Screen.MousePointer = DEFAULT
        Unload Me: Exit Function
    End If
End If

If glbtermopen Then
    SQLQ = "Select " & FldList & " from Term_HREMP_OTHER"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    
Else
    SQLQ = "Select " & FldList & " from HREMP_OTHER "
    SQLQ = SQLQ & " where ER_EMPNBR = " & glbLEE_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

If rsDATA.EOF Then
    rsDATA.AddNew
    rsDATA("ER_COMPNO") = "001"
    rsDATA("ER_EMPNBR") = glbLEE_ID
    If glbtermopen Then
        rsDATA("TERM_SEQ") = glbTERM_Seq
    End If
    'Ticket #20638 Franks 07/18/2011
    rsDATA("ER_LDATE") = Date
    rsDATA("ER_LTIME") = Time$
    rsDATA("ER_LUSER") = glbUserID
    rsDATA.Update
End If

Data1.RecordSource = SQLQ
Data1.Refresh
Call Display_Value

'If rsDATA.BOF And rsDATA.EOF Then
'   MsgBox "Sorry, Employee Removed prior to your access"
'Else
   EERetrieve = True
'End If


Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HREMP", "SELECT")
Call RollBack  '23June99 - js

Exit Function

End Function

Private Sub cmdImport_Click()
    glbDocNewRecord = False
    glbDocName = "OtherInfo"
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmEmpOther")
End Sub

Private Sub dlpPassDate_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub dlpPermitDate_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Form_Activate()
    glbOnTop = "FRMEMPOTHER"
    Call SET_UP_MODE
    'txtCitizen.SetFocus    'Ticket #19045 - causing an error when the user has only Inquire right on this screen
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEMPOTHER"
End Sub

Private Sub ScreenLabels()
lblTitle(1).Caption = lStr("Citizenship")
lblTitle(2).Caption = lStr("Passport Country")
lblTitle(3).Caption = lStr("Passport Expiration Date")
lblTitle(4).Caption = lStr("Visa/Work Permit #")
lblTitle(5).Caption = lStr("Visa/Work Permit Expiration Date")
lblTitle(6).Caption = lStr("Passport Number")
lblTitle(7).Caption = lStr("Other Text 1")
lblTitle(8).Caption = lStr("Other Text 2")
lblTitle(9).Caption = lStr("Other Text 3")
lblTitle(10).Caption = lStr("Other Text 4")
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim I
Dim xSIN

glbOnTop = "FRMEMPOTHER"

Screen.MousePointer = HOURGLASS

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

Screen.MousePointer = DEFAULT

'Ticket #20609 Franks 09/06/2011, move this to ScreenLabels
'For I = 0 To 6
'    lblTitle(I).Caption = lStr(lblTitle(I).Caption)
'Next I
Call ScreenLabels

If glbWFC Then 'Ticket #27336 Franks 08/04/2015
    Call WFCScreen
End If

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

'Release 8.0 - Ticket #22682: Get Employee # of the User - View Own security
If Not glbtermopen Then
    If glbUserEmpNo = glbLEE_ID And Not gSec_OthInfo_ViewOwn Then
        MsgBox "You cannot view your own Other Information.", vbCritical, "info:HR - Security"
        'glbLEE_ID = 0      'Ticket #25208
        Screen.MousePointer = DEFAULT
        Unload Me: Exit Sub
    End If
End If

Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

Call ST_UPD_MODE(False)

'Ticket #24543 - Macaulay Child Development Centre
If glbCompSerial = "S/N - 2420W" Then
    'SIN begins with 9 then Work Visa # and Expiration Date is mandatory
    xSIN = GetEmpData(glbLEE_ID, "ED_SIN", "")
    If Left(xSIN, 1) = "9" Then
        lblTitle(4).FontBold = True
        lblTitle(5).FontBold = True
    End If
End If

Call INI_Controls(Me)

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

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Call NextForm
End Sub

Private Sub imgSec_Click()
    Dim SQLQ
    glbDocName = "OtherInfo"
    SQLQ = getSQL("frmEmpOther")
    Call FillMemoFile(SQLQ, "OtherInfo")
End Sub

'Private Sub imgEmail_Click(Index As Integer)
'    Call txtEEMail_DblClick(Index)
'End Sub

Private Sub lblEEID_Change()
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    Me.Caption = "Other Information - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID

lblEEnum = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub

Private Sub medCTele_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCTele_ValidationError(Index As Integer, InvalidText As String, StartPosition As Integer)
    glbBYPASS380 = True
End Sub

Public Sub cmdNew_Click()

fglbNew = True

Call SET_UP_MODE

On Error GoTo AddN_Err


'xAction = "A"   '24June99 js - added from VB3
Call Set_Control("B", Me)
rsDATA.AddNew


If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"

Exit Sub
AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err


Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRDEPEND", "Add")
Call RollBack

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

txtCitizen.Enabled = TF
txtPassCountry.Enabled = TF
dlpPassDate.Enabled = TF
txtPermit.Enabled = TF
dlpPermitDate.Enabled = TF

'txtDOR2ADDRESS(1).Enabled = TF
'medHEALTHCARD.Enabled = TF

End Sub

Private Sub medDTele_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medECellPhone_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medEPageNbr_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medHEALTHCARD_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medVERSION_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtContName_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub
                 
Function chkEmpOther()
Dim X%

On Error GoTo chkEmpOther_Err

chkEmpOther = False


If Len(dlpPassDate) > 0 Then
    If Not IsDate(dlpPassDate) Then
        MsgBox "Invalid Passport Expiration Date"
        dlpPassDate.SetFocus
        Exit Function
    End If
End If

If Len(dlpPermitDate) > 0 Then
    If Not IsDate(dlpPermitDate) Then
        MsgBox "Invalid Visa/Work Permit Expiration Date"
        dlpPermitDate.SetFocus
        Exit Function
    End If
End If

'Ticket #24543 - Macaulay Child Development Centre
If glbCompSerial = "S/N - 2420W" Then
    'SIN begins with 9 then Work Visa # and Expiration Date is mandatory
    If lblTitle(4).FontBold = True And lblTitle(5).FontBold = True Then
        If Len(Trim(txtPermit.Text)) = 0 Or Len(Trim(dlpPermitDate.Text)) = 0 Then
            MsgBox "'Visa/Work Permit #' and 'Visa/Work Permit Expiration Date' cannot be blank."
            If Len(Trim(txtPermit.Text)) = 0 Then
                txtPermit.SetFocus
            Else
                dlpPermitDate.SetFocus
            End If
            Exit Function
        End If
    End If
End If

chkEmpOther = True

Exit Function

chkEmpOther_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEmpOther", "HREMP", "edit/Add")
Call RollBack '17Aug99 js

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


Private Sub txtDOR2ADDRESS_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtDORADDRESS_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtEDoctor_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtEEMail_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub
Private Function FldList()
Dim SQLQ
SQLQ = ""
SQLQ = SQLQ & "ER_COMPNO,ER_EMPNBR, ER_CITIZENSHIP, ER_PASSCOUNTRY, ER_PASSCOUNTRYDATE, "
SQLQ = SQLQ & "ER_VISAPERMITNO, ER_VISAPERMITDATE,ER_PASSPORTNO,"
SQLQ = SQLQ & "ER_TEXT1, ER_TEXT2,ER_TEXT3,ER_TEXT4," 'Ticket #20609 Franks 09/06/2011
SQLQ = SQLQ & "ER_LDATE, ER_LTIME, ER_LUSER"
If glbWFC Then 'Ticket #27336 Franks 08/04/2015
    SQLQ = SQLQ & ",ER_PAYROLL_ID4, ER_PAYROLL_ID5"
End If
If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"
FldList = SQLQ
End Function

Public Sub Display_Value()
    Dim SQLQ
    If rsDATA.EOF Or rsDATA.BOF Then
        Call Set_Control("B", Me)
        Call SET_UP_MODE
        Exit Sub
    End If
    
If glbtermopen Then
    SQLQ = "Select " & FldList & " from Term_HREMP_OTHER"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    
Else
    SQLQ = "Select " & FldList & " from HREMP_OTHER "
    SQLQ = SQLQ & " where ER_EMPNBR = " & glbLEE_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If
'Data1.RecordSource = SQLQ
'Data1.Refresh
If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
Call Set_Control("R", Me, rsDATA)
Call SET_UP_MODE

Me.cmdModify_Click

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
RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_OtherInformation
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = True
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

'Release 8.1
glbDocName = "OtherInfo" 'George on Jan 26,2006 #10266
If gsAttachment_DB Then 'George on Jan 24,2006 #10266
    Call DispimgIcon(Me, "frmEmpOther")
End If

If gsAttachment_DB Then
    If Data1.Recordset.BOF And Data1.Recordset.EOF Then
        cmdImport.Visible = False
    Else
        cmdImport.Visible = True
    End If
    If Not (gSec_Upd_OtherInformation And Not glbtermopen) Then
        cmdImport.Visible = False
    End If
End If

End Sub

Private Sub txtCitizen_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtPassCountry_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtPermit_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Function AUDITOTHER()
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD, xDiv, xPT
Dim X
Dim UpdateAudit As Boolean
On Error GoTo AUDIT_ERR
AUDITOTHER = False

xADD = False

If oCitizen <> txtCitizen.Text Then UpdateAudit = True
If oPassCountry <> txtPassCountry.Text Then UpdateAudit = True
If oPassDate <> dlpPassDate.Text Then UpdateAudit = True
If oPassNo <> txtPassNo.Text Then UpdateAudit = True
If oPermit <> txtPermit.Text Then UpdateAudit = True
If oPermitDate <> dlpPermitDate.Text Then UpdateAudit = True

If UpdateAudit Then
    GoTo MODUPD
Else
    GoTo MODNOUPD
End If

MODUPD:
    rsTA.Open "SELECT * FROM HRAUDIT2 WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    
    rsTB.Open "select ED_DIV,ED_PT,ED_PAYROLL_ID,ED_SECTION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
    
    If Not rsTB.EOF Then
        If IsNull(rsTB("ED_DIV")) Then xDiv = "" Else xDiv = rsTB("ED_DIV")
    Else
        xDiv = ""
    End If
    
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_HIRECODE_TABL") = "EDHC": rsTA("AU_ORG_TABL") = "EDOR"
    rsTA("AU_NEWEMP") = "N"
    rsTA("AU_DIVUPL") = xDiv
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbLEE_ID
    If Not IsNull(rsTB("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsTB("ED_PAYROLL_ID")
    If Not IsNull(rsTB("ED_SECTION")) Then rsTA("AU_SECTION") = rsTB("ED_SECTION")
    If oCitizen <> txtCitizen.Text Then rsTA("AU_CITIZENSHIP") = txtCitizen.Text
    If oPassCountry <> txtPassCountry.Text Then rsTA("AU_PASSCOUNTRY") = txtPassCountry.Text
    If oPassDate <> dlpPassDate.Text Then
        If IsDate(dlpPassDate.Text) Then rsTA("AU_PASSCOUNTRYDATE") = dlpPassDate.Text
    End If
    If oPassNo <> txtPassNo.Text Then rsTA("AU_PASSPORTNO") = txtPassNo.Text
    If oPermit <> txtPermit.Text Then rsTA("AU_VISAPERMITNO") = txtPermit.Text
    If oPermitDate <> dlpPermitDate.Text Then
        If IsDate(dlpPermitDate.Text) Then rsTA("AU_VISAPERMITDATE") = dlpPermitDate.Text
    End If
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "M"
    rsTA.Update
    rsTA.Close
    rsTB.Close
    
MODNOUPD:
AUDITOTHER = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack
Resume Next
End Function

Private Sub txtUserText_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Function updFollow(xType, xDateType)  'Laura on 11/2/97
Dim newline As String
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim rsFollow As New ADODB.Recordset
Dim rsTT As New ADODB.Recordset
Dim Edit1 As Integer
Dim oDATE, xDate

'Don't need a message for follow up - Jerry asked for v7.6

newline = Chr$(13) & Chr$(10)
updFollow = False

'Add Follow Up for Passport Date Expiration as well.
If xDateType = "Passport" Then
    oDATE = oPassDate
    xDate = dlpPassDate.Text
ElseIf xDateType = "WorkPermit" Then
    oDATE = oPermitDate
    xDate = dlpPermitDate.Text
End If


On Error GoTo CrFollow_Err

If IsDate(oDATE) Then     'DATE Renewal IS NOW MANDATORY
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
    If xDateType = "Passport" Then
        SQLQ = SQLQ & " AND EF_FREAS = 'PE'"
    Else
        SQLQ = SQLQ & " AND EF_FREAS = 'WP'"
    End If
    SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(oDATE)
    dynHRAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If dynHRAT.BOF And dynHRAT.EOF Then
        Edit1 = False
    Else
        Edit1 = True    ' returns true if found records
    End If
Else
    Edit1 = False
End If

If xType = "U" Then
    'New record and date is entered -> create follow up
    rsTB.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If fglbNew And IsDate(xDate) Then
        SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
        If xDateType = "Passport" Then
            SQLQ = SQLQ & " AND EF_FREAS = 'PE'"
        Else
            SQLQ = SQLQ & " AND EF_FREAS = 'WP'"
        End If
        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(xDate)
        rsFollow.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsFollow.EOF Then
            'Create the Code if not already existing
            If xDateType = "Passport" Then
                rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='PE'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            Else
                rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='WP'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            End If
            If rsTT.EOF Then
                rsTT.AddNew
                rsTT("TB_COMPNO") = "001"
                rsTT("TB_NAME") = "FURE"
                If xDateType = "Passport" Then
                    rsTT("TB_KEY") = "PE"
                    rsTT("TB_DESC") = "Passport Expiration"
                Else
                    rsTT("TB_KEY") = "WP"
                    rsTT("TB_DESC") = "Work Permit Expiration"
                End If
                rsTT("TB_LUSER") = glbUserID
                rsTT("TB_LDATE") = Date
                rsTT("TB_LTIME") = Time$
                rsTT.Update
            End If
            rsTT.Close
            Set rsTT = Nothing
            
            'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
            'follow up record
            If xDateType = "Passport" Then
                Call Grant_FollowUpCode_Security(glbUserID, "PE", "Passport Expiration")
            Else
                Call Grant_FollowUpCode_Security(glbUserID, "WP", "Work Permit Expiration")
            End If
            
            'Add by Frank for no duplicated record of HR_FOLLOW_UP End
            rsTB.AddNew
            rsTB("EF_COMPNO") = "001"
            rsTB("EF_EMPNBR") = glbLEE_ID
            rsTB("EF_FDATE") = CVDate(xDate)
            rsTB("EF_FREAS_TABL") = "FURE"
            'Ticket #24257 - Do not update Admin By for them only
            If glbCompSerial <> "S/N - 2262W" Then
                rsTB("EF_ADMINBY_TABL") = "EDAB"
                rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
            End If
            If xDateType = "Passport" Then
                rsTB("EF_FREAS") = "PE"
            Else
                rsTB("EF_FREAS") = "WP"
            End If
            rsTB("EF_COMMENTS") = ""
            rsTB("EF_LDATE") = Date
            rsTB("EF_LTIME") = Time$
            rsTB("EF_LUSER") = glbUserID
            rsTB.Update
            ' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
            Call Pause(0.5)
            'Msg = "A Follow Up Record was created!"
            'MsgBox Msg
        End If
        rsFollow.Close
        rsTB.Close
        updFollow = True
        Exit Function
    End If
    
    'Updating existing record but Follow Up record do not exists and the Date is valid -> create Follow Up
    If fglbNew = False And Edit1 = False And IsDate(xDate) Then
        ' 5/2/2001 Add by Frank for no duplicated record of HR_FOLLOW_UP Begin
        SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
        If xDateType = "Passport" Then
            SQLQ = SQLQ & " AND EF_FREAS = 'PE' "
        Else
            SQLQ = SQLQ & " AND EF_FREAS = 'WP' "
        End If
        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(xDate)
        rsFollow.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsFollow.EOF Then
            'Create the Code if not already existing
            If xDateType = "Passport" Then
                rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='PE'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            Else
                rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='WP'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            End If
            If rsTT.EOF Then
                rsTT.AddNew
                rsTT("TB_COMPNO") = "001"
                rsTT("TB_NAME") = "FURE"
                If xDateType = "Passport" Then
                    rsTT("TB_KEY") = "PE"
                    rsTT("TB_DESC") = "Passport Expiration"
                Else
                    rsTT("TB_KEY") = "WP"
                    rsTT("TB_DESC") = "Work Permit Expiration"
                End If
                rsTT("TB_LUSER") = glbUserID
                rsTT("TB_LDATE") = Date
                rsTT("TB_LTIME") = Time$
                rsTT.Update
            End If
            rsTT.Close
            Set rsTT = Nothing
        
            'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
            'follow up record
            If xDateType = "Passport" Then
                Call Grant_FollowUpCode_Security(glbUserID, "PE", "Passport Expiration")
            Else
                Call Grant_FollowUpCode_Security(glbUserID, "WP", "Work Permit Expiration")
            End If
        
            'Add by Frank for no duplicated record of HR_FOLLOW_UP End
            rsTB.AddNew
            rsTB("EF_COMPNO") = "001"
            rsTB("EF_EMPNBR") = glbLEE_ID
            rsTB("EF_FDATE") = CVDate(xDate)
            rsTB("EF_FREAS_TABL") = "FURE"
            'Ticket #24257 - Do not update Admin By for them only
            If glbCompSerial <> "S/N - 2262W" Then
                rsTB("EF_ADMINBY_TABL") = "EDAB"
                rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
            End If
            If xDateType = "Passport" Then
                rsTB("EF_FREAS") = "PE"
            Else
                rsTB("EF_FREAS") = "WP"
            End If
            rsTB("EF_COMMENTS") = ""
            rsTB("EF_LDATE") = Date
            rsTB("EF_LTIME") = Time$
            rsTB("EF_LUSER") = glbUserID
            rsTB.Update
            ' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
            Call Pause(0.5)
            'Msg = "A Follow Up Record was created!"
            'MsgBox Msg
        End If
        rsFollow.Close
        rsTB.Close
        updFollow = True
        Exit Function
    End If
  
    'Updating existing record and Follow Up record found, the Date is valid -> update existing Follow Up record
    If fglbNew = False And Edit1 = True And IsDate(xDate) Then  ' edited record
        'EOF?
        dynHRAT.MoveFirst
        Do Until dynHRAT.EOF
            'dynHRAT.Edit
            dynHRAT("EF_COMPNO") = "001"
            dynHRAT("EF_EMPNBR") = glbLEE_ID
            dynHRAT("EF_FDATE") = CVDate(xDate)
            If xDateType = "Passport" Then
                dynHRAT("EF_FREAS") = "PE"
            Else
                dynHRAT("EF_FREAS") = "WP"
            End If
            dynHRAT("EF_COMMENTS") = ""
            dynHRAT("EF_LDATE") = Date
            dynHRAT("EF_LTIME") = Time$
            dynHRAT("EF_LUSER") = glbUserID
            dynHRAT.Update
            ' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
            Call Pause(0.5)
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        If oDATE <> xDate Then
            'Msg = "A Follow Up Record was updated!"
            'MsgBox Msg
        End If
        updFollow = True
        Edit1 = True
        Exit Function
    End If
    
    'Updating existing record and the Follow Up exist, the Date is not valid -> delete the Follow Up record
    If fglbNew = False And Edit1 = True And Not IsDate(xDate) Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
        'Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    End If
Else
    If Edit1 = True Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
       ' Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    Else
        updFollow = True
    End If
End If

If xDate = "" Then
    updFollow = True
End If
  
Exit Function

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Function
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Other Information", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Function

Private Sub WFCScreen() 'Ticket #27336 Franks 08/04/2015
    lblTitle(11).Caption = "U.A.N" '"Other Text 5"
    lblTitle(12).Caption = "V.I.C" '"Other Text 6"
    lblTitle(11).Visible = True
    lblTitle(12).Visible = True
    txtUserText(5).DataField = "ER_PAYROLL_ID4"
    txtUserText(6).DataField = "ER_PAYROLL_ID5"
    txtUserText(5).Visible = True
    txtUserText(6).Visible = True
End Sub
