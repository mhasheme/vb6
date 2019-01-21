VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmDEPNDTS 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Dependents"
   ClientHeight    =   8490
   ClientLeft      =   435
   ClientTop       =   1770
   ClientWidth     =   11295
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
   ScaleHeight     =   8490
   ScaleWidth      =   11295
   WindowState     =   2  'Maximized
   Begin VB.ComboBox comDeptTxt1 
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
      Left            =   1920
      TabIndex        =   9
      Tag             =   "00-Depended Text 1"
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtUserText 
      Appearance      =   0  'Flat
      DataField       =   "DP_TEXT4"
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
      Left            =   6240
      MaxLength       =   20
      TabIndex        =   21
      Tag             =   "00-Dependent Text 4"
      Top             =   5460
      Width           =   1620
   End
   Begin VB.TextBox txtUserText 
      Appearance      =   0  'Flat
      DataField       =   "DP_TEXT3"
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
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   10
      Tag             =   "00-Dependent Text 3"
      Top             =   5460
      Width           =   1620
   End
   Begin VB.TextBox txtUserText 
      Appearance      =   0  'Flat
      DataField       =   "DP_TEXT2"
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
      Left            =   6240
      MaxLength       =   20
      TabIndex        =   20
      Tag             =   "00-Dependent Text 2"
      Top             =   5120
      Width           =   1620
   End
   Begin VB.TextBox txtUserText 
      Appearance      =   0  'Flat
      DataField       =   "DP_TEXT1"
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
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   8
      Tag             =   "00-Dependent Text 1"
      Top             =   5120
      Width           =   1620
   End
   Begin VB.ComboBox comOther 
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
      Left            =   7080
      TabIndex        =   19
      Tag             =   "00-Relationship of Dependent to Employee"
      Top             =   4410
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox comMedical 
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
      Left            =   7080
      TabIndex        =   17
      Tag             =   "00-Relationship of Dependent to Employee"
      Top             =   4080
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ComboBox comDental 
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
      ItemData        =   "Fedepnts.frx":0000
      Left            =   7080
      List            =   "Fedepnts.frx":0002
      TabIndex        =   15
      Tag             =   "00-Relationship of Dependent to Employee"
      Top             =   3750
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
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
      Height          =   1155
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Tag             =   "00-Comments"
      Top             =   6000
      Width           =   6645
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "Fedepnts.frx":0004
      Height          =   2085
      Left            =   240
      OleObjectBlob   =   "Fedepnts.frx":0018
      TabIndex        =   0
      Top             =   480
      Width           =   9975
   End
   Begin INFOHR_Controls.DateLookup dlpEligDte 
      DataField       =   "DP_SDATE"
      Height          =   285
      Left            =   5930
      TabIndex        =   11
      Tag             =   "40-Benefit Eligibility Date"
      Top             =   2760
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpEndDte 
      DataField       =   "DP_EDATE"
      Height          =   315
      Left            =   5930
      TabIndex        =   12
      Tag             =   "40-Benefit End Date"
      Top             =   3070
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDOB 
      DataField       =   "DP_DOB"
      Height          =   285
      Left            =   1605
      TabIndex        =   3
      Tag             =   "40-Birth Date of Dependent"
      Top             =   3420
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1135
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   6480
      Top             =   7920
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
      TabIndex        =   50
      Top             =   7830
      Width           =   11295
      _Version        =   65536
      _ExtentX        =   19923
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DP_LTIME"
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
      Left            =   5520
      MaxLength       =   25
      TabIndex        =   49
      Top             =   6360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox txtFName 
      Appearance      =   0  'Flat
      DataField       =   "Dp_FName"
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
      Left            =   1920
      MaxLength       =   25
      TabIndex        =   2
      Tag             =   "00-First Name of Dependent"
      Top             =   3090
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   7560
      TabIndex        =   48
      Top             =   2640
      Width           =   1215
      Begin VB.OptionButton optSex 
         Caption         =   "Female"
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
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Tag             =   "40-Gender"
         Top             =   330
         Width           =   825
      End
      Begin VB.OptionButton optSex 
         Caption         =   "Male"
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
         Left            =   120
         TabIndex        =   23
         Tag             =   "40-Gender"
         Top             =   60
         Width           =   825
      End
   End
   Begin VB.TextBox txtOther 
      Appearance      =   0  'Flat
      DataField       =   "DP_OTHER"
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
      Left            =   6240
      MaxLength       =   1
      TabIndex        =   18
      Tag             =   "00-COB Other"
      Top             =   4410
      Width           =   855
   End
   Begin VB.TextBox txtMedical 
      Appearance      =   0  'Flat
      DataField       =   "DP_MEDICAL"
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
      Left            =   6240
      MaxLength       =   1
      TabIndex        =   16
      Tag             =   "00-COB Medical"
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox txtDental 
      Appearance      =   0  'Flat
      DataField       =   "DP_DENTAL"
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
      Left            =   6240
      MaxLength       =   1
      TabIndex        =   14
      Tag             =   "00-COB Dental"
      Top             =   3750
      Width           =   855
   End
   Begin VB.TextBox txtDeptNo 
      Appearance      =   0  'Flat
      DataField       =   "DP_DEPNO"
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
      Left            =   6240
      MaxLength       =   4
      TabIndex        =   13
      Tag             =   "00-Dependent Number"
      Top             =   3420
      Width           =   855
   End
   Begin MSMask.MaskEdBox MedSIN 
      DataField       =   "DP_SIN"
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Tag             =   "00-Social Security Number"
      Top             =   3750
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
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
      Left            =   1920
      TabIndex        =   6
      Tag             =   "00-Status i.e. Student "
      Top             =   4440
      Width           =   1575
   End
   Begin VB.ComboBox ComSmoker 
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "00-Smoker Yes/No"
      Top             =   4740
      Width           =   855
   End
   Begin VB.ComboBox comRelation 
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
      Left            =   1920
      TabIndex        =   5
      Tag             =   "00-Relationship of Dependent to Employee"
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox txtSurname 
      Appearance      =   0  'Flat
      DataField       =   "DP_SNAME"
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
      Left            =   1920
      MaxLength       =   25
      TabIndex        =   1
      Tag             =   "00-Surname of Dependent"
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DP_LDATE"
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
      Left            =   3840
      MaxLength       =   25
      TabIndex        =   25
      Top             =   6360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DP_LUSER"
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
      Left            =   7200
      MaxLength       =   25
      TabIndex        =   27
      Top             =   6360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   11295
      _Version        =   65536
      _ExtentX        =   19923
      _ExtentY        =   873
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
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   19
         Left            =   120
         TabIndex        =   59
         Top             =   123
         Width           =   1005
      End
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
         Left            =   6360
         TabIndex        =   52
         Top             =   105
         Width           =   1305
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   5520
         TabIndex        =   31
         Top             =   4320
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
         Left            =   1440
         TabIndex        =   30
         Top             =   105
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
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
         Left            =   3120
         TabIndex        =   29
         Top             =   105
         Width           =   1740
      End
   End
   Begin VB.TextBox txtComRelation 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      DataField       =   "DP_RELATE"
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
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image imgHelp 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   7920
      Picture         =   "Fedepnts.frx":5BF4
      Stretch         =   -1  'True
      Top             =   5130
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblWFCTxt2 
      Caption         =   "(format: mm/dd/yyyy)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   8280
      TabIndex        =   58
      Top             =   5133
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dependent Text 4"
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
      Index           =   18
      Left            =   4080
      TabIndex        =   57
      Top             =   5460
      Width           =   2010
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dependent Text 3"
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
      Index           =   17
      Left            =   240
      TabIndex        =   56
      Top             =   5460
      Width           =   1650
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dependent Text 2"
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
      Index           =   16
      Left            =   4080
      TabIndex        =   55
      Top             =   5120
      Width           =   2010
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dependent Text 1"
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
      Index           =   15
      Left            =   240
      TabIndex        =   54
      Top             =   5120
      Width           =   1530
   End
   Begin VB.Label lblWFCCobOther 
      Caption         =   "If data is entered, email explanation to Corp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6240
      TabIndex        =   53
      Top             =   4725
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dependent Number"
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
      Index           =   10
      Left            =   4080
      TabIndex        =   44
      Top             =   3420
      Width           =   1755
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dependent Comment"
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
      Index           =   14
      Left            =   240
      TabIndex        =   51
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COB Other"
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
      Index           =   13
      Left            =   4080
      TabIndex        =   47
      Top             =   4410
      Width           =   1005
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COB Medical"
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
      Index           =   12
      Left            =   4080
      TabIndex        =   46
      Top             =   4080
      Width           =   930
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COB Dental"
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
      Index           =   11
      Left            =   4080
      TabIndex        =   45
      Top             =   3750
      Width           =   840
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Benefit End Date"
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
      Index           =   9
      Left            =   4080
      TabIndex        =   43
      Top             =   3090
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Benefit Eligible Date"
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
      Index           =   8
      Left            =   4080
      TabIndex        =   42
      Top             =   2760
      Width           =   1425
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S. S. N. "
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
      Index           =   7
      Left            =   240
      TabIndex        =   41
      Top             =   3750
      Width           =   600
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
      Left            =   240
      TabIndex        =   40
      Top             =   4410
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dependent Smoker"
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
      TabIndex        =   39
      Top             =   4740
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Relationship"
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
      TabIndex        =   38
      Top             =   4080
      Width           =   1065
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   37
      Top             =   3090
      Width           =   1155
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
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
      TabIndex        =   36
      Top             =   3420
      Width           =   1245
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Surname"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   35
      Top             =   2760
      Width           =   750
   End
   Begin VB.Label lblSex 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "M"
      DataField       =   "DP_SEX"
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
      Left            =   9000
      TabIndex        =   32
      Top             =   2760
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "DP_EMPNBR"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1080
      TabIndex        =   33
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "DP_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7800
      TabIndex        =   34
      Top             =   6960
      Visible         =   0   'False
      Width           =   690
   End
End
Attribute VB_Name = "frmDEPNDTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xAction
Dim OtxtSurname, OtxtFName, OtxtDOB, OmedSIN, OtxtRelation, OtxtComStatus
Dim OtxtEligDate, OtxtEndDte, OtxtComSmoker, OtxtDeptNo
Dim OtxtDental, OtxtMedical, OtxtOther, OOptSex
Dim txtComSmoker, txtComStatus
Dim RSDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim MailBody
Dim fglbNew As Integer
Dim AbortTerm As Boolean

Public Sub cmdCancel_Click()

On Error GoTo Can_Err
Dim x As Variant
   
xAction = " "
If Not (RSDATA.EOF And RSDATA.BOF) Then RSDATA.CancelUpdate
Call Display_Value

fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(True)
Me.vbxTrueGrid.Refresh
Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRDEPEND", "Cancel")
Call RollBack

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMDEPNDTS" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
 '   Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdDelete_Click()
Dim a As Integer, Msg As String, x

xAction = "D" '24June99 js - added from VB3

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "this Dependent?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub
If Not glbtermopen Then
    If Not AUDITDEPNTS() Then MsgBox "ERROR : AUDIT FILE" '24June99 js - added from VB3
End If

If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    RSDATA.Delete
    gdbAdoIhr001X.CommitTrans
Else
    gdbAdoIhr001.BeginTrans
    RSDATA.Delete
    gdbAdoIhr001.CommitTrans
End If

Data1.Refresh
If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If
fglbNew = False
Call SET_UP_MODE

'Call ST_UPD_MODE(True)
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRDEPEND", "Delete")
Call RollBack

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdModify_Click()

On Error GoTo Mod_Err

xAction = "C" '24June99 js-added from VB3
Call SET_UP_MODE
'Call ST_UPD_MODE(True)

'txtSurname.SetFocus
'Data1.Recordset.Edit

OtxtSurname = txtSurname         '24June99 js-added from VB3
OtxtFName = txtFName             '

If optSex(0).Value Then   '25June99 js
    lblSex.Caption = "M"  '
Else                      '
    lblSex.Caption = "F"  '
End If                    '

OOptSex = lblSex.Caption  '
OtxtDOB = dlpDOB.Text            '24June99
OtxtRelation = txtComRelation    '
OtxtComSmoker = ComSmoker.Text   '
OtxtComStatus = ComStatus.Text ' txtComStatus     '
OmedSIN = MedSIN                 '
OtxtEligDate = dlpEligDte.Text        '
OtxtEndDte = dlpEndDte.Text           '

'If glbCompSerial = "S/N - 2219W" Or glbCompSerial = "S/N - 2274W" Then
  OtxtDeptNo = txtDeptNo                '
  OtxtMedical = txtMedical              '
  OtxtDental = txtDental                '
  OtxtOther = txtOther                  '
'End If                                  '
'txtSurname.SetFocus

'Ticket #22417 - Goodmans LLP
If glbCompSerial = "S/N - 2290W" Then
    If txtDental.Text = "Y" Or txtMedical = "Y" Then
        lblWFCTxt2.Visible = True
        imgHelp.Visible = True
    Else
        lblWFCTxt2.Visible = False
        imgHelp.Visible = False
    End If
End If

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HRDEPEND", "Modify")
Call RollBack

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdNew_Click()

'Call ST_UPD_MODE(True)

fglbNew = True

Call SET_UP_MODE

On Error GoTo AddN_Err


xAction = "A"   '24June99 js - added from VB3
Call Set_Control("B", Me)
RSDATA.AddNew


If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"
txtSurname = RTrim$(glbLEE_SName) 'RAUBREY 4/2/97 'default the last name
optSex(0).Value = True
lblSex = "M"
comRelation = "" '???
ComStatus = ""
ComSmoker.ListIndex = 0
dlpEligDte.Text = ""
dlpEndDte.Text = ""
txtFName.SetFocus

If glbWFC Then 'Ticket #16287
    If glbEmpCountry = "CANADA" Then
        comRelation.ListIndex = 0
        ComStatus.ListIndex = 0
        ComSmoker.ListIndex = 0
        txtDeptNo.Text = 0
        'Ticket #17284
        'txtDental.Text = "N"
        'txtMedical.Text = "N"
        'txtOther.Text = "N"
        comDental.ListIndex = -1
        comMedical.ListIndex = -1
        comOther.ListIndex = -1
    End If
End If

'Ticket #22417 - Goodmans LLP
If glbCompSerial = "S/N - 2290W" Then
    comDeptTxt1.ListIndex = 0
    txtUserText(1).Text = Left(comDeptTxt1.Text, 1)
    
    comDental.ListIndex = -1
    comMedical.ListIndex = -1
End If

Exit Sub
AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err


Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRDEPEND", "Add")
Call RollBack

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdOK_Click()
Dim bk As Variant
Dim xBranch

On Error GoTo Add_Err

If Not chk_FeDepnts() Then Exit Sub '25June99 js-added function

If gsEMAIL_ONDEPENDENT Then
    MailBody = ""
    If NewHireForms.count = 0 Then 'Non new hire
        'If fglbNew Or chkCurrent Then
        If fglbNew Then
            MailBody = "The Dependent information has been changed." & vbCrLf & vbCrLf
            MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
            MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
            If glbCompSerial = "S/N - 2382W" Then  'Samuel
                xBranch = GetEmpData(glbLEE_ID, "ED_SECTION", "")
                If Len(xBranch) > 0 Then
                    xBranch = GetTABLDesc("EDSE", xBranch)
                    MailBody = MailBody & "Branch: " & xBranch & vbCrLf
                End If
            End If
            MailBody = MailBody & "Dependent Surname: " & txtSurname.Text & vbCrLf
            MailBody = MailBody & "Dependent First Name: " & txtFName.Text & vbCrLf
        End If
    End If
End If

If Not glbtermopen Then
    If Not AUDITDEPNTS() Then MsgBox "ERROR : AUDIT FILE"
End If

Call UpdUStats(Me) ' update user's stats (who did it and when)
Call UpdDepends

Call Set_Control("U", Me, RSDATA)
If glbtermopen Then
    RSDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    RSDATA.Update
    gdbAdoIhr001X.CommitTrans
Else
    gdbAdoIhr001.BeginTrans
    RSDATA.Update
    gdbAdoIhr001.CommitTrans
End If
Data1.Refresh


'Call ST_UPD_MODE(True)
fglbNew = False

Call SET_UP_MODE

xAction = " "

If gsEMAIL_ONDEPENDENT Then
    If Len(MailBody) > 0 Then
        Screen.MousePointer = DEFAULT
        If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #21601 Franks 02/24/2012
            Call EmailSendingForSamuel
        Else
            Call imgEmail_Click
        End If
    End If
End If

If NextFormIF("dependent") Then
    Call cmdNew_Click
End If

Exit Sub

Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRDEPEND", "Update")
Call RollBack

Unload Me

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String, xReport

RHeading = lblEEName & "'s Dependents"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Public Sub cmdView_Click()
Dim RHeading As String, xReport

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = lblEEName & "'s Dependents"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading

''Me.vbxCrystal.Password = gstrAccPWord$
''Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub comDental_LostFocus()
    'If Len(comDental.Text) > 0 Then
        txtDental.Text = Left(comDental.Text, 1)
        
        'Ticket #22417 - Goodmans LLP
        If glbCompSerial = "S/N - 2290W" Then
            If txtDental.Text = "Y" Or txtMedical = "Y" Then
                lblWFCTxt2.Visible = True
                imgHelp.Visible = True
            Else
                lblWFCTxt2.Visible = False
                imgHelp.Visible = False
            End If
        End If
    'End If
End Sub


Private Sub comDeptTxt1_LostFocus()
    txtUserText(1).Text = Left(comDeptTxt1.Text, 1)
End Sub

Private Sub comMedical_LostFocus()
    'If Len(comMedical.Text) > 0 Then
        txtMedical.Text = Left(comMedical.Text, 1)
    
        'Ticket #22417 - Goodmans LLP
        If glbCompSerial = "S/N - 2290W" Then
            If txtDental.Text = "Y" Or txtMedical = "Y" Then
                lblWFCTxt2.Visible = True
                imgHelp.Visible = True
            Else
                lblWFCTxt2.Visible = False
                imgHelp.Visible = False
            End If
        End If
    
    'End If
End Sub

Private Sub comOther_LostFocus()
    'If Len(comOther.Text) > 0 Then
        txtOther.Text = Left(comOther.Text, 1)
    'End If
End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub comRelation_GotFocus()

Call SetPanHelp(ActiveControl)
MDIMain.panHelp(2).Caption = "10"    'laura jan 05, 1998

txtComRelation = comRelation.Text 'js-2/12/99

End Sub


Private Sub comRelation_Click()
Dim tlen As Integer

tlen = Len(comRelation.Text)

If tlen > 10 Then tlen = 10

If tlen >= 1 Then
    txtComRelation.Text = Left$(comRelation.Text, tlen)
Else
    txtComRelation.Text = " "
End If

End Sub

Private Sub comRelation_KeyPress(KeyAscii As Integer)
If Len(comRelation) > 9 Then
    KeyAscii = 0
    Exit Sub
End If
End Sub

Private Sub comRelation_LostFocus()
'Added by Bryan 20/07/05 Ticket #8963, allows typing in relationship instead of selecting.
Dim tlen As Integer

tlen = Len(comRelation.Text)

If tlen > 10 Then tlen = 10

If tlen >= 1 Then
    txtComRelation.Text = Left$(comRelation.Text, tlen)
Else
    txtComRelation.Text = " "
End If
End Sub

Private Sub ComSmoker_GotFocus()

Call SetPanHelp(ActiveControl)

'---24June99 js - added from VB3
txtComSmoker = ComSmoker.Text 'js-2/12/99
'---

End Sub

Private Sub ComStatus_GotFocus()

Call SetPanHelp(ActiveControl)
'---24June99 js-added from VB3
txtComStatus = ComStatus.Text 'js-2/12/99
'---

End Sub





Public Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError
Screen.MousePointer = HOURGLASS

If glbtermopen Then  'Lucy July 4, 2000
    SQLQ = "Select * from Term_HRDEPEND "
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "Select * from HRDEPEND"
    SQLQ = SQLQ & " where DP_EMPNBR = " & glbLEE_ID
End If

Data1.RecordSource = SQLQ
Data1.Refresh

EERetrieve = True

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DEPRetrieve", "HRDEPEND", "SELECT")
Call RollBack

Exit Function

End Function

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

If ErrorNumber = 3021 Then  ' no record present on a close
    'Response = 0
    ErrorNumber = 0
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRDEPNTS", "SELECT")
End If

End Sub

Private Sub Form_Activate()
    glbOnTop = "FRMDEPNDTS"
    If Not Data1.Recordset.EOF Then 'ADD BY FRANK 05/29/01 FOR TRUE DBGRID DISPLAY (FIRST TIME OPEN FORM)
        If Data1.Recordset("DP_SEX") = "M" Then
          optSex(0) = True
        Else
            optSex(1) = True
        End If
    End If
Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMDEPNDTS"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

'~~~24June99 js-added from VB3~~~
If glbCompSerial = "S/N - 2219W" Or glbCompSerial = "S/N - 2274W" Then   'js-2/11/99-added condition
'  txtDeptNo.Visible = True              '
'  txtDental.Visible = True              '
'  txtMedical.Visible = True             '
'  txtOther.Visible = True               '
'
'  lblTitle(10).Visible = True           '
'  lblTitle(11).Visible = True           '
'  lblTitle(12).Visible = True           '
'  lblTitle(13).Visible = True           '
  txtDeptNo.Text = "0000"               '
End If                                  '
'If glbCompSerial = "S/N - 2375W" Then   'TIMMIS
    memComments.Visible = True
    lblTitle(14).Visible = True
    memComments.DataField = "DP_COMMENTS"
'End If                                  '

'~~~~~~
glbOnTop = "FRMDEPNDTS"
'If glbtermopen Then
'Data1.ConnectionString = glbAdoIHRAUDIT
'Else
'Data1.ConnectionString = glbAdoIHRDB
'End If

xAction = " "  '24June99 js - added from VB3

If glbCompSerial = "S/N - 2375W" Then   'City of Timmis
    lblTitle(6).Caption = "COB"
    ComSmoker.Tag = "00-COB Yes/No"
End If

If glbtermopen Then  'Lucy July 4, 2000
    Data1.ConnectionString = glbAdoIHRAUDIT
    Data1.RecordSource = "Term_HRDEPEND"
Else
    Data1.ConnectionString = glbAdoIHRDB
    Data1.RecordSource = "HRDEPEND"
End If

Screen.MousePointer = HOURGLASS
'Call setCaption(lblTitle(2))
Call LocCaptions

Screen.MousePointer = DEFAULT
If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If
lblEENum.Caption = ShowEmpnbr(lblEEID)

If glbWFC Then 'Ticket #16287
    If glbEmpCountry = "CANADA" Then
        lblTitle(3).FontBold = True
        lblTitle(4).FontBold = True
        lblTitle(5).FontBold = True
        lblTitle(6).FontBold = True
        lblTitle(8).FontBold = True
        lblTitle(10).FontBold = True
        
        'Ticket #17284 - begin
        'lblTitle(11).FontBold = True
        'lblTitle(12).FontBold = True
        'lblTitle(13).FontBold = True
        lblWFCCobOther.Visible = True
        comDental.Clear
        comDental.AddItem "1-COB Spouse Only"
        comDental.AddItem "2-COB Spouse and Employee Only"
        comDental.AddItem "3-COB Spouse, Employee and Dependants"
        comDental.AddItem "4-COB Spouse and Dependants Only"
        comDental.AddItem "N-NO COB coverage"
        comDental.AddItem ""
        comDental.Left = txtDental.Left
        comDental.Visible = True

        comMedical.Clear
        comMedical.AddItem "1-COB Spouse Only"
        comMedical.AddItem "2-COB Spouse and Employee Only"
        comMedical.AddItem "3-COB Spouse, Employee and Dependants"
        comMedical.AddItem "4-COB Spouse and Dependants Only"
        comMedical.AddItem "N-NO COB coverage"
        comMedical.AddItem ""
        comMedical.Left = txtMedical.Left
        comMedical.Visible = True
        
        comOther.Clear
        comOther.AddItem "1-COB Spouse Only"
        comOther.AddItem "2-COB Spouse and Employee Only"
        comOther.AddItem "3-COB Spouse, Employee and Dependants"
        comOther.AddItem "4-COB Spouse and Dependants Only"
        comOther.AddItem "N-NO COB coverage"
        comOther.AddItem ""
        comOther.Left = txtOther.Left
        comOther.Visible = True
        
        'Ticket #17284 - end
        
        'Ticket #22108 Franks 07/12/2012
        lblWFCTxt2.Visible = True 'canada only
        
        'Ticket #22411 Franks 08/08/2012
        imgHelp.Visible = True 'canada only
    End If
End If

'Ticket #22417 - Goodmans LLP
If glbCompSerial = "S/N - 2290W" Then
    comDeptTxt1.Clear
    comDeptTxt1.AddItem "1-Health & Dental"
    comDeptTxt1.AddItem "2-Health Only"
    comDeptTxt1.AddItem "3-Dental Only"
    comDeptTxt1.AddItem "4-No Health or Dental"
    comDeptTxt1.Left = txtUserText(1).Left
    comDeptTxt1.Top = txtUserText(1).Top
    comDeptTxt1.Visible = True

    comDental.Clear
    comDental.AddItem "N-No COB Coverage"
    comDental.AddItem "Y-COB Coverage"
    comDental.Left = txtDental.Left
    comDental.Visible = True

    comMedical.Clear
    comMedical.AddItem "N-No COB Coverage"
    comMedical.AddItem "Y-COB Coverage"
    comMedical.Left = txtMedical.Left
    comMedical.Visible = True
    
    'Ticket #22745
    lblTitle(15).FontBold = True
End If

Call LdcomRel

Call SetMasks  '
'lblSex.Caption = Data1.Recordset("DP_SEX")
Screen.MousePointer = HOURGLASS
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    frmDEPNDTS.Caption = "Dependents - " & Left$(glbLEE_SName, 5)
    frmDEPNDTS.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    If glbLinamar Then  'Ticket #14775
        frmDEPNDTS.lblEEProdLine = glbLEE_ProdLine
    Else
        frmDEPNDTS.lblEEProdLine = ""
    End If
End If
Call ST_UPD_MODE(False)
If Not gSec_Upd_Dependents Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
End If
    
'If show SIN
If glbCompSerial = "S/N - 2344W" Then 'cascade see ticket #5515
    vbxTrueGrid.Columns(4).Caption = "Green Shield"
Else
    If Not gSec_Show_SIN_SSN Then
        MedSIN.Visible = False
        vbxTrueGrid.Columns(4).Visible = False
    End If
End If
Call INI_Controls(Me)
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
Call Display_Value
Screen.MousePointer = DEFAULT

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



Private Sub Label1_Click()

End Sub

Private Sub imgHelp_Click()
Dim MsgStr As String
    'Ticket #22417 - Goodmans LLP
    If glbCompSerial = "S/N - 2290W" Then
        MsgStr = "This date field must be entered when " & lblTitle(11).Caption & " or " & lblTitle(12).Caption & ", is 'Y - COB Coverage'."
    Else
        MsgStr = "For COB Other, this date field must be completed. The field must contain the date of birth of the child's birth mother or father to determine who is the primary and secondary payer"
    End If
    MsgBox MsgStr, vbInformation
End Sub

Private Sub lblEEID_Change()
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    Me.Caption = "Dependents - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblEENum.Caption = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub

Private Sub lblSex_Change()
    If lblSex.Caption = "M" Then  'change added by RAUBREY 4/2/97
      optSex(0) = True
    End If
    If lblSex.Caption = "F" Then
        optSex(1) = True
    End If
End Sub

Private Sub medSIN_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub
Private Sub optSex_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optSex_Click(Index As Integer) 'js

If optSex(0) Then
    lblSex.Caption = "M"
Else
    lblSex.Caption = "F"
End If

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

'cmdModify.Enabled = FT
'cmdDelete.Enabled = FT
'cmdNew.Enabled = FT
'cmdCancel.Enabled = TF
'cmdOK.Enabled = TF
'cmdClose.Enabled = FT
'cmdPrint.Enabled = FT

comRelation.Enabled = TF
ComSmoker.Enabled = TF
ComStatus.Enabled = TF
MedSIN.Enabled = TF
optSex(1).Enabled = TF
optSex(0).Enabled = TF
dlpDOB.Enabled = TF
txtFName.Enabled = TF
txtSurname.Enabled = TF
dlpEligDte.Enabled = TF
dlpEndDte.Enabled = TF
memComments.Enabled = TF
'vbxTrueGrid.Enabled = FT
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
 '   cmdModify.Enabled = False
 '   cmdDelete.Enabled = False
End If
'If glbCompSerial = "S/N - 2219W" Or glbCompSerial = "S/N - 2274W" Then
    txtDeptNo.Enabled = TF
    txtDental.Enabled = TF
    txtMedical.Enabled = TF
    txtOther.Enabled = TF
    txtUserText(1).Enabled = TF
    txtUserText(2).Enabled = TF
    txtUserText(3).Enabled = TF
    txtUserText(4).Enabled = TF
'End If

'Ticket #22417 - Goodmans LLP
If glbCompSerial = "S/N - 2290W" Then
    If txtDental.Text = "" Then
        comDental.ListIndex = -1
    End If
    If txtMedical.Text = "" Then
        comMedical.ListIndex = -1
    End If
    If txtUserText(1).Text = "" Then
        comDeptTxt1.ListIndex = -1
    End If
    comDental.Enabled = TF
    comMedical.Enabled = TF
    comDeptTxt1.Enabled = TF
End If
End Sub

Private Sub txtDental_Change()
'---24June99 js - added from VB3
Dim lower

If Len(Trim(txtDental.Text)) <> 0 Then
  lower = (txtDental.Text)
  txtDental.Text = UCase(lower)
Else
    txtDental.Text = ""
End If
'---
If glbWFC And glbEmpCountry = "CANADA" Then 'Ticket #17284
    If txtDental.Text = "1" Then comDental.ListIndex = 0
    If txtDental.Text = "2" Then comDental.ListIndex = 1
    If txtDental.Text = "3" Then comDental.ListIndex = 2
    If txtDental.Text = "4" Then comDental.ListIndex = 3
    If txtDental.Text = "N" Then comDental.ListIndex = 4
    If txtDental.Text = "" Then comDental.ListIndex = 5
End If

'Ticket #22417 - Goodmans LLP
If glbCompSerial = "S/N - 2290W" Then
    If txtDental.Text = "N" Then comDental.ListIndex = 0
    If txtDental.Text = "Y" Then comDental.ListIndex = 1
End If

End Sub

Private Sub txtDental_GotFocus()

'~~24June99 js - added from VB3
  OtxtDental = txtDental.Text
  Call SetPanHelp(ActiveControl) 'js-15Mar99-added
'~~

End Sub

Private Sub txtDeptNo_Change()

'~~24June99 js - added VB3
  OtxtDeptNo = txtDeptNo.Text
'~~

End Sub


Private Sub txtDeptNo_GotFocus()
    Call SetPanHelp(ActiveControl) 'js-15Mar99-added for panel description
End Sub
'Private Sub txtDOB_Change()
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtDOB_DblClick()
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub txtDOB_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub txtDOB_KeyPress(KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub
'Private Sub txtEligDte_Change()
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtEligDte_DblClick()
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub txtEligDte_GotFocus()
'    Call SetPanHelp(ActiveControl) 'js--1/8/99 - added code for panhelp
'                                   '             description
'End Sub
'Private Sub txtEligDte_KeyPress(KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub
'Private Sub txtEndDte_Change()
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtEndDte_DblClick()
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub txtEndDte_GotFocus()
'    Call SetPanHelp(ActiveControl) 'js--1/8/99 - added code for panhelp
'                                   '             description
'End Sub
'Private Sub txtEndDte_KeyPress(KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub
Private Sub txtFName_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtGender_Change()

End Sub

Private Sub txtMedical_Change()

'~~24June99 js - added from VB3
Dim lower

If Len(Trim(txtMedical.Text)) <> 0 Then
  lower = (txtMedical.Text)
  txtMedical.Text = UCase(lower)
Else
    txtMedical.Text = ""
End If
'~~

If glbWFC And glbEmpCountry = "CANADA" Then 'Ticket #17284
    If txtMedical.Text = "1" Then comMedical.ListIndex = 0
    If txtMedical.Text = "2" Then comMedical.ListIndex = 1
    If txtMedical.Text = "3" Then comMedical.ListIndex = 2
    If txtMedical.Text = "4" Then comMedical.ListIndex = 3
    If txtMedical.Text = "N" Then comMedical.ListIndex = 4
    If txtMedical.Text = "" Then comMedical.ListIndex = 5
End If

'Ticket #22417 - Goodmans LLP
If glbCompSerial = "S/N - 2290W" Then
    If txtMedical.Text = "N" Then comMedical.ListIndex = 0
    If txtMedical.Text = "Y" Then comMedical.ListIndex = 1
End If

End Sub

Private Sub txtMedical_GotFocus()

'24June99 js-added from VB3
  Call SetPanHelp(ActiveControl) 'js-15Mar99-added

End Sub

Private Sub txtOther_Change()
'~~24June99 js-added from VB3
Dim lower

If Len(Trim(txtOther.Text)) <> 0 Then
  lower = (txtOther.Text)
  txtOther.Text = UCase(lower)
Else
    txtOther.Text = ""
End If
'~~

If glbWFC And glbEmpCountry = "CANADA" Then 'Ticket #17284
    If txtOther.Text = "1" Then comOther.ListIndex = 0
    If txtOther.Text = "2" Then comOther.ListIndex = 1
    If txtOther.Text = "3" Then comOther.ListIndex = 2
    If txtOther.Text = "4" Then comOther.ListIndex = 3
    If txtOther.Text = "N" Then comOther.ListIndex = 4
    If txtOther.Text = "" Then comOther.ListIndex = 5
End If

End Sub

Private Sub txtOther_GotFocus()
'~~24June99 js-added from VB3
  Call SetPanHelp(ActiveControl) 'js-15Mar99-added
End Sub

Private Sub txtComRelation_Change()
    comRelation.Text = txtComRelation.Text
    
If glbWFC And glbEmpCountry = "CANADA" Then 'Ticket #13448
    If txtComRelation.Text = "Child" Then comRelation.ListIndex = 0
    If txtComRelation.Text = "Spouse" Then comRelation.ListIndex = 1
    'Ticket #22009 Franks 05/11/2012 - remove Husband and Wife
    'If txtComRelation.Text = "Husband" Then comRelation.ListIndex = 1
    'If txtComRelation.Text = "Spouse" Then comRelation.ListIndex = 2
    'If txtComRelation.Text = "Wife" Then comRelation.ListIndex = 3
ElseIf glbCompSerial = "S/N - 2290W" Then   'Ticket #22745 - Goodmans
    If txtComRelation.Text = "Child" Then comRelation.ListIndex = 0
Else
    If txtComRelation.Text = "Aunt" Then comRelation.ListIndex = 0
    If txtComRelation.Text = "Brother" Then comRelation.ListIndex = 1
    If txtComRelation.Text = "Children" Then comRelation.ListIndex = 2
    If txtComRelation.Text = "Common Law" Then comRelation.ListIndex = 3
    If txtComRelation.Text = "Couple" Then comRelation.ListIndex = 4
    If txtComRelation.Text = "Daughter" Then comRelation.ListIndex = 5
    If txtComRelation.Text = "Estate" Then comRelation.ListIndex = 6
    If txtComRelation.Text = "Ex-Spouse" Then comRelation.ListIndex = 7
    If txtComRelation.Text = "Father" Then comRelation.ListIndex = 8
    If txtComRelation.Text = "Fiancee" Then comRelation.ListIndex = 9
    If txtComRelation.Text = "Fiance" Then comRelation.ListIndex = 10
    If txtComRelation.Text = "Husband" Then comRelation.ListIndex = 11
    If txtComRelation.Text = "Mother" Then comRelation.ListIndex = 12
    If txtComRelation.Text = "Other" Then comRelation.ListIndex = 13
    If txtComRelation.Text = "Parents" Then comRelation.ListIndex = 14
    If txtComRelation.Text = "Sister" Then comRelation.ListIndex = 15
    If txtComRelation.Text = "Son" Then comRelation.ListIndex = 16
    If txtComRelation.Text = "Spouse" Then comRelation.ListIndex = 17
    If txtComRelation.Text = "Uncle" Then comRelation.ListIndex = 18
    If txtComRelation.Text = "Wife" Then comRelation.ListIndex = 19
End If
End Sub

Private Sub txtSurname_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtUserText_Change(Index As Integer)
    'Ticket #22417 - Goodmans LLP
    If glbCompSerial = "S/N - 2290W" And Index = 1 Then
        If txtUserText(1).Text = "1" Then comDeptTxt1.ListIndex = 0
        If txtUserText(1).Text = "2" Then comDeptTxt1.ListIndex = 1
        If txtUserText(1).Text = "3" Then comDeptTxt1.ListIndex = 2
        If txtUserText(1).Text = "4" Then comDeptTxt1.ListIndex = 3
        If txtUserText(1).Text = "N" Then comDeptTxt1.ListIndex = 4
    End If
End Sub

Private Sub txtUserText_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
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
        
        If glbtermopen Then  'Lucy July 4, 2000
            SQLQ = "Select * from Term_HRDEPEND "
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select * from HRDEPEND"
            SQLQ = SQLQ & " where DP_EMPNBR = " & glbLEE_ID
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

'''Sam add July 02 * Remove ADO
Call Display_Value

'ComSmoker.ListIndex = 0
'ComStatus.ListIndex = 0  'move ahead of the exit sub so there is a default SBH Aug 04 1998...
    
If Data1.Recordset.RecordCount <> 0 Then 'SBH removed exit sub and moved from row change...
    If IsNull(Data1.Recordset("DP_SMOKER")) Then
        ComSmoker.ListIndex = 0
    Else
        If Data1.Recordset("DP_SMOKER") Then
            ComSmoker.ListIndex = 1
        Else        'jaddy 11/3/99
            ComSmoker.ListIndex = 0        'jaddy 11/3/99
        End If        'jaddy 11/3/99
    End If
    ComStatus.Text = "Other()"
    If glbWFC And glbEmpCountry = "CANADA" Then
        If Data1.Recordset("DP_STATUS") = "" Then ComStatus.Text = "Other( )"
        If Data1.Recordset("DP_STATUS") = "S" Then ComStatus.Text = "Overage Student(S)"       'jaddy 11/3/99
        If Data1.Recordset("DP_STATUS") = "H" Then ComStatus.Text = "Handicap(H)"        'jaddy 11/3/99
        If Data1.Recordset("DP_STATUS") = "M" Then ComStatus.Text = "Twin/Multiple Birth(M)"        'jaddy 11/3/99
    Else
        If Data1.Recordset("DP_STATUS") = "1" Then ComStatus.Text = "Student(1)"        'jaddy 11/3/99
        If Data1.Recordset("DP_STATUS") = "2" Then ComStatus.Text = "Disability(2)"        'jaddy 11/3/99
        If Data1.Recordset("DP_STATUS") = "Y" Then ComStatus.Text = "Fast Yes(Y)"        'jaddy 11/3/99
        If Data1.Recordset("DP_STATUS") = "N" Then ComStatus.Text = "Fast No(N)"        'jaddy 11/3/99
        If Data1.Recordset("DP_STATUS") = "3" Then ComStatus.Text = "Fast(3)"        'jaddy 11/3/99
        If Data1.Recordset("DP_STATUS") = "S" Then ComStatus.Text = "Fast(S)"        'jaddy 11/3/99
    End If
End If
End Sub
'~~23June99 - js - added from VB3
Function AUDITDEPNTS()   'js-2/12/99-added function
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD, xDiv, xPT
Dim xDependNum, SQLQ
Dim UpdateAudit As Boolean

On Error GoTo AUDIT_ERR
AUDITDEPNTS = False

'Ticket #13520 - Begin
UpdateAudit = False
If OtxtFName <> txtFName.Text Then UpdateAudit = True
If OtxtSurname <> txtSurname.Text Then UpdateAudit = True
If OOptSex <> lblSex.Caption Then UpdateAudit = True
If IsDate(dlpDOB.Text) Then
    If OtxtDOB <> dlpDOB.Text Then UpdateAudit = True
End If                              '
If OtxtRelation <> comRelation.Text Then UpdateAudit = True
If OtxtComSmoker <> ComSmoker.Text Then UpdateAudit = True
If OtxtComStatus <> ComStatus.Text Then UpdateAudit = True
If OmedSIN <> MedSIN Then UpdateAudit = True
If IsDate(dlpEligDte.Text) Then
    If OtxtEligDate <> dlpEligDte.Text Then UpdateAudit = True
End If
If IsDate(dlpEndDte.Text) Then
    If OtxtEndDte <> dlpEndDte.Text Then UpdateAudit = True
End If
If OtxtDeptNo <> txtDeptNo Then UpdateAudit = True
If OtxtDental <> txtDental Then UpdateAudit = True
If OtxtMedical <> txtMedical Then UpdateAudit = True
If OtxtOther <> txtOther Then UpdateAudit = True

If Not UpdateAudit Then GoTo MODNOUPD
'Ticket #13520 - End

rsTB.Open "select ED_DIV,ED_PT FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
  'xDiv = rsTB("ED_DIV")
  'xPT = rsTB("ED_PT")
    If IsNull(rsTB("ED_PT")) Then
        xPT = ""
    Else
        xPT = rsTB("ED_PT")
    End If
    If IsNull(rsTB("ED_DIV")) Then
        xDiv = ""
    Else
        xDiv = rsTB("ED_DIV")
    End If
Else
  xDiv = ""
  xPT = ""
End If
'Frank Sep 29, 2006, Jerry asked from WFC to comment out
'If glbCompSerial = "S/N - 2373W" Then  'Distric of Muskoka
    'Get the total number of dependents
    SQLQ = "SELECT DP_EMPNBR FROM HRDEPEND WHERE DP_EMPNBR = " & glbLEE_ID
    rsTA.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xDependNum = 0
    If Not rsTA.EOF Then
        rsTA.MoveLast
        rsTA.MoveFirst
        xDependNum = rsTA.RecordCount
    End If
    rsTA.Close
'End If
Dim strFields As String
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, AU_DOLENT_TABL, "
strFields = strFields & "AU_EARN_TABL, AU_EARN_TABL, AU_NEWEMP, AU_DEPFNAME, AU_DEPSNAME, AU_DEPSEX, AU_DEPDOB, AU_DEPRELATE, AU_DEPSMOKER, AU_DEPSTATUS, "
strFields = strFields & "AU_DEPSIN, AU_DEPSDATE, AU_DEPNO, AU_DENTAL, AU_MEDICAL, AU_OTHER, AU_COMPNO, AU_EMPNBR, AU_DIVUPL, AU_PTUPL, AU_DEPEDATE, "
strFields = strFields & "AU_LDATE, AU_TYPE, AU_DEPEND_NUM, AU_LUSER, AU_LTIME, AU_UPLOAD,AU_PAYROLL_ID "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False
  
rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR"
    rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_NEWEMP") = "N"
    
    Dim rsEmp As New ADODB.Recordset
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_DOH FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
    End If
    'rsEmp.Close
    
    'If OtxtFName <> txtFName.Text Then
    '    rsTA("AU_DEPFNAME") = txtFName
    'End If
    'If OtxtSurname <> txtSurname.Text Then
    '    rsTA("AU_DEPSNAME") = txtSurname
    'End If
    'Always Pass Surname and First Name since they are the key (WFC Manulife needs it) Frank Ticket #13448
    rsTA("AU_DEPFNAME") = txtFName
    rsTA("AU_DEPSNAME") = txtSurname
    If OOptSex <> lblSex.Caption Then
        If lblSex.Caption = "M" Then
            rsTA("AU_DEPSEX") = "M"
        Else
            rsTA("AU_DEPSEX") = "F"
        End If
    End If
    
    'If OtxtDOB <> txtDOB Then
    If IsDate(dlpDOB.Text) Then            '11Aug99 js
        If OtxtDOB <> dlpDOB.Text Then       '
            rsTA("AU_DEPDOB") = dlpDOB.Text  '
        End If                          '
    Else                                '
        rsTA("AU_DEPDOB") = Null        '
    End If                              '
    
    If OtxtRelation <> comRelation.Text Then
        rsTA("AU_DEPRELATE") = comRelation.Text
    End If
    If OtxtComSmoker <> ComSmoker.Text Then
        If ComSmoker.Text = "Yes" Then
            rsTA("AU_DEPSMOKER") = "-1"
        Else
            rsTA("AU_DEPSMOKER") = "0"
        End If
    End If
    
    If OtxtComStatus <> ComStatus.Text Then
        If ComStatus.Text = "Other( )" Then
            rsTA("AU_DEPSTATUS") = " "
        End If
        If glbWFC And glbEmpCountry = "CANADA" Then
            If ComStatus.Text = "Other( )" Then rsTA("AU_DEPSTATUS") = ""
            If ComStatus.Text = "Overage Student(S)" Then rsTA("AU_DEPSTATUS") = "S"
            If ComStatus.Text = "Handicap(H)" Then rsTA("AU_DEPSTATUS") = "H"
            If ComStatus.Text = "Twin/Multiple Birth(M)" Then rsTA("AU_DEPSTATUS") = "M"
        Else
            If ComStatus.Text = "Student(1)" Then rsTA("AU_DEPSTATUS") = "1"
            If ComStatus.Text = "Disability(2)" Then rsTA("AU_DEPSTATUS") = "2"
            If ComStatus.Text = "Fast Yes(Y)" Then rsTA("AU_DEPSTATUS") = "Y"
            If ComStatus.Text = "Fast No(N)" Then rsTA("AU_DEPSTATUS") = "N"
            If ComStatus.Text = "Fast(3)" Then rsTA("AU_DEPSTATUS") = "3"
            If ComStatus.Text = "Fast(S)" Then rsTA("AU_DEPSTATUS") = "S"
        End If
    End If
    If OmedSIN <> MedSIN Then
        rsTA("AU_DEPSIN") = MedSIN
    End If
    
    If IsDate(dlpEligDte.Text) Then
        If OtxtEligDate <> dlpEligDte.Text Then
            rsTA("AU_DEPSDATE") = dlpEligDte.Text
        End If
    Else
        rsTA("AU_DEPSDATE") = Null
    End If
    
    If IsDate(dlpEndDte.Text) Then
        If OtxtEndDte <> dlpEndDte.Text Then
            rsTA("AU_DEPEDATE") = dlpEndDte.Text
        End If
    Else
        rsTA("AU_DEPEDATE") = Null
    End If
        
    'If glbCompSerial = "S/N - 2219W" Or glbCompSerial = "S/N - 2274W" Then
          If OtxtDeptNo <> txtDeptNo Then
            rsTA("AU_DEPNO") = txtDeptNo
          End If
          If OtxtDental <> txtDental Then
            rsTA("AU_DENTAL") = txtDental
          End If
          If OtxtMedical <> txtMedical Then
            rsTA("AU_MEDICAL") = txtMedical
          End If
          If OtxtOther <> txtOther Then
            rsTA("AU_OTHER") = txtOther
          End If
    'End If
    
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_DIVUPL") = xDiv
    rsTA("AU_PTUPL") = xPT
    If glbWFC Then
        rsTA("AU_DEPEND_NUM") = xDependNum + 1
    End If
    rsTA("AU_LDATE") = Date 'Ticket #14582 Default the date to today, otherwise the au_ldate maybe is null
    If xAction = "A" Then ' NEW                     'js-9Mar99
      If Len(dlpEligDte.Text) > 0 Then              '
          If CVDate(dlpEligDte.Text) < CVDate(Date) Then 'Ticket #14769
            rsTA("AU_LDATE") = Date
          Else
            rsTA("AU_LDATE") = dlpEligDte.Text        '
          End If
      Else                                          '
          rsTA("AU_LDATE") = Date '
      End If                                        '
      rsTA("AU_TYPE") = "A"                         '
      'If glbCompSerial = "S/N - 2373W" Then  'Distric of Muskoka
        rsTA("AU_DEPEND_NUM") = xDependNum + 1
      'End If
    End If                                          '
    If xAction = "C" Then   'EDIT                   '
      If Len(dlpEligDte.Text) > 0 Then              '
        If OtxtEligDate <> dlpEligDte.Text Then          '
          rsTA("AU_LDATE") = dlpEligDte.Text        '
        End If                                      '
      Else                                          '
        rsTA("AU_LDATE") = Date '
      End If                                        '
      rsTA("AU_TYPE") = "M"                         '
    End If                                          '
    If xAction = "D" Then 'Delete                   '
      rsTA("AU_LDATE") = Date  '
      rsTA("AU_TYPE") = "D"                         '
      'If glbCompSerial = "S/N - 2373W" Then  'Distric of Muskoka
        'If xDependNum > 0 Then
            rsTA("AU_DEPEND_NUM") = xDependNum - 1
        'End If
      'End If
    End If                                          '
    'Ticket #23407 Franks 03/15/2013 - begin
    If IsDate(rsEmp("ED_DOH")) Then
        If CVDate(rsEmp("ED_DOH")) > CVDate(Date) Then
            rsTA("AU_LDATE") = rsEmp("ED_DOH")
        End If
    End If
    rsEmp.Close
    'Ticket #23407 Franks 03/15/2013 - end
    rsTA("AU_LUSER") = glbUserID
    If xAction <> "A" And xAction <> "C" And xAction <> "D" Then rsTA("AU_LDATE") = Date
    If glbWFC Then 'For Manulife Transaction, if Benefit End Date enterred, send it to Manulife after this date
        If IsDate(dlpEndDte.Text) Then
            If CVDate(dlpEndDte.Text) > CVDate(Date) Then 'Ticket #21979 Franks 05/01/2012
                If OtxtEndDte <> dlpEndDte.Text Then
                    rsTA("AU_LDATE") = dlpEndDte.Text
                End If
            End If
        End If
    End If
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
rsTA.Update

GoTo MODNOUPD

MODNOUPD:
AUDITDEPNTS = True
Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack

'~~
End Function

Sub SetMasks()
If glbCompSerial = "S/N - 2344W" Then  'cascade see ticket #5515
   lblTitle(7) = "Green Shield" '"S.I.N"                   '
   MedSIN.Tag = "10-Green Shield" '
   MedSIN.MaxLength = 9
   MedSIN.Mask = "&&&&&&&&&"

Else
    If glbEmpCountry = "CANADA" Then                   'laura Oct 28, 1997
        lblTitle(7) = "S.I.N"                       '
        MedSIN.Tag = "10-Social Insurance Number"   '
        MedSIN.MaxLength = 11
        MedSIN.Mask = "###-###-###"
    ElseIf glbEmpCountry = "BAHAMAS" Then              '
        lblTitle(7) = "National Ins."               '
        MedSIN.Tag = "10-National Insurance Number" '
        MedSIN.MaxLength = 11
        MedSIN.Mask = "########"
    ElseIf glbEmpCountry = "U.S.A." Then
        lblTitle(7) = "S.S.N"                   '
        MedSIN.Tag = "10-Social Security Number" '
        MedSIN.MaxLength = 11
        MedSIN.Mask = "###-##-####"
    ElseIf glbEmpCountry = "MEXICO" Then                '
        lblTitle(7) = "National Ins."                   '
        MedSIN.Tag = "10-National Insurance Number" '
        MedSIN.MaxLength = 15
        MedSIN.Mask = "&&&&&&&&&&&&&&&"
    ElseIf glbEmpCountry = "GERMANY" Then                '
        lblTitle(7) = "National Ins."                   '
        MedSIN.Tag = "10-National Insurance Number" '
        MedSIN.MaxLength = 15
        MedSIN.Mask = "&&&&&&&&&&&&&&&"
    Else
        lblTitle(7) = "National Ins." '"S.I.N"                   '
        MedSIN.Tag = "10-Social Insurance Number" '
        MedSIN.MaxLength = 15 ' 11
        MedSIN.Mask = "&&&&&&&&&&&&&&&"
    End If
End If
End Sub

Private Function chk_FeDepnts() '25June99 js-added function
Dim mSIN As String  '24June99 js -added from VB3
Dim xYear As Double
Dim xlocDays As Integer
Dim xMsg As String
Dim a As Integer, Msg As String, x
Dim tmpFlag As Boolean

On Error GoTo FeDepnts_ERR

chk_FeDepnts = False

If Len(txtSurname.Text) <= 0 Then
    MsgBox "You must enter the dependent's last name."
    txtSurname.SetFocus
    Exit Function
End If

If Len(txtFName.Text) <= 0 Then
    MsgBox lStr("You must enter the dependent's first name.")
    txtFName.SetFocus
    Exit Function
End If

If Len(dlpDOB.Text) > 0 Then
    If Not IsDate(dlpDOB.Text) Then
        MsgBox "Not a valid Birth Date"
        dlpDOB.SetFocus
        Exit Function
    End If
End If

'Ticket #24338 - Date Validation
If IsDate(dlpDOB.Text) Then
    'If CVDate(Format(dlpDOB, "mm/dd/yyyy")) > CVDate(Format(Now, "mm/dd/yyyy")) Then
    'Ticket #26814 Franks 03/16/2015 - the function above not work for date format dd/MM/yyyy
    If CVDate(dlpDOB.Text) > CVDate(Now) Then
        MsgBox "Birth Date cannot be greater than today"
        dlpDOB.SetFocus
        Exit Function
    End If
End If

'~~~~~24June99 js -added from VB3~~~~~
If Len(dlpEndDte.Text) > 0 Then
    If Not IsDate(dlpEndDte.Text) Then
        MsgBox "Not a valid Benefit End Date"
        dlpEndDte.SetFocus
        Exit Function
    End If
End If

If Len(dlpEligDte.Text) > 0 Then
    If Not IsDate(dlpEligDte.Text) Then
        MsgBox "Not a valid Benefit Eligibility Date"
        dlpEligDte.Text = ""
        dlpEligDte.SetFocus
    Exit Function
        'commented out by request from LINDA - 19July99 js
    'Else
        '  dlpEndDte = DateAdd("yyyy", 19, dlpDOB) 'js-26Apr99
    End If
Else
    If Len(dlpEligDte.Text) = 0 Then
        'Ticket #22854 Franks 01/08/2013 - begin
        'dlpEndDte.Text = ""
        If glbCompSerial = "S/N - 2290W" Then 'Goodmans
            MsgBox "Manulife requires an Eligibility Date for dependents." & Chr(10) & " If this employee qualifies for Manulife, please enter a date."
            dlpEligDte.SetFocus
            Exit Function
        End If
        If IsDate(dlpEndDte.Text) Then
            MsgBox "Eligible Date must be entered in order to save an End Date."
            dlpEligDte.SetFocus
            Exit Function
        End If
        'Ticket #22854 Franks 01/08/2013 - end
    End If
End If

If glbCompSerial <> "S/N - 2344W" Then  'cascade see ticket #5515
    If Len(MedSIN) > 0 Then
        If glbEmpCountry = "BAHAMAS" Then
            If Len(MedSIN) <> 8 Then
                MsgBox "Invalid National Ins - if Unassigned set to 99999999"
                MedSIN.SetFocus
                Exit Function
            End If
        Else
            If MedSIN <> "999999999" Then
                mSIN = CStr(MedSIN)
                If glbEmpCountry = "CANADA" Then
                    If Not SIN_chk(mSIN) Then
                        MsgBox "Invalid SIN (999-999-999)"
                        Exit Function
                    End If
                End If
                If glbEmpCountry = "U.S.A." Then
                    If Len(MedSIN) >= 1 Then
                        If Not SIN_chk_USA(mSIN) Then
                            MsgBox "Invalid SSN - (999-99-9999)"
                            MedSIN.SetFocus
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

'If glbCompSerial = "S/N - 2219W" Or glbCompSerial = "S/N - 2274W" Then   'by jgr 3/19/99
    If Not IsNumeric(txtDeptNo.Text) And Len(txtDeptNo.Text) > 0 Then ' Or Val(txtDeptNo.Text) = 0 Then
        MsgBox "Please enter numeric dependent information."
        '-------Jaddy 8/5/99
        'If OtxtDeptNo <> 0 Then
        '    txtDeptNo.Text = OtxtDeptNo
        'Else
        '    txtDeptNo.Text = 0   'by jgr 3/19/99
        'End If
        '------
        txtDeptNo.SetFocus
        Exit Function
    Else

        '-------Jaddy 8/5/99
        'If Len(txtDeptNo.Text) < 4 Then
        '    MsgBox "Please enter a 4 digit numeric value for dependent information."
        '    txtDeptNo.Text = OtxtDeptNo
        '    txtDeptNo.SetFocus
        'End If
        '-------
    End If
'End If      'by jgr 3/19/99

If Not optSex(0) And Not optSex(1) Then
    MsgBox "Must Enter either Male or Female"
    optSex(0).SetFocus
    Exit Function
End If

'Ticket #14031
If glbWFC Then
    If ComStatus.Text = "Overage Student(S)" Then
        If IsDate(dlpDOB.Text) Then
            xYear = DateDiff("d", dlpDOB.Text, Date)
            xYear = Round(xYear / 365, 1)
            If xYear < 18 Then
                MsgBox "If the status is Overage Student(S), he/she must be over 18 years old."
                ComStatus.SetFocus
                Exit Function
            End If
        End If
    End If
    If Len(txtSurname.Text) > 0 Then 'Ticket #14154
        If Len(InvalidCharInStr(txtSurname.Text, glbWFCNameChars)) > 0 Then
                MsgBox "Invalid character '" & InvalidCharInStr(txtSurname.Text, glbWFCNameChars) & "' in name field. "
                txtSurname.SetFocus
                Exit Function
        End If
    End If
    If Len(txtFName.Text) > 0 Then 'Ticket #14154
        If Len(InvalidCharInStr(txtFName.Text, glbWFCNameChars)) > 0 Then
                MsgBox "Invalid character '" & InvalidCharInStr(txtFName.Text, glbWFCNameChars) & "' in name field. "
                txtFName.SetFocus
                Exit Function
        End If
    End If

    If glbEmpCountry = "CANADA" Then 'Ticket #16287
        If Len(dlpDOB.Text) = 0 Then
            MsgBox "Date of Birth is Mandatory field"
            dlpDOB.SetFocus
            Exit Function
        End If
        If Len(dlpEligDte.Text) = 0 Then
            MsgBox "Benefit Eligible Date is Mandatory field"
            dlpEligDte.SetFocus
            Exit Function
        End If
        If Len(txtDeptNo.Text) = 0 Then
            MsgBox "Dependent Number is Mandatory field"
            txtDeptNo.SetFocus
            Exit Function
        End If
        'Ticket #17284
        'If Len(txtDental.Text) = 0 Then
        '    MsgBox "COB Dental is Mandatory field"
        '    txtDental.SetFocus
        '    Exit Function
        'End If
        'If Len(txtMedical.Text) = 0 Then
        '    MsgBox "COB Medical is Mandatory field"
        '    txtMedical.SetFocus
        '    Exit Function
        'End If
        'If Len(txtOther.Text) = 0 Then
        '    MsgBox "COB Other is Mandatory field"
        '    txtOther.SetFocus
        '    Exit Function
        'End If

        'Ticket #22009 Franks 05/11/2012 - begin
        '"   Can't enter a new SPOUSE record unless the old Spouse has a Benefit End Date
        If fglbNew And comRelation.Text = "Spouse" And Not glbtermopen Then
            If isSpouseExistWithEndDate(glbLEE_ID) Then
                MsgBox "Can't enter a new SPOUSE record unless the old Spouse has a Benefit End Date"
                comRelation.SetFocus
                Exit Function
            End If
        End If
        '"   Benefit Eligible Date cannot be less than 30 from "today" without a password being entered.
        'If no password Benefit Eligibility Date = Today and create a Follow Up record with the following details
        If Not glbtermopen Then   'CANADA only
        
            'Ticket #22108 Franks 07/12/2012 - begin ===================================
            'If glbWFC Then 'canada only
                If txtUserText(2).Enabled Then
                    If Left(comOther.Text, 1) = "3" Then
                        If Len(txtUserText(2).Text) = 0 Then
                            MsgBox "Dependent Text 2 is required if " & lblTitle(13).Caption & " is '3'"
                            txtUserText(2).SetFocus
                            Exit Function
                        End If
                    End If
                    If Not isValidDepTxt2(txtUserText(2).Text) Then
                        MsgBox "Invalid format for Dependent Text 2. "
                        txtUserText(2).SetFocus
                        Exit Function
                    End If
                End If
                '"   New edit warning message if the user is changing the Eligible or End dates:
                tmpFlag = False
                If IsDate(OtxtEligDate) Then
                    If Len(dlpEligDte.Text) = 0 Then
                        tmpFlag = True
                    Else
                        If Not CVDate(OtxtEligDate) = CVDate(dlpEligDte.Text) Then tmpFlag = True
                    End If
                End If
                If IsDate(OtxtEndDte) Then
                    If Len(dlpEndDte.Text) = 0 Then
                        tmpFlag = True
                    Else
                        If Not CVDate(OtxtEndDte) = CVDate(dlpEndDte.Text) Then tmpFlag = True
                    End If
                End If
                If tmpFlag Then 'changed
                    xMsg = "A change to the eligibility or end dates could cause the dependent to lose coverage at Manulife. "
                    'xMsg = xMsg & Chr(10) & "The Benefit Eligible Date will default to TODAY."
                    xMsg = xMsg & Chr(10) & "Are you sure you want to make this change? "
    
                    a% = MsgBox(xMsg, 36, "Confirm")
                    If a% <> 6 Then Exit Function
                End If
            'End If
            'Ticket #22108 Franks 07/12/2012 - end ========================================
        
        'If fglbNew And Not glbtermopen Then  'CANADA only
            If IsDate(dlpEligDte.Text) Then
                xlocDays = 0
                If fglbNew Then
                    xlocDays = DateDiff("d", CVDate(dlpEligDte.Text), CVDate(Date))
                Else
                    If IsDate(OtxtEligDate) Then
                        If Not CVDate(OtxtEligDate) = CVDate(dlpEligDte.Text) Then  'Change only
                            xlocDays = DateDiff("d", CVDate(dlpEligDte.Text), CVDate(Date))
                        End If
                    End If
                End If
                'If xlocDays = 0 And xlocDays < 30 Then
                If xlocDays > 90 Then '30 -> 90
                    'xMsg = "Benefit Eligible Date cannot be less than 30 from TODAY without a password being entered."
                    'Ticket #22108 Franks 07/12/2012
                    'Ticket #25562 Franks 06/17/2014 - "   Change the check from 30 days to 90 days.
                    xMsg = "Benefit Eligible Date cannot be more than 90 days prior from today without a password being entered."
                    xMsg = xMsg & Chr(10) & "Please enter a password after click OK button "
                    xMsg = xMsg & Chr(10) & "You can enter Benefit Eligible Date within 90 Days or contact Corporate for Retro Approval "
                    MsgBox xMsg
                    glbAccessPswd = False
                    frmAccessPswd.Show 1
                    If glbAccessPswd = False Then   'Access Denied
                        'Exit Sub
                        xMsg = "The password is not correct."
                        xMsg = xMsg & Chr(10) & "The Benefit Eligible Date will default to TODAY."
                        xMsg = xMsg & Chr(10) & "A follow up record will be created with reason 'BEDI'."
                        xMsg = xMsg & Chr(10) & "Are your sure you want to do it? "
                        
                        a% = MsgBox(xMsg, 36, "Confirm")
                        If a% <> 6 Then Exit Function
                        
                        dlpEligDte.Text = Date
                        dlpEligDte.SetFocus
                        'Exit Function
                        'create a follow up & send email
                        Call WFC30daysFollowup(glbLEE_ID)
                        
                        If gsEMAIL_ONDEPEND30DAYS4_WFC Then
                            Call cmdEmailWFC30days
                            
                            If AbortTerm = True Then
                                Screen.MousePointer = vbDefault
                                MDIMain.panHelp(0).FloodType = 1
                                MDIMain.panHelp(0).Caption = "Benefit Eligible Date Change Email Aborted"
                                MsgBox "Error sending email.  Benefit Eligible Date Change Email aborted.", vbCritical + vbOKOnly, "Error"
                                'Exit Sub
                            End If
                        End If
                        
                    End If
                End If
            End If
        
            If Not fglbNew Then 'Ticket #22411 Franks 08/07/2012
            'modified only
            'Canada only
                Call WFCCOB_Chg
            End If
        
        End If
        'Ticket #22009 Franks 05/11/2012 - end
        
'        'Ticket #22108 Franks 07/12/2012 - begin
'        If glbWFC Then 'canada only
'            If txtUserText(2).Enabled Then
'                If Left(comOther.Text, 1) = "3" Then
'                    If Len(txtUserText(2).Text) = 0 Then
'                        MsgBox "Dependent Text 2 is required if " & lblTitle(13).Caption & " is '3'"
'                        txtUserText(2).SetFocus
'                        Exit Function
'                    End If
'                End If
'                If Not isValidDepTxt2(txtUserText(2).Text) Then
'                        MsgBox "Invalid format for Dependent Text 2. "
'                        txtUserText(2).SetFocus
'                        Exit Function
'                End If
'            End If
'            '"   New edit warning message if the user is changing the Eligible or End dates:
'            tmpFlag = False
'            If IsDate(OtxtEligDate) Then
'                If Len(dlpEligDte.Text) = 0 Then
'                    tmpFlag = True
'                Else
'                    If Not CVDate(OtxtEligDate) = CVDate(dlpEligDte.Text) Then tmpFlag = True
'                End If
'            End If
'            If IsDate(OtxtEndDte) Then
'                If Len(dlpEndDte.Text) = 0 Then
'                    tmpFlag = True
'                Else
'                    If Not CVDate(OtxtEndDte) = CVDate(dlpEndDte.Text) Then tmpFlag = True
'                End If
'            End If
'            If tmpFlag Then 'changed
'                xMsg = "A change to the eligibility or end dates could cause the dependent to lose coverage at Manulife. "
'                'xMsg = xMsg & Chr(10) & "The Benefit Eligible Date will default to TODAY."
'                xMsg = xMsg & Chr(10) & "Are you sure you want to make this change? "
'
'                a% = MsgBox(xMsg, 36, "Confirm")
'                If a% <> 6 Then Exit Function
'            End If
'        End If
'        'Ticket #22108 Franks 07/12/2012 - end
    End If
End If

'Ticket #22417 - Goodmans LLP
If glbCompSerial = "S/N - 2290W" Then
    'Ticket #22745
    If comDeptTxt1.ListIndex = -1 Then
        MsgBox lblTitle(15).Caption & " is required."
        comDeptTxt1.SetFocus
        Exit Function
    End If

    If (txtDental.Text = "Y" Or txtMedical = "Y") And lblWFCTxt2.Visible = True Then
        If Len(txtUserText(2).Text) = 0 Then
            MsgBox lblTitle(16).Caption & " is required if " & lblTitle(11).Caption & " or " & lblTitle(12).Caption & " is 'Y - COB Coverage'."
            txtUserText(2).SetFocus
            Exit Function
        End If
        
        If Len(txtUserText(2).Text) > 0 Then
            If Not isValidDepTxt2(txtUserText(2).Text) Then
                MsgBox "Invalid date for " & lblTitle(16).Caption
                txtUserText(2).SetFocus
                Exit Function
            End If
        End If
    End If
End If

GoTo MODNOUPD

MODNOUPD:
chk_FeDepnts = True
Exit Function

FeDepnts_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Checking Dependents Form", "chk_FeDepnts", "Check")
Call RollBack

End Function

Sub LdcomRel() '25June99 js-added from VB3
If glbWFC And glbEmpCountry = "CANADA" Then 'Ticket #13448
    comRelation.AddItem "Child"
    'comRelation.AddItem "Husband" ''Ticket #22009 Franks 05/11/2012
    comRelation.AddItem "Spouse"
    'comRelation.AddItem "Wife" 'Ticket #22009 Franks 05/11/2012
ElseIf glbCompSerial = "S/N - 2290W" Then   'Ticket #22745 - Goodmans
    comRelation.AddItem "Child"
Else
    comRelation.AddItem "Aunt"
    comRelation.AddItem "Brother"
    comRelation.AddItem "Children"
    comRelation.AddItem "Common Law"
    comRelation.AddItem "Couple"
    comRelation.AddItem "Daughter"
    comRelation.AddItem "Estate"
    comRelation.AddItem "Ex-Spouse"
    comRelation.AddItem "Father"
    comRelation.AddItem "Fiancee"
    comRelation.AddItem "Fiance"
    comRelation.AddItem "Husband"
    comRelation.AddItem "Mother"
    comRelation.AddItem "Other"
    comRelation.AddItem "Parents"
    comRelation.AddItem "Sister"
    comRelation.AddItem "Son"
    comRelation.AddItem "Spouse"
    comRelation.AddItem "Uncle"
    comRelation.AddItem "Wife"
End If

ComSmoker.AddItem "No"
ComSmoker.AddItem "Yes"

If glbWFC And glbEmpCountry = "CANADA" Then 'Ticket #13448
    ComStatus.AddItem "Other( )" '"Child( )"
    ComStatus.AddItem "Overage Student(S)"
    ComStatus.AddItem "Handicap(H)"
    ComStatus.AddItem "Twin/Multiple Birth(M)"
Else
    ComStatus.AddItem "Other( )"
    ComStatus.AddItem "Student(1)"
    ComStatus.AddItem "Disability(2)"
    ComStatus.AddItem "Fast Yes(Y)"
    ComStatus.AddItem "Fast No(N)"
    ComStatus.AddItem "Fast(3)"
    ComStatus.AddItem "Fast(S)"
End If

End Sub

Private Function UpdDepends()
RSDATA!DP_STATUS = "" 'Jaddy 11/3/99
If glbWFC And glbEmpCountry = "CANADA" Then
    If ComStatus.Text = "Other( )" Then RSDATA!DP_STATUS = ""
    If ComStatus.Text = "Overage Student(S)" Then RSDATA!DP_STATUS = "S"
    If ComStatus.Text = "Handicap(H)" Then RSDATA!DP_STATUS = "H" '
    If ComStatus.Text = "Twin/Multiple Birth(M)" Then RSDATA!DP_STATUS = "M"   '
Else
    If ComStatus.Text = "Student(1)" Then RSDATA!DP_STATUS = "1"
    If ComStatus.Text = "Disability(2)" Then RSDATA!DP_STATUS = "2" '
    If ComStatus.Text = "Fast Yes(Y)" Then RSDATA!DP_STATUS = "Y"   '
    If ComStatus.Text = "Fast No(N)" Then RSDATA!DP_STATUS = "N"    '
    If ComStatus.Text = "Fast(3)" Then RSDATA!DP_STATUS = "3"       '
    If ComStatus.Text = "Fast(S)" Then RSDATA!DP_STATUS = "S"       '
End If

If ComSmoker.Text = "Yes" Then
    RSDATA!DP_SMOKER = "-1"
Else
    RSDATA!DP_SMOKER = "0"
End If
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

''' Sam add July 2002 * Remove ADO
Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        
        If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
            If glbtermopen Then
                RSDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
            Else
                RSDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            End If
            Call Me.cmdModify_Click
            Call SET_UP_MODE
        Exit Sub
    End If
    If glbtermopen Then  'Lucy July 4, 2000
        SQLQ = "Select * from Term_HRDEPEND "
    Else
        SQLQ = "Select * from HRDEPEND"
    End If
    SQLQ = SQLQ & " where DP_ID = " & Data1.Recordset!DP_ID

   'SQLQ = "select * from HR_DIVISION WHERE DIV='" & Data1.Recordset!Div & "'"
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    If glbtermopen Then
        RSDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    If RSDATA.EOF Or RSDATA.BOF Then Exit Sub
    
    Call Set_Control("R", Me, RSDATA)
    Call Me.cmdModify_Click
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
    RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
    UpdateRight = gSec_Upd_Dependents
End Property

Public Property Get Addable() As Boolean
    Addable = True
End Property

Public Property Get Updateble() As Boolean
    Updateble = True
End Property

Public Property Get Deleteble() As Boolean
    'Deleteble =  True
    If glbWFC Then
        Deleteble = gSec_Del_Dependents
    Else
        Deleteble = gSec_Upd_Dependents
    End If
    
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
ElseIf RSDATA.EOF Then
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

Private Sub LocCaptions()
    'Label
    lblTitle(2).Caption = lStr(lblTitle(2).Caption)
    lblTitle(5).Caption = lStr(lblTitle(5).Caption)
    lblTitle(6).Caption = lStr(lblTitle(6).Caption)
    lblTitle(8).Caption = lStr(lblTitle(8).Caption)
    lblTitle(9).Caption = lStr(lblTitle(9).Caption)
    lblTitle(10).Caption = lStr(lblTitle(10).Caption)
    lblTitle(11).Caption = lStr(lblTitle(11).Caption)
    lblTitle(12).Caption = lStr(lblTitle(12).Caption)
    lblTitle(13).Caption = lStr(lblTitle(13).Caption)
    lblTitle(14).Caption = lStr(lblTitle(14).Caption)
    
    'Ticket #20609 Franks 09/06/2011
    lblTitle(15).Caption = lStr("Dependent Text 1")
    lblTitle(16).Caption = lStr("Dependent Text 2")
    lblTitle(17).Caption = lStr("Dependent Text 3")
    lblTitle(18).Caption = lStr("Dependent Text 4")
    
    vbxTrueGrid.Columns(6).Caption = lStr(vbxTrueGrid.Columns(6).Caption)
    vbxTrueGrid.Columns(7).Caption = lStr(vbxTrueGrid.Columns(7).Caption)
    vbxTrueGrid.Columns(8).Caption = lStr(vbxTrueGrid.Columns(8).Caption)
    vbxTrueGrid.Columns(9).Caption = lStr(vbxTrueGrid.Columns(9).Caption)
    
End Sub

Public Sub imgEmail_Click()
Dim xEmail
Dim xToEmail As String
On Error GoTo Email_Err
    If gsEMAIL_ONDEPENDENT Then
        If Not UserEmailExist Then
            Exit Sub
        End If
        'xEmail = GetCurEmpEmail
        'xEmail = GetComPreferEmail("EMAIL_ONPERFORMANCE")
        
        'Ticket #20317 - Send email to More Emails list as well.
        xToEmail = GetComPreferEmail("EMAIL_ONDEPENDENT", glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONDEPENDENT")
        End If
            
        'If Len(xEmail) > 0 Then    'Hemu - (Ticket #11562) - Jerry asked to remove the check for email address presence.
            frmSendEmail.txtTo.Text = xToEmail
            If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18352, do not cc it to employee
            Else
                frmSendEmail.txtCC.Text = GetCurEmpEmail 'xEmail
            End If
            'Ticket #18578
            frmSendEmail.txtSubject.Text = "info:HR Dependent Change Notice - " & lblEEName.Caption
            frmSendEmail.txtBody.Text = MailBody
            frmSendEmail.Show 1
        'Else
            'If Len(glbLEE_SName) = 0 Then
            '    MsgBox "There is no email on Status/Dates screen for employee. "
            'Else
            '    MsgBox "There is no email on Status/Dates screen for employee " & glbLEE_SName & ", " & glbLEE_FName & ". "
            'End If
        '    MsgBox "There is no email address for the 'Email Notification on " & lstr("Performance") & " ' on Company Preference screen. "
        'End If


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

Sub EmailSendingForSamuel()
Dim xEmail
Dim xToEmail As String
Dim xEmailSubject As String, xBranch  As String

On Error GoTo Email_Err
    'If gsEMAIL_ONDEPENDENT Then
        If Not UserEmailExist Then
            Exit Sub
        End If

        xToEmail = GetComPreferEmail("EMAIL_ONDEPENDENT", glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONDEPENDENT")
        End If
        If Len(xToEmail) > 0 Then
            frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONSALARY")
            xBranch = GetEmpData(glbLEE_ID, "ED_SECTION", "")
            If Len(xBranch) > 0 Then
                xBranch = GetTABLDesc("EDSE", xBranch)
                xBranch = xBranch & " - "
            End If
            xEmailSubject = "info:HR Dependent Change Notice - " & xBranch & lblEEName.Caption
            frmSendEmail.txtSubject.Text = xEmailSubject
        
            frmSendEmail.txtBody.Text = MailBody
            'frmSendEmail.Show 1
            MDIMain.panHelp(0).FloodType = 0
            MDIMain.panHelp(0).Caption = "Sending email..."
            frmSendEmail.Tag = ""
            frmSendEmail.cmdSend_Click
            Do
                DoEvents
            Loop Until frmSendEmail.Tag <> ""   ' MC - dkostka - 05/03/01 - Changed from = "DONE" to <> ""
            ' AC - dkostka - 05/03/01 - Added checking to make sure the email went through,
            If frmSendEmail.Tag = "DONE" Then
                Unload frmSendEmail
            Else
                Unload frmSendEmail
            End If
            MDIMain.panHelp(0).Caption = ""
            MDIMain.panHelp(0).FloodType = 1
        End If

    'End If
    Exit Sub

Email_Err:
    'If Err.Number = 364 Then
    '    Exit Sub
    'End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail", "SMTP", "EmailSendingForSamuel")
    'Resume Next
    Exit Sub

End Sub

Private Function isSpouseExistWithEndDate(xEmpNo)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    SQLQ = "SELECT * FROM HRDEPEND WHERE DP_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND DP_RELATE = 'Spouse' "
    SQLQ = SQLQ & "AND DP_EDATE IS NULL "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        retVal = True
    End If
    rsTemp.Close
    isSpouseExistWithEndDate = retVal
End Function

Private Sub WFC30daysFollowup(xEmpNo)
    Call updFollow("U", "BEDI", Date)
    
End Sub

Private Function updFollow(xType, xCode, xDate)  'Laura on 11/2/97
Dim newline As String
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim rsFollow As New ADODB.Recordset
Dim Edit1 As Integer
'Don't need a message for follow up - Jerry asked for v7.6

newline = Chr$(13) & Chr$(10)
updFollow = False

On Error GoTo CrFollow_Err

If xType = "U" Then
    
    rsTB.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    'If fglbNew And IsDate(xDATE) Then
    If IsDate(xDate) Then
        SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND EF_FREAS = '" & xCode & "'"
        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(xDate)
        rsFollow.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsFollow.EOF Then
        ' Add by Frank for no duplicated record of HR_FOLLOW_UP End
            rsTB.AddNew
            rsTB("EF_COMPNO") = "001"
            rsTB("EF_EMPNBR") = glbLEE_ID
            rsTB("EF_FDATE") = CVDate(xDate) ' CVDate(dlpDate(1).Text)
            rsTB("EF_FREAS_TABL") = "FURE"
            'Ticket #24257 - Do not update Admin By for them only
            If glbCompSerial <> "S/N - 2262W" Then
                rsTB("EF_ADMINBY_TABL") = "EDAB"
                rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
            End If
            rsTB("EF_FREAS") = xCode '"SREV"
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
        'Exit Function
    End If
  
End If

updFollow = True

  
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Function

Sub cmdEmailWFC30days() 'Ticket #22061 Franks 05/24/2012
    Dim rsPen As New ADODB.Recordset
    Dim MailBody As String
    Dim SecCode As String, SecDesc As String
    Dim PenType As String
    Dim UnionCode As String
    Dim SalHrl As String
    Dim xEmpNo As Double
    Dim SQLQ As String
    Dim xStr As String
    Dim xTmpVal As Double
    Dim xCredSer As Double
    Dim xContSer As Double
    Dim DBEarns, DBCR, DBCS, DBCalDB, DBCashout
    Dim DBEarnsHly, DBCRHly, DBCSHly, DBCalDBHly, DBCashoutHly
    Dim DCEarns, DCER, DCEE, DCCashout
    Dim xSalFlag As Boolean
    Dim xHlyFlag As Boolean
    Dim xDBList As String
    
    On Error GoTo ErrorHandler
    
    'xDBList = "AND (LEFT(PE_PENSIONTYPE,2) = 'DB' OR LEFT(PE_PENSIONTYPE,3) = 'IDL' OR LEFT(PE_PENSIONTYPE,3) = 'UPG' OR LEFT(PE_PENSIONTYPE,3) = 'PRE' OR PE_PENSIONTYPE = 'DBSUP' OR PE_PENSIONTYPE = 'MON' ) "

    'Exit Sub
    Load frmSendEmail

    frmSendEmail.txtSubject.Text = "info:HR Dependent Benefit Eligible Date Change Notice - " & lblEEName.Caption
    frmSendEmail.txtTo.Text = "pension@woodbridgegroup.com"
    'MailBody = "The employee Benefit Eligible Date has been retired." & vbCrLf & vbCrLf
    MailBody = "Employee #: " & lblEENum.Caption & vbCrLf
    MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
    SecCode = GetEmpData(lblEENum.Caption, "ED_SECTION", "")
    If Len(SecCode) > 0 Then
        MailBody = MailBody & "Plant: " & GetTABLDesc("EDSE", SecCode) & vbCrLf
    End If
    MailBody = MailBody & vbCrLf
    
    MailBody = MailBody & "Dependent Name: " & txtSurname.Text & ", " & txtFName.Text & vbCrLf
    MailBody = MailBody & "Relationship: " & comRelation.Text & vbCrLf
    MailBody = MailBody & "Previous Benefit Eligible Date: " & CVDate(OtxtEligDate) & vbCrLf
    MailBody = MailBody & "New Benefit Eligible Date: " & dlpEligDte.Text & vbCrLf & vbCrLf
    MailBody = MailBody & "By User: " & GetUserDesc(glbUserID) & vbCrLf
    MailBody = MailBody & "Transaction Date: " & CVDate(Date) & vbCrLf
    
    frmSendEmail.txtBody.Text = MailBody

    MDIMain.panHelp(0).FloodType = 0
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
        AbortTerm = False
    Else
        Unload frmSendEmail
        AbortTerm = True
    End If
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(0).FloodType = 1

exH:
    Exit Sub
    
ErrorHandler:
    If Err.Number = 364 Then Exit Sub
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Resume exH

End Sub

Private Function isValidDepTxt2(xTxt)
Dim retVal As Boolean
Dim xMon, xDay, xYear
    retVal = True
    If Len(xTxt) > 0 Then
        'o   mm/dd/yyyy
        If Not Len(xTxt) = 10 Then retVal = False
        If Not Mid(xTxt, 3, 1) = "/" Then retVal = False
        If Not Mid(xTxt, 6, 1) = "/" Then retVal = False
        xMon = Mid(xTxt, 1, 2)
        xDay = Mid(xTxt, 4, 2)
        xYear = Mid(xTxt, 7, 4)
        If Not IsNumeric(xMon) Then retVal = False
        If Not IsNumeric(xDay) Then retVal = False
        If Not IsNumeric(xYear) Then retVal = False
        If Val(xMon) > 12 Then retVal = False
        If Val(xDay) > 31 Then retVal = False
        If Val(xYear) > 2075 Then retVal = False
        'If Val(xYear) < 1950 Then retVal = False
        'Ticket #23405 Franks 03/11/2013
        If Val(xYear) < 1920 Then retVal = False
    End If
    
    isValidDepTxt2 = retVal
End Function

Private Sub WFCCOB_Chg() 'Ticket #22411
Dim Msg$
Dim xOldVal
Dim SQLQ As String
Dim DenFlag As Boolean
Dim MedFlag As Boolean
Dim OthFlag As Boolean
Dim xUptNo As Integer

    If glbtermopen Then Exit Sub
    DenFlag = False
    MedFlag = False
    OthFlag = False
    If Not (OtxtDental = txtDental.Text) Then
        'If Len(OtxtDental) > 0 Then DenFlag = True
        DenFlag = True
    End If
    If Not (OtxtMedical = txtMedical.Text) Then
        'If Len(OtxtMedical) > 0 Then MedFlag = True
        MedFlag = True
    End If
    If Not (OtxtOther = txtOther.Text) Then
        'If Len(OtxtOther) > 0 Then OthFlag = True
        OthFlag = True
    End If
    
    If Not DenFlag And Not MedFlag And Not OthFlag Then
        'no change
        Exit Sub
    End If
    
    Screen.MousePointer = DEFAULT
    If DenFlag And Not MedFlag And Not OthFlag Then 'dental only
        Msg$ = "Please enter a new Effective Date for the Dental benefit." & vbNewLine
        xUptNo = 1
    End If
    If MedFlag And Not DenFlag And Not OthFlag Then 'medical only
        Msg$ = "Please enter a new Effective Date for the Health benefit." & vbNewLine
        xUptNo = 2
    End If
    If MedFlag And DenFlag Then  'dental and medical
        'Msg$ = "Please enter a new Effective Date for the Dental & Health benefit." & vbNewLine
        'Ticket #29748 Franks 03/14/2017
        Msg$ = "Please enter a new Effective Date for both Dental and Health Benefits." & vbNewLine
        xUptNo = 3
    End If
    If OthFlag Then   'Other change
        'Msg$ = "Please enter a new Effective Date for the Dental && Health benefit." & vbNewLine
        'Ticket #29748 Franks 03/14/2017
        Msg$ = "Please enter a new Effective Date for both Dental and Health Benefits." & vbNewLine
        xUptNo = 3
    End If
    
    glbChgTermDate = ""
    frmMsgTerm.PenTermDate = "WFCCOB_Change"
    frmMsgTerm.lblNote1.Caption = Msg$
    frmMsgTerm.lblNote1.Top = 300: frmMsgTerm.lblNote1.Visible = True
    frmMsgTerm.lblTitle(0).Top = 1080: frmMsgTerm.dlpTermDate.Top = 1080
    frmMsgTerm.dlpTermDate = glbChgTermDate
    frmMsgTerm.Show 1
    If IsDate(glbChgTermDate) Then
        'update the Benefit Effctive Date
        If xUptNo = 1 Or xUptNo = 3 Then 'Dental benefit
            'Call WFCCOB_BenDates(glbLEE_ID, "DN", glbChgTermDate)
            'Ticket #29748 Franks 03/14/2017
            Call WFCCOB_BenDates(glbLEE_ID, "DEN", glbChgTermDate)
        End If
        If xUptNo = 2 Or xUptNo = 3 Then 'medical benefit
            Call WFCCOB_BenDates(glbLEE_ID, "EHC", glbChgTermDate)
        End If
    End If

End Sub

Private Sub WFCCOB_BenDates(xEmpNo, xBenCode, xDate)
Dim rsBenT As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM HRBENFT "
    SQLQ = SQLQ & "WHERE BF_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND BF_BCODE = '" & xBenCode & "' "
    rsBenT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsBenT.EOF Then
        If Not IsNull(rsBenT("BF_EDATE")) Then
            If Not CVDate(rsBenT("BF_EDATE")) = CVDate(xDate) Then
                rsBenT("BF_EDATE") = xDate
                rsBenT.Update
                'update Audit
                Call WFCAUDITBENF("M", xBenCode, xDate)
            End If
        End If
    End If
End Sub

Private Function WFCAUDITBENF(ACTX, xBenCode, xEDate)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String

On Error GoTo AUDIT_ERR
'AUDITBENF = False

rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    If IsNull(rsTB("ED_PT")) Then
        xPT = ""
    Else
        xPT = rsTB("ED_PT")
    End If
    If IsNull(rsTB("ED_DIV")) Then
        xDiv = ""
    Else
        xDiv = rsTB("ED_DIV")
    End If
Else
    xPT = ""
    xDiv = ""
End If
'strfields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_COVER, AU_EDATE, AU_MAXDOL, AU_PPAMT, "
strFields = strFields & "AU_MTHCCOST, AU_MTHECOST, AU_BCODE, AU_BNAME, AU_BRELATE, AU_BDOB, AU_TAXBEN, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, "
strFields = strFields & "AU_OLDPPMT, AU_MAXDOL, AU_PER, AU_BAMT, AU_UNITCOST,AU_CEASEDATE, "
strFields = strFields & "AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE,AU_OLDLOC,AU_OLDWHRS "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False

If ACTX = "D" Then GoTo MODUPD
'GoTo MODNOUPD

MODUPD:

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv
rsTA("AU_BCODE") = xBenCode
rsTA("AU_EDATE") = xEDate
'If glbSoroc Or glbSyndesis Then
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
    End If
    rsEmp.Close
'End If

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = glbLEE_ID
rsTA("AU_LDATE") = Date
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "Y" ' "N" this is for CANADA employee only
rsTA("AU_TYPE") = ACTX
rsTA.Update

'If glbWFC And glbEmpCountry = "CANADA" Then 'Ticket #15818, do not pass benefit to payroll
'    Call WFCCNDBeneAuditFlag(glbLEE_ID)
'End If

MODNOUPD:
'AUDITBENF = True
Exit Function
AUDIT_ERR:

'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

