VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmESEMINARS 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Continuing Education "
   ClientHeight    =   10950
   ClientLeft      =   180
   ClientTop       =   960
   ClientWidth     =   11970
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
   ScaleHeight     =   10950
   ScaleWidth      =   11970
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   7455
      LargeChange     =   315
      Left            =   11400
      Max             =   100
      SmallChange     =   315
      TabIndex        =   96
      Top             =   2040
      Width           =   300
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feesemnr.frx":0000
      Height          =   1665
      Left            =   120
      OleObjectBlob   =   "feesemnr.frx":0014
      TabIndex        =   0
      Top             =   480
      Width           =   10815
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7740
      Top             =   10920
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
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
      TabIndex        =   47
      Top             =   10410
      Width           =   11970
      _Version        =   65536
      _ExtentX        =   21114
      _ExtentY        =   952
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
      Begin VB.CommandButton cmdSuccessionPlan 
         Appearance      =   0  'Flat
         Caption         =   "&Succession Planning"
         Height          =   375
         Left            =   1680
         TabIndex        =   98
         Tag             =   "Load Continuing Education screen"
         Top             =   120
         Width           =   2205
      End
      Begin VB.CommandButton cmdRetest 
         Appearance      =   0  'Flat
         Caption         =   "&Retest"
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Tag             =   "Load Beneficiary screen"
         Top             =   120
         Width           =   1365
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   7140
         Top             =   165
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   "c:\newihr\rgedsem.rpt"
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
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   11970
      _Version        =   65536
      _ExtentX        =   21114
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
      Begin VB.CommandButton cmdMissingCourses 
         Caption         =   "Missing Training records..."
         Height          =   375
         Left            =   9360
         TabIndex        =   97
         Top             =   60
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblEEProdLine 
         AutoSize        =   -1  'True
         Caption         =   "Product Line"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6480
         TabIndex        =   49
         Top             =   165
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   46
         Top             =   160
         Width           =   1005
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1320
         TabIndex        =   45
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2850
         TabIndex        =   44
         Top             =   135
         Width           =   720
      End
   End
   Begin VB.PictureBox pcCourseInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   120
      ScaleHeight     =   8295
      ScaleWidth      =   11175
      TabIndex        =   50
      Top             =   2160
      Width           =   11175
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "ES_LUSER"
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
         Left            =   7170
         MaxLength       =   25
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   8070
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "ES_LTIME"
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
         Left            =   5490
         MaxLength       =   25
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   8070
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "ES_LDATE"
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
         Left            =   3690
         MaxLength       =   25
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   8070
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.TextBox txtCourseName 
         Appearance      =   0  'Flat
         DataField       =   "ES_COURSE"
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
         Left            =   1635
         MaxLength       =   125
         TabIndex        =   4
         Tag             =   "01-Course Name"
         Top             =   660
         Width           =   3855
      End
      Begin VB.TextBox txtExtName 
         Appearance      =   0  'Flat
         DataField       =   "ES_EXTNAME"
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
         Left            =   1635
         MaxLength       =   60
         TabIndex        =   5
         Tag             =   "00-Description of Course"
         Top             =   990
         Width           =   3855
      End
      Begin VB.TextBox txtKeyword 
         Appearance      =   0  'Flat
         DataField       =   "ES_KEYWORD"
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
         Left            =   7020
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "00-Enter Keyword"
         Top             =   990
         Width           =   1095
      End
      Begin VB.TextBox txtCourseHRS 
         Appearance      =   0  'Flat
         DataField       =   "ES_HOURS"
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
         Left            =   7020
         MaxLength       =   5
         TabIndex        =   18
         Tag             =   "11-Number of Scheduled Course Hours "
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Frame frmAttendance 
         Caption         =   "Attendance Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   0
         TabIndex        =   58
         Top             =   4770
         Width           =   4635
         Begin VB.TextBox txtAttHrs 
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
            Height          =   285
            Left            =   1140
            MaxLength       =   5
            TabIndex        =   36
            Tag             =   "10-Number of Hours Spent on Course"
            Top             =   600
            Width           =   885
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   4
            Left            =   840
            TabIndex        =   35
            Tag             =   "01-Attendance Reason - Code"
            Top             =   240
            Width           =   2500
            _ExtentX        =   4419
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ADRE"
         End
         Begin Threed.SSCheck chkIncentive 
            Height          =   195
            Left            =   3360
            TabIndex        =   37
            Tag             =   "Incentive -  Attendance Management"
            Top             =   330
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Incentive  "
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin Threed.SSCheck chkSeniority 
            Height          =   195
            Left            =   3360
            TabIndex        =   38
            Tag             =   "Hours to be added to employee's seniority."
            Top             =   630
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Seniority    "
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin VB.Label lbltitle 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Reason"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   9
            Left            =   90
            TabIndex        =   60
            Top             =   300
            Width           =   660
         End
         Begin VB.Label lbltitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "  Hours "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   59
            Top             =   630
            Width           =   615
         End
      End
      Begin VB.Frame frmSkills 
         Caption         =   "Skills Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   4890
         TabIndex        =   55
         Top             =   4770
         Width           =   4245
         Begin VB.TextBox txtSkillsExp 
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
            Height          =   285
            Left            =   1520
            MaxLength       =   2
            TabIndex        =   40
            Tag             =   "10-Skill Rating (0-99)"
            Top             =   615
            Width           =   885
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   5
            Left            =   1200
            TabIndex        =   39
            Tag             =   "00-Skills obtained resulting from course - Code"
            Top             =   240
            Width           =   2500
            _ExtentX        =   4419
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDSK"
         End
         Begin VB.Label lbltitle 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Skill  "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   11
            Left            =   210
            TabIndex        =   57
            Top             =   300
            Width           =   375
         End
         Begin VB.Label lbltitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Exp. Factor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   12
            Left            =   210
            TabIndex        =   56
            Top             =   645
            Width           =   990
         End
      End
      Begin VB.CheckBox chkPresentor 
         Alignment       =   1  'Right Justify
         Caption         =   "Presenter "
         DataField       =   "ES_PRESENTOR"
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
         Left            =   6000
         TabIndex        =   31
         Tag             =   "40-Presenter -y/n"
         Top             =   3705
         Width           =   1220
      End
      Begin VB.Frame frmTotal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7020
         TabIndex        =   54
         Top             =   3330
         Width           =   1095
         Begin MSMask.MaskEdBox medContTotal 
            Height          =   285
            Left            =   0
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   0
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "Currency"
            PromptChar      =   "_"
         End
      End
      Begin VB.TextBox txtAccount 
         Appearance      =   0  'Flat
         DataField       =   "ES_ACCOUNTNO"
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
         Left            =   7020
         MaxLength       =   25
         TabIndex        =   16
         Tag             =   "00-Account Number"
         Top             =   660
         Width           =   1965
      End
      Begin VB.TextBox txtCompanyName 
         Appearance      =   0  'Flat
         DataField       =   "ES_COMPANYNAME"
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "00-Company Name"
         Top             =   1650
         Width           =   3855
      End
      Begin VB.TextBox txtTrainerName 
         Appearance      =   0  'Flat
         DataField       =   "ES_TRAINNER"
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   8
         Tag             =   "00-Trainer Name"
         Top             =   1995
         Width           =   3855
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   270
         Left            =   9660
         TabIndex        =   34
         Tag             =   "Import Document for attachment"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtCEUCred 
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
         Height          =   285
         Left            =   7020
         MaxLength       =   5
         TabIndex        =   33
         Tag             =   "11-CEU Credit"
         Top             =   3990
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame frmCourseDesc 
         Height          =   1575
         Left            =   0
         TabIndex        =   51
         Top             =   5880
         Visible         =   0   'False
         Width           =   9135
         Begin VB.TextBox memCourseDesc 
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
            Height          =   1050
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   41
            Tag             =   "00-Comments"
            Top             =   360
            Width           =   4575
         End
         Begin VB.TextBox memCourseLoc 
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
            Height          =   1050
            Left            =   4920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Tag             =   "00-Comments"
            Top             =   360
            Width           =   4155
         End
         Begin VB.Label lbltitle 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Course Description"
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
            Index           =   28
            Left            =   240
            TabIndex        =   53
            Top             =   120
            Width           =   2445
         End
         Begin VB.Label lbltitle 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Course Location"
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
            Index           =   29
            Left            =   4920
            TabIndex        =   52
            Top             =   120
            Width           =   1965
         End
      End
      Begin INFOHR_Controls.CodeLookup clpEmpCur 
         DataField       =   "ES_EMPCUR"
         Height          =   285
         Left            =   8160
         TabIndex        =   20
         Top             =   1650
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECUR"
      End
      Begin INFOHR_Controls.DateLookup dlpRenewal 
         DataField       =   "ES_RENEW"
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Tag             =   "40-Date when course is to be renewed"
         Top             =   3660
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDatComp 
         DataField       =   "ES_DATCOMP"
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Tag             =   "41-Date when course was completed"
         Top             =   3330
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpStartDate 
         DataField       =   "ES_START"
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Tag             =   "41-Date when course was Started"
         Top             =   3000
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpSchDate 
         DataField       =   "ES_SCHEDULED"
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Tag             =   "40-Date when course was Scheduled"
         Top             =   2670
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ES_RESULTS"
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   9
         Tag             =   "00-Results of the Course - Code"
         Top             =   2340
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ESRT"
         MaxLength       =   10
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ES_CONDUCT"
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   6
         Tag             =   "00-Organization/Individual Instructing - Code"
         Top             =   1320
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ESCB"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ES_CTYPE"
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Tag             =   "01-Course Type - Code"
         Top             =   0
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ESCT"
         MaxLength       =   8
      End
      Begin MSMask.MaskEdBox medEECont 
         DataField       =   "ES_TBEMP"
         Height          =   285
         Index           =   0
         Left            =   7020
         TabIndex        =   19
         Tag             =   "20-Amount Employee Paid"
         Top             =   1650
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Currency"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEECont 
         DataField       =   "ES_OTHER"
         Height          =   285
         Index           =   2
         Left            =   7020
         TabIndex        =   21
         Tag             =   "20-Other Expenses Paid"
         Top             =   1995
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Currency"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEECont 
         DataField       =   "ES_TBCO"
         Height          =   285
         Index           =   1
         Left            =   7020
         TabIndex        =   23
         Tag             =   "20-Amount Employer Paid"
         Top             =   2340
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Currency"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEECont 
         DataField       =   "ES_ACCOM"
         Height          =   285
         Index           =   3
         Left            =   7020
         TabIndex        =   25
         Tag             =   "20-Accommodation"
         Top             =   2670
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Currency"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ES_COORDINATED"
         Height          =   285
         Index           =   6
         Left            =   6720
         TabIndex        =   14
         Tag             =   "00-Co-Ordinated By"
         Top             =   0
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ESCC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ES_METHODUSED"
         Height          =   285
         Index           =   7
         Left            =   6720
         TabIndex        =   15
         Tag             =   "00-Method Used"
         Top             =   330
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ESMU"
      End
      Begin INFOHR_Controls.CodeLookup clpOherCur 
         DataField       =   "ES_OTCUR"
         Height          =   285
         Left            =   8160
         TabIndex        =   22
         Top             =   1995
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECUR"
      End
      Begin INFOHR_Controls.CodeLookup clpEmployerCur 
         DataField       =   "ES_EMPLOYCUR"
         Height          =   285
         Left            =   8160
         TabIndex        =   24
         Top             =   2340
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECUR"
      End
      Begin INFOHR_Controls.CodeLookup clpAcomCur 
         DataField       =   "ES_ACOMCUR"
         Height          =   285
         Left            =   8160
         TabIndex        =   26
         Top             =   2670
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECUR"
      End
      Begin INFOHR_Controls.CodeLookup clpTotCur 
         DataField       =   "ES_TOTCUR"
         Height          =   285
         Left            =   8160
         TabIndex        =   30
         Top             =   3330
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECUR"
      End
      Begin MSMask.MaskEdBox medEECont 
         DataField       =   "ES_LEARNING"
         Height          =   285
         Index           =   4
         Left            =   7020
         TabIndex        =   27
         Tag             =   "20-Learning Material"
         Top             =   3000
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Currency"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.CodeLookup clpLearnCur 
         DataField       =   "ES_LEARNINGCUR"
         Height          =   285
         Left            =   8160
         TabIndex        =   28
         Top             =   3000
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECUR"
      End
      Begin VB.Frame frmCourseCode 
         Height          =   375
         Left            =   4800
         TabIndex        =   64
         Top             =   7680
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtMain 
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
            Height          =   285
            Left            =   320
            TabIndex        =   2
            Tag             =   "00-Course Code"
            Top             =   0
            Width           =   990
         End
         Begin VB.Label lblDesc 
            Caption         =   "*** NOT ATTACHED ***"
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
            Left            =   1560
            TabIndex        =   65
            Top             =   30
            Width           =   3075
         End
         Begin VB.Image imgIcon 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   0
            Picture         =   "feesemnr.frx":911C
            Top             =   30
            Width           =   240
         End
      End
      Begin INFOHR_Controls.CodeLookup clpCEUType 
         Height          =   285
         Left            =   1320
         TabIndex        =   32
         Tag             =   "00-CEU Type - Code"
         Top             =   3990
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ESUT"
         MaxLength       =   8
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ES_CRSCODE"
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Tag             =   "00-Course Code"
         Top             =   330
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ESCD"
         MaxLength       =   8
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Course Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   21
         Left            =   60
         TabIndex        =   100
         Top             =   330
         Width           =   1200
      End
      Begin VB.Label lblCEUType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CEU Type"
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
         Left            =   60
         TabIndex        =   99
         Top             =   4035
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblCNum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label20"
         DataField       =   "ES_COMPNO"
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
         Left            =   90
         TabIndex        =   95
         Top             =   8070
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label20"
         DataField       =   "ES_EMPNBR"
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
         Left            =   900
         TabIndex        =   94
         Top             =   8070
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Course Type"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   93
         Top             =   30
         Width           =   1200
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Course Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   92
         Top             =   660
         Width           =   1380
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Course Description"
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
         Left            =   60
         TabIndex        =   91
         Top             =   1005
         Width           =   1575
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Conducted By      "
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
         Left            =   60
         TabIndex        =   90
         Top             =   1350
         Width           =   1320
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Results  "
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
         Left            =   60
         TabIndex        =   89
         Top             =   2340
         Width           =   1125
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Scheduled Date       "
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
         Left            =   60
         TabIndex        =   88
         Top             =   2700
         Width           =   1605
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Completed"
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
         Left            =   60
         TabIndex        =   87
         Top             =   3390
         Width           =   1365
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
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
         Index           =   14
         Left            =   60
         TabIndex        =   86
         Top             =   3060
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Renewal Date"
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
         Left            =   60
         TabIndex        =   85
         Top             =   3720
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Keyword"
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
         Left            =   5880
         TabIndex        =   84
         Top             =   990
         Width           =   975
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Course Hours"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   5520
         TabIndex        =   83
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "   Employee $"
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
         Left            =   5520
         TabIndex        =   82
         Top             =   1650
         Width           =   1350
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Other Expenses $"
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
         Left            =   5220
         TabIndex        =   81
         Top             =   1995
         Width           =   1665
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "     Employer $"
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
         Left            =   5580
         TabIndex        =   80
         Top             =   2340
         Width           =   1305
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total $"
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
         Left            =   5940
         TabIndex        =   79
         Top             =   3330
         Width           =   975
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Accommodation $"
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
         Index           =   20
         Left            =   5400
         TabIndex        =   78
         Top             =   2670
         Width           =   1515
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Account #"
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
         Index           =   19
         Left            =   5970
         TabIndex        =   77
         Top             =   660
         Width           =   915
      End
      Begin VB.Label lbltitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
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
         Index           =   22
         Left            =   60
         TabIndex        =   76
         Top             =   1680
         Width           =   1365
      End
      Begin VB.Label lbltitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Trainer Name"
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
         Index           =   23
         Left            =   60
         TabIndex        =   75
         Top             =   1995
         Width           =   1440
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Method Used"
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
         Index           =   24
         Left            =   5460
         TabIndex        =   74
         Top             =   330
         Width           =   1125
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Co-Ordinated By"
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
         Index           =   25
         Left            =   5460
         TabIndex        =   73
         Top             =   0
         Width           =   1305
      End
      Begin VB.Label lbltitle 
         Alignment       =   2  'Center
         Caption         =   "Currency"
         Height          =   255
         Index           =   26
         Left            =   8160
         TabIndex        =   72
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Learning Material $"
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
         Index           =   27
         Left            =   5400
         TabIndex        =   71
         Top             =   3000
         Width           =   1515
      End
      Begin VB.Image imgNoSec 
         Height          =   240
         Left            =   9240
         Picture         =   "feesemnr.frx":9266
         Top             =   4320
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSec 
         Height          =   240
         Left            =   9240
         Picture         =   "feesemnr.frx":93B0
         Top             =   4320
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblImport 
         Alignment       =   1  'Right Justify
         Caption         =   "Continuing Education"
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
         Height          =   255
         Left            =   6840
         TabIndex        =   70
         Top             =   4320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblCEUCred 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CEU Credit"
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
         Left            =   6105
         TabIndex        =   69
         Top             =   4035
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblPosCode 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         DataField       =   "ES_JOB"
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
         Height          =   195
         Left            =   1680
         TabIndex        =   68
         Top             =   4350
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label LabelPos 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Position :"
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
         Left            =   60
         TabIndex        =   67
         Top             =   4350
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblPosDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "PosDesc"
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
         Height          =   195
         Left            =   2760
         TabIndex        =   66
         Top             =   4350
         Visible         =   0   'False
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmESEMINARS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim fglbNew  As Integer
Dim fglHredsem As String       'added by Laura
Dim fglCursName As String      '
Dim fglExtName As String       '
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim rsGrid As New ADODB.Recordset
Dim Ctrl As Control 'Sam add July 2002 * Remove ADO
Dim OldtxtMain As String

Private Sub chkIncentive_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Function chkSemnr()
Dim oCode As String, OCodeD As String
Dim Msg, a%

chkSemnr = False

If Len(clpCode(1).Text) = 0 Then
    MsgBox lStr("Course Type is a required field")
    If Not glbWFC Then clpCode(1).SetFocus
    Exit Function
End If

If clpCode(1).Caption = "Unassigned" Then
    MsgBox lStr("Course Type code must be valid")
    If Not glbWFC Then clpCode(1).SetFocus
    Exit Function
End If

'Release 8.0 - Ticket #24441: Make Course Code Mandatory
If clpCode(0).Visible Then
    If Len(clpCode(0).Text) = 0 Then
        MsgBox lStr("Course Code is a required field")
        clpCode(0).SetFocus
        Exit Function
    End If
ElseIf txtMain.Visible Then
    If Len(txtMain.Text) = 0 Then
        MsgBox lStr("Course Code is a required field")
        txtMain.SetFocus
        Exit Function
    End If
End If

If clpCode(0).Visible Then
    If Len(clpCode(0).Text) > 0 Then
        If clpCode(0).Caption = "Unassigned" Then
            MsgBox lStr("Course Code must be valid")
            clpCode(0).SetFocus
            Exit Function
        End If
    End If
End If

If Len(txtCourseName) < 1 Then
    MsgBox lStr("Course Name is a required field")
    If txtCourseName.Enabled Then 'Ticket #25148 Franks 03/03/2014
        txtCourseName.SetFocus
    End If
    Exit Function
End If

If glbWFC Then 'Ticket #13520
    If clpCode(0).Visible Then
        If Len(clpCode(0).Text) = 0 Then
            MsgBox lStr("Course Code is a required field")
            clpCode(0).SetFocus
            Exit Function
        End If
    ElseIf txtMain.Visible Then
        If Len(txtMain.Text) = 0 Then
            MsgBox lStr("Course Code is a required field")
            txtMain.SetFocus
            Exit Function
        End If
    End If
    
    If IsDate(dlpSchDate.Text) Then '#30081 Franks 06/23/2017
    Else
        If Len(clpCode(3).Text) = 0 Then
            MsgBox lStr("Results is a required field")
            clpCode(3).SetFocus
            Exit Function
        End If
        If Not IsDate(dlpDatComp) Then
            MsgBox lStr("Date Completed is a required field")
            dlpDatComp.SetFocus
            Exit Function
        End If
    End If
    If Len(txtCourseHRS.Text) = 0 Then
        MsgBox lStr("Course Hours is a required field")
        txtCourseHRS.SetFocus
        Exit Function
    Else
        If Val(txtCourseHRS.Text) = 0 Then
            MsgBox lStr("Course Hours is a required field")
            txtCourseHRS.SetFocus
            Exit Function
        End If
    End If
    If Val(medContTotal.Text) = 0 Then
        Msg = "The Total $ is 0.  "
        Msg = Msg & "Are You Sure there is no cost for this course?"
        
        a% = MsgBox(Msg, 36, "Confirm")
        If a% <> 6 Then Exit Function
    End If
End If

If Len(clpCode(2).Text) > 0 Then
    If clpCode(2).Caption = "Unassigned" Then
        MsgBox lStr("Conducted By code must be valid")
        If Not glbWFC Then clpCode(2).SetFocus
        Exit Function
    End If
End If

If Len(clpCode(3).Text) > 0 Then
    If clpCode(3).Caption = "Unassigned" Then
        MsgBox lStr("Results code must be valid")
        clpCode(3).SetFocus
        Exit Function
    End If
End If

'Ticket #22701 - County of Lanark
If glbCompSerial = "S/N - 2172W" Then
    If Len(txtCourseHRS) = 0 Or Val(txtCourseHRS) = 0 Then
        MsgBox "Course Hours is requried field"
        txtCourseHRS.SetFocus
        Exit Function
    End If
End If

If Len(medEECont(0)) < 1 Then
    medEECont(0) = 0
Else
    If Not IsNumeric(medEECont(0)) Then
        MsgBox "Employee's Contribution must be numeric"
        medEECont(0).SetFocus
        Exit Function
    End If
End If

If Len(medEECont(1)) < 1 Then
    medEECont(1) = 0
Else
    If Not IsNumeric(medEECont(1)) Then
        MsgBox "Employer's Contribution must be numeric"
        medEECont(1).SetFocus
        Exit Function
    End If
End If

If Len(medEECont(2)) < 1 Then
    medEECont(2) = 0
Else
    If Not IsNumeric(medEECont(2)) Then
        MsgBox lStr("Other Expenses must be numeric")
        medEECont(2).SetFocus
        Exit Function
    End If
End If

If Len(medEECont(3)) < 1 Then
    medEECont(3) = 0
Else
    If Not IsNumeric(medEECont(3)) Then
        MsgBox lStr("Accommodation must be numeric")
        medEECont(3).SetFocus
        Exit Function
    End If
End If

If Len(clpCode(4).Text) > 0 Then
    If clpCode(4).Caption = "Unassigned" Then
        MsgBox "Attendance Reason code must be valid"
        clpCode(4).SetFocus
        Exit Function
    End If
    If Len(txtAttHrs) <= 0 Then
        MsgBox "Hours is Required Field"
        txtAttHrs.SetFocus
        Exit Function
        txtAttHrs = 0
    End If
End If

If Len(clpCode(5).Text) > 0 Then
    If clpCode(5).Caption = "Unassigned" Then
        MsgBox "Skill code must be valid"
        clpCode(5).SetFocus
        Exit Function
    End If
End If

If Len(txtAttHrs) >= 1 Then
    If Not IsNumeric(txtAttHrs) Then
        MsgBox "Attendance Hours must be numeric"
        txtAttHrs.SetFocus
        Exit Function
    End If
    If clpCode(4).Text = "" Then
        MsgBox "Attendance Reason cannot be blank if Attendance Hours is entered"
        clpCode(4).SetFocus
        Exit Function
    End If
End If

If Len(txtSkillsExp) >= 1 Then
    If Not IsNumeric(txtSkillsExp) Then
        MsgBox "Skills Exp. must be numeric"
        txtSkillsExp.SetFocus
        Exit Function
    End If
    If clpCode(5).Text = "" Then
        MsgBox "Skill code cannot be blank if Exp. Factor is entered"
        clpCode(5).SetFocus
        Exit Function
    End If
End If
'~~~ added by raubrey 7/29/97 for new data elements
If Len(txtCourseHRS) > 0 Then
    If Not IsNumeric(txtCourseHRS) Then
        MsgBox lStr("Course Hours must be numeric")
        txtCourseHRS.SetFocus
        Exit Function
    End If
End If

If Len(dlpSchDate.Text) > 0 Then
    If Not IsDate(dlpSchDate.Text) Then
        MsgBox lStr("Scheduled date is invalid")
        dlpSchDate.SetFocus
        Exit Function
    End If
End If

If Len(dlpStartDate.Text) > 0 Then
    If Not IsDate(dlpStartDate.Text) Then
        MsgBox lStr("Start date is invalid")
        dlpStartDate.SetFocus
        Exit Function
    End If
End If

If Len(dlpDatComp.Text) > 0 Then
    If Not IsDate(dlpDatComp.Text) Then
        MsgBox lStr("Date Completed is invalid")
        dlpDatComp.SetFocus
        Exit Function
    End If
End If

If clpCode(4).Text <> "" Or clpCode(5).Text <> "" Then
    If Not IsDate(dlpStartDate.Text) And Not IsDate(dlpDatComp.Text) Then
        MsgBox lStr("Start date") & " or " & lStr("Date Completed") & " is required if Attendance/Skills Data is entered"
        dlpStartDate.SetFocus
        Exit Function
    End If
End If


'''If Len(txtStartDate) = 0 Then
'''    If IsDate(dlpDatComp) Then
'''        txtStartDate = txtDatComp
'''    End If
'''End If
'''
'''If Len(txtDatComp) = 0 Then
'''    If IsDate(txtStartDate) Then
'''        txtDatComp = txtStartDate
'''    End If
'''End If
'''
''''---------- Added on 98/05/09 by Andy
'''If Len(txtSchDate) = 0 Then
'''    If Len(txtDatComp) = 0 Then
'''        MsgBox "Date Completed is a required field" '16Aug99 js - changed message
'''        txtDatComp.SetFocus
'''        Exit Function
'''    End If
'''End If
''''-----------

''If Len(dlpSchDate.Text) = 0 Then
'    If Len(dlpStartDate.Text) = 0 Then
'        MsgBox "Must enter either Scheduled Date or Start Date"
'        dlpStartDate.SetFocus
'        Exit Function
'    End If
''End If

If Len(dlpRenewal.Text) > 0 Then
    If Not IsDate(dlpRenewal.Text) Then
        MsgBox lStr("Renewal date is invalid")
        dlpRenewal.SetFocus
        Exit Function
    End If
End If
If glbLinamar Then
    If Val(txtCourseHRS) = 0 Then
        MsgBox lStr("Course Hours is requried field")
        txtCourseHRS.SetFocus
        Exit Function
    End If
    If Len(txtAccount) > 0 Then
        If Not IsNumeric(txtAccount) Then
            MsgBox lStr("Account # must be numeric")
            txtAccount.SetFocus
            Exit Function
        End If
    End If
End If
If Len(clpEmpCur.Text) > 0 Then
    If clpEmpCur.Caption = "Unassigned" Then
        MsgBox "Employee Currency code must be valid"
        clpEmpCur.SetFocus
        Exit Function
    End If
End If
If Len(clpOherCur.Text) > 0 Then
    If clpOherCur.Caption = "Unassigned" Then
        MsgBox "Other Expenses Currency code must be valid"
        clpOherCur.SetFocus
        Exit Function
    End If
End If
If Len(clpEmployerCur.Text) > 0 Then
    If clpEmployerCur.Caption = "Unassigned" Then
        MsgBox "Employer Currency code must be valid"
        clpEmployerCur.SetFocus
        Exit Function
    End If
End If
If Len(clpAcomCur.Text) > 0 Then
    If clpAcomCur.Caption = "Unassigned" Then
        MsgBox "Accomodations Currency code must be valid"
        clpAcomCur.SetFocus
        Exit Function
    End If
End If
If Len(clpTotCur.Text) > 0 Then
    If clpTotCur.Caption = "Unassigned" Then
        MsgBox "Total Currency code must be valid"
        clpTotCur.SetFocus
        Exit Function
    End If
End If

'Ticket #20447 - Jerry said to open up for everyone
'Frank Apr 13, 2007 Ticket #12859 - City of Chatham-Kent
'If glbCompSerial = "S/N - 2188W" Then
    If Len(clpCEUType.Text) > 0 Then
        If clpCEUType.Caption = "Unassigned" Then
            MsgBox "CEU Type code must be valid"
            If Not glbWFC Then clpCEUType.SetFocus
            Exit Function
        End If
    End If
    If Len(txtCEUCred.Text) > 0 Then
        If Not IsNumeric(txtCEUCred.Text) Then
            MsgBox "CEU Credit is not number"
            txtCEUCred.SetFocus
            Exit Function
        End If
    End If
'End If


'~~~
If Not ChkDup() Then
        Exit Function
End If
chkSemnr = True

End Function

Private Sub chkPresentor_GotFocus()
  Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkSeniority_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpCEUType_GotFocus()
  Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpCode_Change(Index As Integer)
If Index = 4 Then Call ATTCode_Desc(Index)
'If Index = 0 Then Call CrsName_Desc

End Sub

Private Sub clpCode_GotFocus(Index As Integer)
If Index = 0 Then
    'If glbWFC Then 'Ticket #13520
    Dim xStr As String
        xStr = Trim(GetTransDiv(clpCode(1).Text))
        clpCode(Index).TransDiv = xStr '"'*','9BSL','9FIP'" 'xStr
    'End If
End If
End Sub

Private Function GetTransDiv(xFirstVal As String)
Dim rsTran As New ADODB.Recordset
Dim SQLQ As String
Dim xFinal As String

    xFinal = ""
    If Not Len(xFirstVal) = 0 Then
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ESCD' "
        SQLQ = SQLQ & "AND TB_USR1 ='" & xFirstVal & "' "
        rsTran.Open SQLQ, gdbAdoIhr001, adOpenStatic
        'xFinal = "'*'"
        If Not rsTran.EOF Then
            xFinal = "'*'"
        End If
        Do While Not rsTran.EOF
            xFinal = xFinal & ",'" & rsTran("TB_KEY") & "'"
            rsTran.MoveNext
        Loop
    End If
    GetTransDiv = xFinal
End Function

Private Sub clpCode_LostFocus(Index As Integer)
If Index = 0 Then Call CrsName_Desc
If Index = 0 Then Call CourseCode_Type
End Sub

Private Sub CourseCode_Type()
Dim Response%, Msg$, Title$, DgDef As Double
Dim SQLQ As String, RType
Dim RSTABL As New ADODB.Recordset
'''On Error GoTo Dept_GL_Err

If Len(clpCode(0).Text) > 0 Then
    RSTABL.Open "SELECT TB_NAME,TB_KEY,TB_USR1 FROM HRTABL WHERE TB_NAME = 'ESCD' AND TB_KEY='" & clpCode(0).Text & "'", gdbAdoIhr001
    If Not RSTABL.EOF Then
        If IsNull(RSTABL("TB_USR1")) Then
            RType = ""
        Else
            RType = RSTABL("TB_USR1")
        End If
        If Len(RType) > 0 Then
            If clpCode(1).Text <> RType Then
                    Msg$ = lStr("Do you want the associated Course Type?")
                    Title$ = "info:HR"
                    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2  ' Describe dialog.
                    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                    If Response% = IDYES Then clpCode(1).Text = RType
            End If
        End If
    End If
    RSTABL.Close
End If

Exit Sub

Dept_GL_Err:
If Err = 94 Then
    ' clpGLNum.Text = ""
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Course Code Snap", "Course Code", "SELECT")
Call RollBack '21June99 js
End Sub

Sub cmdCancel_Click()
Dim I%
Dim x
'''On Error GoTo Can_Err


'Data1.Recordset.CancelBatch
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh
fglbNew = False
''' Sam add July 2002 * Remove ADO
If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
'rsDATA.CancelUpdate
Call Display_Value



If fglbNew Then
    clpCode(4).Text = ""
    clpCode(5).Text = ""
    txtAttHrs.Text = ""
    txtSkillsExp.Text = ""
    chkSeniority.Value = False 'added by RAUBREY 4/4/97
    chkIncentive.Value = False
End If


Call UpConttotal
'fglbNew = False
'Call SET_UP_MODE
'Call ST_UPD_MODE(True)  ' reset screen's attributes

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREDSEM", "Cancel")
Call RollBack '23July99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMESEMINARS" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, x

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

'''On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

'-- added on 11/1/97 by Laura
fglHredsem = dlpRenewal.Text
If fglHredsem <> "" Then
    'Friesens - Ticket #16189 and City of Chatham-Kent - Ticket #16794
    If glbCompSerial <> "S/N - 2279W" And glbCompSerial <> "S/N - 2188W" Then
        If Not updFollow("D") Then
            Exit Sub
        End If
    End If
End If
'--
'7.9 - Enhancement - For all the clients now
'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
    'Call procedure to clear the Course Taken Date in Training List if Renewal Date is still there.
    'If IsDate(dlpRenewal.Text) And IsDate(dlpDatComp.Text) Then
    '    Call Clear_Course_Taken_Date(dlpRenewal.Text, dlpDatComp.Text, txtMain.Text, lblPosCode.Caption)
    'End If
    If IsDate(dlpRenewal.Text) And IsDate(dlpDatComp.Text) Then
        Call Undo_Training_List_Rec_on_ContEdu_Delete(lblPosCode.Caption, txtMain.Text, dlpRenewal.Text, dlpDatComp.Text)
    ElseIf IsDate(dlpDatComp.Text) Then
        Call Undo_Training_List_Rec_on_ContEdu_Delete(lblPosCode.Caption, txtMain.Text)
    End If
'End If

If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001X.CommitTrans
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.Execute "Delete from Term_HRDOC_EDSEM where ES_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " and ES_DOCKEY=" & glbDocKey & " " '
    End If
    Data1.Refresh

    Set rsGrid = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True

Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001.CommitTrans
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.Execute "delete from HRDOC_EDSEM where ES_TYPE='" & UCase(glbDocName) & "' AND ES_EMPNBR = " & glbLEE_ID & " and ES_DOCKEY=" & glbDocKey & " "
    End If
    Data1.Refresh
    
    Set rsGrid = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True
    
End If
'If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
'End If

Call UpConttotal

fglbNew = False

Call SET_UP_MODE
'Call ST_UPD_MODE(True)


Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HREDSEM", "Delete")
Call RollBack '23July99 js

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()

'''On Error GoTo Mod_Err


If dlpRenewal.Text <> "" Then
  fglHredsem = dlpRenewal.Text
Else
  fglHredsem = ""
End If

fglCursName = txtCourseName
fglExtName = txtExtName

If Not glbtermopen Then
    If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
        If Len(Trim(dlpDatComp)) = 0 And Len(Trim(dlpStartDate)) = 0 Then
            Call AttSkill(True)
        End If
    End If
End If

If gSec_Inq_SUCCESSION Then
    cmdSuccessionPlan.Visible = True
    cmdSuccessionPlan.Enabled = True
Else
    cmdSuccessionPlan.Visible = False
End If

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HREDSEM", "Modify")
Call RollBack '23July99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
Dim x%
fglbNew = True

'Call ST_UPD_MODE(True)
'''On Error GoTo AddN_Err

Call SET_UP_MODE

If gsAttachment_DB Then
    'glbJob = ""
    'glbSDate = "01/01/1900"
    lblImport.Visible = True 'False
    imgSec.Visible = False
    imgNoSec.Visible = True 'False
    cmdImport.Visible = True 'False
End If

Call Set_Control("B", Me)

'7.9 - Enhancement - For all the clients now
'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
    lblPosDesc.Caption = ""
'End If

If gSec_Inq_SUCCESSION Then
    cmdSuccessionPlan.Enabled = False
Else
    cmdSuccessionPlan.Visible = False
End If

rsDATA.AddNew

If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID

lblCNum.Caption = "001"

For x% = 0 To 4 '3
    medEECont(x%) = 0
Next

Call UpConttotal

chkPresentor = 0
If glbWFC Then 'Ticket #24708 Franks 11/27/2013
    'clpCode(0).Enabled = True
    'clpCode(0).SetFocus
    'txtCourseName.SetFocus
Else
    clpCode(1).Enabled = True
    clpCode(1).SetFocus
End If

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HREDSEM", "Add")
Call RollBack '23July99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
'Dim xChange1, xChange2
Dim x
Dim xID As Long

'''On Error GoTo Add_Err

' burlington ticket#12619, County of Wellington (Ticket #21712)
If Not (glbCompSerial = "S/N - 2351W" Or glbWFC Or glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Or glbCompSerial = "S/N - 2262W") Then
    If Len(dlpDatComp.Text) = 0 Then dlpDatComp.Text = Date
End If

If Not chkSemnr() Then Exit Sub

If fglbNew Or frmAttendance.Enabled = True Then
    If clpCode(4).Caption <> "Unassigned" And clpCode(4) <> "" Then
        If Not UpdAttandance() Then Exit Sub
    End If
    If clpCode(5).Caption <> "Unassigned" And clpCode(5) <> "" Then
        If Not UpdSkills() Then Exit Sub
    End If
End If

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    Call UpdUStats(Me) ' update user's stats (who did it and when)
    Call Set_Control("U", Me, rsDATA)
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    xID = rsDATA("ES_ID")
    Data1.Refresh
    
    Set rsGrid = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True
Else
    'Friesens - Ticket #16189 and City of Chatham-Kent - Ticket #16794
    If glbCompSerial <> "S/N - 2279W" And glbCompSerial <> "S/N - 2188W" Then
        'Ticket #22682: Release 8.0 - Set older Follow Up records as Completed first if uncompleted
        'follow up records are found for Salary, before adding a new follow up record.
        If fglbNew Then
            glbFollowUpList = "EDUC"
            If Older_FollowUp_Records_Found(glbFollowUpList) Then
                frmFollowUpList.Show 1
            End If
        End If
    
        If Not updFollow("U") Then
            Exit Sub
        End If
    End If
    '7.9 - Enhancement - For all the clients now
    'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
    'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        If IsDate(dlpRenewal.Text) And IsDate(dlpDatComp.Text) Then
            'Call function to compute the Renewal Date if this Course exists in the
            'Employee's Training list and follow up record exists
            Call Update_TrainingList_FollowUp(dlpDatComp.Text, txtMain.Text, lblPosCode.Caption)
        ElseIf Trim(dlpRenewal.Text) = "" And IsDate(dlpDatComp.Text) Then
            'Only Complete Date is entered, there is no Renewal Date
            'Delete the Training record and the follow up record for this course and job
            Call Delete_TrainRec_FollowUp(txtMain.Text, lblPosCode.Caption)
        End If
    'End If
    
    Call UpdUStats(Me) ' update user's stats (who did it and when)
    Call Set_Control("U", Me, rsDATA)
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    xID = rsDATA("ES_ID")
    Data1.Refresh
    
    Set rsGrid = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True
    
End If

If gsAttachment_DB Then
    If glbDocNewRecord Then 'New Record only
        If Len(glbDocImpFile) > 0 Then
            glbDocKey = xID
            Call AttachmentAdd(glbLEE_ID, glbDocImpFile, glbDocType, glbDocDesc)
        End If
    End If
    glbDocImpFile = ""
End If


If fglbNew Or frmAttendance.Enabled = True Then
    clpCode(4).Text = ""
    txtAttHrs = ""
    chkSeniority.Value = False 'added by RAUBREY 4/4/97
    chkIncentive.Value = False
    clpCode(5).Text = ""
    txtSkillsExp = ""
    txtCompanyName.Text = ""
    txtTrainerName.Text = ""
End If

fglbNew = False

Call SET_UP_MODE
'Call ST_UPD_MODE(True)

If NextFormIF("Course") Then
    Call cmdNew_Click
End If

Exit Sub

Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREDSEM", "Update")
Call RollBack '23July99 js

End Sub

'Private Sub cmdOK_GotFocus()
 '  Call SetPanHelp(ActiveControl)
''End Sub

Sub cmdPrint_Click()
Dim RHeading As String, xReport, dscGroup$

'cmdPrint.Enabled = False
RHeading = lblEEName.Caption & "'s Continuing Education"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
dscGroup$ = "PgHeading" & "= '" & Replace(RHeading, "'", "' + chr(39) + '") & "'" 'LAURA DEC 10, 1997
Me.vbxCrystal.Formulas(0) = dscGroup$                     '
Me.vbxCrystal.BoundReportHeading = RHeading               '

If Not glbtermopen Then
    xReport = glbIHRREPORTS & "rgedsem.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
    End If
    Me.vbxCrystal.SelectionFormula = "{HREDSEM.ES_EMPNBR}=" & glbLEE_ID & " "
End If

If glbtermopen Then
    xReport = glbIHRREPORTS & "rgedse1.rpt"

    Me.vbxCrystal.ReportFileName = xReport
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
    End If
    Me.vbxCrystal.SelectionFormula = "{Term_HREDSEM.TERM_SEQ}=" & glbTERM_Seq & " "
End If
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1
'cmdPrint.Enabled = True

End Sub

Sub cmdView_Click()
Dim RHeading As String, xReport, dscGroup$

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

'cmdPrint.Enabled = False
RHeading = lblEEName.Caption & "'s Continuing Education"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
dscGroup$ = "PgHeading" & "= '" & Replace(RHeading, "'", "' + chr(39) + '") & "'" 'LAURA DEC 10, 1997
Me.vbxCrystal.Formulas(0) = dscGroup$                     '
Me.vbxCrystal.BoundReportHeading = RHeading               '

If Not glbtermopen Then
    xReport = glbIHRREPORTS & "rgedsem.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
    End If
    Me.vbxCrystal.SelectionFormula = "{HREDSEM.ES_EMPNBR}=" & glbLEE_ID & " "
End If

If glbtermopen Then
    xReport = glbIHRREPORTS & "rgedse1.rpt"

    Me.vbxCrystal.ReportFileName = xReport
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
    End If
    Me.vbxCrystal.SelectionFormula = "{Term_HREDSEM.TERM_SEQ}=" & glbTERM_Seq & " "
End If
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
'cmdPrint.Enabled = True

End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub ATTCode_Desc(Indx As Integer)
Dim SQLQ As String
Dim rsTA As New ADODB.Recordset
'''On Error GoTo AttCode_desc_Err


If Len(clpCode(Indx).Text) > 0 And Indx = 4 Then
    
    SQLQ = "SELECT TB_INDICATOR FROM HRTABL WHERE TB_NAME='ADRE' AND TB_KEY = '" & clpCode(Indx).Text & "'"
    rsTA.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsTA.EOF Then
        If rsTA("TB_INDICATOR") = True Then
            chkIncentive.Value = True
        Else
            chkIncentive.Value = False
        End If
    End If
End If

Exit Sub

AttCode_desc_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Code Snap", "TABL", "SELECT")
Call RollBack '23July99 js

End Sub

Private Sub cmdImport_Click()
    glbDocNewRecord = fglbNew
    glbDocName = "EdSem"
    If fglbNew Then
        glbDocKey = 0
    Else
        glbDocKey = rsDATA("ES_ID")
    End If
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmESEMINARS")
End Sub

Private Sub cmdMissingCourses_Click()
    frmEMissTrainLst.Show 1
    
    If EERetrieve_Null_on_Top() = False Then Exit Sub
    
End Sub

Private Sub cmdRetest_Click()
Unload Me
Load frmESEMRETEST
End Sub

Private Sub cmdSuccessionPlan_Click()
    Unload frmESEMINARS
    Set frmESEMINARS = Nothing 'carmen may 00
    Load frmESuccession
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRDEPNTS", "SELECT")
Call RollBack '23July99 js

End Sub

Function EERetrieve()
Dim SQLQ As String
EERetrieve = False

'''On Error GoTo EERError

Screen.MousePointer = HOURGLASS

If glbtermopen Then         'Lucy July 5, 2000
    SQLQ = "Select * from Term_HREDSEM"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY ES_CTYPE ASC, ES_DATCOMP DESC, ES_EMPNBR"
Else
    SQLQ = "Select * from HREDSEM"
    SQLQ = SQLQ & " where ES_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY ES_CTYPE ASC, ES_DATCOMP DESC, ES_EMPNBR"
End If


Data1.RecordSource = SQLQ
Data1.Refresh

Set rsGrid = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Call UpConttotal

EERetrieve = True

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DEPRetrieve", "HREDSEM", "SELECT")
Call RollBack '23July99 js

Exit Function

End Function

Private Sub dlpDatComp_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub dlpDatComp_LostFocus()
    '7.9 - Enhancement - For all the clients now
    'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
    'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        If IsDate(dlpDatComp) And Len(Trim(txtMain.Text)) > 0 Then
            'Call function to compute the Renewal Date
            dlpRenewal = Compute_Renewal_Date(dlpDatComp.Text, txtMain.Text, lblPosCode.Caption)
        End If
    'End If
End Sub

Private Sub dlpRenewal_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub dlpStartDate_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub dlpStartDate_LostFocus()
    'County of Wellington (Ticket #21712)
    'Not Burlington ticket# 12619 ' Not Bird Packaging Limited Ticket #13701
    'Frank 10/20/2003
    'As Jerry request, default Date Complete to Start Date if it's blank
    If glbCompSerial <> "S/N - 2351W" And glbCompSerial <> "S/N - 2387W" And glbCompSerial <> "S/N - 2262W" Then
        If Len(dlpDatComp) = 0 Then
            If IsDate(dlpStartDate) Then
                dlpDatComp = dlpStartDate
            End If
        End If
        '7.9 - Enhancement - For all the clients now
        'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
        'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
            If IsDate(dlpDatComp) And Len(Trim(txtMain.Text)) > 0 Then
                'Call function to compute the Renewal Date
                dlpRenewal = Compute_Renewal_Date(dlpDatComp.Text, txtMain.Text, lblPosCode.Caption)
            End If
        'End If
    End If
End Sub

Private Sub Form_Activate()
    glbOnTop = "FRMESEMINARS"
    Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMESEMINARS"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer, x%       ' records found

glbOnTop = "FRMESEMINARS"

'If Course Code Master setup, then use it instead of Course Code lookup from HRTABL
'Ticket #12204
If getCrsCodeMasterFlag Then
    clpCode(0).DataField = ""
    clpCode(0).Visible = False
    txtMain.DataField = "ES_CRSCODE"
    lblDesc.Caption = ""
    frmCourseCode.Visible = True
    frmCourseCode.Top = clpCode(0).Top
    frmCourseCode.Left = clpCode(0).Left
    frmCourseCode.BorderStyle = 0
End If

If glbtermopen Then         'Lucy July 5, 2000
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

Screen.MousePointer = DEFAULT

If glbCompSerial = "S/N - 2214W" Then
   lbltitle(6).Caption = "Tuition $"
   lbltitle(7).Caption = "Travel $"
   lbltitle(15).Caption = "Daily Allowance $"
End If

'7.9 - Enhancement - For all the clients now
'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
    'Ticket #20518 - Show the Retest button when Course Code Master is blank
    If CourseCodeMaster_Blank Then
        cmdRetest.Visible = True
    Else
        cmdRetest.Visible = False
    End If
    
    LabelPos.Visible = True
    lblPosCode.Visible = True
    lblPosDesc.Visible = True
    cmdMissingCourses.Visible = True
'End If

'Ticket #14526 Show mem fields for Course Desc and Course Location
If glbCompSerial = "S/N - 2188W" Then 'City of Chatham-Kent
    frmCourseDesc.Visible = True
    memCourseDesc.DataField = "ES_COURSE_DESC"
    memCourseLoc.DataField = "ES_COURSE_LOCATION"
    cmdMissingCourses.Top = 7560
End If

'Hemu - 05/29/2003 Begin - Ticket # 4204
'If glbCompSerial = "S/N - 2161W" Then
    clpCode(1).TextBoxWidth = 1200
    clpCode(0).TextBoxWidth = 1200
    clpCode(1).MaxLength = 8
    clpCode(0).MaxLength = 8
'Else
'    clpCode(1).TextBoxWidth = 870
'    clpCode(0).TextBoxWidth = 870
'    clpCode(1).MaxLength = 4
'    clpCode(0).MaxLength = 4
'End If
'Hemu - 05/29/2003 End

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

If glbCompSerial = "S/N - 2347W" Then  'For Surrey Place
    lbltitle(6).Caption = "Meals/Other $"
    lbltitle(15).Caption = "Travel $"
    lbltitle(7).Caption = "Registration $"
    medEECont(0).Tag = "20-Amount Meals/Other"
    medEECont(2).Tag = "20-Amount Travel"
    medEECont(1).Tag = "20-Amount Registration"
End If
If glbOttawaCCAC Then
    lbltitle(6) = "Travel $"
    lbltitle(7) = "Registration $"
    medEECont(0).Tag = "20-Amount Travel"
    medEECont(1).Tag = "20-Amount Registration"
    
End If
If glbLinamar Then
    lbltitle(17).FontBold = True
End If

'Ticket #22701 - County of Lanark
If glbCompSerial = "S/N - 2172W" Then
    lbltitle(17).FontBold = True
End If


If gSec_Inq_SUCCESSION Then
    cmdSuccessionPlan.Visible = True
    cmdSuccessionPlan.Enabled = True
Else
    cmdSuccessionPlan.Visible = False
End If

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

'Ticket #20447 - Jerry said to open up for everyone
'Bryan added Mar 26, 2007 Ticket#12859
'If glbCompSerial = "S/N - 2188W" Then
    lblCEUType.Visible = True
    lblCEUCred.Visible = True
    txtCEUCred.Visible = True
    txtCEUCred.DataField = "ES_CEUCREDIT"
    clpCEUType.Visible = True
    clpCEUType.DataField = "ES_CEUTYPE"
'End If

If glbWFC Then 'Ticket #13520
    lbltitle(21).FontBold = True
    lbltitle(4).FontBold = True 'Ticket #15396
    lbltitle(5).FontBold = True 'Ticket #15396
    lbltitle(17).FontBold = True 'Ticket #15396
End If

'Release 8.0 - Ticket #24441: Make Course Code Mandatory
lbltitle(21).FontBold = True

Screen.MousePointer = HOURGLASS

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    frmESEMINARS.Caption = "Continuing Education - " & Left$(glbLEE_SName, 5)
    frmESEMINARS.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = ShowEmpnbr(lblEEID)

Call Display_Value

Call ST_UPD_MODE(True)

If Not gSec_Upd_Education_Seminars Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
End If

Call INI_Controls(Me)

For x% = 0 To 25
    Call setCaption(lbltitle(x%))
Next

Call setCaption(lbltitle(27))
Call setCaption(lblCEUType)
clpCEUType.Tag = "00-" & lblCEUType.Caption & " - Code"
Call setCaption(lblCEUCred)
txtCEUCred.Tag = "11-" & lblCEUCred.Caption

Label1.Caption = lStr(Label1.Caption)
chkPresentor.Caption = lStr(chkPresentor.Caption)

For x% = 0 To 17
    'Call setCaption(lbltitle(x%))
    vbxTrueGrid.Columns(x%).Caption = lStr((vbxTrueGrid.Columns(x%).Caption))
Next

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
Dim c As Long

'''On Error GoTo Eh

If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    If Me.Height >= 11000 Then
        scrControl.Value = 0
        
        pcCourseInfo.Top = 2200
        
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        scrControl.Left = Me.ScaleWidth - scrControl.Width
        scrControl.Height = Me.Height - 2500
        
        scrControl.Max = 2200
        
    End If


'    'Horizontal Scroll
'    scrHScroll.Width = Me.Width - 200
'    If Me.Width >= 11190 Then '9700 Then
'        scrHScroll.Value = 0
'        scrHScroll.Visible = False
'    Else
'        scrHScroll.Visible = True
'        scrHScroll.Top = Me.Height - 700
'        scrHScroll.Width = Me.Width - 120
'    End If
    
End If

exH:
    Exit Sub
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form_Resize", "Continuing Education", "edit/Add")
    Resume exH
End Sub

Private Sub Form_Unload(Cancel As Integer)

    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmESEMINARS = Nothing 'carmen may 00
    Call NextForm
End Sub

Private Sub imgIcon_Click()
Call txtMain_DblClick
End Sub


Private Sub medEECont_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medEECont_LostFocus(Index As Integer)
    Call UpConttotal
End Sub

Private Sub AttSkill(YN)
frmAttendance.Enabled = YN And gSec_Add_Attendance  'Ticket #22682
frmSkills.Enabled = YN
clpCode(4).Enabled = YN
txtAttHrs.Enabled = YN
chkSeniority.Enabled = YN
chkIncentive.Enabled = YN
clpCode(5).Enabled = YN
txtSkillsExp.Enabled = YN
If YN Then
    lbltitle(11).ForeColor = &H0
    lbltitle(12).ForeColor = &H0
    lbltitle(10).ForeColor = &H0
    lbltitle(9).ForeColor = &H0
    chkSeniority.ForeColor = &H0
    chkIncentive.ForeColor = &H0
Else
    lbltitle(11).ForeColor = &H808080
    lbltitle(12).ForeColor = &H808080
    lbltitle(10).ForeColor = &H808080
    lbltitle(9).ForeColor = &H808080
    chkSeniority.ForeColor = &H808080
    chkIncentive.ForeColor = &H808080
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

fUPMode = TF    ' update mode

'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT

'If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
'medContTotal.Enabled = False
'medEECont(0).Enabled = False
'medEECont(1).Enabled = False
'medEECont(2).Enabled = False
'medEECont(3).Enabled = False
'medEECont(3).Enabled = False
'clpCode(0).Enabled = False
'clpCode(1).Enabled = False
'clpCode(2).Enabled = False
'clpCode(3).Enabled = False
'txtCourseHRS.Enabled = False
'txtCourseName.Enabled = False
'dlpDatComp.Enabled = False
'txtExtName.Enabled = False
'txtKeyword.Enabled = False
'dlpRenewal.Enabled = False
'dlpSchDate.Enabled = False
'dlpStartDate.Enabled = False
'txtAccount.Enabled = False
'chkPresentor.Enabled = False
'Else
medContTotal.Enabled = TF
medEECont(0).Enabled = TF
medEECont(1).Enabled = TF
medEECont(2).Enabled = TF
medEECont(3).Enabled = TF
medEECont(3).Enabled = TF
clpCode(0).Enabled = TF
txtMain.Enabled = TF
If glbWFC Then 'Ticket #24708 Franks 11/27/2013
    '"   When adding a new Continuing Education record, Course Type, Conducted By, Method Used and CEU Type cannot be entered by the user.
    '"   Make the same for edit.
    clpCode(1).Enabled = False
    'clpCode(2).Enabled = False
    clpCEUType.Enabled = False
    clpCode(7).Enabled = False
    lbltitle(1).Enabled = False
    'lbltitle(3).Enabled = False
    lblCEUType.Enabled = False
    lbltitle(24).Enabled = False
    ''Ticket #24767 Franks 12/11/2013 - "   The Course Name is read only
    lbltitle(2).Enabled = False
    txtCourseName.Enabled = False
    ''Ticket #24767 Franks 12/11/2013 - "   The Coordinated By is read only
    lbltitle(25).Enabled = False
    clpCode(6).Enabled = False
Else
    clpCode(1).Enabled = TF
    clpCode(2).Enabled = TF
    clpCEUType.Enabled = TF
    clpCode(7).Enabled = TF
    txtCourseName.Enabled = TF
    clpCode(6).Enabled = TF
End If
clpCode(3).Enabled = TF
txtCourseHRS.Enabled = TF
txtCompanyName.Enabled = TF
txtTrainerName.Enabled = TF
dlpDatComp.Enabled = TF
txtExtName.Enabled = TF
txtKeyword.Enabled = TF
dlpRenewal.Enabled = TF
dlpSchDate.Enabled = TF
dlpStartDate.Enabled = TF
txtAccount.Enabled = TF
chkPresentor.Enabled = TF

txtCEUCred.Enabled = TF
clpEmpCur.Enabled = TF
clpOherCur.Enabled = TF
clpEmployerCur.Enabled = TF
clpAcomCur.Enabled = TF
medEECont(4).Enabled = TF
clpLearnCur.Enabled = TF
clpTotCur.Enabled = TF

'vbxTrueGrid.Enabled = FT

If glbCompSerial = "S/N - 2225W" Then  'PowerStream Inc. (Markham Hydro) - Ticket #13925
    txtCourseName.Enabled = False
End If

glbDocName = "EdSem"
If gsAttachment_DB Then
    If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
        If rsDATA.RecordCount > 0 Then
            If Not IsNull(rsDATA("ES_DOCKEY")) Then
                glbDocKey = rsDATA("ES_DOCKEY")
            Else
                glbDocKey = 0
            End If
        Else
            If Not IsNull(Data1.Recordset("ES_DOCKEY")) Then
                glbDocKey = Data1.Recordset("ES_DOCKEY")
            Else
                glbDocKey = 0
            End If
        End If
    End If
    
    Call DispimgIcon(Me, "frmESEMINARS")
    If gSec_Upd_Education_Seminars And Not glbtermopen Then
        If Data1.Recordset.BOF And Data1.Recordset.EOF Then
            cmdImport.Visible = False
        Else
            cmdImport.Visible = True
        End If
    End If
End If

'End If
If Not glbtermopen Then Call AttSkill(fglbNew)

End Sub

Private Sub scrControl_Change()
pcCourseInfo.Top = 2200 - scrControl.Value
End Sub

Private Sub txtAccount_GotFocus()
  Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtAttHrs_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtCEUCred_GotFocus()
  Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtCompanyName_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtCourseHRS_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtCourseHRS_LostFocus()
If Not IsNumeric(txtCourseHRS) Then txtCourseHRS = 0
If glbWFC Then 'Ticket #15522
    If fglbNew Then
        If glbUNION = "NONE" Or glbUNION = "EXEC" Then
            medEECont(1).Text = txtCourseHRS * 50
        Else
            medEECont(1).Text = txtCourseHRS * 35
        End If
    End If
End If
End Sub

Private Sub txtCourseName_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtExtName_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtKeyword_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtMain_Change()
    lblDesc = Replace(GetCrsCodeDesc(txtMain), "&", "&&")
    If Len(txtMain) > 0 Then
        If lblDesc.Caption = "" Then lblDesc.Caption = "Unassigned"
    End If
End Sub

Private Sub txtMain_DblClick()
    OldtxtMain = txtMain.Text
    glbCrsCode = txtMain.Text
    glbCrsCodeDesc = lblDesc.Caption
    Call Get_CourseCode(False, clpCode(1).Text) 'Ticket #13520, added clpCode(1).Text
    txtMain.Text = glbCrsCode
    lblDesc.Caption = glbCrsCodeDesc
End Sub

Private Sub txtMain_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtMain_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtMain_LostFocus()
    If Not (OldtxtMain = txtMain.Text) Then
        If glbCrsCodeStrArr(17) = "*" Then ' * - means the changes come from Lookup
            Call SaveArrayInFields
            
            'Ticket #22075 - The total was not getting updated.
            Call UpConttotal
        Else
            'Ticket #22075 - Populate with Course Code Master data if record found.
            If Get_CourseCode_Master_Data(clpCode(1).Text, txtMain.Text) Then
                If glbCrsCodeStrArr(17) = "*" Then ' * - means the changes come from Lookup
                    Call SaveArrayInFields
                    
                    'Ticket #22075 - The total was not getting updated.
                    Call UpConttotal
                    
                    If Len(Trim(txtCourseName)) = 0 Then
                        txtCourseName = Replace(lblDesc.Caption, "&&", "&")
                    End If
                End If
            End If
        End If
    End If
    If Len(lblDesc.Caption) > 0 Then
        If Len(Trim(txtCourseName)) = 0 Then
            txtCourseName = Replace(lblDesc.Caption, "&&", "&")
        End If
    End If
End Sub

Private Sub txtSkillsExp_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub UpConttotal()
Dim x%, xTotal
xTotal = ""
For x% = 0 To 4 '3
    If IsNumeric(medEECont(x%)) Then xTotal = Val(xTotal) + Val(medEECont(x%))
Next
medContTotal = xTotal

End Sub

Private Function UpdAttandance()
Dim SQLQ As String
Dim rsTB As New ADODB.Recordset, xCrtJOB
Dim iRow As Integer, Msg As String, Edat As String
Dim dynHRAT As New ADODB.Recordset
Dim iRec As Integer, xDOA
Dim newline As String
Dim xSHIFT, xSuper

iRec = False
UpdAttandance = False

'''On Error GoTo CrAtt_Err

If glbtermopen Then
    rsTB.Open "SELECT JH_EMPNBR FROM Term_JOB_HISTORY WHERE TERM_SEQ=" & glbTERM_Seq, gdbAdoIhr001, adOpenStatic
    
    SQLQ = "INSERT INTO Term_ATTENDANCE "
Else
    rsTB.Open "SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenStatic
    
    SQLQ = "INSERT INTO HR_ATTENDANCE "
End If
xCrtJOB = Not rsTB.EOF
rsTB.Close

SQLQ = SQLQ & "(AD_COMPNO,AD_EMPNBR,AD_DOA,AD_REASON,AD_HRS,AD_COMM,"
'existing before Langs Farm ticket #14985 multiposition employee:
'If xCrtJOB Then SQLQ = SQLQ & "AD_SHIFT,AD_SUPER,AD_JOB,AD_SALARY,AD_SALCD,AD_DHRS,AD_WHRS,AD_ORG,"

If xCrtJOB Then SQLQ = SQLQ & "AD_SHIFT,AD_SUPER,AD_JOB,AD_SALARY,AD_SALCD,AD_DHRS,AD_WHRS,AD_ORG,"

SQLQ = SQLQ & "AD_LDATE, AD_LTIME, AD_LUSER"
If glbtermopen Then SQLQ = SQLQ & ", TERM_SEQ "
SQLQ = SQLQ & ") "

'existing before Langs Farm ticket #14985 multiposition employee changes:
'SQLQ = SQLQ & "SELECT '001' AS AD_COMPNO, ED_EMPNBR AS AD_EMPNBR,"

'Simona - begin - Langs Farm ticket #14985 multiposition employee case
If glbMulti Then
    SQLQ = SQLQ & "SELECT TOP 1 '001' AS AD_COMPNO, ED_EMPNBR AS AD_EMPNBR,"
Else
    SQLQ = SQLQ & "SELECT '001' AS AD_COMPNO, ED_EMPNBR AS AD_EMPNBR,"
End If
'Simona - end - Langs Farm ticket #14985

If dlpDatComp.Text = "" Then
    SQLQ = SQLQ & Date_SQL(dlpStartDate.Text) & " AS AD_DOA,"
Else
    SQLQ = SQLQ & Date_SQL(dlpDatComp.Text) & " AS AD_DOA,"
End If

SQLQ = SQLQ & "'" & clpCode(4).Text & "' AS AD_REASON,"
SQLQ = SQLQ & CDbl(txtAttHrs) & " AS AD_HRS,"
SQLQ = SQLQ & "'" & Replace(txtCourseName, "'", "''") & Chr(13) & Chr(10)
SQLQ = SQLQ & Replace(txtExtName, "'", "''") & "' AS AD_COMM,"

If xCrtJOB Then SQLQ = SQLQ & "JH_SHIFT,JH_REPTAU,JH_JOB,SH_SALARY,SH_SALCD,JH_DHRS,JH_WHRS,JH_ORG,"

SQLQ = SQLQ & Date_SQL(Date) & " AS AD_LDATE,"
SQLQ = SQLQ & "'" & Time$ & "' AS AD_LTIME,"
SQLQ = SQLQ & "'" & glbUserID & "' AS AD_LUSER "

If glbtermopen Then
    SQLQ = SQLQ & "," & glbTERM_Seq & " AS TERM_SEQ "
    SQLQ = SQLQ & "FROM "
    
    If glbOracle Then
        If xCrtJOB Then
            SQLQ = SQLQ & "Term_HREMP, Term_JOB_HISTORY, Term_SALARY_HISTORY where Term_HREMP.ED_EMPNBR=Term_JOB_HISTORY.JH_EMPNBR) "
            SQLQ = SQLQ & " and Term_JOB_HISTORY.JH_EMPNBR=Term_SALARY_HISTORY.SH_EMPNBR "
            SQLQ = SQLQ & "AND Term_JOB_HISTORY.JH_JOB=Term_SALARY_HISTORY.SH_JOB "
            SQLQ = SQLQ & "AND Term_JOB_HISTORY.JH_CURRENT=Term_SALARY_HISTORY.SH_CURRENT "
        Else
            SQLQ = SQLQ & "Term_HREMP "
        End If
    Else
        If xCrtJOB Then
            SQLQ = SQLQ & "(Term_HREMP LEFT JOIN Term_JOB_HISTORY ON Term_HREMP.ED_EMPNBR=Term_JOB_HISTORY.JH_EMPNBR) "
            SQLQ = SQLQ & "LEFT JOIN Term_SALARY_HISTORY ON Term_JOB_HISTORY.JH_EMPNBR=Term_SALARY_HISTORY.SH_EMPNBR "
            SQLQ = SQLQ & "AND Term_JOB_HISTORY.JH_JOB=Term_SALARY_HISTORY.SH_JOB "
            SQLQ = SQLQ & "AND Term_JOB_HISTORY.JH_CURRENT=Term_SALARY_HISTORY.SH_CURRENT "
        Else
            SQLQ = SQLQ & "Term_HREMP "
        End If
    End If
    
    SQLQ = SQLQ & "WHERE Term_HREMP.TERM_SEQ=" & glbTERM_Seq
    
    If xCrtJOB Then SQLQ = SQLQ & " AND (Term_JOB_HISTORY.JH_CURRENT<>0 )"
    
    gdbAdoIhr001X.Execute SQLQ
Else
    SQLQ = SQLQ & "FROM "
    
    If glbOracle Then
        If xCrtJOB Then
            SQLQ = SQLQ & "HREMP, HR_JOB_HISTORY, HR_SALARY_HISTORY where HREMP.ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR  "
            SQLQ = SQLQ & "and HR_JOB_HISTORY.JH_EMPNBR=HR_SALARY_HISTORY.SH_EMPNBR "
            SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_JOB=HR_SALARY_HISTORY.SH_JOB "
            SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_CURRENT=HR_SALARY_HISTORY.SH_CURRENT "
        Else
            SQLQ = SQLQ & "HREMP "
        End If
    Else
        If xCrtJOB Then
            SQLQ = SQLQ & "(HREMP LEFT JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR) "
            SQLQ = SQLQ & "LEFT JOIN HR_SALARY_HISTORY ON HR_JOB_HISTORY.JH_EMPNBR=HR_SALARY_HISTORY.SH_EMPNBR "
            SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_JOB=HR_SALARY_HISTORY.SH_JOB "
            SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_CURRENT=HR_SALARY_HISTORY.SH_CURRENT "
        Else
            SQLQ = SQLQ & "HREMP "
        End If
    End If
    
    If glbOracle Then
        SQLQ = SQLQ & " and ED_EMPNBR=" & glbLEE_ID
    Else
        SQLQ = SQLQ & "WHERE ED_EMPNBR=" & glbLEE_ID
    End If
    
    If xCrtJOB Then SQLQ = SQLQ & " AND (JH_CURRENT<>0 )"
        
    gdbAdoIhr001.Execute SQLQ
    
End If


UpdAttandance = True

Exit Function

CrAtt_Err:
If Err = 3022 Then
    MsgBox "Duplicate Attendance record exists. This Attendance record will not be added."
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Function
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Attendance Insert", "HRATTEND", "INSERT")
Resume Next

End Function

Private Function updFollow(xType) 'created by Laura on 11/1/97
Dim newline As String
Dim SQLQ As String
Dim Msg As String, Edat As String
Dim iRec As Integer
Dim rsTB As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim Edit1 As Integer
Dim rsTT As New ADODB.Recordset

'Don't need a message for follow up - Jerry asked for v7.6
newline = Chr$(13) & Chr$(10)
updFollow = False


'''On Error GoTo CrFollow_Err

newline = Chr$(13) & Chr$(10)

If IsDate(fglHredsem) Then 'Jaddy 11/15
    ' DATE Renewal IS NOW MANDATORY
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & Val(glbLEE_ID)
    SQLQ = SQLQ & " AND EF_FREAS = 'EDUC'"
    SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(fglHredsem)
    
    'Hemu - Because on updating the follow up record with changed Course Name and Descrip.
    '       it overwrites all the Comments of the EDUC and EF_FDATE for that employee
    SQLQ = SQLQ & " AND EF_COMMENTS ='" & Replace(fglCursName, "'", "''") & newline & fglExtName & "'"
    'Hemu End
    
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
    
    rsTB.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    If fglbNew And IsDate(dlpRenewal.Text) Then 'Jaddy 11/15
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = glbLEE_ID
        rsTB("EF_FDATE") = CVDate(dlpRenewal.Text)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
        End If
        rsTB("EF_FREAS") = "EDUC"
        rsTB("EF_COMMENTS") = txtCourseName & newline & txtExtName
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = glbUserID
        rsTB.Update
        rsTB.Close
        
        rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='EDUC'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
        If rsTT.EOF Then
            rsTT.AddNew
            rsTT("TB_COMPNO") = "001"
            rsTT("TB_NAME") = "FURE"
            rsTT("TB_KEY") = "EDUC"
            rsTT("TB_DESC") = "Continuing Education Review"
            rsTT("TB_LUSER") = glbUserID
            rsTT("TB_LDATE") = Date
            rsTT("TB_LTIME") = Time$
            rsTT.Update
        End If
        rsTT.Close
        
        'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
        'follow up record
        Call Grant_FollowUpCode_Security(glbUserID, "EDUC", "Continuing Education Review")
        
        updFollow = True
        'Msg = "A Follow Up Record was created!"
        'MsgBox Msg
        Exit Function
    End If
    If fglbNew = False And Edit1 = False And IsDate(dlpRenewal.Text) Then 'Jaddy 11/15
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = glbLEE_ID
        rsTB("EF_FDATE") = CVDate(dlpRenewal.Text)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
        End If
        rsTB("EF_FREAS") = "EDUC"
        rsTB("EF_COMMENTS") = txtCourseName & newline & txtExtName
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = glbUserID
        rsTB.Update
        rsTB.Close
        
        rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='EDUC'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
        If rsTT.EOF Then
            rsTT.AddNew
            rsTT("TB_COMPNO") = "001"
            rsTT("TB_NAME") = "FURE"
            rsTT("TB_KEY") = "EDUC"
            rsTT("TB_DESC") = "Continuing Education Review"
            rsTT("TB_LUSER") = glbUserID
            rsTT("TB_LDATE") = Date
            rsTT("TB_LTIME") = Time$
            rsTT.Update
        End If
        rsTT.Close
        
        'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
        'follow up record
        Call Grant_FollowUpCode_Security(glbUserID, "EDUC", "Continuing Education Review")
        
        updFollow = True
        'Msg = "A Follow Up Record was created!"
        'MsgBox Msg
        Exit Function
    End If
    If fglbNew = False And Edit1 = True And IsDate(dlpRenewal.Text) Then  'Jaddy 11/15 ' edited record
        'EOF?
        dynHRAT.MoveFirst
        Do Until dynHRAT.EOF
            'dynHRAT.Edit
            dynHRAT("EF_COMPNO") = "001"
            dynHRAT("EF_EMPNBR") = glbLEE_ID
            dynHRAT("EF_FDATE") = CVDate(dlpRenewal.Text)
            dynHRAT("EF_FREAS") = "EDUC"
            dynHRAT("EF_COMMENTS") = txtCourseName & newline & txtExtName
            dynHRAT("EF_LDATE") = Date
            dynHRAT("EF_LTIME") = Time$
            dynHRAT("EF_LUSER") = glbUserID
            dynHRAT.Update
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        If fglCursName <> txtCourseName Or fglExtName <> txtExtName Or fglHredsem <> dlpRenewal.Text Then
           'Msg = "A Follow Up Record was updated!"
            'MsgBox Msg
        End If
        updFollow = True
        Edit1 = True
        Exit Function
    End If
    If fglbNew = False And Edit1 = True And (Not IsDate(dlpRenewal.Text)) Then
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
        'Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    Else
        updFollow = True
    End If
End If
    
If Not IsDate(dlpRenewal.Text) Then
    updFollow = True
End If

Exit Function

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Function
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Function

Private Function UpdSkills()
Dim SQLQ As String
Dim iRow As Integer, Msg As String, Edat As String
Dim iRec As Integer
Dim dynHREMPSKL As New ADODB.Recordset
Dim newline As String, xDOA

newline = Chr$(13) & Chr$(10)
iRec = False
UpdSkills = False

'''On Error GoTo CrSklls2_Err

If Len(dlpDatComp.Text) = 0 Then
  ' dkostka - 12/12/2001 - Changed from SchDate to StartDate on request of jerryr, as StartDate is req'd.
    xDOA = dlpStartDate.Text
Else
    xDOA = dlpDatComp.Text
End If

If glbtermopen Then
    SQLQ = "SELECT * FROM Term_EMPSKL WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "SELECT * FROM HREMPSKL WHERE  SE_EMPNBR = " & glbLEE_ID
End If
SQLQ = SQLQ & " AND SE_SKILL = '" & clpCode(5).Text & "'"
If Len(txtSkillsExp) > 0 Then SQLQ = SQLQ & " AND SE_LEVEL = " & CDbl(txtSkillsExp)
SQLQ = SQLQ & " AND SE_DATE = " & Date_SQL(xDOA)

If glbtermopen Then
    dynHREMPSKL.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    dynHREMPSKL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

If dynHREMPSKL.EOF And dynHREMPSKL.BOF Then
    iRec = True
Else
    iRec = False
End If

If iRec = True And fglbNew Then
    dynHREMPSKL.AddNew
    dynHREMPSKL("SE_COMPNO") = "001"
    dynHREMPSKL("SE_EMPNBR") = IIf(glbtermopen, glbTERM_ID, glbLEE_ID)
    dynHREMPSKL("SE_SKILL") = clpCode(5).Text
    dynHREMPSKL("SE_SKILL_TABL") = "EDSK"
    dynHREMPSKL("SE_LEVEL") = Val(txtSkillsExp)
    dynHREMPSKL("SE_COMM1") = txtCourseName & newline & txtExtName
    dynHREMPSKL("SE_DATE") = CVDate(xDOA)
    dynHREMPSKL("SE_LDATE") = Date
    dynHREMPSKL("SE_LTIME") = Time$
    dynHREMPSKL("SE_LUSER") = glbUserID
    If glbtermopen Then dynHREMPSKL("TERM_SEQ") = glbTERM_Seq
    dynHREMPSKL.Update
    dynHREMPSKL.Close
    UpdSkills = True
    Exit Function
End If

If iRec = True And fglbNew = False Then ' edited record on Continuing Education screen but no Skills record
    dynHREMPSKL.AddNew
    dynHREMPSKL("SE_COMPNO") = "001"
    dynHREMPSKL("SE_EMPNBR") = IIf(glbtermopen, glbTERM_ID, glbLEE_ID)
    dynHREMPSKL("SE_SKILL") = clpCode(5).Text
    dynHREMPSKL("SE_SKILL_TABL") = "EDSK"
    dynHREMPSKL("SE_LEVEL") = Val(txtSkillsExp)
    dynHREMPSKL("SE_COMM1") = txtCourseName & newline & txtExtName
    dynHREMPSKL("SE_DATE") = CVDate(xDOA)
    dynHREMPSKL("SE_LDATE") = Date
    dynHREMPSKL("SE_LTIME") = Time$
    dynHREMPSKL("SE_LUSER") = glbUserID
    If glbtermopen Then dynHREMPSKL("TERM_SEQ") = glbTERM_Seq
    dynHREMPSKL.Update
    dynHREMPSKL.Close
    UpdSkills = True
    Exit Function
End If

Msg = "Duplicate Skills record found"
MsgBox Msg
dynHREMPSKL.Close

Exit Function

CrSklls2_Err:
If Err = 3022 Then
    MsgBox "Duplicate Skill record exists. This Skill record will not be added."
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Function
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Skills Insert", "HRSKILLS", "INSERT")
Resume Next

End Function
'Private Sub txtStartDate_KeyPress(KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub

Private Sub txtTrainerName_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    rsGrid.Bookmark = Bookmark
    If IsNull(rsGrid("ES_START")) And IsNull(rsGrid("ES_DATCOMP")) Then
        RowStyle.ForeColor = vbRed
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
        
        If glbtermopen Then         'Lucy July 5, 2000
            SQLQ = "Select * from Term_HREDSEM"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select * from HREDSEM"
            SQLQ = SQLQ & " where ES_EMPNBR = " & glbLEE_ID
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
    
        Set rsGrid = Data1.Recordset.Clone
        vbxTrueGrid.FetchRowStyle = True

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
Dim I%

Call Display_Value
Call UpConttotal

End Sub

Private Function RollBack()
'''On Error GoTo rr
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
rr:
End Function

Sub CrsName_Desc()
    'If clpCode(0).Caption = "Unassigned" Then
    '    txtCourseName = ""
    'Else
    'If Course Code is blank, don't wipe up the Course Name
    If Len(clpCode(0).Caption) > 0 Then
        'Frank 10/20/03
        'As Jerry request, if Course Name exists there, don't replace it
        If Len(Trim(txtCourseName)) = 0 Then
            txtCourseName = Replace(clpCode(0).Caption, "&&", "&")
        End If
    End If
End Sub

''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
Dim SQLQ
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    If glbtermopen Then
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
            rsDATA.CursorLocation = adUseServer
        End If
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    '7.9 - Enhancement - For all the clients now
    'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
    'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        lblPosDesc.Caption = ""
    'End If
Else
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    If glbtermopen Then
        SQLQ = "Select * from Term_HREDSEM"
        SQLQ = SQLQ & " WHERE ES_ID = " & Data1.Recordset!ES_ID
        SQLQ = SQLQ & " ORDER BY ES_CTYPE ASC, ES_DATCOMP DESC, ES_EMPNBR"
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "Select * "
        SQLQ = SQLQ & " from HREDSEM "
        SQLQ = SQLQ & " where ES_ID = " & Data1.Recordset!ES_ID
        SQLQ = SQLQ & " ORDER BY ES_CTYPE ASC, ES_DATCOMP DESC, ES_EMPNBR"
        If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
            rsDATA.CursorLocation = adUseServer
        End If
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
    
    '7.9 - Enhancement - For all the clients now
    'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
    'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        If lblPosCode.Caption <> "" Then
            lblPosDesc.Caption = " " & GetJobData(lblPosCode.Caption, "JB_DESCR", "")
        Else
            lblPosDesc.Caption = ""
        End If
    'End If
End If

Call SET_UP_MODE

Me.cmdModify_Click

End Sub

Private Function ChkDup()
Dim SQLQ, Logx, Msg$, SavReviewDate
Dim Response%
Dim rsTB As New ADODB.Recordset
Dim Title$, Msg1$, DgDef

ChkDup = False

Logx = False
If glbtermopen Then
    SQLQ = "SELECT ES_EMPNBR FROM Term_HREDSEM WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "SELECT ES_EMPNBR FROM HREDSEM WHERE ES_EMPNBR = " & glbLEE_ID
End If
' danielk - 10/30/2002 - if course code was blank it was just using start date - some people use
'                        start date, course type, and course name, leave course code blank
'                        Ticket #3061
If clpCode(0) <> "" Then
    SQLQ = SQLQ & " AND ES_CRSCODE = '" & clpCode(0) & "'"
ElseIf txtCourseName <> "" Then
    'Modified by Frank 11/18/03 Ticket #5109
    'SQLQ = SQLQ & " AND ES_COURSE = '" & txtCourseName.Text & "'"
    'Hemu - Ticket #16697
    If glbSQL Or glbOracle Then
        SQLQ = SQLQ & " AND ES_COURSE = '" & Replace(txtCourseName, "'", "'+char(39)+'") & "'"
    Else
        SQLQ = SQLQ & " AND ES_COURSE = '" & Replace(txtCourseName, "'", "'+chr(39)+'") & "'"
    End If
End If
' danielk - 10/30/2002 - end

' danielk - 10/25/2002 - was crashing if they entered scheduled date but left start date blank, should
'                        accept either one or both.  Ticket #3025
If Len(dlpStartDate.Text) > 0 Then
    ' they're using start date
    'SQLQ = SQLQ & " AND ES_START = " & IIf(glbSQL, "", "CVDATE") & "('" & Format(dlpStartDate, "mmm dd,yyyy") & "')"
    SQLQ = SQLQ & " AND ES_START = " & Date_SQL(dlpStartDate) & " "
Else 'Date_SQL
    ' they're using scheduled date
    'SQLQ = SQLQ & " AND ES_SCHEDULED = " & IIf(glbSQL, "", "CVDATE") & "('" & Format(dlpSchDate, "mmm dd,yyyy") & "')"
    SQLQ = SQLQ & " AND ES_SCHEDULED = " & Date_SQL(dlpSchDate) & " "
End If
' danielk - 10/31/2002 - Was reporting a dupe on edit/ok, should check to make sure the record is not
'                        the actual record the user is editing (if it's an edit), don't report that
'                        one as a dupe.  Ticket #3061
If rsDATA.EditMode <> adEditAdd Then SQLQ = SQLQ & " AND ES_ID<>" & Data1.Recordset("ES_ID")
If glbtermopen Then
    rsTB.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockReadOnly
Else
    rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockReadOnly
End If
If Not rsTB.EOF Then Logx = True
rsTB.Close
If Logx = True Then
    ' danielk - 10/25/2002 - wasn't working if they left start date blank and used scheduled date, fixed.
    '                        Ticket #3025
    ' danielk - 10/30/2002 - see above, ticket #3061
    If Len(clpCode(0)) > 0 Then
        ' using course code
        Msg$ = "Warning: 'Course Code' + "
        clpCode(0).SetFocus
    Else
        ' using course name
        Msg$ = "Warning: 'Course Name' + "
        txtCourseName.SetFocus
    End If
    If Len(dlpStartDate.Text) > 0 Then
        Msg$ = Msg$ & "'Start Date'  Duplicate.  "
    Else
        Msg$ = Msg$ & "'Scheduled Date'  Duplicate.  "
    End If
    ' danielk - 10/30/2002 - end
        Msg1$ = Chr(10) & "Select:" & Chr(10)
        Msg1$ = Msg1$ & "   'OK' to save changes," & Chr(10)
        Msg1$ = Msg1$ & "   'Cancel' to make changes to the transaction"
        DgDef = MB_OKCANCEL + MB_ICONSTOP + MB_DEFBUTTON1
        Title$ = "info:HR - CONTINUING EDUCATION ENTRY"
        Response% = MsgBox(Msg$ & Msg1$, DgDef, Title)
        If Response% = IDCANCEL Then
            ChkDup = False
            Exit Function
        Else
            ChkDup = True
        End If
Else
ChkDup = True
End If
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
    RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
    UpdateRight = gSec_Upd_Education_Seminars
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

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    frmESEMINARS.Caption = "Continuing Education - " & Left$(glbLEE_SName, 5)
    frmESEMINARS.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
 If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
'lblEEID = glbLEE_ID
lblEENum = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub

Private Sub imgSec_Click()
    Dim SQLQ
    SQLQ = getSQL("frmESEMINARS")
    Call FillMemoFile(SQLQ, "EdSem")
End Sub

Private Sub SaveArrayInFields()
Dim K As Integer
    clpCode(1).Text = glbCrsCodeStrArr(1) 'Course Type
    clpCode(6).Text = glbCrsCodeStrArr(2) 'Co-Ordinated By
    txtCompanyName.Text = glbCrsCodeStrArr(3) '
    txtTrainerName.Text = glbCrsCodeStrArr(4) '
    'txtCourseHRS.Text = glbCrsCodeStrArr(5)
    'Ticket #25148 Franks 03/03/2014 -  don't overwrite Course Hours
    If Len(txtCourseHRS.Text) = 0 Then txtCourseHRS.Text = glbCrsCodeStrArr(5)    'Course Hours
    'If txtCourseHRS.Text = 0 Then txtCourseHRS.Text = glbCrsCodeStrArr(5)    'Course Hours
    'Ticket #25519 Franks 05/22/2014 add len to the following code
    If Len(txtCourseHRS.Text) = 0 Then txtCourseHRS.Text = glbCrsCodeStrArr(5)     'Course Hours
    'If glbWFC Then 'Ticket #25178 Franks 03/11/2014 - Do not overwrite the HREDEM record with the zero Dollars.
        If Len(medEECont(0).Text) = 0 Then medEECont(0).Text = glbCrsCodeStrArr(6)    'Employee $
        If medEECont(0).Text = 0 Then medEECont(0).Text = glbCrsCodeStrArr(6)    'Employee $
        If Len(medEECont(2).Text) = 0 Then medEECont(2).Text = glbCrsCodeStrArr(7)    'Other Expenses $
        If medEECont(2).Text = 0 Then medEECont(2).Text = glbCrsCodeStrArr(7)    'Other Expenses $
        If Len(medEECont(1).Text) = 0 Then medEECont(1).Text = glbCrsCodeStrArr(8)    'Employer $
        If medEECont(1).Text = 0 Then medEECont(1).Text = glbCrsCodeStrArr(8)    'Employer $
        If Len(medEECont(3).Text) = 0 Then medEECont(3).Text = glbCrsCodeStrArr(9)    'Accommodation $
        If medEECont(3).Text = 0 Then medEECont(3).Text = glbCrsCodeStrArr(9)    'Accommodation $
        If Len(medEECont(4).Text) = 0 Then medEECont(4).Text = glbCrsCodeStrArr(10)    'Learning Material $
        If medEECont(4).Text = 0 Then medEECont(4).Text = glbCrsCodeStrArr(10)    'Learning Material $
    'Else
    '    medEECont(0).Text = glbCrsCodeStrArr(6) 'Employee $
    '    medEECont(2).Text = glbCrsCodeStrArr(7)  'Other Expenses $
    '    medEECont(1).Text = glbCrsCodeStrArr(8) 'Employer $
    '    medEECont(3).Text = glbCrsCodeStrArr(9) 'Accommodation $
    '    medEECont(4).Text = glbCrsCodeStrArr(10) 'Learning Material $
    'End If
    clpEmpCur.Text = glbCrsCodeStrArr(11) 'Currency
    clpOherCur.Text = glbCrsCodeStrArr(12) 'Currency
    clpEmployerCur.Text = glbCrsCodeStrArr(13) 'Currency
    clpAcomCur.Text = glbCrsCodeStrArr(14) 'Currency
    clpLearnCur.Text = glbCrsCodeStrArr(15) 'Currency
    clpTotCur.Text = glbCrsCodeStrArr(16) 'Currency
    'Ticket #24708 Franks 11/27/2013
    'clpCode(2).Text = glbCrsCodeStrArr(18) 'Conducted By
    clpCode(6).Text = glbCrsCodeStrArr(18) 'Coordinated By
    clpCEUType.Text = glbCrsCodeStrArr(19) 'CEU Type
    clpCode(7).Text = glbCrsCodeStrArr(20) 'Method Used
    For K = 1 To 20 '17
        glbCrsCodeStrArr(K) = ""
    Next K
End Sub

Private Function Compute_Renewal_Date(xDateComplete, xCourse, xJob)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsCourseMst As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim SQLQ As String
    Dim xCurRenPrd, xPrvRenPrd
    Dim xCurRenTyp, xPrvRenTyp, xDWMY As String
    Dim flgCrsFound As Boolean
    
    'To display the renewal date on the screen
    
    'Initialise
    xCurRenTyp = ""
    xPrvRenTyp = ""
    xCurRenPrd = 0
    xPrvRenPrd = 0
    flgCrsFound = False
    
    'Find out if this Course is Unique for each Position
    'If Unique for Each Position - then retrieve Renewal Period from Required Courses screen
    'If not Unique for Each Position - then retrieve Renewal Period from Course Code Mst screen
    SQLQ = "SELECT ES_CRSCODE,ES_UNIQUE_FOR_POS,ES_RENEW_CRS_CUR,ES_CUR_PRD_DWMY, ES_RENEW_CRS_PRV,ES_PRV_PRD_DWMY FROM HR_COURSECODE_MASTER"
    SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & xCourse & "'"
    rsCourseMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsCourseMst.EOF Then
        flgCrsFound = True
        'If Unique for Each Position - retrieve from Required Courses screen
        If rsCourseMst("ES_UNIQUE_FOR_POS") <> 0 Then
            'Unique for each Position
            SQLQ = "SELECT PC_CRSCODE,PC_RENEW_CRS_CUR,PC_CUR_PRD_DWMY,PC_RENEW_CRS_PRV,PC_PRV_PRD_DWMY FROM HR_JOB_COURSE "
            SQLQ = SQLQ & " WHERE PC_JOB = '" & xJob & "'"
            SQLQ = SQLQ & " AND PC_CRSCODE = '" & xCourse & "'"
            rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsReqCourse.EOF Then
                If Not IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) And Not IsNull(rsReqCourse("PC_CUR_PRD_DWMY")) And rsReqCourse("PC_RENEW_CRS_CUR") <> 0 And rsReqCourse("PC_CUR_PRD_DWMY") <> "" Then
                    xCurRenTyp = rsReqCourse("PC_CUR_PRD_DWMY")
                    xCurRenPrd = rsReqCourse("PC_RENEW_CRS_CUR")
                End If
                If Not IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) And Not IsNull(rsReqCourse("PC_PRV_PRD_DWMY")) And rsReqCourse("PC_RENEW_CRS_PRV") <> 0 And rsReqCourse("PC_PRV_PRD_DWMY") <> "" Then
                    xPrvRenTyp = rsReqCourse("PC_PRV_PRD_DWMY")
                    xPrvRenPrd = rsReqCourse("PC_RENEW_CRS_PRV")
                End If
            Else
                flgCrsFound = False
            End If
            rsReqCourse.Close
            Set rsReqCourse = Nothing
        Else
            'Not Unique for Each Position
            If Not IsNull(rsCourseMst("ES_RENEW_CRS_CUR")) And Not IsNull(rsCourseMst("ES_CUR_PRD_DWMY")) And rsCourseMst("ES_RENEW_CRS_CUR") <> 0 And rsCourseMst("ES_CUR_PRD_DWMY") <> "" Then
                xCurRenTyp = rsCourseMst("ES_CUR_PRD_DWMY")
                xCurRenPrd = rsCourseMst("ES_RENEW_CRS_CUR")
            End If
            If Not IsNull(rsCourseMst("ES_RENEW_CRS_PRV")) And Not IsNull(rsCourseMst("ES_PRV_PRD_DWMY")) And rsCourseMst("ES_RENEW_CRS_PRV") <> 0 And rsCourseMst("ES_PRV_PRD_DWMY") <> "" Then
                xPrvRenTyp = rsCourseMst("ES_PRV_PRD_DWMY")
                xPrvRenPrd = rsCourseMst("ES_RENEW_CRS_PRV")
            End If
        End If
    Else
        flgCrsFound = False
    End If
    rsCourseMst.Close
    Set rsCourseMst = Nothing
    
    If flgCrsFound Then
        'Compute Renewal Date for Training List record and Follow Up record as well
        SQLQ = "SELECT * FROM HR_TRAIN"
        SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
        If xJob <> "" And Not IsNull(xJob) Then
            SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
        Else
            SQLQ = SQLQ & " AND (TR_JOB IS NULL OR TR_JOB = '')"
        End If
        SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourse & "'"
        rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsHRTrain.EOF Then
            'Find out the Type of Position to get the Renewal Period for Renewal Date computation
            If rsHRTrain("TR_POS_TYPE") = "C" Or rsHRTrain("TR_POS_TYPE") = "T" And Not IsNull(rsHRTrain("TR_POS_TYPE")) Then
                If xCurRenTyp = "" Then
                    'No Current Renewal Period and so no Renewal Date
                    'Training List and Follow Up for this course should be deleted
                    Compute_Renewal_Date = ""
                Else
                    Select Case xCurRenTyp
                        Case "D"
                            xDWMY = "d"
                        Case "W"
                            xDWMY = "ww"
                        Case "M"
                            xDWMY = "m"
                        Case "Y"
                            xDWMY = "yyyy"
                    End Select
                    Compute_Renewal_Date = DateAdd(xDWMY, xCurRenPrd, CVDate(xDateComplete))
                End If
            ElseIf rsHRTrain("TR_POS_TYPE") = "P" And Not IsNull(rsHRTrain("TR_POS_TYPE")) Then
                If xPrvRenTyp = "" Then
                    'No Previous Renewal Period and so no Renewal Date
                    'Training List and Follow Up for this course should be deleted
                    Compute_Renewal_Date = ""
                Else
                    Select Case xPrvRenTyp
                        Case "D"
                            xDWMY = "d"
                        Case "W"
                            xDWMY = "ww"
                        Case "M"
                            xDWMY = "m"
                        Case "Y"
                            xDWMY = "yyyy"
                    End Select
                    Compute_Renewal_Date = DateAdd(xDWMY, xPrvRenPrd, CVDate(xDateComplete))
                End If
            Else
                'it's an independant course, not required by any current, temp or
                'tracked positions of the employee
                'Compute the date based on the Current Renewal Period
                If xCurRenTyp = "" Then
                    'No Current Renewal Period and so no Renewal Date
                    'Training List record of this course should be deleted
                    Compute_Renewal_Date = ""
                Else
                    Select Case xCurRenTyp
                        Case "D"
                            xDWMY = "d"
                        Case "W"
                            xDWMY = "ww"
                        Case "M"
                            xDWMY = "m"
                        Case "Y"
                            xDWMY = "yyyy"
                    End Select
                    Compute_Renewal_Date = DateAdd(xDWMY, xCurRenPrd, CVDate(xDateComplete))
                End If
            End If
        End If
        rsHRTrain.Close
        Set rsHRTrain = Nothing
    Else
        Compute_Renewal_Date = ""
    End If
    
End Function

Private Sub Update_TrainingList_FollowUp(xDateComplete, xCourse, xJob)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsCourseMst As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim SQLQ As String
    Dim xCurRenPrd, xPrvRenPrd
    Dim xCurRenTyp, xPrvRenTyp, xDWMY As String
    Dim flgCrsFound As Boolean
    Dim xComments As String
    
'    'Initialise
'    xCurRenTyp = ""
'    xPrvRenTyp = ""
'    xCurRenPrd = 0
'    xPrvRenPrd = 0
'    flgCrsFound = False
'
'    'Delete Training List and Follow Up record for the Courses which does not have appropriate
'    'Renewal period or not found in the Course Code Mst or Required Courses screen
'
'    'Find out if this Course is Unique for each Position
'    'If Unique for Each Position - then retrieve Renewal Period from Required Courses screen
'    'If not Unique for Each Position - then retrieve Renewal Period from Course Code Mst screen
'    SQLQ = "SELECT ES_CRSCODE,ES_UNIQUE_FOR_POS,ES_RENEW_CRS_CUR,ES_CUR_PRD_DWMY, ES_RENEW_CRS_PRV,ES_PRV_PRD_DWMY FROM HR_COURSECODE_MASTER"
'    SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & xCourse & "'"
'    rsCourseMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If Not rsCourseMst.EOF Then
'        flgCrsFound = True
'        'If xJob <> "" And Not IsNull(xJob) Then
'            'If Unique for Each Position - retrieve from Required Courses screen
'            If rsCourseMst("ES_UNIQUE_FOR_POS") Then
'                'Unique for Each Position course
'                SQLQ = "SELECT PC_CRSCODE, PC_RENEW_CRS_CUR,PC_CUR_PRD_DWMY,PC_RENEW_CRS_PRV,PC_PRV_PRD_DWMY FROM HR_JOB_COURSE "
'                SQLQ = SQLQ & " WHERE PC_JOB = '" & xJob & "'"
'                SQLQ = SQLQ & " AND PC_CRSCODE = '" & xCourse & "'"
'                rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                If Not rsReqCourse.EOF Then
'                    If Not IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) And Not IsNull(rsReqCourse("PC_CUR_PRD_DWMY")) And rsReqCourse("PC_RENEW_CRS_CUR") <> 0 And rsReqCourse("PC_CUR_PRD_DWMY") <> "" Then
'                        xCurRenTyp = rsReqCourse("PC_CUR_PRD_DWMY")
'                        xCurRenPrd = rsReqCourse("PC_RENEW_CRS_CUR")
'                    Else
'                        'No current renewal period
'                        flgCrsFound = False
'                    End If
'                    If Not IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) And Not IsNull(rsReqCourse("PC_PRV_PRD_DWMY")) And rsReqCourse("PC_RENEW_CRS_PRV") <> 0 And rsReqCourse("PC_PRV_PRD_DWMY") <> "" Then
'                        xPrvRenTyp = rsReqCourse("PC_PRV_PRD_DWMY")
'                        xPrvRenPrd = rsReqCourse("PC_RENEW_CRS_PRV")
'                    Else
'                        'No previous renwal period
'                        flgCrsFound = False
'                    End If
'                Else
'                    'Course not required by any job
'                    flgCrsFound = False
'                End If
'                rsReqCourse.Close
'                Set rsReqCourse = Nothing
'            Else
'                'Not Unique for each position required course
'                If Not IsNull(rsCourseMst("ES_RENEW_CRS_CUR")) And Not IsNull(rsCourseMst("ES_CUR_PRD_DWMY")) And rsCourseMst("ES_RENEW_CRS_CUR") <> 0 And rsCourseMst("ES_CUR_PRD_DWMY") <> "" Then
'                    xCurRenTyp = rsCourseMst("ES_CUR_PRD_DWMY")
'                    xCurRenPrd = rsCourseMst("ES_RENEW_CRS_CUR")
'                Else
'                    'No current renewal period
'                    flgCrsFound = False
'                End If
'                If Not IsNull(rsCourseMst("ES_RENEW_CRS_PRV")) And Not IsNull(rsCourseMst("ES_PRV_PRD_DWMY")) And rsCourseMst("ES_RENEW_CRS_PRV") <> 0 And rsCourseMst("ES_PRV_PRD_DWMY") <> "" Then
'                    xPrvRenTyp = rsCourseMst("ES_PRV_PRD_DWMY")
'                    xPrvRenPrd = rsCourseMst("ES_RENEW_CRS_PRV")
'                Else
'                    'No previous renewal period
'                    flgCrsFound = False
'                End If
'            End If
'        'Else
'        '    flgCrsFound = False
'        'End If
'    Else
'        'Course not found in the Course Code Master screen
'        flgCrsFound = False
'    End If
'    rsCourseMst.Close
'    Set rsCourseMst = Nothing
'
'    If flgCrsFound Then
        'Update Training List record with Completed Date and new Renewal Date
        'Update Follow Up record as well
        SQLQ = "SELECT * FROM HR_TRAIN"
        SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
        If xJob <> "" And Not IsNull(xJob) Then
            SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
        Else
            SQLQ = SQLQ & " AND (TR_JOB IS NULL OR TR_JOB = '')"
        End If
        SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourse & "'"
        rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsHRTrain.EOF Then
            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & Replace(xComments, "'", "''") & "%' AND EF_FDATE = " & Date_SQL(rsHRTrain("TR_RENEW"))
                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsFollowUp.EOF Then
                    rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                Else
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & Replace(xComments, "'", "''") & "%' "
                    SQLQ = SQLQ & " AND EF_COMPLETED <>1 ORDER BY EF_LDATE DESC"
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    Else
                        'Follow Up record missing, add a new one
                        rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, IIf(Trim(dlpRenewal.Text) <> "", CVDate(dlpRenewal.Text), Null), rsHRTrain("TR_CRSCODE"), rsHRTrain("TR_JOB"))
                    End If
                End If
                rsFollowUp.Close
                Set rsFollowUp = Nothing
            End If
            rsHRTrain("TR_RENEW") = IIf(Trim(dlpRenewal.Text) <> "", CVDate(dlpRenewal.Text), Null)
            rsHRTrain("TR_COURSE_TAKEN") = CVDate(xDateComplete)
            rsHRTrain("TR_LDATE") = Date
            rsHRTrain("TR_LUSER") = glbUserID
            rsHRTrain("TR_LTIME") = Time$
            rsHRTrain.Update
            
            'Update Follow Up record for Required Courses which has Job Code associated.
            'If (xJob <> "" And Not IsNull(xJob)) Or Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                'Update Follow Up record - Effective Date
                SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsFollowUp.EOF Then
                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                    'rsFollowUp("EF_COMMENTS") = "Course: " & rsCourseCode("ES_CRSCODE") & " - " & GetTABLDesc("ESCD", rsCourseCode("ES_CRSCODE")) & " for Position: " & rsEmpJobs("TW_JOB")
                    rsFollowUp("EF_LDATE") = Date
                    rsFollowUp("EF_LUSER") = glbUserID
                    rsFollowUp("EF_LTIME") = Time$
                    rsFollowUp.Update
                End If
                rsFollowUp.Close
                Set rsFollowUp = Nothing
            'End If
        End If
        rsHRTrain.Close
        Set rsHRTrain = Nothing
'    Else
'        'This could be an independent course or course with no renewal periods
'        'or course not found in Course Code master screen or not required by any job
'        'Search Training List for this course record and delete it
'        SQLQ = "SELECT * FROM HR_TRAIN"
'        SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
'        If xJob <> "" And Not IsNull(xJob) Then
'            SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
'        End If
'        SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourse & "'"
'        rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'        If Not rsHRTrain.EOF Then
'            'Delete the required course's Follow Up records which has Position code associated.
'            If xJob <> "" And Not IsNull(xJob) Then
'                'Update the Follow Up record as course completed
'                SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1"
'                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
'                gdbAdoIhr001.Execute SQLQ
'            End If
'            rsHRTrain.Delete
'        End If
'        rsHRTrain.Close
'        Set rsHRTrain = Nothing
'    End If
    
End Sub

Private Sub Delete_TrainRec_FollowUp(xCourse, xJob)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim SQLQ As String
    Dim xComments As String
    
    'Search Training List for this course record and delete it
    SQLQ = "SELECT * FROM HR_TRAIN"
    SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
    If xJob <> "" And Not IsNull(xJob) Then
        SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
    Else
        SQLQ = SQLQ & " AND (TR_JOB IS NULL OR TR_JOB = '')"
    End If
    SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourse & "'"
    rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRTrain.EOF Then
        If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
            xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
            SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
            SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
            SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & Replace(xComments, "'", "''") & "%' AND EF_FDATE = " & Date_SQL(rsHRTrain("TR_RENEW"))
            rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsFollowUp.EOF Then
                rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                rsHRTrain.Update
            End If
            rsFollowUp.Close
            Set rsFollowUp = Nothing
        End If
        
        'Update the required courses Follow Up records which has Position code associated
        'as Complete because Completed Date is entered but no Renewal Date.
        'Follow Up records have been created for independent courses as well
        'If xJob <> "" And Not IsNull(xJob) And Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
            'Get the Follow Up record to delete first
            'SQLQ = "DELETE FROM HR_FOLLOW_UP"
            'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
            'Mark instead as completed
            If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) And rsHRTrain("TR_FOLLOWUP_ID") <> "" Then
                SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(dlpDatComp.Text) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                gdbAdoIhr001.Execute SQLQ
            End If
        'End If
        
        'Delete the Training List record - no Renewal Date entered.
        rsHRTrain.Delete
    End If
    rsHRTrain.Close
    Set rsHRTrain = Nothing

End Sub

Private Sub Clear_Course_Taken_Date(xRenewalDt, xDateComp, xCourse, xJob)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim SQLQ As String
    
    'Search Training List for this course record and delete it
    SQLQ = "SELECT * FROM HR_TRAIN"
    SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
    If xJob <> "" And Not IsNull(xJob) Then
        SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
    Else
        SQLQ = SQLQ & " AND (TR_JOB IS NULL OR TR_JOB = '')"
    End If
    SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourse & "'"
    SQLQ = SQLQ & " AND TR_COURSE_TAKEN = " & Date_SQL(xDateComp)
    SQLQ = SQLQ & " AND TR_RENEW = " & Date_SQL(xRenewalDt)
    
    rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRTrain.EOF Then
        rsHRTrain("TR_COURSE_TAKEN") = Null
        rsHRTrain("TR_LDATE") = Date
        rsHRTrain("TR_LUSER") = glbUserID
        rsHRTrain("TR_LTIME") = Time$
        rsHRTrain.Update
    End If
    rsHRTrain.Close
    Set rsHRTrain = Nothing

End Sub

Private Sub Undo_Training_List_Rec_on_ContEdu_Delete(xJob, xCourse, Optional xRenewalDt, Optional xCompleteDt)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsCourseMst As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsEmpJobs As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim SQLQ As String
    Dim xEduRec, xCurRenPrd, xPrvRenPrd, xFlwRenPrd As Integer
    Dim xCurRenTyp, xPrvRenTyp, xFlwRenTyp, xDWMY As String
    Dim flgCrsTakenBefore, flgUnqForPos As Boolean
    Dim xComments As String
    Dim xOrgDate As Date
    

    'Initialise
    xEduRec = 0
    xFlwRenPrd = 0
    xFlwRenTyp = ""
    flgCrsTakenBefore = False
    
    'Course Record being deleted is the one WITH Course Renewal Date
    If Not IsMissing(xRenewalDt) Then
    
        'Check if the course is unique for each position
        SQLQ = "SELECT ES_CRSCODE,ES_UNIQUE_FOR_POS FROM HR_COURSECODE_MASTER"
        SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & xCourse & "'"
        rsCourseMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsCourseMst.EOF Then
            'Course found
            flgUnqForPos = IIf(IsNull(rsCourseMst("ES_UNIQUE_FOR_POS")), False, rsCourseMst("ES_UNIQUE_FOR_POS"))
        Else
            flgUnqForPos = False
        End If
        rsCourseMst.Close
        Set rsCourseMst = Nothing
    
        SQLQ = "SELECT * FROM HR_TRAIN"
        SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
        If xJob <> "" And Not IsNull(xJob) Then
            SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
        Else
            SQLQ = SQLQ & " AND (TR_JOB IS NULL OR TR_JOB = '')"
        End If
        SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourse & "'"
        rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsHRTrain.EOF Then
            'Check if Course was taken previously
            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE, ES_DATCOMP, ES_RENEW FROM HREDSEM"
            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
            SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourse & "'"
            If flgUnqForPos <> 0 Then
                'Unique for each position course then check if the course was taken for the right position
                If xJob <> "" And Not IsNull(xJob) Then
                    SQLQ = SQLQ & " AND ES_JOB = '" & xJob & "'"
                End If
            End If
            SQLQ = SQLQ & " AND ES_ID <> " & Data1.Recordset("ES_ID")
            SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsContEdu.EOF Then
                'Course Taken records found
                
                rsContEdu.MoveFirst
            
                'Update Training List with the last Course Taken Date and Renewal Date.
                If Not IsNull(rsContEdu("ES_RENEW")) Then
                    rsHRTrain("TR_RENEW") = CVDate(rsContEdu("ES_RENEW"))
                Else
                    rsHRTrain("TR_RENEW") = CVDate(xCompleteDt)
                End If
                If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                    rsHRTrain("TR_COURSE_TAKEN") = CVDate(rsContEdu("ES_DATCOMP"))
                    'rsHRTrain("TR_COURSE_TAKEN") = IIf(Not IsNull(rsContEdu("ES_DATCOMP")), CVDate(rsContEdu("ES_DATCOMP")), Null)
                Else
                    rsHRTrain("TR_COURSE_TAKEN") = ""
                End If
                rsHRTrain("TR_LDATE") = Date
                rsHRTrain("TR_LUSER") = glbUserID
                rsHRTrain("TR_LTIME") = Time$
                
                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & Replace(xComments, "'", "''") & "%' AND EF_FDATE = " & Date_SQL(rsContEdu("ES_RENEW"))
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                End If
                
                rsHRTrain.Update
                
                If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    'Update Follow Up record - Effective Date
                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                        'rsFollowUp("EF_COMMENTS") = "Course: " & rsCourseCode("ES_CRSCODE") & " - " & GetTABLDesc("ESCD", rsCourseCode("ES_CRSCODE")) & " for Position: " & rsEmpJobs("TW_JOB")
                        rsFollowUp("EF_LDATE") = Date
                        rsFollowUp("EF_LUSER") = glbUserID
                        rsFollowUp("EF_LTIME") = Time$
                        rsFollowUp.Update
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                End If
            Else
                'No Course Taken records found - (the one being deleted was the first record).
                'Get the Renewal Period and recompute the Renewal Date
                
                If xJob = "" Or IsNull(xJob) Then
                    'If Independant Course - Reset the Renewal Date to Follow Up Period + Position Start Date
                    'Retrieve Renewal Periods from Course Code Master because this is an independant course
                    SQLQ = "SELECT ES_CRSCODE,ES_RENEW_FOLLOWUP,ES_FLWUP_PRD_DWMY FROM HR_COURSECODE_MASTER"
                    SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & xCourse & "'"
                    rsCourseMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsCourseMst.EOF Then
                        'Course found
                        xFlwRenPrd = rsCourseMst("ES_RENEW_FOLLOWUP")
                        xFlwRenTyp = rsCourseMst("ES_FLWUP_PRD_DWMY")
                    End If
                    rsCourseMst.Close
                    Set rsCourseMst = Nothing
                Else
                    'Retrieve renewal period from Required Courses table
                    SQLQ = "SELECT PC_CRSCODE,PC_RENEW_FOLLOWUP,PC_FLWUP_PRD_DWMY FROM HR_JOB_COURSE "
                    SQLQ = SQLQ & " WHERE PC_JOB = '" & xJob & "'"
                    SQLQ = SQLQ & " AND PC_CRSCODE = '" & xCourse & "'"
                    rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsReqCourse.EOF Then
                        'Course found
                        xFlwRenPrd = rsReqCourse("PC_RENEW_FOLLOWUP")
                        xFlwRenTyp = rsReqCourse("PC_FLWUP_PRD_DWMY")
                    End If
                    rsReqCourse.Close
                    Set rsReqCourse = Nothing
                End If
                    
                'Compute Renewal Date
                'Course never taken before - Renewal Date = Follow Up Period + Position Start Date
                Select Case xFlwRenTyp
                    Case "D"
                        xDWMY = "d"
                    Case "W"
                        xDWMY = "ww"
                    Case "M"
                        xDWMY = "m"
                    Case "Y"
                        xDWMY = "yyyy"
                End Select
                
                If IsDate(rsHRTrain("TR_RENEW")) Then
                    xOrgDate = rsHRTrain("TR_RENEW")
                End If
                If Not IsNull(rsHRTrain("TR_SDATE")) Then
                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xFlwRenPrd, CVDate(rsHRTrain("TR_SDATE")))
                Else
                    rsHRTrain("TR_RENEW") = CVDate(xCompleteDt)
                End If
                
                rsHRTrain("TR_COURSE_TAKEN") = Null     'Course never taken before
                rsHRTrain("TR_LDATE") = Date
                rsHRTrain("TR_LUSER") = glbUserID
                rsHRTrain("TR_LTIME") = Time$
                
                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & Replace(xComments, "'", "''") & "%' "
                    If IsDate(xOrgDate) Then SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(xOrgDate)
                    SQLQ = SQLQ & " ORDER BY EF_FDATE DESC"
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsFollowUp.MoveFirst
                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                End If
                
                rsHRTrain.Update
                
                If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    'Update Follow Up record - Effective Date
                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                        'rsFollowUp("EF_COMMENTS") = "Course: " & rsCourseCode("ES_CRSCODE") & " - " & GetTABLDesc("ESCD", rsCourseCode("ES_CRSCODE")) & " for Position: " & rsEmpJobs("TW_JOB")
                        rsFollowUp("EF_LDATE") = Date
                        rsFollowUp("EF_LUSER") = glbUserID
                        rsFollowUp("EF_LTIME") = Time$
                        rsFollowUp.Update
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                End If
            End If
            rsContEdu.Close
            Set rsContEdu = Nothing
        End If
        rsHRTrain.Close
        Set rsHRTrain = Nothing
    Else
        'No Course Renewal Date means there is no corresponding Training List record and Follow Up record,
        'which means we may have to create one.
        'Check if course was taken before
            '- if not taken, then check which Current or Tracked Position require this course
                'if Position found - get the Follow Up Renewal period for that Position Course
                'if Position not found - do not add Training List record for this course.
            '- If taken, then check which Current or Tracked Position require this course
                'if Position found - based on the type of Position, Current or Tracked, get the
                    'Renewal Period and compute Renewasl Date based on Course Taken date
                'if Position not found - do not add Training List record for this course
                    'clear the Renewal Date from the last Course Taken record.
                
        'Check if the course is unique for each position
        SQLQ = "SELECT ES_CRSCODE,ES_UNIQUE_FOR_POS FROM HR_COURSECODE_MASTER"
        SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & xCourse & "'"
        rsCourseMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsCourseMst.EOF Then
            'Course found
            flgUnqForPos = IIf(IsNull(rsCourseMst("ES_UNIQUE_FOR_POS")), False, rsCourseMst("ES_UNIQUE_FOR_POS"))
        Else
            flgUnqForPos = False
        End If
        rsCourseMst.Close
        Set rsCourseMst = Nothing
        
        'Course Taken before?
        SQLQ = "SELECT ES_EMPNBR,ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
        SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourse & "'"
        If flgUnqForPos <> 0 Then
            'Unique for each position course then check if the course was taken for the right position
            If xJob <> "" And Not IsNull(xJob) Then
                SQLQ = SQLQ & " AND ES_JOB = '" & xJob & "'"
            End If
        End If
        SQLQ = SQLQ & " AND ES_ID <> " & Data1.Recordset("ES_ID")
        SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
        rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsContEdu.EOF Then
            'Course Taken before
            rsContEdu.MoveFirst
            flgCrsTakenBefore = True
        Else
            flgCrsTakenBefore = False
        End If
        
        
        'Check which Current or Tracked Position required this Course
        'Get list of Current/Temporary and Tracked Positions of this employee
        SQLQ = "SELECT JH_ID AS TW_ID, JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
        SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
        SQLQ = SQLQ & " AND JH_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & xCourse & "')"
        
        'if Unique for each Position record that means each Current/Tracked Job requiring this
        'course will have it's own Training List record - of course depending on the Renewal Period
        'Retrieve only the job assigned to the deleted Course Taken record.
        If flgUnqForPos <> 0 Then
            If xJob <> "" And Not IsNull(xJob) Then
                SQLQ = SQLQ & " AND JH_JOB = '" & xJob & "'"
            End If
        End If
        
        SQLQ = SQLQ & " UNION "
        SQLQ = SQLQ & " SELECT TW_ID, TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
        SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
        SQLQ = SQLQ & " AND TW_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & xCourse & "')"
        
        'if Unique for each Position record that means each Current/Tracked Job requiring this
        'course will have it's own Training List record - of course depending on the Renewal Period
        'Retrieve only the job assigned to the deleted Course Taken record.
        If flgUnqForPos <> 0 Then
            If xJob <> "" And Not IsNull(xJob) Then
                SQLQ = SQLQ & " AND TW_JOB = '" & xJob & "'"
            End If
        End If
        
        SQLQ = SQLQ & " ORDER BY TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
        rsEmpJobs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmpJobs.EOF Then
            rsEmpJobs.MoveFirst
        
            Do While Not rsEmpJobs.EOF
                'Get the renewal periods of the course
                SQLQ = "SELECT PC_CRSCODE,PC_RENEW_CRS_CUR,PC_CUR_PRD_DWMY,PC_RENEW_CRS_PRV,PC_PRV_PRD_DWMY,PC_RENEW_FOLLOWUP,PC_FLWUP_PRD_DWMY FROM HR_JOB_COURSE "
                SQLQ = SQLQ & " WHERE PC_JOB = '" & rsEmpJobs("TW_JOB") & "'"
                SQLQ = SQLQ & " AND PC_CRSCODE = '" & xCourse & "'"
                rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsReqCourse.EOF Then
                    'Course found
                    xCurRenPrd = rsReqCourse("PC_RENEW_CRS_CUR")
                    xCurRenTyp = rsReqCourse("PC_CUR_PRD_DWMY")
                    xPrvRenPrd = rsReqCourse("PC_RENEW_CRS_PRV")
                    xPrvRenTyp = rsReqCourse("PC_PRV_PRD_DWMY")
                    xFlwRenPrd = IIf(IsNull(rsReqCourse("PC_RENEW_FOLLOWUP")), 99, rsReqCourse("PC_RENEW_FOLLOWUP"))
                    xFlwRenTyp = IIf(IsNull(rsReqCourse("PC_FLWUP_PRD_DWMY")), "Y", rsReqCourse("PC_FLWUP_PRD_DWMY"))
                End If
                rsReqCourse.Close
                Set rsReqCourse = Nothing
                
                'if Unique for each Position Course check if the Training List existing for this Job
                'already exists - then skip to next Employee Position requiring this course
                If flgUnqForPos <> 0 And xJob <> "" And Not IsNull(xJob) Then
                    SQLQ = "SELECT * FROM HR_TRAIN"
                    SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                    SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourse & "'"
                    rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsHRTrain.EOF Then
                        'Skip to next Employee Job because for this Job, the training list record already
                        'exist for this unique for each position course.
                        GoTo next_EmpPosition
                    Else
                        'Continue with the rest of the process
                    End If
                    rsHRTrain.Close
                    Set rsHRTrain = Nothing
                End If
                
                SQLQ = "SELECT * FROM HR_TRAIN WHERE 1=2"
                rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                
                'Course Taken before?
                If flgCrsTakenBefore = True Then
                    'Course Taken before
                    'Compute the Renewal Period for this course and add a Training List and Follow Up record
                    If (rsEmpJobs("POS_TYPE") = "CURRENT" Or rsEmpJobs("POS_TYPE") = "TEMPORARY") And rsEmpJobs("TW_CURRENT") Then
                        'Primary Current/Temporary Position
                        'Based on Current Renewal Period if found
                        If IsNull(xCurRenPrd) Or xCurRenPrd = 0 Or xCurRenPrd = "" Then
                            'No Renewal Period found, clear last course taken record's Renewal Date
                            'There won't be Training List record, because there was no Renewal Date on the
                            'deleted Course Taken record.
                            rsContEdu("ES_RENEW") = Null
                            rsContEdu("ES_LDATE") = Date
                            rsContEdu("ES_LUSER") = glbUserID
                            rsContEdu("ES_LTIME") = Time$
                            rsContEdu.Update
                        
                            'If flgUnqForPos Then
                            '    'Go to next position
                            '    GoTo next_EmpPosition
                            'Else
                                'Exit loop - only the first position gets this course
                                Exit Do
                           ' End If
                        Else
                            'Compute renewal date
                            Select Case xCurRenTyp
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            'Add a new Training List record with Renewal Date based on Current Renewal Period and
                            'Course Taken Date
                            rsHRTrain.AddNew
                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xCurRenPrd, CVDate(rsContEdu("ES_DATCOMP")))
                            rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                            
                            'Update Continuing Education record as well
                            rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                            rsContEdu("ES_JOB") = rsEmpJobs("TW_JOB")
                            rsContEdu("ES_LDATE") = Date
                            rsContEdu("ES_LUSER") = glbUserID
                            rsContEdu("ES_LTIME") = Time$
                            rsContEdu.Update
                        End If
                    Else
                        'Previous position
                        'Based on Previous Renewal period if found
                        If IsNull(xPrvRenPrd) Or xPrvRenPrd = 0 Or xPrvRenPrd = "" Then
                            'No Renewal Period found, clear last course taken record's Renewal Date
                            'There won't be Training List record, because there was no Renewal Date on the
                            'deleted Course Taken record.
                            rsContEdu("ES_RENEW") = Null
                            rsContEdu("ES_LDATE") = Date
                            rsContEdu("ES_LUSER") = glbUserID
                            rsContEdu("ES_LTIME") = Time$
                            rsContEdu.Update
                            
                            Exit Do
                        Else
                            'Compute renewal date
                            Select Case xPrvRenTyp
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            'Add a new Training List record with Renewal Date based on Prev Renewal Period
                            'Course Taken Date
                            rsHRTrain.AddNew
                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xPrvRenPrd, CVDate(rsContEdu("ES_DATCOMP")))
                            rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                            
                            'Update Continuing Education record as well
                            rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                            rsContEdu("ES_JOB") = rsEmpJobs("TW_JOB")
                            rsContEdu("ES_LDATE") = Date
                            rsContEdu("ES_LUSER") = glbUserID
                            rsContEdu("ES_LTIME") = Time$
                            rsContEdu.Update
                        End If
                    End If
                Else
                    'Course not taken before
                    'Compute renewal date based on Follow Up Period
                    Select Case xFlwRenTyp
                        Case "D"
                            xDWMY = "d"
                        Case "W"
                            xDWMY = "ww"
                        Case "M"
                            xDWMY = "m"
                        Case "Y"
                            xDWMY = "yyyy"
                    End Select
                    'Add a new Training List record with Renewal Date based on Follow Up Period
                    rsHRTrain.AddNew
                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xFlwRenPrd, CVDate(rsEmpJobs("TW_SDATE")))
                End If
                
                rsHRTrain("TR_COMPNO") = "001"
                rsHRTrain("TR_EMPNBR") = glbLEE_ID
                rsHRTrain("TR_CRSCODE") = xCourse
                
                rsHRTrain("TR_JOB") = rsEmpJobs("TW_JOB")
                rsHRTrain("TR_SDATE") = rsEmpJobs("TW_SDATE")
                If (rsEmpJobs("POS_TYPE") = "CURRENT") And rsEmpJobs("TW_CURRENT") Then
                    rsHRTrain("TR_POS_TYPE") = "C"
                ElseIf (rsEmpJobs("POS_TYPE") = "TEMPORARY") And rsEmpJobs("TW_CURRENT") Then
                    rsHRTrain("TR_POS_TYPE") = "T"
                Else
                    rsHRTrain("TR_POS_TYPE") = "P"
                End If
                rsHRTrain("TR_LDATE") = Date
                rsHRTrain("TR_LTIME") = Time$
                rsHRTrain("TR_LUSER") = glbUserID
                
                'Add a Follow Up record for this Training course
                SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE 1 = 2"
                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                rsFollowUp.AddNew
                rsFollowUp("EF_COMPNO") = "001"
                rsFollowUp("EF_EMPNBR") = glbLEE_ID
                rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                rsFollowUp("EF_FREAS_TABL") = "FURE"
                'Ticket #24257 - Do not update Admin By for them only
                If glbCompSerial <> "S/N - 2262W" Then
                    rsFollowUp("EF_ADMINBY_TABL") = "EDAB"
                    rsFollowUp("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
                End If
                rsFollowUp("EF_FREAS") = "EDUC"
                rsFollowUp("EF_COMMENTS") = "Course: " & xCourse & " - " & GetTABLDesc("ESCD", xCourse) & " for Position: " & rsEmpJobs("TW_JOB")
                rsFollowUp("EF_LDATE") = Date
                rsFollowUp("EF_LTIME") = Time$
                rsFollowUp("EF_LUSER") = glbUserID
                rsFollowUp.Update
                
                rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                rsHRTrain.Update
                
                rsFollowUp.Close
                Set rsFollowUp = Nothing
            
                'Update Position record with Follow Up ID
                'if the course code is TRAIN
                If xCourse = "TRAIN" Then
                    'Search HR_JOB_HISTORY or HR_TEMP_WORK table for this Position record
                    'and update with Follow Up Id
                    If (rsEmpJobs("POS_TYPE") = "CURRENT") Then
                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & rsEmpJobs("TW_ID")
                    Else
                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & rsEmpJobs("TW_ID")
                    End If
                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsTJob.EOF Then
                        If (rsEmpJobs("POS_TYPE") = "CURRENT") Then
                            rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                        Else
                            rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                        End If
                        rsTJob.Update
                    End If
                    rsTJob.Close
                    Set rsTJob = Nothing
                End If
                
                rsHRTrain.Close
                Set rsHRTrain = Nothing
                'If flgUnqForPos Then
                '    'Go to next position
                '    GoTo next_EmpPosition
                'Else
                    'Exit loop - only the first position gets this course
                    Exit Do
                'End If
                
next_EmpPosition:
                rsEmpJobs.MoveNext
            Loop
        Else
            'No Current/Temporary or Tracked Positions require this course
            
            'Course Taken before?
            If flgCrsTakenBefore = True Then
                'Course Taken before
                'Clear renewal date if found in the last course taken record
                'There won't be Training List record because the Course Taken record been deleted does
                'not have Renewal Date.
                rsContEdu("ES_RENEW") = Null
                rsContEdu("ES_LDATE") = Date
                rsContEdu("ES_LUSER") = glbUserID
                rsContEdu("ES_LTIME") = Time$
                rsContEdu.Update
            Else
                'Do not do anything, just let the Cont Education record delete.
                'No Current or Tracked Position of this employee require this course.
            End If
            
        End If
        rsEmpJobs.Close
        Set rsEmpJobs = Nothing
        
        rsContEdu.Close
        Set rsContEdu = Nothing
    End If

End Sub

Private Function EERetrieve_Null_on_Top()
    Dim SQLQ As String
    EERetrieve_Null_on_Top = False
    
    '''On Error GoTo EERError_Null
    
    Screen.MousePointer = HOURGLASS
    
    If glbtermopen Then         'Lucy July 5, 2000
        SQLQ = "Select * from Term_HREDSEM"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " ORDER BY ES_CTYPE ASC, ES_DATCOMP DESC, ES_EMPNBR"
    Else
        SQLQ = "Select * from HREDSEM"
        SQLQ = SQLQ & " where ES_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " ORDER BY ES_ID DESC, ES_CTYPE ASC, ES_DATCOMP DESC, ES_EMPNBR"
    End If
        
    Data1.RecordSource = SQLQ
    Data1.Refresh
    
    Set rsGrid = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True
    
    Call UpConttotal

    EERetrieve_Null_on_Top = True
    
    Screen.MousePointer = DEFAULT
    
    Exit Function
    
EERError_Null:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Continuing Education Retrieve", "HREDSEM", "SELECT")
    Call RollBack '23July99 js

Exit Function

End Function

Private Function CourseCodeMaster_Blank() As Boolean
    Dim rsCourseCode As New ADODB.Recordset
    Dim SQLQ As String
    
    SQLQ = "SELECT * FROM HR_COURSECODE_MASTER"
    rsCourseCode.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsCourseCode.EOF Then
        CourseCodeMaster_Blank = False
    Else
        CourseCodeMaster_Blank = True
    End If
    rsCourseCode.Close
    Set rsCourseCode = Nothing
        
End Function

