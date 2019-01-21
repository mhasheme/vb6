VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmESuccession 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Succession Planning"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11730
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
   Icon            =   "feSuccession.frx":0000
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10950
   ScaleWidth      =   11730
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   9015
      LargeChange     =   300
      Left            =   10920
      Max             =   4000
      SmallChange     =   300
      TabIndex        =   29
      Top             =   2160
      Width           =   300
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11730
      _Version        =   65536
      _ExtentX        =   20690
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
         Left            =   6960
         TabIndex        =   27
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblEENumber 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
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
         TabIndex        =   21
         Top             =   135
         Width           =   720
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8280
      Top             =   11520
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
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   26
      Top             =   10290
      Width           =   11730
      _Version        =   65536
      _ExtentX        =   20690
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
      Begin VB.CommandButton cmdContEdu 
         Appearance      =   0  'Flat
         Caption         =   "&Continuing Education"
         Height          =   375
         Left            =   1680
         TabIndex        =   28
         Tag             =   "Load Continuing Education screen"
         Top             =   120
         Width           =   2205
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6495
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
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EU_LDATE"
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
      Left            =   3120
      MaxLength       =   25
      TabIndex        =   17
      Top             =   11700
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EU_LTIME"
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
      Left            =   4800
      MaxLength       =   25
      TabIndex        =   18
      Top             =   11700
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EU_LUSER"
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
      Left            =   6480
      MaxLength       =   25
      TabIndex        =   19
      Top             =   11700
      Visible         =   0   'False
      Width           =   1590
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feSuccession.frx":030A
      Height          =   1545
      Left            =   0
      OleObjectBlob   =   "feSuccession.frx":031E
      TabIndex        =   30
      Top             =   480
      Width           =   10815
   End
   Begin VB.Frame panDetails 
      BorderStyle     =   0  'None
      Height          =   10815
      Left            =   0
      TabIndex        =   31
      Top             =   2040
      Width           =   10815
      Begin VB.Frame Frame1 
         Height          =   4605
         Left            =   120
         TabIndex        =   32
         Top             =   0
         Width           =   10695
         Begin VB.TextBox txtDegree 
            Appearance      =   0  'Flat
            DataField       =   "EU_DEGREE3"
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
            Index           =   2
            Left            =   9960
            MaxLength       =   25
            TabIndex        =   41
            Top             =   3840
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtMajor 
            Appearance      =   0  'Flat
            DataField       =   "EU_MAJOR3"
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
            Index           =   2
            Left            =   9960
            MaxLength       =   25
            TabIndex        =   40
            Top             =   3600
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtDegree 
            Appearance      =   0  'Flat
            DataField       =   "EU_DEGREE2"
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
            Index           =   1
            Left            =   6960
            MaxLength       =   25
            TabIndex        =   39
            Top             =   3840
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtMajor 
            Appearance      =   0  'Flat
            DataField       =   "EU_MAJOR2"
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
            Index           =   1
            Left            =   6960
            MaxLength       =   25
            TabIndex        =   38
            Top             =   3600
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtDegree 
            Appearance      =   0  'Flat
            DataField       =   "EU_DEGREE1"
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
            Left            =   3960
            MaxLength       =   25
            TabIndex        =   37
            Top             =   3840
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtMajor 
            Appearance      =   0  'Flat
            DataField       =   "EU_MAJOR1"
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
            Left            =   3960
            MaxLength       =   25
            TabIndex        =   36
            Top             =   3600
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtDept 
            Appearance      =   0  'Flat
            DataField       =   "EU_DEPTNO"
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
            Left            =   3960
            MaxLength       =   25
            TabIndex        =   35
            Top             =   3240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtLoc 
            Appearance      =   0  'Flat
            DataField       =   "EU_LOC"
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
            Left            =   5640
            MaxLength       =   25
            TabIndex        =   34
            Top             =   3240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdPhoto 
            Appearance      =   0  'Flat
            Caption         =   "&Photo Off"
            Height          =   330
            Left            =   8880
            TabIndex        =   33
            Tag             =   "Print the reports marked with an 'x'"
            Top             =   2920
            Width           =   1380
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EU_LANG1_SPOKEN"
            DataSource      =   " "
            Height          =   285
            Index           =   0
            Left            =   3240
            TabIndex        =   42
            Tag             =   "00-Language - Code"
            Top             =   5220
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDL1"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EU_LANG1_WRITTEN"
            DataSource      =   " "
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   43
            Tag             =   "00-Language - Code"
            Top             =   4965
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDL1"
         End
         Begin INFOHR_Controls.CodeLookup clpPosCode 
            DataField       =   "EU_JOB"
            Height          =   285
            Index           =   0
            Left            =   3240
            TabIndex        =   44
            Tag             =   "01-Position code"
            Top             =   4965
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   6
            LookupType      =   5
            Enabled         =   0   'False
         End
         Begin INFOHR_Controls.DateLookup dlpReviewDate 
            DataField       =   "EU_PERF1_DATE"
            Height          =   285
            Index           =   0
            Left            =   7440
            TabIndex        =   45
            Tag             =   "41-Performance Review Date"
            Top             =   5220
            Visible         =   0   'False
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   503
            TextBoxWidth    =   1215
            Enabled         =   0   'False
         End
         Begin INFOHR_Controls.DateLookup dlpReviewDate 
            DataField       =   "EU_PERF2_DATE"
            Height          =   285
            Index           =   1
            Left            =   7440
            TabIndex        =   46
            Tag             =   "41-Performance Review Date"
            Top             =   4920
            Visible         =   0   'False
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   503
            TextBoxWidth    =   1215
            Enabled         =   0   'False
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EU_PERF2_RATE"
            Height          =   285
            Index           =   7
            Left            =   585
            TabIndex        =   47
            Tag             =   "00-Performance Rating - Code "
            Top             =   4965
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "SDPC"
            Enabled         =   0   'False
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EU_LANG2_WRITTEN"
            DataSource      =   " "
            Height          =   285
            Index           =   3
            Left            =   4680
            TabIndex        =   48
            Tag             =   "00-Language - Code"
            Top             =   4965
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDL1"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EU_LANG3_SPOKEN"
            DataSource      =   " "
            Height          =   285
            Index           =   4
            Left            =   6120
            TabIndex        =   49
            Tag             =   "00-Language - Code"
            Top             =   4965
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDL1"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EU_LANG3_WRITTEN"
            DataSource      =   " "
            Height          =   285
            Index           =   5
            Left            =   6120
            TabIndex        =   50
            Tag             =   "00-Language - Code"
            Top             =   5220
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDL1"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EU_PERF1_RATE"
            Height          =   285
            Index           =   6
            Left            =   1800
            TabIndex        =   51
            Tag             =   "00-Performance Rating - Code "
            Top             =   5220
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "SDPC"
            Enabled         =   0   'False
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EU_LANG2_SPOKEN"
            DataSource      =   " "
            Height          =   285
            Index           =   2
            Left            =   4680
            TabIndex        =   52
            Tag             =   "00-Language - Code"
            Top             =   5220
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDL1"
         End
         Begin VB.Label lblLang1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Complete"
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
            TabIndex        =   95
            Top             =   4230
            Width           =   1050
         End
         Begin VB.Label lblLang1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Degree"
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
            TabIndex        =   94
            Top             =   3930
            Width           =   525
         End
         Begin VB.Label lblLang1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Major Study"
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
            TabIndex        =   93
            Top             =   3630
            Width           =   840
         End
         Begin VB.Label PicNotF 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Photo not Available"
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   8040
            TabIndex        =   92
            Top             =   1200
            Width           =   2115
            WordWrap        =   -1  'True
         End
         Begin VB.Image picPhoto 
            BorderStyle     =   1  'Fixed Single
            Height          =   2535
            Left            =   7680
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2625
         End
         Begin VB.Label lblLang1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Position Start Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   91
            Top             =   570
            Width           =   1620
         End
         Begin VB.Label lblCode 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   21
            Left            =   1440
            TabIndex        =   90
            Top             =   3000
            Width           =   2940
         End
         Begin VB.Label lblCode 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EU_DOH"
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
            Index           =   20
            Left            =   1440
            TabIndex        =   89
            Top             =   2700
            Width           =   1095
         End
         Begin VB.Label lblCode 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   18
            Left            =   5640
            TabIndex        =   88
            Top             =   3000
            Width           =   2940
         End
         Begin VB.Label lblMajor 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   2
            Left            =   7440
            TabIndex        =   87
            Top             =   3600
            Width           =   2940
         End
         Begin VB.Label lblDegree 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   2
            Left            =   7440
            TabIndex        =   86
            Top             =   3900
            Width           =   2940
         End
         Begin VB.Label lblYear 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EU_YEAR3"
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
            Index           =   2
            Left            =   7440
            TabIndex        =   85
            Top             =   4200
            Width           =   1095
         End
         Begin VB.Label lblMajor 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   0
            Left            =   1440
            TabIndex        =   84
            Top             =   3600
            Width           =   2940
         End
         Begin VB.Label lblDegree 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   0
            Left            =   1440
            TabIndex        =   83
            Top             =   3900
            Width           =   2940
         End
         Begin VB.Label lblMajor 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   1
            Left            =   4440
            TabIndex        =   82
            Top             =   3600
            Width           =   2940
         End
         Begin VB.Label lblDegree 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   1
            Left            =   4440
            TabIndex        =   81
            Top             =   3900
            Width           =   2940
         End
         Begin VB.Label lblYear 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EU_YEAR1"
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
            Index           =   0
            Left            =   1440
            TabIndex        =   80
            Top             =   4200
            Width           =   1095
         End
         Begin VB.Label lblYear 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EU_YEAR2"
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
            Index           =   1
            Left            =   4440
            TabIndex        =   79
            Top             =   4200
            Width           =   1095
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Location"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   21
            Left            =   4680
            TabIndex        =   78
            Top             =   3030
            Width           =   750
         End
         Begin VB.Label lblLang1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Formal Education / Certificates"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   77
            Top             =   3360
            Width           =   2655
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Department"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   76
            Top             =   3030
            Width           =   990
         End
         Begin VB.Label lblLang1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   75
            Top             =   2730
            Width           =   885
         End
         Begin VB.Label lblCode 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   9
            Left            =   4800
            TabIndex        =   74
            Top             =   1170
            Width           =   1095
         End
         Begin VB.Label lblCode 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   8
            Left            =   1440
            TabIndex        =   73
            Top             =   1170
            Width           =   1095
         End
         Begin VB.Label lblCode 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   7
            Left            =   4800
            TabIndex        =   72
            Top             =   1450
            Width           =   1935
         End
         Begin VB.Label lblCode 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   6
            Left            =   1440
            TabIndex        =   71
            Top             =   1450
            Width           =   1935
         End
         Begin VB.Label lblCode 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   5
            Left            =   5640
            TabIndex        =   70
            Top             =   2350
            Width           =   1935
         End
         Begin VB.Label lblCode 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   4
            Left            =   5640
            TabIndex        =   69
            Top             =   2070
            Width           =   1935
         End
         Begin VB.Label lblCode 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   3
            Left            =   3540
            TabIndex        =   68
            Top             =   2350
            Width           =   1935
         End
         Begin VB.Label lblCode 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   2
            Left            =   3540
            TabIndex        =   67
            Top             =   2070
            Width           =   1935
         End
         Begin VB.Label lblCode 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   1
            Left            =   1440
            TabIndex        =   66
            Top             =   2350
            Width           =   1935
         End
         Begin VB.Label lblCode 
            BorderStyle     =   1  'Fixed Single
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
            Index           =   0
            Left            =   1440
            TabIndex        =   65
            Top             =   2070
            Width           =   1935
         End
         Begin VB.Label lblPosCode 
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   64
            Top             =   240
            Width           =   5775
         End
         Begin VB.Label lblLang1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Spoken"
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
            Left            =   240
            TabIndex        =   63
            Top             =   2100
            Width           =   555
         End
         Begin VB.Label lblLang1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Written"
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
            TabIndex        =   62
            Top             =   2380
            Width           =   510
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Rating"
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
            TabIndex        =   61
            Top             =   1480
            Width           =   615
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Review Date"
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
            TabIndex        =   60
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   690
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Current Review"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   58
            Top             =   900
            Width           =   1695
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Rating"
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
            Left            =   3720
            TabIndex        =   57
            Top             =   1480
            Width           =   735
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Review Date"
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
            Left            =   3720
            TabIndex        =   56
            Top             =   1200
            Width           =   1110
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Last Review"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   3540
            TabIndex        =   55
            Top             =   900
            Width           =   1695
         End
         Begin VB.Label lblLang1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Languages"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   54
            Top             =   1800
            Width           =   945
         End
         Begin VB.Label lblCode 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EU_SDATE"
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
            Index           =   22
            Left            =   1800
            TabIndex        =   53
            Top             =   540
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   120
         TabIndex        =   96
         Top             =   4560
         Width           =   10695
         Begin VB.TextBox memComments 
            Appearance      =   0  'Flat
            DataField       =   "EU_COMMENTS"
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
            Height          =   855
            Index           =   3
            Left            =   2085
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Tag             =   "00-Comments"
            Top             =   2640
            Width           =   6885
         End
         Begin VB.TextBox memComments 
            Appearance      =   0  'Flat
            DataField       =   "EU_CSR_DEVL"
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
            Height          =   495
            Index           =   2
            Left            =   2085
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Tag             =   "00-Comments"
            Top             =   2070
            Width           =   6885
         End
         Begin VB.TextBox memComments 
            Appearance      =   0  'Flat
            DataField       =   "EU_CSR_WEAK"
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
            Height          =   495
            Index           =   1
            Left            =   2085
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Tag             =   "00-Comments"
            Top             =   1500
            Width           =   6885
         End
         Begin VB.TextBox memComments 
            Appearance      =   0  'Flat
            DataField       =   "EU_CSR_STRE"
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
            Height          =   495
            Index           =   0
            Left            =   2085
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Tag             =   "00-Comments"
            Top             =   930
            Width           =   6885
         End
         Begin VB.CheckBox chkRelocate 
            DataField       =   "EU_RELOC"
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
            Left            =   2085
            TabIndex        =   7
            Top             =   3675
            Width           =   495
         End
         Begin VB.TextBox txtLocation 
            Appearance      =   0  'Flat
            DataField       =   "EU_Location"
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
            Left            =   3780
            MaxLength       =   25
            TabIndex        =   8
            Top             =   3675
            Width           =   4935
         End
         Begin VB.TextBox txtReviewer 
            Appearance      =   0  'Flat
            DataField       =   "EU_REVIEWER"
            Height          =   285
            Left            =   2490
            TabIndex        =   97
            Tag             =   "00-Employee Number of individual's reviewer"
            Top             =   5280
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkOutsideRVW 
            DataField       =   "EU_OUTREVIEW"
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
            Left            =   2085
            TabIndex        =   15
            Top             =   5685
            Width           =   180
         End
         Begin VB.CheckBox chkLastReview 
            Caption         =   "Last Succession Plan"
            DataField       =   "EU_LAST_RVW"
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
            Left            =   7080
            TabIndex        =   2
            Top             =   565
            Width           =   2175
         End
         Begin INFOHR_Controls.DateLookup dlpReviewDate 
            DataField       =   "EU_CSR_DATE"
            Height          =   285
            Index           =   2
            Left            =   3705
            TabIndex        =   1
            Tag             =   "41-Performance Review Date"
            Top             =   570
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   503
            TextBoxWidth    =   1215
            Enabled         =   0   'False
         End
         Begin INFOHR_Controls.CodeLookup clpPosCode 
            DataField       =   "EU_JOBPREF1"
            Height          =   285
            Index           =   1
            Left            =   1770
            TabIndex        =   9
            Tag             =   "01-Position code"
            Top             =   4080
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   6
            LookupType      =   5
            Enabled         =   0   'False
         End
         Begin INFOHR_Controls.CodeLookup clpPosCode 
            DataField       =   "EU_JOBPREF2"
            Height          =   285
            Index           =   2
            Left            =   6120
            TabIndex        =   10
            Tag             =   "01-Position code"
            Top             =   4080
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   6
            LookupType      =   5
            Enabled         =   0   'False
         End
         Begin INFOHR_Controls.CodeLookup clpPosCode 
            DataField       =   "EU_JOBPREF3"
            Height          =   285
            Index           =   3
            Left            =   1770
            TabIndex        =   11
            Tag             =   "01-Position code"
            Top             =   4417
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   6
            LookupType      =   5
            Enabled         =   0   'False
         End
         Begin INFOHR_Controls.CodeLookup clpPosCode 
            DataField       =   "EU_JOBPREF4"
            Height          =   285
            Index           =   4
            Left            =   6120
            TabIndex        =   12
            Tag             =   "01-Position code"
            Top             =   4410
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   6
            LookupType      =   5
            Enabled         =   0   'False
         End
         Begin INFOHR_Controls.CodeLookup clpPosCode 
            DataField       =   "EU_JOBPREF5"
            Height          =   285
            Index           =   5
            Left            =   1770
            TabIndex        =   13
            Tag             =   "01-Position code"
            Top             =   4755
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   6
            LookupType      =   5
            Enabled         =   0   'False
         End
         Begin INFOHR_Controls.EmployeeLookup elpReviewer 
            Height          =   285
            Left            =   1770
            TabIndex        =   14
            Tag             =   "10-Employee Number of individual's reviewer"
            Top             =   5280
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   503
            ShowUnassigned  =   1
            RefreshDescriptionWhen=   2
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EU_PROMOT"
            DataSource      =   " "
            Height          =   285
            Index           =   8
            Left            =   1770
            TabIndex        =   0
            Tag             =   "00-Language - Code"
            Top             =   240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "OPRO"
         End
         Begin INFOHR_Controls.CodeLookup clpOutsideRVWR 
            DataField       =   "EU_OUTREVIEWER"
            DataSource      =   " "
            Height          =   285
            Left            =   4440
            TabIndex        =   16
            Tag             =   "00-Outside Reviewer - Code"
            Top             =   5640
            Visible         =   0   'False
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "OTRW"
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Summary Comments"
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
            Index           =   9
            Left            =   480
            TabIndex        =   110
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label lblOutsideRVWR 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Outside Reviewer"
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
            Left            =   3000
            TabIndex        =   109
            Top             =   5685
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.Label lblPromot 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Promotability"
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
            Left            =   120
            TabIndex        =   108
            Top             =   285
            Width           =   885
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Current Succession Plan"
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
            Left            =   120
            TabIndex        =   107
            Top             =   610
            Width           =   1815
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Development Plan"
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
            Index           =   13
            Left            =   480
            TabIndex        =   106
            Top             =   2070
            Width           =   1335
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Weaknesses"
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
            Left            =   480
            TabIndex        =   105
            Top             =   1500
            Width           =   1335
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Strengths"
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
            Index           =   15
            Left            =   480
            TabIndex        =   104
            Top             =   930
            Width           =   1335
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   16
            Left            =   3120
            TabIndex        =   103
            Top             =   615
            Width           =   495
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Relocate ?"
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
            Left            =   120
            TabIndex        =   102
            Top             =   3720
            Width           =   1575
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Index           =   18
            Left            =   3000
            TabIndex        =   101
            Top             =   3720
            Width           =   1575
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Position Preferences"
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
            Left            =   120
            TabIndex        =   100
            Top             =   4125
            Width           =   1650
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Reviewer"
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
            Left            =   120
            TabIndex        =   99
            Top             =   5325
            Width           =   1650
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Outside Review"
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
            Left            =   120
            TabIndex        =   98
            Top             =   5685
            Width           =   1650
         End
      End
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EU_EMPNBR"
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
      Left            =   2400
      TabIndex        =   24
      Top             =   11820
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EU_COMPNO"
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   570
      TabIndex        =   25
      Top             =   11820
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmESuccession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Dim fUPMode As Integer
Dim fglbNewSalRec%
Dim glbNew
Dim fglbNew
Dim RSDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim Ctrl As Control 'Sam add July 2002 * Remove ADO
Dim glbPicDir, glbPicBMP

Private Function chkSuccession()
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, Msg As String
Dim RsSuccession As New ADODB.Recordset

chkSuccession = False

On Error GoTo chkSuccession_Err

'??? Which one is required field?
If Len(dlpReviewDate(2)) = 0 Then
    MsgBox "Date is a required field."
    dlpReviewDate(2).SetFocus
    Exit Function
End If

If Not glbtermopen Then
    SQLQ = "SELECT * FROM HR_SUCCESSION WHERE EU_EMPNBR = " & glbLEE_ID
Else
    SQLQ = "SELECT * FROM Term_HR_SUCCESSION WHERE TERM_SEQ = " & glbTERM_Seq
End If
SQLQ = SQLQ & " AND EU_CSR_DATE = " & Date_SQL(dlpReviewDate(2)) & " "
If Not fglbNew Then
    SQLQ = SQLQ & " AND EU_ID <> " & Data1.Recordset("EU_ID") & " "
End If
'If Not fglbNew Then sqlq = sqlq & " AND EL_LANGNO = " & Val(txtLangNum)

SQLQ = SQLQ & " order by EU_ID desc"

RsSuccession.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not RsSuccession.EOF Then
        MsgBox "Duplicate record existed - not entered"
        dlpReviewDate(2).SetFocus
        RsSuccession.Close
        Exit Function
End If
'If Not RsSuccession.EOF Then
'    If Not IsNull(RsSuccession("EL_LANGNO")) Then
'        If fglbNew Then
'        '    txtLangNum = RsEHST("DE_DOCNO")
'        'Else
'            txtLangNum = RsSuccession("EL_LANGNO") + 1
'        End If
'    End If
'Else
'    txtLangNum = 1
'End If
RsSuccession.Close

If elpReviewer = "0" Then elpReviewer = ""
If Len(elpReviewer) > 0 Then
    If elpReviewer.Caption = "Unassigned" Then
        MsgBox "Employee # not valid. Check # and re-enter!"
        elpReviewer.SetFocus
        Exit Function
    End If
End If

chkSuccession = True

Exit Function

chkSuccession_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkSuccession", "HR_SUCCESSION", "edit/Add")
Resume Next

End Function

Sub cmdCancel_Click()
Dim x

On Error GoTo Can_Err
'data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'data1.Refresh
fglbNew = False
''' Sam add July 2002 * Remove ADO

RSDATA.CancelUpdate

Call Display_Value

Call txtReviewer_Change

'Call ST_UPD_MODE(True)  ' reset screen's attributes
'Call SET_UP_MODE
Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_SUCCESSION", "Cancel")
Call RollBack '23July99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMESuccession" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, x

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub


If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    RSDATA.Delete
    gdbAdoIhr001X.CommitTrans
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    RSDATA.Delete
    gdbAdoIhr001.CommitTrans
    Data1.Refresh
End If
If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_SUCCESSION", "Delete")
Call RollBack '23July99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
    Dim SQLQ As String
    Dim rsTemp As New ADODB.Recordset
    Dim intI, x, xfld
    intI = 0
    fglbNew = True
    
    'Call ST_UPD_MODE(True)
    
    On Error GoTo AddN_Err

    Call Set_Control("B", Me)
    
    'Retrieve Current Job Title
    SQLQ = "select JH_JOB,JH_SDATE from HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE JH_EMPNBR=" & glbLEE_ID & " AND JH_CURRENT <> 0"
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If rsTemp.EOF Then
        clpPosCode(0).Text = ""
        lblCode(22).Caption = ""        'Ticket #23222
    Else
        clpPosCode(0).Text = rsTemp("JH_JOB")
        lblCode(22).Caption = rsTemp("JH_SDATE")    'Ticket #23222
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    
    'Ticket #23222 - Date of Hire, Dept, Location
    lblCode(20).Caption = GetEmpData(glbLEE_ID, "ED_DOH")
    txtDept.Text = GetEmpData(glbLEE_ID, "ED_DEPTNO")
    lblCode(21).Caption = GetDeptName(txtDept.Text, "DF_NAME")
    txtLoc.Text = GetEmpData(glbLEE_ID, "ED_LOC")
    lblCode(18).Caption = GetTABLDesc("EDLC", txtLoc.Text)
    
    If glbOracle Then
        SQLQ = "SELECT "
    Else
        SQLQ = "SELECT top 2 "
    End If
    If glbtermopen Then
        SQLQ = SQLQ & " Term_PERFORM_HISTORY.* "
    Else
        SQLQ = SQLQ & " HR_PERFORM_HISTORY.* "
    End If

'    If glbtermopen Then
'        SQLQ = SQLQ & " Term_PERFORM_HISTORY.*,"
'    Else
'        SQLQ = SQLQ & " HR_PERFORM_HISTORY.*,"
'    End If
'    For x = 0 To 2
'        xfld = "REPTAU" & IIf(x = 0, "", x + 1)
'        If glbLinamar Then
'            SQLQ = SQLQ & " CASE WHEN PH_" & xfld & " IS NOT NULL AND LEN(PH_" & xfld & ")>2 "
'            SQLQ = SQLQ & " THEN RIGHT(PH_" & xfld & ",3)+'-'+"
'            SQLQ = SQLQ & " LEFT(PH_" & xfld & ",LEN(PH_" & xfld & ")-3) "
'            SQLQ = SQLQ & " ELSE STR(PH_" & xfld & ") END "
'            SQLQ = SQLQ & " AS " & xfld & IIf(x = 2, "", ",")
'        Else
'            If glbOracle Then
'                SQLQ = SQLQ & "PH_" & xfld & " AS " & xfld & IIf(x = 2, "", ",")
'            Else
'                SQLQ = SQLQ & "STR(PH_" & xfld & ") AS " & xfld & IIf(x = 2, "", ",")
'            End If
'
'        End If
'    Next
    If glbtermopen Then
        SQLQ = SQLQ & " FROM Term_PERFORM_HISTORY "
        SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
    Else
        SQLQ = SQLQ & " FROM  HR_PERFORM_HISTORY"
        SQLQ = SQLQ & " WHERE PH_EMPNBR = " & glbLEE_ID
    End If
    
    If glbOracle Then
        SQLQ = SQLQ & " AND ROWNUM <=2 "
    End If
    SQLQ = SQLQ & " ORDER BY PH_PREVIEW DESC"
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Do While Not rsTemp.EOF
        If Not IsNull(rsTemp("PH_PREVIEW")) Then dlpReviewDate(intI).Text = rsTemp("PH_PREVIEW")
        If Not IsNull(rsTemp("PH_PCODE")) Then clpCode(intI + 6).Text = rsTemp("PH_PCODE")
        intI = intI + 1
        If intI > 1 Then Exit Do
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    
    If glbtermopen Then
        SQLQ = "Select * from Term_HR_LANGUAGE"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = "Select * "
        SQLQ = SQLQ & " from HR_LANGUAGE "
        SQLQ = SQLQ & " where EL_EMPNBR = " & glbLEE_ID
    End If
    SQLQ = SQLQ & " order by EL_ID ASC"
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    intI = 0
    Do While Not rsTemp.EOF
        If Not IsNull(rsTemp("EL_LANG_SPOKEN")) Then clpCode(intI).Text = rsTemp("EL_LANG_SPOKEN")
        If Not IsNull(rsTemp("EL_LANG_WRITTEN")) Then clpCode(intI + 1).Text = rsTemp("EL_LANG_WRITTEN")
        intI = intI + 2
        If intI > 4 Then Exit Do
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    
    'Ticket #23222 - Formal Education
    If glbtermopen Then
        SQLQ = "Select top 3 * from Term_EDU"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = "Select top 3 * "
        SQLQ = SQLQ & " from HREDU "
        SQLQ = SQLQ & " where EU_EMPNBR = " & glbLEE_ID
    End If
    SQLQ = SQLQ & " order by EU_YEAR DESC"
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    intI = 0
    Do While Not rsTemp.EOF
        If Not IsNull(rsTemp("EU_MAJOR")) Then txtMajor(intI).Text = rsTemp("EU_MAJOR")
        lblMajor(intI).Caption = GetTABLDesc("EUMJ", txtMajor(intI).Text)
        
        If Not IsNull(rsTemp("EU_DEGREE")) Then txtDegree(intI).Text = rsTemp("EU_DEGREE")
        lblDegree(intI).Caption = GetTABLDesc("EUDE", txtDegree(intI).Text)
        
        If Not IsNull(rsTemp("EU_YEAR")) Then lblYear(intI).Caption = rsTemp("EU_YEAR")
        
        intI = intI + 1
        If intI > 2 Then Exit Do
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    
    Call SET_UP_MODE
    
RSDATA.AddNew


If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"

Me.clpCode(0).Enabled = True
'Me.clpCode(0).SetFocus
Me.clpCode(1).Enabled = True
'ComPromot.ListIndex = 0
chkLastReview.Value = 1
glbNew = True
Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_SUCCESSION", "Add")
Resume Next
End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
    Dim x
    On Error GoTo Add_Err
    
    If Not chkSuccession() Then Exit Sub

'    txtPromot = ComPromot.Text


Call UpdUStats(Me)
Call Set_Control("U", Me, RSDATA)

If glbtermopen Then
    RSDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    If fglbNew Then gdbAdoIhr001.Execute "Update Term_HR_SUCCESSION set EU_LAST_RVW = 0 where TERM_SEQ = " & glbTERM_Seq
    RSDATA.Update
    gdbAdoIhr001X.CommitTrans
Else
    gdbAdoIhr001.BeginTrans
    If fglbNew Then gdbAdoIhr001.Execute "Update HR_SUCCESSION set EU_LAST_RVW = 0 where EU_EMPNBR=" & glbLEE_ID
    RSDATA.Update
    gdbAdoIhr001.CommitTrans
End If
Data1.Refresh
fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(False)
If NextFormIF("Succession") Then
    Call cmdNew_Click
End If

Exit Sub

Add_Err:
If Err = 3022 Then
    'Data1.UpdateControls  ' no dups
    Data1.Recordset.CancelUpdate
    Data1.Recordset.Resync
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_SUCCESSION", "Update")
Resume Next
Unload Me

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Succession"
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

RHeading = lblEEName & "'s Succession"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub chkOutsideRVW_Click()
    If chkOutsideRVW.Value = 0 Then
        lblOutsideRVWR.Visible = False
        clpOutsideRVWR.Visible = False
        clpOutsideRVWR.Text = ""
        clpOutsideRVWR.Caption = ""
    Else
        lblOutsideRVWR.Visible = True
        clpOutsideRVWR.Visible = True
    End If
End Sub

Private Sub cmdContEdu_Click()
    Unload frmESuccession
    Set frmESuccession = Nothing 'carmen may 00
    Load frmESEMINARS
End Sub

Private Sub cmdPhoto_Click()
    Call SubPicture
End Sub

'Private Sub ComPromot_Change()
'        txtPromot = ComPromot.Text
'
'End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HR_SUCCESSION", "SELECT")

End Sub

Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError
Screen.MousePointer = HOURGLASS

'Ticket #23923 - Get Employee # of the User - View Own security
If Not glbtermopen Then
    If glbUserEmpNo = glbLEE_ID And Not gSec_SP_ViewOwn Then
        MsgBox "You cannot view your own Succession Planning information.", vbCritical, "info:HR - Security"
        'glbLEE_ID = 0      'Ticket #25208
        Screen.MousePointer = DEFAULT
        Unload Me: Exit Function
    End If
End If


If glbtermopen Then
    SQLQ = "Select * from Term_HR_SUCCESSION"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "Select * "
    SQLQ = SQLQ & " from HR_SUCCESSION "
    SQLQ = SQLQ & " where EU_EMPNBR = " & glbLEE_ID
End If
SQLQ = SQLQ & " ORDER BY EU_CSR_DATE DESC"

Data1.RecordSource = SQLQ
Data1.Refresh

EERetrieve = True

'Ticket #23222 - Photo
If glbSQL Or glbOracle Then
    If cmdPhoto.Caption = "&Photo Off" Then
        picPhoto.Visible = False
        PicNotF.Visible = True
        
        If glbtermopen Then
            Call FillPhoto(Val(glbTERM_ID))
        Else
            Call FillPhoto(Val(glbLEE_ID))
        End If
    Else
        picPhoto.Visible = False
        PicNotF.Visible = False
    End If
    'Ticket #23389 Franks 03/07/2013 - commented following lines
    'If glbWFC Then 'Ticket #21119 Franks 11/14/2011
    '    If IsNull(Data1.Recordset("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = Data1.Recordset("ED_VADIM1")
    'End If
Else
    If Len(glbPicDir) < 1 Then
        picPhoto.Visible = False
    Else
        If cmdPhoto.Caption = "&Photo Off" Then
            picPhoto.Visible = False
            PicNotF.Visible = True
            Call LoadPhoto(Val(glbLEE_ID))
        Else
            picPhoto.Visible = False
            PicNotF.Visible = False
        End If
    End If
End If

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SuccessionRetrieve", "HR_SUCCESSION", "SELECT")
Resume Next

Exit Function

End Function

Private Sub elpReviewer_Change()
    txtReviewer.Text = getEmpnbr(elpReviewer.Text)
End Sub

Private Sub Form_Activate()
    glbOnTop = "FRMESuccession"
Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMESuccession"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "FRMESuccession"

If glbtermopen Then  'Lucy July 4, 2000
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

Screen.MousePointer = HOURGLASS

Call setCaption(lblTitle(20))
Call setCaption(lblTitle(21))

'Ticket #25332: Hiding it because neither Jerry or Frank knows why this is there and there is no corresponding
'database or data entry field.
vbxTrueGrid.Columns(3).Visible = False

Screen.MousePointer = DEFAULT

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

'Ticket #23923 - Get Employee # of the User - View Own security
If Not glbtermopen Then
    If glbUserEmpNo = glbLEE_ID And Not gSec_SP_ViewOwn Then
        MsgBox "You cannot view your own Succession Planning information.", vbCritical, "info:HR - Security"
        'glbLEE_ID = 0      'Ticket #25208
        Screen.MousePointer = DEFAULT
        Unload Me: Exit Sub
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

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "Succession Plan - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENUM.Caption = ShowEmpnbr(lblEEID)

'Ticket #23222 - Photo
Call PhotoFormLoad

Call addItems
Call Display_Value
Call ST_UPD_MODE(False)

If Not gSec_Upd_Basic Then
'    cmdNew.Enabled = False
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
'If glbLinamar Then clpCode(1).TextBoxWidth = 2000

Call INI_Controls(Me)

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

Private Sub Form_Resize()
On Error GoTo Eh
Dim C As Long

If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= 13700 Then
        scrControl.Value = 0
        panDetails.Top = 2040
        scrControl.Visible = False
    Else
        scrControl.Left = Me.Width - 600    '400
        scrControl.Visible = True
        scrControl.Height = Me.Height - 4040
        If Me.Height < 10050 Then
            scrControl.Max = 8000
        Else
            scrControl.Max = 4000
        End If
        
        If Me.Height - scrControl.Top - 480 > 0 Then
        '    scrControl.Height = Me.Height - 2040
        End If
    End If
End If

exH:
    Exit Sub
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HR_SUCCESSION", "form resize")
    Resume exH
End Sub

Private Sub Form_Unload(Cancel As Integer)

    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Call NextForm
End Sub

Private Sub medSKLevel_GotFocus()
    Call SetPanHelp(ActiveControl)
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
 
'If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
'End If

clpCode(0).Enabled = TF
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
clpCode(4).Enabled = TF
clpCode(5).Enabled = TF
clpCode(8).Enabled = TF
'ComPromot.Enabled = TF
dlpReviewDate(2).Enabled = TF
memComments(0).Enabled = TF
memComments(1).Enabled = TF
memComments(2).Enabled = TF
memComments(3).Enabled = TF
chkRelocate.Enabled = TF
txtLocation.Enabled = TF
clpPosCode(1).Enabled = TF
clpPosCode(2).Enabled = TF
clpPosCode(3).Enabled = TF
clpPosCode(4).Enabled = TF
clpPosCode(5).Enabled = TF
txtReviewer.Enabled = TF
elpReviewer.Enabled = TF
chkOutsideRVW.Enabled = TF
clpOutsideRVWR.Enabled = TF

lblPosCode.Caption = clpPosCode(0).Caption
lblCode(0).Caption = clpCode(0).Caption
lblCode(1).Caption = clpCode(1).Caption
lblCode(2).Caption = clpCode(2).Caption
lblCode(3).Caption = clpCode(3).Caption
lblCode(4).Caption = clpCode(4).Caption
lblCode(5).Caption = clpCode(5).Caption
lblCode(6).Caption = clpCode(6).Caption
lblCode(7).Caption = clpCode(7).Caption
lblCode(8).Caption = dlpReviewDate(0).Text
lblCode(9).Caption = dlpReviewDate(1).Text


'txtLangNum.Enabled = TF

End Sub


Private Sub medSKLevel_KeyPress(KeyAscii As Integer)
If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
    KeyAscii = 0
    Exit Sub
End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False

End Sub

Private Sub txtSKComment_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub
'Private Sub txtSKDate_Change()
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtSKDate_DblClick()
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub txtSKDate_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub txtSKDate_KeyPress(KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub

Private Sub scrControl_Change()
    panDetails.Top = 2100 - scrControl.Value
End Sub

Private Sub txtDegree_Change(Index As Integer)
    lblDegree(Index).Caption = GetTABLDesc("EUDE", txtDegree(Index).Text)
End Sub

Private Sub txtDept_Change()
    lblCode(21).Caption = GetDeptName(txtDept.Text, "DF_NAME")
End Sub

Private Sub txtLoc_Change()
    lblCode(18).Caption = GetTABLDesc("EDLC", txtLoc.Text)
End Sub

Private Sub txtMajor_Change(Index As Integer)
        lblMajor(Index).Caption = GetTABLDesc("EUMJ", txtMajor(Index).Text)
End Sub

Private Sub txtReviewer_Change()
    elpReviewer = ShowEmpnbr(txtReviewer.Text)
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
        
        If glbtermopen Then
            SQLQ = "Select * from Term_HR_SUCCESSION"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select * "
            SQLQ = SQLQ & " from HR_SUCCESSION "
            SQLQ = SQLQ & " where EU_EMPNBR = " & glbLEE_ID
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
Dim Skll As String, Skllvl As String, SklDte As String
Dim tdcode$
Dim SQLQ As String

On Error GoTo Tab1_Err
'Sam add july 2002 * remove ado
Call Display_Value
If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
If Data1.Recordset("EU_OUTREVIEW") <> 0 Then
    clpOutsideRVWR.Visible = True
    lblOutsideRVWR.Visible = True
End If
End If
Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_SUCCESSION", "Add")
Resume Next

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
''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()

Dim SQLQ

'Ticket #23923 - Get Employee # of the User - View Own security
If glbLEE_ID = 0 Then Screen.MousePointer = DEFAULT: Unload Me: Exit Sub

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    If glbtermopen Then
        RSDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        RSDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
Else
    If glbtermopen Then
        SQLQ = "Select * from Term_HR_SUCCESSION"
        SQLQ = SQLQ & " WHERE EU_ID = " & Data1.Recordset!EU_ID
        If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
        RSDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "Select * "
        SQLQ = SQLQ & " from HR_SUCCESSION "
        SQLQ = SQLQ & " where EU_ID = " & Data1.Recordset!EU_ID
        If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
        RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    If RSDATA.EOF Or RSDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, RSDATA)
End If

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
UpdateRight = gSec_Upd_SUCCESSION
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

Private Sub lblEEID_Change()
On Error Resume Next
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    frmESuccession.Caption = "Succession - " & Left$(glbLEE_SName, 5)
    frmESuccession.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
 If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
'lblEEID = glbLEE_ID
lblEENUM = ShowEmpnbr(lblEEID)
'ComPromot.ListIndex = 0
'If rsDATA("EU_PROMOT") = "1 year" Then ComPromot.ListIndex = 1
'If rsDATA("EU_PROMOT") = "2 years" Then ComPromot.ListIndex = 2
'If rsDATA("EU_PROMOT") = "3-5 years" Then ComPromot.ListIndex = 3
'If rsDATA("EU_PROMOT") = "+5 years" Then ComPromot.ListIndex = 4
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub

Private Sub addItems()
Dim ctylist, x
'ComPromot.AddItem "a.  Now"
'ComPromot.AddItem "b.  1 year"
'ComPromot.AddItem "c.  2 years"
'ComPromot.AddItem "d.  3-5 years"
'ComPromot.AddItem "e.  +5 years"
'ComPromot.AddItem "Now"
'ComPromot.AddItem "1 year"
'ComPromot.AddItem "2 years"
'ComPromot.AddItem "3-5 years"
'ComPromot.AddItem "+5 years"
'ComPromot.ListIndex = 0
End Sub

'Ticket #23222 - Photo

Sub SubPicture()
Dim xPIC
Dim Msg As String
Dim xHeight, xWidth
On Error GoTo cmdPic_ERR

If glbtermopen Then Exit Sub

If glbSQL Or glbOracle Then
    If cmdPhoto.Caption = "&Photo Off" Then
      picPhoto.Visible = False
      PicNotF.Visible = False
      cmdPhoto.Caption = "&Photo"
    Else
      picPhoto.Visible = False
      PicNotF.Visible = True
      cmdPhoto.Caption = "&Photo Off"
      Call FillPhoto(Val(lblEEID))
    End If
Else
    If Len(glbPicDir) < 1 Then
      picPhoto.Visible = False
      Exit Sub
    End If
    If cmdPhoto.Caption = "&Photo Off" Then
      picPhoto.Visible = False
      PicNotF.Visible = False
      picPhoto = LoadPicture()
      cmdPhoto.Caption = "&Photo"
    Else
      picPhoto.Visible = False
      PicNotF.Visible = True
      cmdPhoto.Caption = "&Photo Off"
      Call LoadPhoto(Val(lblEEID))
    End If
End If
Exit Sub

cmdPic_ERR:
If Err Then
  PicNotF.Visible = True
  Exit Sub
End If

End Sub

Private Function FillPhoto(zEMPNBR As Long)
    On Error GoTo ErrHandler:
    Dim rsPHOTO As New ADODB.Recordset
    Dim byteChunk() As Byte

    Dim Offset As Long
    Dim Totalsize As Long
    Dim Remainder As Long

    Dim FieldSize As Long
    Dim FileNumber As Integer
    Const HeaderSize As Long = 78
    Const ChunkSize As Long = 100
    Dim TempFile As String
    Dim TempDir As String * 255
    
    GetTempPath 255, TempDir
    TempFile = Replace(Replace(TempDir, Chr(0), "") & "\tempfile.tmp", "\\", "\")
    
    picPhoto.Picture = Nothing
    If zEMPNBR = 0 Then Exit Function
    rsPHOTO.Open "select * from HR_PHOTO WHERE PT_EMPNBR=" & zEMPNBR, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If rsPHOTO.EOF Then Exit Function
    
    
    FileNumber = FreeFile
    Open TempFile For Binary Access Write As FileNumber
    
    ReDim byteChunk(rsPHOTO("PT_PHOTO").ActualSize)
    byteChunk() = rsPHOTO("PT_PHOTO").GetChunk(rsPHOTO("PT_PHOTO").ActualSize)
    Put FileNumber, , byteChunk()

    Close FileNumber
    picPhoto.Picture = LoadPicture(TempFile)
    Kill (TempFile)
    rsPHOTO.Close
    Dim xHeight, xWidth
    picPhoto.Stretch = False
    xHeight = picPhoto.Height
    xWidth = picPhoto.Width
    picPhoto.Stretch = True
    picPhoto.Height = 2325
    picPhoto.Width = (xWidth * picPhoto.Height) / xHeight
    picPhoto.Stretch = True
    picPhoto.Visible = True
    PicNotF.Visible = False
    Exit Function
    
ErrHandler:
    MsgBox Err.Description, , "Error "
    
End Function

Private Function LoadPhoto(zEMPNBR As Long)
Dim xHeight, xWidth
glbPicBMP = glbPicDir & zEMPNBR & ".JPG"

'Hemu
If Not IsNull(glbPicBMP) Then
    If Not (Dir(glbPicBMP) = "") Then
        picPhoto = LoadPicture(glbPicBMP)
    Else
        Exit Function
    End If
Else
    Exit Function
End If
'If Not IsNull(glbPicBMP) Then picPhoto = LoadPicture(glbPicBMP)
'Hemu

picPhoto.Stretch = False
xHeight = picPhoto.Height
xWidth = picPhoto.Width
picPhoto.Stretch = True
picPhoto.Height = 2325
picPhoto.Width = (xWidth * picPhoto.Height) / xHeight
picPhoto.Stretch = True
picPhoto.Visible = True
PicNotF.Visible = False
End Function

Private Sub PhotoFormLoad()
Dim xPIC
    xPIC = glbIHRREPORTS & "IHRPICS.MTR"
    If (Dir(xPIC) = "" And Not glbOracle And Not glbSQL) Or glbtermopen Then
      PicNotF.Visible = False
      cmdPhoto.Enabled = False 'Jaddy 10/28/99
      picPhoto.Visible = False
      glbPicDir = ""
      cmdPhoto.Caption = "&Photo"
    Else
      PicNotF.Visible = True
      cmdPhoto.Enabled = True 'Jaddy 10/28/99
      picPhoto.Visible = False
      glbPicDir = glbIHRREPORTS
    End If

    If glbSQL Or glbOracle Then
        If cmdPhoto.Caption = "&Photo Off" Then
          picPhoto.Visible = False
          PicNotF.Visible = True
          Call FillPhoto(Val(glbLEE_ID))
        Else
          picPhoto.Visible = False
          PicNotF.Visible = False
        End If
    Else
        If Len(glbPicDir) < 1 Then
          picPhoto.Visible = False
        Else
            If cmdPhoto.Caption = "&Photo Off" Then
              picPhoto.Visible = False
              PicNotF.Visible = True
              Call LoadPhoto(Val(glbLEE_ID))
            Else
                picPhoto.Visible = False
                PicNotF.Visible = False
            End If
        End If
    End If
    
End Sub

