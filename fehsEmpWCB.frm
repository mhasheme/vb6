VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEHSEMPWCB 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Claim/Medical"
   ClientHeight    =   9645
   ClientLeft      =   30
   ClientTop       =   900
   ClientWidth     =   13860
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
   ScaleHeight     =   9645
   ScaleWidth      =   13860
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   0
      Max             =   50
      SmallChange     =   4
      TabIndex        =   68
      Top             =   9240
      Width           =   13095
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fehsEmpWCB.frx":0000
      Height          =   1455
      Left            =   0
      OleObjectBlob   =   "fehsEmpWCB.frx":0014
      TabIndex        =   31
      Top             =   480
      Width           =   12975
   End
   Begin VB.VScrollBar scrControl 
      Height          =   6915
      LargeChange     =   315
      Left            =   13080
      Max             =   100
      SmallChange     =   315
      TabIndex        =   67
      Top             =   2040
      Visible         =   0   'False
      Width           =   300
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   9600
      Top             =   8280
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
      Caption         =   "Ado2"
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EC_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   8280
      MaxLength       =   25
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   9360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EC_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   10050
      MaxLength       =   25
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   9360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EC_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   12030
      MaxLength       =   25
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   9360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   13860
      _Version        =   65536
      _ExtentX        =   24447
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
         Left            =   7200
         TabIndex        =   74
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
         Top             =   135
         Width           =   720
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   120
      Top             =   9300
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
   Begin VB.Frame ScrFrame 
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   0
      TabIndex        =   41
      Top             =   2040
      Width           =   13095
      Begin VB.ComboBox comIncidentNo 
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Tag             =   "01-Incident Number"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "First Physician"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   120
         TabIndex        =   54
         Top             =   1980
         Width           =   10635
         Begin VB.TextBox txtMedInfo 
            Appearance      =   0  'Flat
            DataField       =   "EC_PHYSADDR"
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
            Left            =   6000
            MaxLength       =   40
            TabIndex        =   11
            Tag             =   "00-Address of physician/health care worker"
            Top             =   600
            Width           =   4485
         End
         Begin VB.TextBox txtMedInfo 
            Appearance      =   0  'Flat
            DataField       =   "EC_PHYSNM"
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
            Left            =   6000
            MaxLength       =   40
            TabIndex        =   9
            Tag             =   "00-Name of Physician/health care worker"
            Top             =   270
            Width           =   4485
         End
         Begin VB.TextBox txtPHYSAddress 
            Appearance      =   0  'Flat
            DataField       =   "EC_PHYS1_EMAIL"
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
            Left            =   1630
            TabIndex        =   12
            Tag             =   "Email Address"
            Top             =   900
            Width           =   2475
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_PHYSCOD"
            Height          =   285
            Index           =   2
            Left            =   1320
            TabIndex        =   8
            Tag             =   "01-Type of doctor/health care provided- Code"
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ECPT"
         End
         Begin MSMask.MaskEdBox MedPhone 
            DataField       =   "EC_DOCPHONE"
            Height          =   285
            Index           =   1
            Left            =   1630
            TabIndex        =   10
            Tag             =   "00-Phone of Physician/health care worker"
            Top             =   570
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   19
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "(###) ###-####    Ext(####)"
            Mask            =   "(###) ###-####    Ext(####)"
            PromptChar      =   "_"
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            Height          =   285
            Index           =   2
            Left            =   1320
            TabIndex        =   13
            Tag             =   "40-Medical Visit"
            Top             =   1220
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1105
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            Height          =   285
            Index           =   3
            Left            =   6000
            TabIndex        =   14
            Tag             =   "40-Employer Notified"
            Top             =   1215
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1105
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employer Notified"
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
            Left            =   4560
            TabIndex        =   76
            Top             =   1260
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Medical Visit"
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
            TabIndex        =   75
            Top             =   1265
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
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
            Left            =   120
            TabIndex        =   59
            Top             =   600
            Width           =   465
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " Address"
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
            Left            =   5250
            TabIndex        =   58
            Top             =   660
            Width           =   615
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " Name"
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
            Left            =   5250
            TabIndex        =   57
            Top             =   330
            Width           =   465
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   56
            Top             =   285
            Width           =   360
         End
         Begin VB.Label lblAddress 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Email "
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
            Index           =   0
            Left            =   120
            TabIndex        =   55
            Top             =   900
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Second Physician"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   120
         TabIndex        =   48
         Top             =   3510
         Width           =   5265
         Begin VB.TextBox txtMedInfo 
            Appearance      =   0  'Flat
            DataField       =   "EC_PHYSNM2"
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
            Index           =   3
            Left            =   1380
            MaxLength       =   40
            TabIndex        =   16
            Tag             =   "00-Name of Physician/health care worker"
            Top             =   600
            Width           =   3585
         End
         Begin VB.TextBox txtMedInfo 
            Appearance      =   0  'Flat
            DataField       =   "EC_PHYSADDR2"
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
            Index           =   4
            Left            =   1380
            MaxLength       =   40
            TabIndex        =   18
            Tag             =   "00-Address of physician/health care worker"
            Top             =   1260
            Width           =   3585
         End
         Begin VB.TextBox txtPHYSAddress 
            Appearance      =   0  'Flat
            DataField       =   "EC_PHYS2_EMAIL"
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
            Left            =   1380
            TabIndex        =   19
            Tag             =   "Email Address"
            Top             =   1590
            Width           =   2475
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_PHYSCOD2"
            Height          =   285
            Index           =   4
            Left            =   1065
            TabIndex        =   15
            Tag             =   "00-Type of doctor/health care provided- Code"
            Top             =   270
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ECPT"
         End
         Begin MSMask.MaskEdBox MedPhone 
            DataField       =   "EC_DOCPHONE2"
            Height          =   285
            Index           =   2
            Left            =   1380
            TabIndex        =   17
            Tag             =   "00-Phone of Physician/health care worker"
            Top             =   930
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   19
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "(###) ###-####    Ext(####)"
            Mask            =   "(###) ###-####    Ext(####)"
            PromptChar      =   "_"
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            Height          =   285
            Index           =   4
            Left            =   1440
            TabIndex        =   20
            Tag             =   "40-Medical Visit"
            Top             =   1920
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1105
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            Height          =   285
            Index           =   5
            Left            =   1440
            TabIndex        =   21
            Tag             =   "40-Employer Notified"
            Top             =   2240
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1105
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employer Notified"
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
            Index           =   22
            Left            =   60
            TabIndex        =   78
            Top             =   2240
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Medical Visit"
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
            Index           =   21
            Left            =   60
            TabIndex        =   77
            Top             =   1920
            Visible         =   0   'False
            Width           =   1125
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   11
            Left            =   60
            TabIndex        =   53
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            Left            =   60
            TabIndex        =   52
            Top             =   630
            Width           =   420
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
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
            TabIndex        =   51
            Top             =   1290
            Width           =   570
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
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
            TabIndex        =   50
            Top             =   960
            Width           =   465
         End
         Begin VB.Label lblAddress 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Email "
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
            Index           =   1
            Left            =   60
            TabIndex        =   49
            Top             =   1620
            Width           =   825
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Third Physician"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   5520
         TabIndex        =   42
         Top             =   3510
         Width           =   5265
         Begin VB.TextBox txtMedInfo 
            Appearance      =   0  'Flat
            DataField       =   "EC_PHYSADDR3"
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
            Index           =   5
            Left            =   1470
            MaxLength       =   40
            TabIndex        =   25
            Tag             =   "00-Address of physician/health care worker"
            Top             =   1260
            Width           =   3585
         End
         Begin VB.TextBox txtMedInfo 
            Appearance      =   0  'Flat
            DataField       =   "EC_PHYSNM3"
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
            Index           =   6
            Left            =   1470
            MaxLength       =   40
            TabIndex        =   23
            Tag             =   "00-Name of Physician/health care worker"
            Top             =   600
            Width           =   3585
         End
         Begin VB.TextBox txtPHYSAddress 
            Appearance      =   0  'Flat
            DataField       =   "EC_PHYS3_EMAIL"
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
            Left            =   1470
            TabIndex        =   26
            Tag             =   "Email Address"
            Top             =   1590
            Width           =   2475
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_PHYSCOD3"
            Height          =   285
            Index           =   5
            Left            =   1155
            TabIndex        =   22
            Tag             =   "00-Type of doctor/health care provided- Code"
            Top             =   270
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ECPT"
         End
         Begin MSMask.MaskEdBox MedPhone 
            DataField       =   "EC_DOCPHONE3"
            Height          =   285
            Index           =   3
            Left            =   1470
            TabIndex        =   24
            Tag             =   "00-Phone of Physician/health care worker"
            Top             =   930
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   19
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "(###)###-#### Ext(####)"
            Mask            =   "(###)###-#### Ext(####)"
            PromptChar      =   "_"
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            Height          =   285
            Index           =   6
            Left            =   1500
            TabIndex        =   27
            Tag             =   "40-Medical Visit"
            Top             =   1920
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1105
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            Height          =   285
            Index           =   7
            Left            =   1500
            TabIndex        =   28
            Tag             =   "40-Employer Notified"
            Top             =   2235
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1105
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employer Notified"
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
            Left            =   120
            TabIndex        =   80
            Top             =   2235
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Medical Visit"
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
            Index           =   23
            Left            =   120
            TabIndex        =   79
            Top             =   1965
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
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
            Left            =   120
            TabIndex        =   47
            Top             =   990
            Width           =   465
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
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
            Left            =   120
            TabIndex        =   46
            Top             =   1320
            Width           =   570
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            TabIndex        =   45
            Top             =   660
            Width           =   420
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   18
            Left            =   120
            TabIndex        =   44
            Top             =   315
            Width           =   360
         End
         Begin VB.Label lblAddress 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Email "
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
            Index           =   2
            Left            =   120
            TabIndex        =   43
            Top             =   1605
            Width           =   825
         End
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_RATEGRP"
         Height          =   285
         Index           =   3
         Left            =   1125
         TabIndex        =   30
         Tag             =   "00-Rate Group Code"
         Top             =   6120
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECGP"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_WCBRES"
         Height          =   285
         Index           =   1
         Left            =   9165
         TabIndex        =   3
         Tag             =   "01-Results of Claim - Code"
         Top             =   780
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECRS"
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "EC_WCBCDTE"
         Height          =   285
         Index           =   1
         Left            =   9150
         TabIndex        =   7
         Tag             =   "40-Date Claim was closed"
         Top             =   1550
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin VB.TextBox txtMedInfo 
         Appearance      =   0  'Flat
         DataField       =   "EC_HOSP"
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
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   29
         Tag             =   "00-Name of Hospital"
         Top             =   5700
         Width           =   5475
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "EC_WCBFDTE"
         Height          =   285
         Index           =   0
         Left            =   5205
         TabIndex        =   2
         Tag             =   "41-Date Filed with WSIB"
         Top             =   780
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "EC_APPROVE_DATE"
         Height          =   285
         Index           =   8
         Left            =   1470
         TabIndex        =   4
         Tag             =   "41-Date Filed with WSIB"
         Top             =   1155
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "EC_FROM_DATE"
         Height          =   285
         Index           =   9
         Left            =   5190
         TabIndex        =   5
         Tag             =   "41-Date Filed with WSIB"
         Top             =   1200
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "EC_TO_DATE"
         Height          =   285
         Index           =   10
         Left            =   9150
         TabIndex        =   6
         Tag             =   "41-Date Filed with WSIB"
         Top             =   1160
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin MSMask.MaskEdBox medClaimNo 
         DataField       =   "EC_WCBNBR"
         Height          =   285
         Left            =   1815
         TabIndex        =   1
         Tag             =   "00-Claim Number"
         Top             =   800
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AAAAAAAAA-AA"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc Data2 
         Height          =   330
         Left            =   9600
         Top             =   5760
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
         Caption         =   "Ado2"
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
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Claim Number"
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
         Left            =   255
         TabIndex        =   86
         Top             =   845
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
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
         Left            =   8400
         TabIndex        =   85
         Top             =   1205
         Width           =   585
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
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
         Left            =   4320
         TabIndex        =   84
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Approved Date"
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
         Index           =   26
         Left            =   255
         TabIndex        =   83
         Top             =   1205
         Width           =   1080
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Claim Status"
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
         Left            =   8145
         TabIndex        =   82
         Top             =   825
         Width           =   870
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Incident Number"
         Height          =   255
         Index           =   25
         Left            =   240
         TabIndex        =   81
         Top             =   390
         Width           =   1575
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Incident Number:"
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
         Left            =   3495
         TabIndex        =   72
         Top             =   420
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "File Closed On"
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
         Left            =   7935
         TabIndex        =   71
         Top             =   1595
         Width           =   1020
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Filed On"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   4335
         TabIndex        =   70
         Top             =   825
         Width           =   720
      End
      Begin VB.Label lblWCB 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CLAIM INFORMATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   0
         Width           =   1920
      End
      Begin VB.Label lblMed 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MEDICAL INFORMATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   165
         TabIndex        =   66
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hospital"
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
         TabIndex        =   65
         Top             =   5760
         Width           =   570
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Group"
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
         Left            =   120
         TabIndex        =   64
         Top             =   6150
         Width           =   1020
      End
      Begin VB.Label lblUserDesc 
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
         Left            =   6480
         TabIndex        =   63
         Top             =   6120
         Width           =   2295
      End
      Begin VB.Label lblUpdateBy 
         Caption         =   "Updated By"
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
         Left            =   5280
         TabIndex        =   62
         Top             =   6120
         Width           =   975
      End
      Begin VB.Label lblUpdDateDesc 
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
         Left            =   6480
         TabIndex        =   61
         Top             =   6360
         Width           =   1935
      End
      Begin VB.Label lblUpdateDate 
         Caption         =   "Updated Date"
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
         Left            =   5280
         TabIndex        =   60
         Top             =   6360
         Width           =   1095
      End
      Begin VB.Label lblIncidentNo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         DataField       =   "EC_CASE"
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
         Left            =   4815
         TabIndex        =   73
         Tag             =   "11-Unique ID of Incident"
         Top             =   420
         Visible         =   0   'False
         Width           =   90
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
      Left            =   7560
      TabIndex        =   39
      Top             =   9360
      Visible         =   0   'False
      Width           =   615
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
      Left            =   6060
      TabIndex        =   40
      Top             =   9360
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmEHSEMPWCB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim fglbNew

Private Function chkHSWCB()
Dim SQLQ As String, Msg As String, dd#, x%

chkHSWCB = False

On Error GoTo chkHSWCB_Err
'Jdy 4/10/00
'If Len(medClaimNo) <= 0 Then
'    MsgBox "Claim Number is a required field"
'    medClaimNo.SetFocus
'    Exit Function
'End If
If Len(lblIncidentNo.Caption) < 1 Then
    MsgBox "Incident Number is a required field"
    comIncidentNo.SetFocus
    Exit Function
End If
If Not IfIncidentNo(Val(lblIncidentNo.Caption)) Then
    MsgBox "Incident Number Not Valid"
    comIncidentNo.SetFocus
    Exit Function
End If

If Len(dlpDate(0).Text) >= 1 Then
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "Date filed is not a valid date."
        dlpDate(0).SetFocus
        Exit Function
    End If
Else
    MsgBox "Date filed is required."
    dlpDate(0).SetFocus
    Exit Function
End If

If Len(dlpDate(1).Text) >= 1 Then
    If Not IsDate(dlpDate(1).Text) Then
        MsgBox "Date file closed is not a valid date."
        dlpDate(1).SetFocus
        Exit Function
    End If
End If

'If Len(clpCode(2).Text) < 1 Then   'As per new release documentation
'    MsgBox "Physician/Health Care code is a required field"
'     clpCode(2).SetFocus
'    Exit Function
'End If
'Jdy 4/10/00
'If Len( clpCode(3)) < 1 Then
'    MsgBox "Rate Group is a required field"
'     clpCode(3).SetFocus
'    Exit Function
'End If

If Not IsEmail(txtPHYSAddress(0).Text) Then
    MsgBox "Email address must be in xxx@yyy.zzz format.", vbExclamation + vbOKOnly, "Invalid Email Address"
    Exit Function
End If

If Not IsEmail(txtPHYSAddress(1).Text) Then
    MsgBox "Email address must be in xxx@yyy.zzz format.", vbExclamation + vbOKOnly, "Invalid Email Address"
    Exit Function
End If

If Not IsEmail(txtPHYSAddress(2).Text) Then
    MsgBox "Email address must be in xxx@yyy.zzz format.", vbExclamation + vbOKOnly, "Invalid Email Address"
    Exit Function
End If

For x% = 1 To 3
    If Len(clpCode(x%).Text) > 0 And clpCode(x%).Caption = "Unassigned" Then
        If x% = 1 Then MsgBox "Incident Result code must be valid"
        If x% = 2 Then MsgBox "Physician Type code must be valid"
        If x% = 3 Then MsgBox "Rate Group code must be valid"
         clpCode(x%).SetFocus
        Exit Function
    End If
Next

chkHSWCB = True

Exit Function

chkHSWCB_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSWCB", "HR_OHS_CLAIM_MEDICAL", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

'Private Sub cmdCAction_Click()
'frmEHSCorrective.Show
'Unload Me
'End Sub

'Private Sub cmdCAction_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdCancel_Click()
Dim x
On Error GoTo Can_Err

If Not (rsDATA.EOF And rsDATA.BOF) Then rsDATA.CancelUpdate
Call Display_Value
fglbNew = False
'Call ST_UPD_MODE(True)  ' reset screen's attributes
Call SET_UP_MODE
Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_OHS_CLAIM_MEDICAL", "Cancel")
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
If glbOnTop = "FRMEHSEMPWCB" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdContact_Click()
'frmEHSContact.Show
'Unload Me
'End Sub

'Private Sub cmdContact_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdIncident_Click()
'frmEHSINCIDENT.Show
'Unload Me
'End Sub

'Private Sub cmdIncident_GotFocus()
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

On Error GoTo Mod_Err

'Call ST_UPD_MODE(True)
Call SET_UP_MODE
'medClaimNo.SetFocus
Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_OHS_CLAIM_MEDICAL", "Modify")
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

Sub cmdOK_Click()
Dim x

On Error GoTo Add_Err

If Not chkHSWCB() Then Exit Sub
rsDATA.Requery
If fglbNew Then
    rsDATA.AddNew
End If
Call UpdUStats(Me) ' update user's stats (who did it and when)
Call Set_Control("U", Me, rsDATA)
If rsDATA!EC_WCBNBR = "" Then rsDATA!EC_WCBNBR = Null
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
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
x = NextFormIF("Claim/Medical")
Exit Sub

Add_Err:
If Err = 3022 Then
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OHS_CLAIM_MEDICAL", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

'Private Sub cmdOK_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String
RHeading = lblEEName & "'s WSIB/Medical"
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

RHeading = lblEEName & "'s WSIB/Medical"
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

'Private Sub cmdWSIB_Click()
'frmEHSEMPWCBC.Show
'Unload Me
'End Sub

Function EERetrieve()
Dim SQLQ As String

EERetrieve = False
Screen.MousePointer = HOURGLASS

On Error GoTo EERError
If glbtermopen Then
    SQLQ = "SELECT " & FldList & " FROM Term_OHS_CLAIM_MEDICAL "
    SQLQ = SQLQ & "WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "SELECT " & FldList & " FROM HR_OHS_CLAIM_MEDICAL "
    SQLQ = SQLQ & "WHERE EC_EMPNBR = " & glbLEE_ID
End If
Data1.RecordSource = SQLQ
Data1.Refresh

If glbtermopen Then
    SQLQ = "SELECT EC_EMPNBR, EC_CASE, EC_OCCDATE from Term_HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
Else
    SQLQ = "SELECT EC_EMPNBR, EC_CASE, EC_OCCDATE from HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " WHERE EC_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
End If

Data2.RecordSource = SQLQ
Data2.Refresh

comIncidentNo.Clear
Do Until Data2.Recordset.EOF                  'JDY
  comIncidentNo.AddItem Data2.Recordset("EC_CASE") 'JDY
  Data2.Recordset.MoveNext                    'JDY
Loop

EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "OCH Retrieve", "HR_OHS_CLAIM_MEDICAL", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

Exit Function

End Function

Private Sub comIncidentNo_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comIncidentNo_LostFocus()
lblIncidentNo.Caption = comIncidentNo
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMEHSEMPWCB"
End Sub

Private Sub Form_GotFocus()
glbOnTop = "FRMEHSEMPWCB"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim xInt As Integer

glbOnTop = "FRMEHSEMPWCB"
If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
    Data2.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
    Data2.ConnectionString = glbAdoIHRDB
End If


Screen.MousePointer = DEFAULT

If glbLinHS Then 'Ticket #12401
    glbLinEmpNo = glbLEE_ID
    If Not glbtermopen Then
        If Len(glbDiv) = 0 Then Call Get_Div(False) 'frmDIVISIONS.Show 1
        If Len(glbDiv) = 0 Then Unload Me: Exit Sub
    Else
        If Len(glbDiv) = 0 Then Call Get_Div(False) 'frmDIVISIONS.Show 1
        If Len(glbDiv) = 0 Then Unload Me: Exit Sub
    End If
    glbLinHSDivNo = Val("999999" & glbDiv)
    glbLEE_ID = glbLinHSDivNo
    glbLEE_SName = glbDivDesc
Else
    If glbLinamar Then
        If glbLEE_ID <> 0 Then
            If Left(Trim(str(glbLEE_ID)), 6) = "999999" Then
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

If glbLinamar Then 'Ticket #15172
    Frame1.Caption = "First Health Professional"
    Frame1.Height = 1545
    lblTitle(19).Visible = True
    lblTitle(20).Visible = True
    dlpDate(2).Visible = True
    dlpDate(2).DataField = "EC_PHYS1_VISIT"
    dlpDate(3).Visible = True
    dlpDate(3).DataField = "EC_PHYS1_NOTIFIED"
    
    Frame2.Caption = "Second Health Professional"
    Frame2.Top = 3630 '+ 30
    Frame2.Height = 2595
    lblTitle(21).Visible = True
    lblTitle(22).Visible = True
    dlpDate(4).Visible = True
    dlpDate(4).DataField = "EC_PHYS2_VISIT"
    dlpDate(5).Visible = True
    dlpDate(5).DataField = "EC_PHYS2_NOTIFIED"
    
    Frame3.Caption = "Third Health Professional"
    Frame3.Top = 3630 '+ 30
    Frame3.Height = 2595
    lblTitle(23).Visible = True
    lblTitle(24).Visible = True
    dlpDate(6).Visible = True
    dlpDate(6).DataField = "EC_PHYS3_VISIT"
    dlpDate(7).Visible = True
    dlpDate(7).DataField = "EC_PHYS3_NOTIFIED"
    
    xInt = 1500
    lblTitle(7).Top = 4920 + xInt
    txtMedInfo(2).Top = 4860 + xInt
    lblTitle(9).Top = 5310 + xInt
    clpCode(3).Top = 5280 + xInt
    lblUpdateBy.Top = 5280 + xInt
    lblUserDesc.Top = 5280 + xInt
    lblUpdateDate.Top = 5520 + xInt
    lblUpdDateDesc.Top = 5520 + xInt
    
Else
    Frame1.Height = 1425
    Frame2.Height = 1995
    Frame3.Height = 1995
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

comIncidentNo.Clear
Do Until Data2.Recordset.EOF                  'JDY
  comIncidentNo.AddItem Data2.Recordset("EC_CASE") 'JDY
  Data2.Recordset.MoveNext                    'JDY
Loop

Me.vbxTrueGrid.SetFocus

If glbLinHS Then
    If Len(glbDivDesc) > 0 Then   ' dont do on add new until in
        Me.Caption = "WSIB/Medical - " & glbDivDesc
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv
    lblEENumber.Caption = lStr("Division")
Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
        Me.Caption = "WSIB/Medical - " & Left$(glbLEE_SName, 5)
        Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    lblEENum.Caption = ShowEmpnbr(lblEEID)
End If

Call ST_UPD_MODE(False)
Call Display_Value
If Not gSec_Upd_Health_Safety Then
'    cmdModify.Enabled = False
End If

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False



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
If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= 8350 Then
        scrControl.Value = 0
        ScrFrame.Top = 2040
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        If Me.Height < 7500 Then
            scrControl.Max = 5000
        Else
            scrControl.Max = 1500
        End If
        scrControl.Left = Me.Width - scrControl.Width - 120
        If Me.Height - scrControl.Top - 780 > 0 Then
            scrControl.Height = Me.Height - scrControl.Top - 780
        End If
    End If
    
    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 120
    'ScrFrame.Height = Me.ScaleHeight - (scrHScroll.Height + 200)
    If Me.Width >= 9750 Then
        scrHScroll.Value = 0
        scrHScroll.Visible = False
    Else
        scrHScroll.Visible = True
        If Me.Width < 7500 Then
            scrHScroll.Max = 200
        Else
            scrHScroll.Max = 30
        End If
        scrHScroll.Top = Me.Height - 800
        scrHScroll.Width = Me.Width - 120
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

MDIMain.panHelp(0).Caption = "Select function from the menu."
Set frmEHSEMPWCB = Nothing 'carmen may 00
Call NextForm
End Sub




Private Sub lblIncidentNo_Change()
    If Not (Val(lblIncidentNo.Caption) = 0) Then
        comIncidentNo = lblIncidentNo
    Else
        comIncidentNo = ""
    End If
End Sub

Private Sub medClaimNo_GotFocus()
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
'cmdPrint.Enabled = FT

'cmdCAction.Enabled = FT
'cmdIncident.Enabled = FT
'cmdTCause.Enabled = FT
'cmdContact.Enabled = FT
'cmdInjLoc.Enabled = FT
'cmdWSIB.Enabled = FT

comIncidentNo.Enabled = TF
medClaimNo.Enabled = TF
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
dlpDate(0).Enabled = TF
dlpDate(1).Enabled = TF
dlpDate(8).Enabled = TF
dlpDate(9).Enabled = TF
dlpDate(10).Enabled = TF
txtMedInfo(0).Enabled = TF
txtMedInfo(1).Enabled = TF
txtMedInfo(2).Enabled = TF
Frame1.Enabled = TF
Frame2.Enabled = TF
Frame3.Enabled = TF
'vbxTrueGrid.Enabled = FT
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'   cmdModify.Enabled = False
Else
'Me.cmdModify_Click
End If
End Sub

Private Sub MedPhone_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub TDBGrid1_Click()

End Sub

Private Sub scrControl_Change()
ScrFrame.Top = 2040 - scrControl.Value
End Sub

Private Sub scrHScroll_Change()
ScrFrame.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
End Sub

Private Sub txtMedInfo_GotFocus(Index As Integer)
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
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        If glbtermopen Then
            SQLQ = "SELECT " & FldList & " FROM Term_OHS_CLAIM_MEDICAL "
            SQLQ = SQLQ & "WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "SELECT " & FldList & " FROM HR_OHS_CLAIM_MEDICAL "
            SQLQ = SQLQ & "WHERE EC_EMPNBR = " & glbLEE_ID
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
Dim tdcode$, Y
Dim SQLQ As String



On Error GoTo Tab1_Err
'Y = Fnd_Match_Data1()
Call Display_Value

 ' set description for code
 ' set description for code


If Data1.Recordset.RecordCount <> 0 Then
    If Not IsNull(Data1.Recordset("EC_WCBCDTE")) Then
        dlpDate(1).Text = Data1.Recordset("EC_WCBCDTE")
    Else
        dlpDate(1).Text = ""
    End If
End If

Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_OHS_CLAIM_MEDICAL", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub
Private Function FldList()
Dim SQLQ
SQLQ = ""
SQLQ = SQLQ & "EC_COMPNO, EC_EMPNBR, EC_CASE, EC_WCBNBR, "
SQLQ = SQLQ & "EC_WCBFDTE, EC_WCBCDTE, EC_WCBRES, EC_PHYSCOD,"
SQLQ = SQLQ & "EC_PHYSNM, EC_PHYSADDR, EC_HOSP, EC_RATEGRP,"
SQLQ = SQLQ & "EC_LDATE, EC_LTIME, EC_LUSER, EC_PHYSCOD2,"
SQLQ = SQLQ & "EC_PHYSNM2, EC_PHYSADDR2, EC_DOCPHONE,"
SQLQ = SQLQ & "EC_DOCPHONE2, EC_PHYSCOD3, EC_PHYSNM3,"
SQLQ = SQLQ & "EC_PHYS1_EMAIL, EC_PHYS2_EMAIL, EC_PHYS3_EMAIL,"
SQLQ = SQLQ & "EC_PHYSADDR3 , EC_DOCPHONE3"
If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"
If glbLinamar Then
SQLQ = SQLQ & ",EC_PHYS1_VISIT,EC_PHYS1_NOTIFIED,EC_PHYS2_VISIT,EC_PHYS2_NOTIFIED,EC_PHYS3_VISIT,EC_PHYS3_NOTIFIED"
SQLQ = SQLQ & ",EC_APPROVE_DATE,EC_FROM_DATE,EC_TO_DATE"
End If
FldList = SQLQ
End Function

''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
    Dim SQLQ
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
 
    If glbtermopen Then
        SQLQ = "SELECT " & FldList & " FROM Term_OHS_CLAIM_MEDICAL "
        SQLQ = SQLQ & "WHERE EC_CASE = " & Data1.Recordset!EC_CASE
        SQLQ = SQLQ & " AND EC_WCBFDTE = " & Date_SQL(Data1.Recordset!EC_WCBFDTE)
        
        'If glbWFC Then
        '  SQLQ = SQLQ & " AND TERM_SEQ = " & glbTERM_Seq
        'End If
        SQLQ = SQLQ & " AND TERM_SEQ = " & glbTERM_Seq
        If Not IsNull(Data1.Recordset!EC_WCBNBR) Then
            SQLQ = SQLQ & " AND EC_WCBNBR ='" & Data1.Recordset!EC_WCBNBR & "'"
        End If
        
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "SELECT " & FldList & " FROM HR_OHS_CLAIM_MEDICAL "
        SQLQ = SQLQ & "WHERE EC_CASE = " & Data1.Recordset!EC_CASE
        If Not IsNull(Data1.Recordset!EC_WCBNBR) Then
            SQLQ = SQLQ & " AND EC_WCBNBR = '" & Data1.Recordset!EC_WCBNBR & "'"
        End If
        SQLQ = SQLQ & " AND EC_WCBFDTE = " & Date_SQL(Data1.Recordset!EC_WCBFDTE)
        'If glbWFC Then
        '      SQLQ = SQLQ & " AND EC_EMPNBR =" & glbLEE_ID
        'End If
        SQLQ = SQLQ & " AND EC_EMPNBR =" & glbLEE_ID
        'SQLQ = SQLQ & " AND EC_WCBNBR ='" & Data1.Recordset!EC_WCBNBR & "'"
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If

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
        Me.Caption = "Claim / Medical Information  - " & glbDivDesc
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv

    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = ""
    End If

Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
        frmEHSEMPWCB.Caption = "Claim / Medical Information - " & Left$(glbLEE_SName, 5)
        frmEHSEMPWCB.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
    'lblEEID = glbLEE_ID
    lblEENum = ShowEmpnbr(lblEEID)
    
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = glbLEE_ProdLine
    Else
        lblEEProdLine = ""
    End If
    
End If
End Sub

Private Function IsEmail(Address As String) As Boolean
    IsEmail = True
    If Len(Address) = 0 Then Exit Function
    ' Make sure there's an @ in the address
    If InStr(Address, "@") = 0 Then IsEmail = False: Exit Function
    ' Make sure they have at least one period after the @
    If InStr(InStr(Address, "@"), Address, ".") = 0 Then IsEmail = False: Exit Function
    ' Make sure they have text before the period
    If Mid(Address, InStr(Address, "@") + 1, 1) = "." Then IsEmail = False: Exit Function
    ' Make sure they have text after the period
    If Right(Address, 1) = "." Then IsEmail = False: Exit Function
End Function

Function IfIncidentNo(InciNo As Double)
  IfIncidentNo = False
  If Data2.Recordset.BOF And Data2.Recordset.EOF Then
     Exit Function
  End If
  Data2.Recordset.MoveFirst
  Data2.Recordset.Find "EC_Case=" & InciNo
  If Data2.Recordset.EOF Then Exit Function
  IfIncidentNo = True

End Function

Sub cmdNew_Click()
    Dim SQLQ As String
    
    If Not gSec_Upd_Health_Safety Then
        MsgBox "You Do Not Have Authority For This Transacaction"
        Exit Sub
    End If
    
    fglbNew = True
    Call SET_UP_MODE
    
    On Error GoTo AddN_Err
    
    fglbNew = True
    
    Call Set_Control("B", Me)
    
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
    
    If glbtermopen Then
        SQLQ = "SELECT EC_EMPNBR, EC_CASE, EC_OCCDATE from Term_HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
    Else
        SQLQ = "SELECT EC_EMPNBR, EC_CASE, EC_OCCDATE from HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & " WHERE EC_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
    End If
    
    Data2.RecordSource = SQLQ
    Data2.Refresh
    
    comIncidentNo.Clear
    Do Until Data2.Recordset.EOF                  'JDY
      comIncidentNo.AddItem Data2.Recordset("EC_CASE") 'JDY
      Data2.Recordset.MoveNext                    'JDY
    Loop
    
    lblCNum.Caption = "001"
    comIncidentNo.SetFocus
    Exit Sub
    
AddN_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_OHS_CLAIM_MEDICAL", "Add")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If

End Sub

Sub cmdDelete_Click()
    Dim a As Integer, Msg As String, INo&, x
    
    If Not gSec_Upd_Health_Safety Then
        MsgBox "You Do Not Have Authority For This Transacaction"
        Exit Sub
    End If
    
    If Data1.Recordset.BOF And Data1.Recordset.EOF Then
        MsgBox "Nothing to Delete"
        Exit Sub
    End If
    
    
    On Error GoTo Del_Err
    
    
    Msg = "Are You Sure You Want To Delete "
    Msg = Msg & Chr(10) & "This Record?  "
    
    a% = MsgBox(Msg, 36, "Confirm Delete")
    If a% <> 6 Then Exit Sub
    
    
    If glbtermopen Then
        gdbAdoIhr001X.BeginTrans
        rsDATA.Delete
        gdbAdoIhr001X.CommitTrans
        Data1.Refresh
    Else
        gdbAdoIhr001.BeginTrans
        rsDATA.Delete
        gdbAdoIhr001.CommitTrans
        Data1.Refresh
    End If
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        Call Display_Value
    End If
        
    Me.vbxTrueGrid.SetFocus
    fglbNew = False
    
    Call SET_UP_MODE
    
    Exit Sub
    
Del_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_OHS_CLAIM_MEDICAL", "Delete")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

