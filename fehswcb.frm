VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEHSWCB 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Claim/Medical"
   ClientHeight    =   10620
   ClientLeft      =   30
   ClientTop       =   900
   ClientWidth     =   12120
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
   ScaleHeight     =   10620
   ScaleWidth      =   12120
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   120
      Max             =   50
      SmallChange     =   4
      TabIndex        =   64
      Top             =   9720
      Width           =   10575
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fehswcb.frx":0000
      Height          =   2325
      Left            =   120
      OleObjectBlob   =   "fehswcb.frx":0014
      TabIndex        =   0
      Top             =   600
      Width           =   11895
   End
   Begin VB.VScrollBar scrControl 
      Height          =   6675
      LargeChange     =   315
      Left            =   11640
      Max             =   100
      SmallChange     =   315
      TabIndex        =   63
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7920
      Top             =   10200
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
      Left            =   2400
      MaxLength       =   25
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   10080
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
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   10080
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EC_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   6150
      MaxLength       =   25
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   10080
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   12120
      _Version        =   65536
      _ExtentX        =   21378
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
         TabIndex        =   72
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   135
         Width           =   720
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   120
      Top             =   10140
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
      Height          =   6735
      Left            =   120
      TabIndex        =   36
      Top             =   3000
      Width           =   11415
      Begin VB.CommandButton cmdWhatsMissing 
         Appearance      =   0  'Flat
         Caption         =   "What is Missing in Form 7?"
         Height          =   375
         Left            =   2640
         TabIndex        =   85
         Tag             =   "Generate Form 7"
         Top             =   6240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdPrintWF7 
         Appearance      =   0  'Flat
         Caption         =   "Generate Form 7"
         Height          =   375
         Left            =   240
         TabIndex        =   84
         Tag             =   "Generate Form 7"
         Top             =   6240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdPageLeft 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   10560
         Picture         =   "fehswcb.frx":6F3C
         Style           =   1  'Graphical
         TabIndex        =   83
         Tag             =   "Grant All Basic"
         Top             =   50
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtFirmAcct 
         Appearance      =   0  'Flat
         DataField       =   "EC_FIRM_ACCT"
         Enabled         =   0   'False
         Height          =   285
         Left            =   5520
         TabIndex        =   82
         Top             =   1440
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtFirmAcctNo 
         Appearance      =   0  'Flat
         DataField       =   "EC_FIRM_ACCT_NUM"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   81
         Top             =   1440
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.ComboBox comFirmAcctNo 
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
         ItemData        =   "fehswcb.frx":737E
         Left            =   1815
         List            =   "fehswcb.frx":7380
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Tag             =   "10-Type of Employee "
         Top             =   960
         Width           =   3480
      End
      Begin VB.Frame Frame1 
         Caption         =   "First Health Care Provider"
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
         TabIndex        =   50
         Top             =   1860
         Width           =   10275
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
            Left            =   5685
            MaxLength       =   80
            TabIndex        =   9
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
            Left            =   5685
            MaxLength       =   40
            TabIndex        =   7
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
            Left            =   1500
            TabIndex        =   10
            Tag             =   "Email Address"
            Top             =   900
            Width           =   2715
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_PHYSCOD"
            Height          =   285
            Index           =   2
            Left            =   1200
            TabIndex        =   6
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
            Left            =   1500
            TabIndex        =   8
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
            Left            =   1190
            TabIndex        =   11
            Tag             =   "40-Medical Visit"
            Top             =   1215
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
            Left            =   5880
            TabIndex        =   12
            Tag             =   "40-Employer Notified"
            Top             =   930
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
            Left            =   4440
            TabIndex        =   74
            Top             =   975
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
            TabIndex        =   73
            Top             =   1260
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
            TabIndex        =   55
            Top             =   615
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
            Left            =   4410
            TabIndex        =   54
            Top             =   645
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
            Left            =   4410
            TabIndex        =   53
            Top             =   315
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
            TabIndex        =   52
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
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   945
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Second Health Care Provider"
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
         TabIndex        =   43
         Top             =   3270
         Width           =   5080
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
            Left            =   1260
            MaxLength       =   40
            TabIndex        =   14
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
            Left            =   1260
            MaxLength       =   80
            TabIndex        =   44
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
            Left            =   1260
            TabIndex        =   16
            Tag             =   "Email Address"
            Top             =   1590
            Width           =   2715
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_PHYSCOD2"
            Height          =   285
            Index           =   4
            Left            =   945
            TabIndex        =   13
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
            Left            =   1260
            TabIndex        =   15
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
            TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   76
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
            Left            =   150
            TabIndex        =   75
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
            Left            =   130
            TabIndex        =   49
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
            Left            =   130
            TabIndex        =   48
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
            Left            =   130
            TabIndex        =   47
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
            Left            =   130
            TabIndex        =   46
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
            Left            =   130
            TabIndex        =   45
            Top             =   1620
            Width           =   825
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Third Health Care Provider"
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
         Left            =   5310
         TabIndex        =   37
         Top             =   3270
         Width           =   5080
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
            Left            =   1350
            MaxLength       =   80
            TabIndex        =   22
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
            Left            =   1350
            MaxLength       =   40
            TabIndex        =   20
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
            Left            =   1350
            TabIndex        =   23
            Tag             =   "Email Address"
            Top             =   1590
            Width           =   2715
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_PHYSCOD3"
            Height          =   285
            Index           =   5
            Left            =   1035
            TabIndex        =   19
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
            Left            =   1350
            TabIndex        =   21
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
            TabIndex        =   24
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
            TabIndex        =   25
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
            TabIndex        =   78
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
            TabIndex        =   77
            Top             =   1920
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
            Top             =   1620
            Width           =   825
         End
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EC_RATEGRP"
         Height          =   285
         Index           =   3
         Left            =   6405
         TabIndex        =   5
         Tag             =   "00-Rate Group Code"
         Top             =   960
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
         Left            =   6405
         TabIndex        =   4
         Tag             =   "01-Results of Claim - Code"
         Top             =   630
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
         Left            =   1500
         TabIndex        =   3
         Tag             =   "40-Date Claim was closed"
         Top             =   630
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "EC_WCBFDTE"
         Height          =   285
         Index           =   0
         Left            =   6405
         TabIndex        =   2
         Tag             =   "41-Date Filed with WSIB"
         Top             =   300
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
         Top             =   300
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
         TabIndex        =   26
         Tag             =   "00-Name of Hospital"
         Top             =   5340
         Width           =   5475
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Firm / Account #"
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
         Index           =   25
         Left            =   255
         TabIndex        =   80
         Top             =   1020
         Width           =   1200
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
         Left            =   6975
         TabIndex        =   71
         Tag             =   "11-Unique ID of Incident"
         Top             =   30
         Width           =   90
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Incident Number:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   5415
         TabIndex        =   70
         Top             =   30
         Width           =   1470
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Result"
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
         Left            =   5415
         TabIndex        =   69
         Top             =   660
         Width           =   450
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
         Left            =   255
         TabIndex        =   68
         Top             =   660
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
         Left            =   5415
         TabIndex        =   67
         Top             =   360
         Width           =   720
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
         TabIndex        =   66
         Top             =   330
         Width           =   975
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
         TabIndex        =   65
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
         TabIndex        =   62
         Top             =   1560
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
         Left            =   240
         TabIndex        =   61
         Top             =   5400
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
         Left            =   5415
         TabIndex        =   60
         Top             =   990
         Width           =   900
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
         TabIndex        =   59
         Top             =   5760
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
         TabIndex        =   58
         Top             =   5760
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
         TabIndex        =   57
         Top             =   6000
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
         TabIndex        =   56
         Top             =   6000
         Width           =   1095
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
      Left            =   1680
      TabIndex        =   34
      Top             =   10080
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
      Left            =   180
      TabIndex        =   35
      Top             =   10080
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmEHSWCB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim fglbNew
Dim xOClaimNo
Dim xOFiledOn

Private Function chkHSWCB()
Dim SQLQ As String, Msg As String, dd#, X%

chkHSWCB = False

On Error GoTo chkHSWCB_Err
'Jdy 4/10/00
'If Len(medClaimNo) <= 0 Then
'    MsgBox "Claim Number is a required field"
'    medClaimNo.SetFocus
'    Exit Function
'End If

If Len(dlpDate(0).Text) >= 1 Then
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "Filed On date is not a valid date."
        dlpDate(0).SetFocus
        Exit Function
    End If
Else
    MsgBox "Filed On date is required."
    dlpDate(0).SetFocus
    Exit Function
End If

If Len(dlpDate(1).Text) >= 1 Then
    If Not IsDate(dlpDate(1).Text) Then
        MsgBox "File Closed On date is not a valid date."
        dlpDate(1).SetFocus
        Exit Function
    End If
End If

'Release 8.0 - This field is required in Form 7 otherwise the Form cannot be generated
If cmdPrintWF7.Visible = True And comFirmAcctNo.ListIndex = -1 Then
    MsgBox "Firm / Account # is a required field for Form 7"
    comFirmAcctNo.SetFocus
    Exit Function
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

'Ticket #22443 - Remove the mandatory
'Jerry said if First Aid Provided then Physician information cannot be left blank
'Form 7
'If Not IsNull(rsDATA!EC_FAPROVIDED) Then
'    If Len(clpCode(2).Text) < 1 Then
'        MsgBox "When Health Care is Provided on Incident screen the 'Type' code under " & Frame1.Caption & " is a required field"
'        clpCode(2).SetFocus
'        Exit Function
'    End If
'
'    If Len(txtMedInfo(0).Text) < 1 Then
'        MsgBox "When Health Care is Provided on Incident screen the 'Name' under " & Frame1.Caption & " is a required field"
'        txtMedInfo(0).SetFocus
'        Exit Function
'    End If
'
'    If Len(txtMedInfo(1).Text) < 1 Then
'        MsgBox "When Health Care is Provided on Incident screen the 'Address' under " & Frame1.Caption & " is a required field"
'        txtMedInfo(1).SetFocus
'        Exit Function
'    End If
'
'    If Not IsDate(dlpDate(2)) And Trim(dlpDate(2).Text) = "" Then
'        MsgBox "When Health Care is Provided on Incident screen the '" & lblTitle(19).Caption & "' under " & Frame1.Caption & " is a required field"
'        'MsgBox lblTitle(19).Caption & " under '" & Frame1.Caption & "' has not been entered."
'        dlpDate(2).SetFocus
'        Exit Function
'    ElseIf Not IsDate(dlpDate(2)) Then
'        MsgBox "Invalid '" & lblTitle(19).Caption & "' under " & Frame1.Caption
'        dlpDate(2).SetFocus
'        Exit Function
'    End If
'End If

If Len(Trim(dlpDate(2).Text)) > 0 Then
    If Not IsDate(dlpDate(2)) Then
        MsgBox "Invalid '" & lblTitle(19).Caption & "' date under " & Frame1.Caption
        dlpDate(2).SetFocus
        Exit Function
    End If
End If

If Len(Trim(dlpDate(3).Text)) > 0 Then
    If Not IsDate(dlpDate(3)) Then
        MsgBox "Invalid '" & lblTitle(20).Caption & "' date under " & Frame1.Caption
        dlpDate(3).SetFocus
        Exit Function
    End If
End If

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

For X% = 1 To 5
    If Len(clpCode(X%).Text) > 0 And clpCode(X%).Caption = "Unassigned" Then
        If X% = 1 Then MsgBox "Incident Result code must be valid"
        If X% = 2 Or X% = 4 Or X% = 5 Then MsgBox "Health Care Provider Type code must be valid"
        If X% = 3 Then MsgBox "Rate Group code must be valid"
        clpCode(X%).SetFocus
        Exit Function
    End If
Next

chkHSWCB = True

Exit Function

chkHSWCB_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSWCB", "HR_OCC_HEALTH_SAFETY", "edit/Add")
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
Dim X
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
If glbOnTop = "FRMEHSWCB" Then glbOnTop = ""

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

xOClaimNo = medClaimNo
xOFiledOn = dlpDate(0).Text

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

Function cmdOK_Click()
Dim X
Dim xID

On Error GoTo Add_Err

cmdOK_Click = False

If Not chkHSWCB() Then Exit Function

rsDATA.Requery

Call UpdUStats(Me) ' update user's stats (who did it and when)

Call Set_Control("U", Me, rsDATA)

If rsDATA!EC_WCBNBR = "" Then rsDATA!EC_WCBNBR = Null

'Ticket #21463 - If the Claim # changes then update the Form 9 table with the new Claim #.
If Len(xOFiledOn) > 0 Then
    If IsDate(xOFiledOn) Then
        If xOClaimNo <> medClaimNo Or CVDate(xOFiledOn) <> CVDate(dlpDate(0).Text) Then
            Call Update_Form9_Fields(xOClaimNo)
        End If
    End If
End If

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    xID = rsDATA!EC_ID
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    xID = rsDATA!EC_ID
End If
Data1.Refresh
Data1.Recordset.Find "EC_ID=" & xID

cmdOK_Click = True
fglbNew = False

xOClaimNo = medClaimNo.Text

'Call ST_UPD_MODE(True)
Call SET_UP_MODE

X = NextFormIF("Claim/Medical")

Exit Function

Add_Err:
If Err = 3022 Then
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
'frmEHSWCBC.Show
'Unload Me
'End Sub

Function EERetrieve()
Dim SQLQ As String

EERetrieve = False
Screen.MousePointer = HOURGLASS

On Error GoTo EERError
If glbtermopen Then
    SQLQ = "SELECT " & FldList & " FROM Term_HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & "WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "SELECT " & FldList & " FROM HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & "WHERE EC_EMPNBR = " & glbLEE_ID
End If
SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
Data1.RecordSource = SQLQ
Data1.Refresh

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

Private Sub cmdPrintWF7_Click()
    'This will generate the WSIB Form 7
    
    'Save the Claim/Info data first
    If Not cmdOK_Click() Then Exit Sub
    
    If gsAttachment_DB = False Then
        MsgBox "Attachment Database not found to store the Form 7." & vbCrLf & vbCrLf & "Form 7 cannot be generated.", vbExclamation, "Cannot generate Form 7"
        Exit Sub
    End If

    'Call CutePDF_Test
    Call Generate_WSIB_Form7
    
    'Call function to retrieve the original file and fields.
    'Call function to retrieve data to populate the fields
        'Make sure all the data entry is done
        'Report Missing Data
    'Generate the form 7
        'Attach the Concerns document
End Sub

Private Sub cmdWhatsMissing_Click()
    glbF7FirmAcct = IIf(IsNull(Data1.Recordset!EC_FIRM_ACCT), "", Data1.Recordset!EC_FIRM_ACCT)
    glbF7FirmAcctNo = IIf(IsNull(Data1.Recordset!EC_FIRM_ACCT_NUM), "", Data1.Recordset!EC_FIRM_ACCT_NUM)
    glbF7CaseNo = Data1.Recordset!EC_CASE
        
    frmEHSF7WhatsMissing.Show 1
End Sub

Private Sub comFirmAcctNo_Change()
    'txtFirmAcct.Text = Left(comFirmAcctNo.Text, 1)
    'txtFirmAcctNo.Text = Mid(comFirmAcctNo.Text, 3, InStr(3, comFirmAcctNo.Text, ":") - 3)
End Sub

Private Sub comFirmAcctNo_Click()
    If comFirmAcctNo.Text <> "" Then
        txtFirmAcct.Text = Left(comFirmAcctNo.Text, 1)
        txtFirmAcctNo.Text = Mid(comFirmAcctNo.Text, 3, InStr(3, comFirmAcctNo.Text, ":") - 3)
    End If
End Sub

Private Sub cmdPageLeft_Click(Index As Integer)
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

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMEHSWCB"
End Sub

Private Sub Form_GotFocus()
glbOnTop = "FRMEHSWCB"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim xInt As Integer

glbOnTop = "FRMEHSWCB"
If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
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
    Frame2.Top = 2790 + 200
    Frame2.Height = 2595
    lblTitle(21).Visible = True
    lblTitle(22).Visible = True
    dlpDate(4).Visible = True
    dlpDate(4).DataField = "EC_PHYS2_VISIT"
    dlpDate(5).Visible = True
    dlpDate(5).DataField = "EC_PHYS2_NOTIFIED"
    
    Frame3.Caption = "Third Health Professional"
    Frame3.Top = 2790 + 200
    Frame3.Height = 2595
    lblTitle(23).Visible = True
    lblTitle(24).Visible = True
    dlpDate(6).Visible = True
    dlpDate(6).DataField = "EC_PHYS3_VISIT"
    dlpDate(7).Visible = True
    dlpDate(7).DataField = "EC_PHYS3_NOTIFIED"
    
    xInt = 800
    lblTitle(7).Top = 4920 + xInt
    txtMedInfo(2).Top = 4860 + xInt
    'lblTitle(9).Top = 5310 + xInt
    'clpCode(3).Top = 5280 + xInt
    lblUpdateBy.Top = 5280 + xInt
    lblUserDesc.Top = 5280 + xInt
    lblUpdateDate.Top = 5520 + xInt
    lblUpdDateDesc.Top = 5520 + xInt
    
Else
    'For WSIB Form 7 - Ticket #20038
    Frame1.Height = 1545
    lblTitle(19).Visible = True
    lblTitle(19).Caption = "Date Seen"
    lblTitle(20).Visible = True
    dlpDate(2).Visible = True
    dlpDate(2).DataField = "EC_PHYS1_VISIT"
    dlpDate(3).Visible = True
    dlpDate(3).DataField = "EC_PHYS1_NOTIFIED"

    Frame2.Top = 2790 + 700
    Frame3.Top = 2790 + 700

    xInt = 800
    lblTitle(7).Top = 4920 + xInt
    txtMedInfo(2).Top = 4860 + xInt
    lblUpdateBy.Top = 5280 + xInt
    lblUserDesc.Top = 5280 + xInt
    lblUpdateDate.Top = 5520 + xInt
    lblUpdDateDesc.Top = 5520 + xInt
    
'    'Populate with Firm/Account #s
    Call Populate_FirmAccountNo

    'Frame1.Height = 1425
    'Frame2.Height = 1995
    'Frame3.Height = 1995
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

If glbtermopen Then
    cmdPrintWF7.Enabled = False
    cmdWhatsMissing.Enabled = False
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
    If Me.Height >= 10170 Then
        scrControl.Value = 0
        ScrFrame.Top = 3000 '2040
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        If Me.Height < 8000 Then
            scrControl.Max = 5000
        Else
            scrControl.Max = 2500
        End If
        scrControl.Left = Me.Width - scrControl.Width - 240
        If Me.Height - scrControl.Top - 780 > 0 Then
            scrControl.Height = Me.Height - scrControl.Top - 780
        End If
    End If
    
    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 120
    'ScrFrame.Height = Me.ScaleHeight - (scrHScroll.Height + 200)
    If Me.Width >= 10750 Then
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
Set frmEHSWCB = Nothing 'carmen may 00
Call NextForm
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

medClaimNo.Enabled = TF
 clpCode(1).Enabled = TF
 clpCode(2).Enabled = TF
 clpCode(3).Enabled = TF
dlpDate(0).Enabled = TF
dlpDate(1).Enabled = TF
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

Private Sub scrControl_Change()
ScrFrame.Top = 3000 - scrControl.Value
End Sub

Private Sub scrHScroll_Change()
ScrFrame.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
End Sub

Private Sub txtFirmAcctNo_Change()
    Dim X As Integer
    For X = 0 To comFirmAcctNo.ListCount - 1
        If Left(comFirmAcctNo.List(X), InStr(1, comFirmAcctNo.List(X), ":") - 1) = txtFirmAcct.Text & "-" & txtFirmAcctNo.Text Then
            comFirmAcctNo.ListIndex = X
            Exit For
        Else
            comFirmAcctNo.ListIndex = -1
        End If
    Next
    'comFirmAcctNo.List(0) = txtFirmAcct.Text & "-" & txtFirmAcctNo.Text
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
            SQLQ = "SELECT " & FldList & " FROM Term_HR_OCC_HEALTH_SAFETY "
            SQLQ = SQLQ & "WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "SELECT " & FldList & " FROM HR_OCC_HEALTH_SAFETY "
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

'Call Populate_FirmAccountNo

 ' set description for code
 ' set description for code

If Data1.Recordset.RecordCount <> 0 Then
    If Not IsNull(Data1.Recordset("EC_WCBCDTE")) Then
        dlpDate(1).Text = Data1.Recordset("EC_WCBCDTE")
    Else
        dlpDate(1).Text = ""
    End If
End If

'Only for clients who have purchased WSIB Module
If gSec_Inq_HSW7CmpMst And gSec_Inq_HSW7Injury And glbWSIBModule Then
    If Data1.Recordset.RecordCount <> 0 Then
        If IsNull(Data1.Recordset("EC_FORM7")) Then
            lblTitle(25).FontBold = False
            cmdPageLeft(0).Visible = False
            cmdPrintWF7.Visible = False
            cmdWhatsMissing.Visible = False
        Else
            If Data1.Recordset("EC_FORM7") Then
                lblTitle(25).FontBold = True    'Release 8.0 - This field is required in Form 7 otherwise the Form cannot be generated
                cmdPageLeft(0).Visible = True
                cmdPrintWF7.Visible = True
                cmdWhatsMissing.Visible = True
            Else
                lblTitle(25).FontBold = False
                cmdPageLeft(0).Visible = False
                cmdPrintWF7.Visible = False
                cmdWhatsMissing.Visible = False
            End If
        End If
    Else
        lblTitle(25).FontBold = False
        cmdPageLeft(0).Visible = False
        cmdPrintWF7.Visible = False
        cmdWhatsMissing.Visible = False
    End If
End If

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

Private Function FldList()
Dim SQLQ
SQLQ = ""
SQLQ = SQLQ & "EC_ID, EC_COMPNO, EC_EMPNBR, EC_CASE, EC_WCBNBR, "
SQLQ = SQLQ & "EC_WCBFDTE, EC_WCBCDTE, EC_WCBRES, EC_PHYSCOD,"
SQLQ = SQLQ & "EC_PHYSNM, EC_PHYSADDR, EC_HOSP, EC_RATEGRP,"
SQLQ = SQLQ & "EC_LDATE, EC_LTIME, EC_LUSER, EC_PHYSCOD2,"
SQLQ = SQLQ & "EC_PHYSNM2, EC_PHYSADDR2, EC_DOCPHONE,"
SQLQ = SQLQ & "EC_DOCPHONE2, EC_PHYSCOD3, EC_PHYSNM3,"
SQLQ = SQLQ & "EC_PHYS1_EMAIL, EC_PHYS2_EMAIL, EC_PHYS3_EMAIL,"
SQLQ = SQLQ & "EC_PHYSADDR3 , EC_DOCPHONE3"
SQLQ = SQLQ & ",EC_FORM7, EC_FIRM_ACCT_NUM, EC_FIRM_ACCT"
SQLQ = SQLQ & ",EC_PHYS1_VISIT,EC_PHYS1_NOTIFIED, EC_FAPROVIDED"
If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"
If glbLinamar Then
    'SQLQ = SQLQ & ",EC_PHYS1_VISIT,EC_PHYS1_NOTIFIED,EC_PHYS2_VISIT,EC_PHYS2_NOTIFIED,EC_PHYS3_VISIT,EC_PHYS3_NOTIFIED"
    SQLQ = SQLQ & ",EC_PHYS2_VISIT,EC_PHYS2_NOTIFIED,EC_PHYS3_VISIT,EC_PHYS3_NOTIFIED"
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
      SQLQ = "SELECT " & FldList & " FROM Term_HR_OCC_HEALTH_SAFETY "
      SQLQ = SQLQ & "WHERE EC_CASE = " & Data1.Recordset!EC_CASE
      If glbWFC Then
        SQLQ = SQLQ & " AND TERM_SEQ = " & glbTERM_Seq
      End If
      If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
      rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
      SQLQ = "SELECT " & FldList & " FROM HR_OCC_HEALTH_SAFETY "
      SQLQ = SQLQ & "WHERE EC_CASE = " & Data1.Recordset!EC_CASE
      If glbWFC Then
            SQLQ = SQLQ & " AND EC_EMPNBR =" & glbLEE_ID
      End If
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
UpdateRight = gSec_Upd_HSClaimMed
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
        frmEHSWCB.Caption = "Claim / Medical Information - " & Left$(glbLEE_SName, 5)
        frmEHSWCB.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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

Private Sub Populate_FirmAccountNo()
    Dim rsEmprMst As New ADODB.Recordset
    Dim SQLQ As String
    
    comFirmAcctNo.Clear
    SQLQ = "SELECT EY_TRADLEGAL_NAME,EY_FIRM_ACCT,EY_FIRM_ACCT_NUM, EY_RATE_GRP_NUM FROM HR_OHS_COMPANY_MASTER"
    rsEmprMst.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsEmprMst.EOF
        'Since there can be more than one same Firm/Account # but different rate. Bill suggseted to put Rate Group on the dropdown list
        'instead of Trade Name.
        'comFirmAcctNo.AddItem rsEmprMst("EY_FIRM_ACCT") & "-" & rsEmprMst("EY_FIRM_ACCT_NUM") & ":" & rsEmprMst("EY_TRADLEGAL_NAME")
        comFirmAcctNo.AddItem rsEmprMst("EY_FIRM_ACCT") & "-" & rsEmprMst("EY_FIRM_ACCT_NUM") & ":" & rsEmprMst("EY_RATE_GRP_NUM")
        rsEmprMst.MoveNext
    Loop
    rsEmprMst.Close
    Set rsEmprMst = Nothing
    
End Sub

Private Sub CutePDF_Test()
Dim strFldNames As String
Dim objMyForm
Dim ErrorMessage, FName
Dim nCount, nI
Dim nReturn

Set objMyForm = CreateObject("CutePDF.Document")  'Create form object
objMyForm.Initialize ("FS21-010-94171023-00658222") 'Initialize object by serial number of the license


'Open an encrypted PDF form file from an URL with password 'cutepdf'
'If objMyForm.openFile("ftp://www.ftpsite.com/Encrypted_Form.pdf", "cutepdf") = False Then
If objMyForm.openFile(glbIHRREPORTS & "Form7.pdf") = False Then
    ErrorMessage = objMyForm.GetLastError()
End If

'Open App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "WSIBForm7_Fields.TXT" For Output As #1


'Retrieve form fields values
'cmbFields.Clear
'nCount = objMyForm.numberOfFields                      'Get total number of fields in the form
'For nI = 1 To nCount Step 1
'    FName = objMyForm.getFieldName(nI)                 'Get field name first
'    'fValue = objMyForm.getFieldValue(fName)           'Then get field value by name
'    'cmbFields.AddItem FName
'    Print #1, FName
'Next


'Set some value into fields
nReturn = objMyForm.setField("txtworkerojobtitleoccupation", "Application Development + Support")
nReturn = objMyForm.setField("txtemployerfaxtimeinposition", "55")

nReturn = objMyForm.setField("SIN.1", "999")
nReturn = objMyForm.setField("SIN.2", "999")
nReturn = objMyForm.setField("SIN.3", "999")
nReturn = objMyForm.setField("cbtradeunion", "no")  'values = yes/no
nReturn = objMyForm.setField("txtworkerreferencenumber", "12345")
nReturn = objMyForm.setField("txtworkerDOBday", "25")
nReturn = objMyForm.setField("txtworkerDOBMonth", "11")
nReturn = objMyForm.setField("txtworkerDOBYear", "72")
nReturn = objMyForm.setField("Telephone.ACode", "416")
nReturn = objMyForm.setField("Telephone.Main", "816 1223")
nReturn = objMyForm.setField("cbworkersex", "male")   'values = female/male
nReturn = objMyForm.setField("txtworkerDOHDay", "05")
nReturn = objMyForm.setField("txtworkerDOHMonth", "03")
nReturn = objMyForm.setField("txtworkerDOHYear", "03")

nReturn = objMyForm.setField("txtworkerfirstname", "HR Systems Strategies Inc.")
nReturn = objMyForm.setField("txtworkersurname", "Mistry")
nReturn = objMyForm.setField("txtworkerfirstname", "Hemu")
nReturn = objMyForm.setField("txtworkeraddress", "117 Country Club Drive")
nReturn = objMyForm.setField("txtworkercitytown", "King City")
nReturn = objMyForm.setField("txtworkerpostalcode", "K2J 9G4")
nReturn = objMyForm.setField("txtworkerprov", "ON")

    'nReturn = objMyForm.setField("cbfirmaccountno3", "accountnumber")
    nReturn = objMyForm.setField("cbfirmaccountno3", "firm")

'
'Save completed form file into a new PDF file
objMyForm.saveFile ("C:\Form04.pdf")

'Save a duplicate copy
'objMyForm.saveFile ("C:\Form05.pdf")

MsgBox "WSIB Form 7 generation complete."

End Sub

Private Sub Generate_WSIB_Form7()
    Dim rsEMP As New ADODB.Recordset
    Dim rsCompMst As New ADODB.Recordset
    Dim rsHS As New ADODB.Recordset
    Dim rsJOB As New ADODB.Recordset
    Dim rsSal As New ADODB.Recordset
    Dim rsStatCat As New ADODB.Recordset
    Dim rsForm7Sec As New ADODB.Recordset
    Dim RsLang As New ADODB.Recordset
    Dim SQLQ As String
    Dim xHourlyRate As Double
    Dim xSunDate, xSatDate
    Dim xSat As Integer
    Dim xUnionFlg As String
    Dim xPathToSaveIn As String
    Dim xFileName As String
    Dim xFileExtension As String
    Dim xNamePos As String
    Dim xENFlg As Boolean
    Dim xFRFlg As Boolean
    Dim xOTHFlg As Boolean
    Dim xOTHLang As String
    
    Dim objMyForm
    Dim ErrorMessage
    Dim nCount, nI
    Dim nReturn
    Dim xConcDocFound As Boolean
    Dim xWrtnDocFound As Boolean
    Dim xYears As Double
    Dim xMonths As Integer
    
    Screen.MousePointer = HOURGLASS

    Set objMyForm = CreateObject("CutePDF.Document")  'Create form object
    objMyForm.Initialize ("FS21-010-94171023-00658222") 'Initialize object by serial number of the license
    
'    'Test lines -------------------------------------------------------------------------------------
    'Dim fieldValue
    'If objMyForm.openFile("U:\HR Systems VB6\Custom Features 7x\Health & Safety\Form7updated.pdf") = False Then  'Open a PDF file's form
    '    ErrorMessage = objMyForm.GetLastError()
    'End If
'    'For x = 0 To 331
'    '    fieldValue = objMyForm.getFieldName(x)
'    '    Debug.Print x & ". " & fieldValue
'    'Next
'    fieldValue = objMyForm.getFieldValue("cbworkerpreferredlanguage3")
'    fieldValue = objMyForm.getFieldType("TXTKADDITIONAL Text")
'    fieldValue = objMyForm.getFieldValue("TXTKADDITIONAL TEXT")
'    fieldValue = objMyForm.setField("TXTKADDITIONAL TEXT", "abc")
'    fieldValue = objMyForm.getFieldValue("LBLweb notice2")
    'fieldValue = objMyForm.getFieldValue("cbE1")
'    fieldValue = objMyForm.getFieldValue("cbE1rtwtypeofwork")
'    fieldValue = objMyForm.getFieldValue("cbE2Modworkconfirmby")
'    fieldValue = objMyForm.getFieldValue("CBF1providedwithlimitations")
'    fieldValue = objMyForm.getFieldValue("CBF2modifiedworkdiscussed")
'    fieldValue = objMyForm.getFieldValue("CBF3modifiedworkoffered")
'    fieldValue = objMyForm.getFieldValue("CBF3modifiedworkofferedwasit")
'    fieldValue = objMyForm.getFieldValue("CBF3modifiedworkdeclinedattachcopy")
'    fieldValue = objMyForm.getFieldValue("cbF4whoarrangedrtw")
'    fieldValue = objMyForm.getFieldValue("cbH2vacationpaypercheque")
'    fieldValue = objMyForm.getFieldValue("cbtimelastworked")
'    fieldValue = objMyForm.getFieldValue("cbh4timelastworkedfrom")
'    fieldValue = objMyForm.getFieldValue("cbh4timelastworkedto")
'    fieldValue = objMyForm.getFieldValue("cbH7advanceearnings")
'    fieldValue = objMyForm.getFieldValue("cbH7advanceearningsamount")
'    fieldValue = objMyForm.getFieldValue("cbIschedule")
'    'End Test lines ---------------------------------------------------------------------------------
    
    'Open an encrypted PDF form file from an URL with password 'cutepdf'
    'If objMyForm.openFile("ftp://www.ftpsite.com/Encrypted_Form.pdf", "cutepdf") = False Then
    'Open the Form 7 template
    If objMyForm.openFile(glbIHRREPORTS & "Form7.pdf") = False Then
        ErrorMessage = objMyForm.GetLastError()
        Screen.MousePointer = DEFAULT
        MsgBox "Cannot find the template for Form 7. WSIB Form 7 cannot be generated.", vbCritical, "info:HR - Missing Form 7 Template"
        Exit Sub
    End If
    
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME, ED_ADDR1, ED_CITY, ED_PROV, ED_PCODE, ED_ORG,"
    SQLQ = SQLQ & " ED_DOB, ED_PHONE, ED_DOH, ED_SEX, ED_SIN, ED_BUSNBR, ED_EMP, ED_PT, ED_TD1DOL,"
    SQLQ = SQLQ & " ED_PROVAMT"
    SQLQ = SQLQ & " FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsEMP.EOF Then
        'Employee missing in the employee file - cannot generate the WSIB Form 7
        MsgBox "Employee data missing cannot generate the WSIB Form 7.", vbExclamation, "info:HR - Form 7 Generation"
        rsEMP.Close
        Set rsEMP = Nothing
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
    
    SQLQ = "SELECT JH_EMPNBR, JH_JOB, JH_SDATE, JH_DHRS FROM HR_JOB_HISTORY"
    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND JH_CURRENT <> 0"
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsJOB.EOF Then
        'Employee Job missing - cannot generate the WSIB Form 7
        MsgBox "Employee Job data missing cannot generate the WSIB Form 7.", vbExclamation, "info:HR - Form 7 Generation"
        rsJOB.Close
        Set rsJOB = Nothing
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If

    SQLQ = "SELECT SH_EMPNBR, SH_EDATE, SH_SALARY, SH_SALCD, SH_WHRS FROM HR_SALARY_HISTORY"
    SQLQ = SQLQ & " WHERE SH_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND SH_CURRENT <> 0"
    rsSal.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsSal.EOF Then
        'Employee Salary missing - cannot generate the WSIB Form 7
        MsgBox "Employee Salary data missing cannot generate the WSIB Form 7.", vbExclamation, "info:HR - Form 7 Generation"
        rsSal.Close
        Set rsSal = Nothing
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
    
    'Employee Language Skill
    xENFlg = False
    xFRFlg = False
    xOTHFlg = False
    SQLQ = "SELECT EL_EMPNBR, EL_LANG_SPOKEN, EL_LANG_WRITTEN FROM HR_LANGUAGE"
    SQLQ = SQLQ & " WHERE EL_EMPNBR = " & glbLEE_ID
    RsLang.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If RsLang.EOF Then
        'Employee Language Skills missing - cannot generate the WSIB Form 7
        'MsgBox "Employee Language skills data missing cannot generate the WSIB Form 7.", vbExclamation, "info:HR - Form 7 Generation"
        'Jerry asked to default to English
        xENFlg = True
        RsLang.Close
        Set RsLang = Nothing
        'Screen.MousePointer = DEFAULT
        'Exit Sub
    Else
        xENFlg = False
        xFRFlg = False
        xOTHFlg = False
        Do While Not RsLang.EOF
            If Not IsNull(RsLang("EL_LANG_WRITTEN")) Then
                If RsLang("EL_LANG_WRITTEN") = "EN" Then
                    xENFlg = True
                ElseIf RsLang("EL_LANG_WRITTEN") = "FR" Then
                    xFRFlg = True
                Else
                    xOTHFlg = True
                    xOTHLang = GetTABLDesc("EDL1", RsLang("EL_LANG_WRITTEN"))
                End If
            End If
            RsLang.MoveNext
        Loop
        RsLang.Close
        Set RsLang = Nothing
    End If
    
    SQLQ = "SELECT * FROM HR_OCC_HEALTH_SAFETY"
    SQLQ = SQLQ & " WHERE EC_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND EC_CASE =" & Data1.Recordset!EC_CASE
    rsHS.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsHS.EOF Then
        'Employee Health & Safety data missing - cannot generate the WSIB Form 7
        MsgBox "Employee Health & Safety data missing cannot generate the WSIB Form 7.", vbExclamation, "info:HR - Form 7 Generation"
        rsHS.Close
        Set rsHS = Nothing
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HR_OHS_FORM7_SECTIONS"
    SQLQ = SQLQ & " WHERE F7_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND F7_CASE =" & Data1.Recordset!EC_CASE
    rsForm7Sec.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsForm7Sec.EOF Then
        'Employee Health & Safety data missing - cannot generate the WSIB Form 7
        MsgBox "Employee Health & Safety data in the Additional Form 7 Sections is missing cannot generate the WSIB Form 7.", vbExclamation, "info:HR - Form 7 Generation"
        rsForm7Sec.Close
        Set rsForm7Sec = Nothing
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HR_OHS_COMPANY_MASTER"
    SQLQ = SQLQ & " WHERE EY_FIRM_ACCT = '" & Data1.Recordset!EC_FIRM_ACCT & "'"
    SQLQ = SQLQ & " AND EY_FIRM_ACCT_NUM ='" & Data1.Recordset!EC_FIRM_ACCT_NUM & "'"
    rsCompMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsCompMst.EOF Then
        'Employee Health & Safety data missing - cannot generate the WSIB Form 7
        MsgBox "Employer data missing cannot generate the WSIB Form 7.", vbExclamation, "info:HR - Form 7 Generation"
        rsCompMst.Close
        Set rsCompMst = Nothing
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
    
    SQLQ = "SELECT SC_WORKER_TYPE, SC_OTHER_DESC FROM HR_EMPLOYEE_MATRIX"
    SQLQ = SQLQ & " WHERE SC_EMP = '" & rsEMP("ED_EMP") & "'"
    SQLQ = SQLQ & " AND SC_PT = '" & rsEMP("ED_PT") & "'"
    rsStatCat.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsStatCat.EOF Then
        'Employee Health & Safety data missing - cannot generate the WSIB Form 7
        MsgBox "Employee Type Matrix missing cannot generate the WSIB Form 7.", vbExclamation, "info:HR - Form 7 Generation"
        rsStatCat.Close
        Set rsStatCat = Nothing
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
           
    MDIMain.panHelp(1).Caption = "Please wait, generating Form 7...."
    MDIMain.panHelp(0).FloodType = 1
       
    MDIMain.panHelp(0).FloodPercent = 5

    'Start assigning values to Form 7
    
    'A. Worker Information
    If Not IsNull(rsHS("EC_JOBDESC")) Then
        nReturn = objMyForm.setField("txtworkerojobtitleoccupation", rsHS("EC_JOBDESC"))
    Else
        nReturn = objMyForm.setField("txtworkerojobtitleoccupation", GetJobData(rsJOB("JH_JOB"), "JB_DESCR"))
    End If
    
    'Ticket #22838 - Angela said the formulat should be same as the one on Status/Dates screen to calculate the year/months
    'of service. I am going to change to that for now as no one else had noticed this change below. So the new change will
    'based on Status/Dates screen and if client wants something else then we will serial # control it.

    'Since we are now storing the Job Start Date incase of an older job - Ticket #21550
    'nReturn = objMyForm.setField("txtemployerfaxtimeinposition", DateDiff("yyyy", rsJOB("JH_SDATE"), Now()) & " year(s)")
    'Ticket #22838
    'nReturn = objMyForm.setField("txtemployerfaxtimeinposition", DateDiff("yyyy", rsHS("EC_JBSDATE"), Now()) & " year(s)")
'    xMonths = DateDiff("m", rsHS("EC_JBSDATE"), Now())
'    xYears = Int(xMonths / 12)
'    If xYears = 0 And xMonths = 0 Then
'        nReturn = objMyForm.setField("txtemployerfaxtimeinposition", DateDiff("w", rsHS("EC_JBSDATE"), Now()) & " weeks")
'    ElseIf xYears = 0 And xMonths > 0 Then
'        nReturn = objMyForm.setField("txtemployerfaxtimeinposition", DateDiff("m", rsHS("EC_JBSDATE"), Now()) & IIf(xMonths < 2, " month", " months"))
'    Else
'        xMonths = xMonths - (xYears * 12)
'        If xMonths = 0 Then
'            nReturn = objMyForm.setField("txtemployerfaxtimeinposition", xYears & IIf(xYears < 2, " year", " years"))
'        Else
'            nReturn = objMyForm.setField("txtemployerfaxtimeinposition", xYears & "." & xMonths & " years")
'        End If
'    End If
    'nReturn = objMyForm.setField("txtemployerfaxtimeinposition", Round(DateDiff("m", rsHS("EC_JBSDATE"), Now()) / 12, 1) & " year(s)")
    
    'As per Status/Dates screen - formula to calculate the years of service.
    If Not IsNull(rsHS("EC_JBSDATE")) Then
        xYears = DateDiff("d", CVDate(rsHS("EC_JBSDATE")), Date)
        xYears = Round(xYears / 365, 1)
        If xYears < 1 Then
            nReturn = objMyForm.setField("txtemployerfaxtimeinposition", xYears & " year")
        ElseIf xYears = 1 Then
            nReturn = objMyForm.setField("txtemployerfaxtimeinposition", xYears & " year")
        Else
            nReturn = objMyForm.setField("txtemployerfaxtimeinposition", xYears & " years")
        End If
    End If
    
    nReturn = objMyForm.setField("SIN.1", Mid(rsEMP("ED_SIN"), 1, 3))
    nReturn = objMyForm.setField("SIN.2", Mid(rsEMP("ED_SIN"), 4, 3))
    nReturn = objMyForm.setField("SIN.3", Mid(rsEMP("ED_SIN"), 7, 3))
    
    nReturn = objMyForm.setField("txtworkersurname", rsEMP("ED_SURNAME"))
    nReturn = objMyForm.setField("txtworkerfirstname", rsEMP("ED_FNAME"))
    nReturn = objMyForm.setField("txtworkeraddress", rsEMP("ED_ADDR1"))
    nReturn = objMyForm.setField("txtworkercitytown", rsEMP("ED_CITY"))
    nReturn = objMyForm.setField("txtworkerprov", rsEMP("ED_PROV"))
    nReturn = objMyForm.setField("txtworkerpostalcode", Mid(Replace(rsEMP("ED_PCODE"), " ", ""), 1, 3) & " " & Mid(Replace(rsEMP("ED_PCODE"), " ", ""), 4, 3))
    
    xUnionFlg = IIf(IsNull(GetCode_Data("EDOR", rsEMP("ED_ORG"), "TB_USR1", "0")), "", GetCode_Data("EDOR", rsEMP("ED_ORG"), "TB_USR1", "0"))
    If xUnionFlg = "0" Or xUnionFlg = "" Then
        nReturn = objMyForm.setField("cbtradeunion", "no")  'values = yes/no
    Else
        nReturn = objMyForm.setField("cbtradeunion", "Yes")  'values = yes/no
    End If
    nReturn = objMyForm.setField("txtworkerreferencenumber", glbLEE_ID)
    nReturn = objMyForm.setField("txtworkerDOBday", Day(rsEMP("ED_DOB")))
    nReturn = objMyForm.setField("txtworkerDOBMonth", month(rsEMP("ED_DOB")))
    nReturn = objMyForm.setField("txtworkerDOBYear", Right(Year(rsEMP("ED_DOB")), 2))
    nReturn = objMyForm.setField("Telephone.ACode", Mid(Replace(Replace(Replace(Replace(rsEMP("ED_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 1, 3))
    nReturn = objMyForm.setField("Telephone.Main", Mid(Replace(Replace(Replace(Replace(rsEMP("ED_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 4, 3) & " " & Mid(Replace(Replace(Replace(Replace(rsEMP("ED_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 7, 4))
    nReturn = objMyForm.setField("cbworkersex", IIf(rsEMP("ED_SEX") = "M", "male", IIf(rsEMP("ED_SEX") = "F", "female", "n"))) 'values = female/male
    nReturn = objMyForm.setField("txtworkerDOHDay", Day(rsEMP("ED_DOH")))
    nReturn = objMyForm.setField("txtworkerDOHMonth", month(rsEMP("ED_DOH")))
    nReturn = objMyForm.setField("txtworkerDOHYear", Right(Year(rsEMP("ED_DOH")), 2))
    
    'Preferred Language
    If xENFlg Then
        nReturn = objMyForm.setField("cbworkerpreferredlanguage3", "Anglais")
    ElseIf xFRFlg Then
        nReturn = objMyForm.setField("cbworkerpreferredlanguage3", "Francais")
    ElseIf xOTHFlg Then
        nReturn = objMyForm.setField("cbworkerpreferredlanguage3", "Autre")
        nReturn = objMyForm.setField("txtworkerotherlanguage", xOTHLang)
    End If
    
    MDIMain.panHelp(0).FloodPercent = 10
    
    'B. Employer Information
    nReturn = objMyForm.setField("txtlegalremployername", rsCompMst("EY_TRADLEGAL_NAME"))
    If rsCompMst("EY_FIRM_ACCT") = "A" Then
        nReturn = objMyForm.setField("cbfirmaccountno3", "no")  'works now
    ElseIf rsCompMst("EY_FIRM_ACCT") = "F" Then
        nReturn = objMyForm.setField("cbfirmaccountno3", "Yes") 'works now
    End If
    nReturn = objMyForm.setField("Employer Firm Num", rsCompMst("EY_FIRM_ACCT_NUM"))
    nReturn = objMyForm.setField("txtemployermailingaddress", rsCompMst("EY_MAIL_ADDRESS"))
    nReturn = objMyForm.setField("textemployercity", rsCompMst("EY_CITY"))
    nReturn = objMyForm.setField("txtemployerprov", rsCompMst("EY_PROV"))
    nReturn = objMyForm.setField("txtemployerpostalcode", rsCompMst("EY_PCODE"))
    nReturn = objMyForm.setField("txtemployertelephoneareacode", Mid(Replace(Replace(Replace(Replace(rsCompMst("EY_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 1, 3))
    nReturn = objMyForm.setField("txtemployertelephonemain", Mid(Replace(Replace(Replace(Replace(rsCompMst("EY_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 4, 3) & " " & Mid(Replace(Replace(Replace(Replace(rsCompMst("EY_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 7, 4))
    If Not IsNull(rsCompMst("EY_FAX")) Then
        nReturn = objMyForm.setField("txtemployerfaxareacode", Mid(Replace(Replace(Replace(Replace(rsCompMst("EY_FAX"), ")", ""), "(", ""), " ", ""), "-", ""), 1, 3))
        nReturn = objMyForm.setField("txtemployerfaxmain", Mid(Replace(Replace(Replace(Replace(rsCompMst("EY_FAX"), ")", ""), "(", ""), " ", ""), "-", ""), 4, 3) & " " & Mid(Replace(Replace(Replace(Replace(rsCompMst("EY_FAX"), ")", ""), "(", ""), " ", ""), "-", ""), 7, 4))
    End If
    If Not IsNull(rsCompMst("EY_RATE_GRP_NUM")) Then
        nReturn = objMyForm.setField("txtemployertatenumber", rsCompMst("EY_RATE_GRP_NUM"))
    End If
    If Not IsNull(rsCompMst("EY_CLASS_UNIT_CODE")) Then
        nReturn = objMyForm.setField("txtemployerclassificationunitnumber", rsCompMst("EY_CLASS_UNIT_CODE"))
    End If
    nReturn = objMyForm.setField("txtemployerbusinessactivity", rsCompMst("EY_BUSINESS_DESC"))
    
    If Not IsNull(rsCompMst("EY_WKER_GRT_20")) Then
        If rsCompMst("EY_WKER_GRT_20") <> 0 Then
            nReturn = objMyForm.setField("cbmorthan20 3", "yes")
        Else
            nReturn = objMyForm.setField("cbmorthan20 3", "no")
        End If
    End If
    
    If Not IsNull(rsCompMst("EY_WKER_BRNC_ADDR")) Then
        nReturn = objMyForm.setField("txtalternateemployeraddress", rsCompMst("EY_WKER_BRNC_ADDR"))
    End If
    If Not IsNull(rsCompMst("EY_WKER_BRNC_CITY")) Then
        nReturn = objMyForm.setField("txtalternateemployercity", rsCompMst("EY_WKER_BRNC_CITY"))
    End If
    If Not IsNull(rsCompMst("EY_WKER_BRNC_PROV")) Then
        nReturn = objMyForm.setField("txtalternateemployerprov", rsCompMst("EY_WKER_BRNC_PROV"))
    End If
    If Not IsNull(rsCompMst("EY_WKER_BRNC_PCODE")) Then
        nReturn = objMyForm.setField("txtalternateemployerpostalcode", rsCompMst("EY_WKER_BRNC_PCODE"))
    End If
    If Not IsNull(rsCompMst("EY_WKER_BRNC_PHONE")) Then
        nReturn = objMyForm.setField("txtalternateemployertelephoneareacode", Mid(Replace(Replace(Replace(Replace(rsCompMst("EY_WKER_BRNC_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 1, 3))
        nReturn = objMyForm.setField("txtalternateemployertelephonemain", Mid(Replace(Replace(Replace(Replace(rsCompMst("EY_WKER_BRNC_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 4, 3) & " " & Mid(Replace(Replace(Replace(Replace(rsCompMst("EY_WKER_BRNC_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 7, 4))
    End If
    
    MDIMain.panHelp(0).FloodPercent = 15
    
    'C. Accident/Illness Dates and Details
    nReturn = objMyForm.setField("txtdoinjuryday", Day(rsHS("EC_OCCDATE")))
    nReturn = objMyForm.setField("txtdoinjurymonth", month(rsHS("EC_OCCDATE")))
    nReturn = objMyForm.setField("txtdoinjuryyear", Right(Year(rsHS("EC_OCCDATE")), 2))
    
    If Not IsNull(rsHS("EC_OCCTM")) Then
        If Not IsNull(rsHS("EC_OCCTM_FORMAT")) And rsHS("EC_OCCTM_FORMAT") <> "" Then
            If UCase(rsHS("EC_OCCTM_FORMAT")) = "AM" Then
                nReturn = objMyForm.setField("txttimeofinjuryAM", rsHS("EC_OCCTM"))
                nReturn = objMyForm.setField("cbtimeofinjury", "AM")
            ElseIf UCase(rsHS("EC_OCCTM_FORMAT")) = "PM" Then
                nReturn = objMyForm.setField("txttimeofinjuryPM", rsHS("EC_OCCTM"))
                nReturn = objMyForm.setField("cbtimeofinjury", "PM")
            End If
        Else
            If Val(Left(rsHS("EC_OCCTM"), 2)) < 12 Then
                nReturn = objMyForm.setField("txttimeofinjuryAM", rsHS("EC_OCCTM"))
                nReturn = objMyForm.setField("cbtimeofinjury", "AM")
            Else
                nReturn = objMyForm.setField("txttimeofinjuryPM", rsHS("EC_OCCTM"))
                nReturn = objMyForm.setField("cbtimeofinjury", "PM")
            End If
        End If
    End If
    
    'Check this how it appears on the pdf form
    nReturn = objMyForm.setField("txtdatereportedday", Day(rsHS("EC_DATENOT")))
    nReturn = objMyForm.setField("txtdatereportedmonth", month(rsHS("EC_DATENOT")))
    nReturn = objMyForm.setField("txtdatereportedyear", Right(Year(rsHS("EC_DATENOT")), 2))
    
    If Not IsNull(rsHS("EC_TIMNOT")) Then
        If Not IsNull(rsHS("EC_TIMNOT_FORMAT")) And rsHS("EC_TIMNOT_FORMAT") <> "" Then
            If UCase(rsHS("EC_TIMNOT_FORMAT")) = "AM" Then
                nReturn = objMyForm.setField("txttiemreportedAM", rsHS("EC_TIMNOT"))
                nReturn = objMyForm.setField("cbtimereported", "Yes")   'works now
            ElseIf UCase(rsHS("EC_TIMNOT_FORMAT")) = "PM" Then
                nReturn = objMyForm.setField("txttiemreportedAMPM", rsHS("EC_TIMNOT"))
                nReturn = objMyForm.setField("cbtimereported", "PM")    'works now
            End If
        Else
            If Val(Left(rsHS("EC_TIMNOT"), 2)) < 12 Then
                nReturn = objMyForm.setField("txttiemreportedAM", rsHS("EC_TIMNOT"))
                nReturn = objMyForm.setField("cbtimereported", "Yes")   'works now
            Else
                nReturn = objMyForm.setField("txttiemreportedAMPM", rsHS("EC_TIMNOT"))
                nReturn = objMyForm.setField("cbtimereported", "PM")    'works now
            End If
        End If
    End If
    
    If Not IsNull(rsHS("EC_EMPNOT")) Then
        'Because the second line allows 12chrs only but the first 45chrs, I am combining both the Name and Position and then splitting by
        '45:12chrs so all the Name and Position is filled in.
        xNamePos = ""
        xNamePos = GetEmpData(rsHS("EC_EMPNOT"), "ED_FNAME") & " " & GetEmpData(rsHS("EC_EMPNOT"), "ED_SURNAME") & ", " & GetJobData(GetJHData(rsHS("EC_EMPNOT"), "JH_JOB", ""), "JB_DESCR")
        'nReturn = objMyForm.setField("txtaccidentreportedto1", GetEmpData(rsHS("EC_EMPNOT"), "ED_FNAME") & " " & GetEmpData(rsHS("EC_EMPNOT"), "ED_SURNAME"))
        'nReturn = objMyForm.setField("txtaccidentreportedto2", GetJobData(GetJHData(rsHS("EC_EMPNOT"), "JH_JOB", ""), "JB_DESCR"))
        nReturn = objMyForm.setField("txtaccidentreportedto1", Left(xNamePos, 45))
        nReturn = objMyForm.setField("txtaccidentreportedto2", Mid(xNamePos, 46, 12))
        
        nReturn = objMyForm.setField("txtaccidentreportedtotelephoneareacode", Mid(Replace(Replace(Replace(Replace(GetEmpData(rsHS("EC_EMPNOT"), "ED_BUSNBR"), ")", ""), "(", ""), " ", ""), "-", ""), 1, 3))
        nReturn = objMyForm.setField("txtaccidentreportedtotelephone", Mid(Replace(Replace(Replace(Replace(GetEmpData(rsHS("EC_EMPNOT"), "ED_BUSNBR"), ")", ""), "(", ""), " ", ""), "-", ""), 4, 3) & " " & Mid(Replace(Replace(Replace(Replace(GetEmpData(rsHS("EC_EMPNOT"), "ED_BUSNBR"), ")", ""), "(", ""), " ", ""), "-", ""), 7, 4))
        nReturn = objMyForm.setField("txtaccidentreportedtotelephoneextension", Mid(Replace(Replace(Replace(Replace(GetEmpData(rsHS("EC_EMPNOT"), "ED_BUSNBR"), ")", ""), "(", ""), " ", ""), "-", ""), 11))
    End If
    
    'Classification Code Table Name 'ECCL'. Fixed codes = SUDN,GOVT,OCCD,FATA - Classification - Incident Data
    If Not IsNull(rsHS("EC_CLASS")) Then
        If rsHS("EC_CLASS") = "SUDN" Then   'Sudden Specific Event/Occurence
            nReturn = objMyForm.setField("cb3wastheaccident sudden", "sudden specific event/occurence") 'works now
        ElseIf rsHS("EC_CLASS") = "GOVT" Then   'Gradually Occurring Over Time
            nReturn = objMyForm.setField("cb3wastheaccident gradually", "gradually occuring over time") 'works now
        ElseIf rsHS("EC_CLASS") = "OCCD" Then   'Occupational Disease
            nReturn = objMyForm.setField("cb3wastheaccident occdis", "occupational disease") 'works
        ElseIf rsHS("EC_CLASS") = "FATA" Then   'Fatality
            nReturn = objMyForm.setField("cb3wastheaccident fatality", "fatality")  'works
        End If
    End If
    
    'Type of Illness/Incident, Table Name 'ECTY' - Type - Incident Data
    'Fixed codes = STRU, OVEX, REPT, FIRE, FALL, HRMF, ASLT, SLIP, MVHI
    If Not IsNull(rsHS("EC_TYPE")) Then
        If rsHS("EC_TYPE") = "STRU" Then
            nReturn = objMyForm.setField("cb4typeofaccident struck", "struck/caught")   'works
        ElseIf rsHS("EC_TYPE") = "OVEX" Then
            nReturn = objMyForm.setField("cb4typeofaccident overexertion", "overexertion")  'works
        ElseIf rsHS("EC_TYPE") = "REPT" Then
            nReturn = objMyForm.setField("cb4typeofaccident repetition", "repetition")  'works
        ElseIf rsHS("EC_TYPE") = "FIRE" Then
            nReturn = objMyForm.setField("cb4typeofaccident fire", "fire/explosion") 'works
        ElseIf rsHS("EC_TYPE") = "FALL" Then
            nReturn = objMyForm.setField("cb4typeofaccident fall", "fall")  'works
        ElseIf rsHS("EC_TYPE") = "HRMF" Then
            nReturn = objMyForm.setField("cb4typeofaccident harmful", "harmful substances/environmental") 'works
        ElseIf rsHS("EC_TYPE") = "ASLT" Then
            nReturn = objMyForm.setField("cb4typeofaccident assault", "assault")    'works
        ElseIf rsHS("EC_TYPE") = "SLIP" Then
            nReturn = objMyForm.setField("cb4typeofaccident slip", "slip/trip") 'works
        ElseIf rsHS("EC_TYPE") = "MVHI" Then
            nReturn = objMyForm.setField("cb4typeofaccident motor veh", "motor vehicl incident") 'worked now
        Else
            nReturn = objMyForm.setField("cb4typeofaccident other", "other")    'works
            nReturn = objMyForm.setField("txt4typeofaccident other", GetTABLDesc("ECTY", rsHS("EC_TYPE")))
        End If
    End If
    
    'Area of Injury
    If Not IsNull(rsHS("EC_HEAD")) Then
        If rsHS("EC_HEAD") = "HEAD" Then
            nReturn = objMyForm.setField("cbareaofinjury head", "head")
        End If
    End If
    If Not IsNull(rsHS("EC_FACE")) Then
        If rsHS("EC_FACE") = "FACE" Then
            nReturn = objMyForm.setField("cbareaofinjury face", "face")
        End If
    End If
            
    If Not IsNull(rsHS("EC_EYES")) Then
        If rsHS("EC_EYES") = "EYE" Then
            nReturn = objMyForm.setField("cbareaofinjury eyes", "eye(s)")
        End If
    End If
            
    If Not IsNull(rsHS("EC_EARS")) Then
        If rsHS("EC_EARS") = "EAR" Then
            nReturn = objMyForm.setField("cbareaofinjury ears", "ear(s)")
        End If
    End If
    
    If Not IsNull(rsHS("EC_OTHER")) Then
        If Len(rsHS("EC_OTHER")) > 0 Then
            nReturn = objMyForm.setField("cbareaofinjury other", "other")
        End If
    End If
            
    If Not IsNull(rsHS("EC_OTHER")) Then
        If Len(rsHS("EC_OTHER")) > 0 Then
            nReturn = objMyForm.setField("txtareaofinjury other desc", GetTABLDesc("ECBS", rsHS("EC_OTHER")))
        End If
    End If
            
            
    If Not IsNull(rsHS("EC_TEETH")) Then
        If rsHS("EC_TEETH") = "TETH" Then
            nReturn = objMyForm.setField("cbareaofinjury teeth", "teeth")
        End If
    End If
            
    If Not IsNull(rsHS("EC_NECK")) Then
        If rsHS("EC_NECK") = "NECK" Then
            nReturn = objMyForm.setField("cbareaofinjury neck", "neck")
        End If
    End If
            
    If Not IsNull(rsHS("EC_CHEST")) Then
        If rsHS("EC_CHEST") = "CHST" Then
            nReturn = objMyForm.setField("cbareaofinjury chest", "chest")
        End If
    End If
            
            
    If Not IsNull(rsHS("EC_UPPER_BACK")) Then
        If rsHS("EC_UPPER_BACK") = "UBCK" Then
            nReturn = objMyForm.setField("cbareaofinjury upperback", "upper back")
        End If
    End If
            
    If Not IsNull(rsHS("EC_LOWER_BACK")) Then
        If rsHS("EC_LOWER_BACK") = "LBCK" Then
            nReturn = objMyForm.setField("cbareaofinjury lower back", "lower back")
        End If
    End If
            
    If Not IsNull(rsHS("EC_ABDOMEN")) Then
        If rsHS("EC_ABDOMEN") = "ABMN" Then
            nReturn = objMyForm.setField("cbareaofinjury abdomen", "abdomen")
        End If
    End If
            
    If Not IsNull(rsHS("EC_PELVIS")) Then
        If rsHS("EC_PELVIS") = "PLVS" Then
            nReturn = objMyForm.setField("cbareaofinjury  pelvis", "pelvis")
        End If
    End If
            
            
    If Not IsNull(rsHS("EC_RGT_SHOULDER")) Then
        If rsHS("EC_RGT_SHOULDER") = "SHDR" Then
            nReturn = objMyForm.setField("cbareaofinjury shoulder right", "shoulder right")
        End If
    End If
            
    If Not IsNull(rsHS("EC_LFT_SHOULDER")) Then
        If rsHS("EC_LFT_SHOULDER") = "SHDL" Then
            nReturn = objMyForm.setField("cbareaofinjury shoulder left", "shoulder left")
        End If
    End If
            
    If Not IsNull(rsHS("EC_RGT_ARM")) Then
        If rsHS("EC_RGT_ARM") = "ARMR" Then
            nReturn = objMyForm.setField("cbareaofinjury arm right", "arm right")
        End If
    End If
            
    If Not IsNull(rsHS("EC_LFT_ARM")) Then
        If rsHS("EC_LFT_ARM") = "ARML" Then
            nReturn = objMyForm.setField("cbareaofinjury arm left", "arm left")
        End If
    End If
            
    If Not IsNull(rsHS("EC_RGT_ELBOW")) Then
        If rsHS("EC_RGT_ELBOW") = "ELBR" Then
            nReturn = objMyForm.setField("cbareaofinjury elbow right", "elbow right")
        End If
    End If
            
    If Not IsNull(rsHS("EC_LFT_ELBOW")) Then
        If rsHS("EC_LFT_ELBOW") = "ELBL" Then
            nReturn = objMyForm.setField("cbareaofinjury elbow left", "elbow left")
        End If
    End If
            
    If Not IsNull(rsHS("EC_RGT_FOREARM")) Then
        If rsHS("EC_RGT_FOREARM") = "FAMR" Then
            nReturn = objMyForm.setField("cbareaofinjury forearm right", "forearm right")
        End If
    End If
            
    If Not IsNull(rsHS("EC_LFT_FOREARM")) Then
        If rsHS("EC_LFT_FOREARM") = "FAML" Then
            nReturn = objMyForm.setField("cbareaofinjury forearm left", "forearm left")
        End If
    End If
            
            
    If Not IsNull(rsHS("EC_RGT_WRIST")) Then
        If rsHS("EC_RGT_WRIST") = "WRTR" Then
            nReturn = objMyForm.setField("cbareaofinjury wrist right", "wrist right")
        End If
    End If
            
    If Not IsNull(rsHS("EC_LFT_WRIST")) Then
        If rsHS("EC_LFT_WRIST") = "WRTL" Then
            nReturn = objMyForm.setField("cbareaofinjury wrist left", "wrist left")
        End If
    End If
            
    If Not IsNull(rsHS("EC_RGT_HAND")) Then
        If rsHS("EC_RGT_HAND") = "HNDR" Then
            nReturn = objMyForm.setField("cbareaofinjury hand right", "hand right")
        End If
    End If
            
    If Not IsNull(rsHS("EC_LFT_HAND")) Then
        If rsHS("EC_LFT_HAND") = "HNDL" Then
            nReturn = objMyForm.setField("cbareaofinjury hand left", "hand left")
        End If
    End If
            
    If Not IsNull(rsHS("EC_RGT_FINGER")) Then
        If rsHS("EC_RGT_FINGER") = "FNGR" Then
            nReturn = objMyForm.setField("cbareaofinjury fingers right", "finger(s) right")
        End If
    End If
            
    If Not IsNull(rsHS("EC_LFT_FINGER")) Then
        If rsHS("EC_LFT_FINGER") = "FNGL" Then
            nReturn = objMyForm.setField("cbareaofinjury fingers left", "finger(s) left")
        End If
    End If
            
            
    If Not IsNull(rsHS("EC_RGT_HIP")) Then
        If rsHS("EC_RGT_HIP") = "HIPR" Then
            nReturn = objMyForm.setField("cbareaofinjury hip right", "hip right")
        End If
    End If
            
    If Not IsNull(rsHS("EC_LFT_HIP")) Then
        If rsHS("EC_LFT_HIP") = "HIPL" Then
            nReturn = objMyForm.setField("cbareaofinjury hip left", "hip left")
        End If
    End If
            
    If Not IsNull(rsHS("EC_RGT_THIGH")) Then
        If rsHS("EC_RGT_THIGH") = "THGR" Then
            nReturn = objMyForm.setField("cbareaofinjury thigh right", "thigh right")
        End If
    End If
            
    If Not IsNull(rsHS("EC_LFT_THIGH")) Then
        If rsHS("EC_LFT_THIGH") = "THGL" Then
            nReturn = objMyForm.setField("cbareaofinjury thigh left", "thigh left")
        End If
    End If
            
    If Not IsNull(rsHS("EC_RGT_KNEE")) Then
        If rsHS("EC_RGT_KNEE") = "KNER" Then
            nReturn = objMyForm.setField("cbareaofinjury knee right", "knee right")
        End If
    End If
            
    If Not IsNull(rsHS("EC_LFT_KNEE")) Then
        If rsHS("EC_LFT_KNEE") = "KNEL" Then
            nReturn = objMyForm.setField("cbareaofinjury knee left", "knee left")
        End If
    End If
            
    If Not IsNull(rsHS("EC_RGT_LOWER_LEG")) Then
        If rsHS("EC_RGT_LOWER_LEG") = "LLGR" Then
            nReturn = objMyForm.setField("cbareaofinjury lower leg right", "lower leg right")
        End If
    End If
            
    If Not IsNull(rsHS("EC_LFT_LOWER_LEG")) Then
        If rsHS("EC_LFT_LOWER_LEG") = "LLGL" Then
            nReturn = objMyForm.setField("cbareaofinjury lower leg left", "lower leg left")
        End If
    End If
            
            
    If Not IsNull(rsHS("EC_RGT_ANKLE")) Then
        If rsHS("EC_RGT_ANKLE") = "ANKR" Then
            nReturn = objMyForm.setField("cbareaofinjury ankle right", "ankle right")
        End If
    End If
            
    If Not IsNull(rsHS("EC_LFT_ANKLE")) Then
        If rsHS("EC_LFT_ANKLE") = "ANKL" Then
            nReturn = objMyForm.setField("cbareaofinjury ankle left", "ankle left")
        End If
    End If
            
    If Not IsNull(rsHS("EC_RGT_FOOT")) Then
        If rsHS("EC_RGT_FOOT") = "FOTR" Then
            nReturn = objMyForm.setField("cbareaofinjury foot right", "foot right")
        End If
    End If
            
    If Not IsNull(rsHS("EC_LFT_FOOT")) Then
        If rsHS("EC_LFT_FOOT") = "FOTL" Then
            nReturn = objMyForm.setField("cbareaofinjury foot left", "foot left")
        End If
    End If
            
    If Not IsNull(rsHS("EC_RGT_TOES")) Then
        If rsHS("EC_RGT_TOES") = "TOER" Then
            nReturn = objMyForm.setField("cbareaofinjury toes right", "toe(s) right")
        End If
    End If
            
    If Not IsNull(rsHS("EC_LFT_TOES")) Then
        If rsHS("EC_LFT_TOES") = "TOEL" Then
            nReturn = objMyForm.setField("cbareaofinjury toes left", "toe(s) left")
        End If
    End If
    
    If Not IsNull(rsHS("EC_COMMENTS")) Then
        nReturn = objMyForm.setField("txt6describewhathappened", rsHS("EC_COMMENTS"))
    End If
    
    
    '7. Did the accident/illness happen on the employer's premises?
    If Not IsNull(rsHS("EC_PREMISES")) Then
        If rsHS("EC_PREMISES") <> 0 Then
            nReturn = objMyForm.setField("cb7accidenthappen", "Yes")
            'nReturn = objMyForm.setField("txt7accidenthappen", rsHS("EC_EMP_PREMISES"))
            'nReturn = objMyForm.setField("txt7accidenthappen", GetTABLDesc("ECPA", rsHS("EC_AREA")))
        Else
            nReturn = objMyForm.setField("cb7accidenthappen", "No")
        End If
        
        'Premises can be entered even if No is selected.
        If Not IsNull(rsHS("EC_EMP_PREMISES")) Then
            nReturn = objMyForm.setField("txt7accidenthappen", rsHS("EC_EMP_PREMISES"))
        End If
    End If
    
    '8. Did the accident/illness happen outside the Province of Ontario
    If Not IsNull(rsHS("EC_OUTSIDE_PROV")) Then
        If rsHS("EC_OUTSIDE_PROV") <> 0 Then
            nReturn = objMyForm.setField("cb8accidenthappenoop", "Yes")
            nReturn = objMyForm.setField("txt8accidenthappenoop", rsHS("EC_OUTSIDE_CITY"))
        Else
            nReturn = objMyForm.setField("cb8accidenthappenoop", "No")
        End If
    End If
    
    '9. Are you aware of any witness or other employees involved in this accident/illnes
    If Not IsNull(rsHS("EC_WITNESS")) Then
        If rsHS("EC_WITNESS") <> 0 Then
            nReturn = objMyForm.setField("cb9witness", "Yes")
            If Not IsNull(rsHS("EC_WITNESS1_EMPNBR")) Then
                'nReturn = objMyForm.setField("txt9witnessnamepos1", GetEmpData(rsHS("EC_WITNESS1_EMPNBR"), "ED_FNAME") & " " & GetEmpData(rsHS("EC_WITNESS1_EMPNBR"), "ED_SURNAME") & " - " & rsHS("EC_WITNESS1"))
                nReturn = objMyForm.setField("txt9witnessnamepos1", GetEmpData(rsHS("EC_WITNESS1_EMPNBR"), "ED_FNAME") & " " & GetEmpData(rsHS("EC_WITNESS1_EMPNBR"), "ED_SURNAME") & "; " & rsHS("EC_WITNESS1"))
            Else
                'Ticket #27175 - Non Employee Witness
                If Not IsNull(rsHS("EC_WITNESS1")) Then 'Ticket #27408 Franks 08/13/2015
                    nReturn = objMyForm.setField("txt9witnessnamepos1", rsHS("EC_WITNESS1"))
                End If
            End If
            
            If Not IsNull(rsHS("EC_WITNESS2_EMPNBR")) Then
                'nReturn = objMyForm.setField("txt9witnessnamepos2", GetEmpData(rsHS("EC_WITNESS2_EMPNBR"), "ED_FNAME") & " " & GetEmpData(rsHS("EC_WITNESS2_EMPNBR"), "ED_SURNAME") & " - " & rsHS("EC_WITNESS2"))
                nReturn = objMyForm.setField("txt9witnessnamepos2", GetEmpData(rsHS("EC_WITNESS2_EMPNBR"), "ED_FNAME") & " " & GetEmpData(rsHS("EC_WITNESS2_EMPNBR"), "ED_SURNAME") & "; " & rsHS("EC_WITNESS2"))
            Else
                'Ticket #27175 - Non Employee Witness
                If Not IsNull(rsHS("EC_WITNESS2")) Then 'Ticket #27408 Franks 08/13/2015
                    nReturn = objMyForm.setField("txt9witnessnamepos2", rsHS("EC_WITNESS2"))
                End If
            End If
        Else
            nReturn = objMyForm.setField("cb9witness", "No")
        End If
    End If
    
    '10. Was any individual, who does not work for your firm,.....
    If Not IsNull(rsHS("EC_INDIV_RESP")) Then
        If rsHS("EC_INDIV_RESP") <> 0 Then
            nReturn = objMyForm.setField("cb10individualrespons", "Yes")
            nReturn = objMyForm.setField("txt10individualresponsnameaddetc", rsHS("EC_INDIV_NAME") & " - " & rsHS("EC_INDIV_PHONE"))
        Else
            nReturn = objMyForm.setField("cb10individualrespons", "No")
        End If
    End If
    
    '11. Are you aware of any prior similar or related problem, injury or condition
    If Not IsNull(rsHS("EC_SIMILAR_INJ")) Then
        If rsHS("EC_SIMILAR_INJ") <> 0 Then
            nReturn = objMyForm.setField("cb11similarincident", "Myself")
            nReturn = objMyForm.setField("txt11similarincident", rsHS("EC_SIMILAR_INJ_DEATAILS"))
        Else
            nReturn = objMyForm.setField("cb11similarincident", "No")
        End If
    End If
    
    '12. If you have any concerns about this claim, attach a written submission on this form
    If Not IsNull(rsHS("EC_ANY_CONCERNS")) Then
        If rsHS("EC_ANY_CONCERNS") <> 0 Then
            nReturn = objMyForm.setField("cb12submission", "Yes")
        Else
            nReturn = objMyForm.setField("cb12submission", "no")
        End If
    End If
    
    MDIMain.panHelp(0).FloodPercent = 30
    
    'D. Health Care
    'If Not IsNull(rsHS("EC_PHYSNM")) Then
    If Not IsNull(rsHS("EC_PHYS1_VISIT")) Then  'Ticket #22368
        nReturn = objMyForm.setField("cbD1workerreceivehealthcare", "Yes")
    
        If Not IsNull(rsHS("EC_PHYS1_VISIT")) Then
            nReturn = objMyForm.setField("txtD1workerreceivehealthcarewhenDay", Day(rsHS("EC_PHYS1_VISIT")))
            nReturn = objMyForm.setField("txtD1workerreceivehealthcarewhenMonth", month(rsHS("EC_PHYS1_VISIT")))
            nReturn = objMyForm.setField("txtD1workerreceivehealthcarewhenYear", Right(Year(rsHS("EC_PHYS1_VISIT")), 2))
        End If
    Else
        nReturn = objMyForm.setField("cbD1workerreceivehealthcare", "No")
    End If
    
    If Not IsNull(rsHS("EC_PHYS1_NOTIFIED")) Then
        nReturn = objMyForm.setField("txtD2employerlearnedofworkerreceivehealthcarewhenDay", Day(rsHS("EC_PHYS1_NOTIFIED")))
        nReturn = objMyForm.setField("txtD2employerlearnedofworkerreceivehealthcarewhenMonth", month(rsHS("EC_PHYS1_NOTIFIED")))
        nReturn = objMyForm.setField("txtD2employerlearnedofworkerreceivehealthcarewhenYear", Right(Year(rsHS("EC_PHYS1_NOTIFIED")), 2))
    End If
    
    'Where was the worker treated for his injury?/First Aid Provided (Incident Data) - Table Name: ECFF
    If Not IsNull(rsHS("EC_FAPROVIDED")) Then
        If rsHS("EC_FAPROVIDED") = "ONST" Then
            nReturn = objMyForm.setField("cbD3workerreceivehealthcarewhere on site", "Yes") 'works now
        ElseIf rsHS("EC_FAPROVIDED") = "AMBU" Then
            nReturn = objMyForm.setField("cbD3workerreceivehealthcarewhere ambulance", "Ambulance") 'worked
        ElseIf rsHS("EC_FAPROVIDED") = "EMRG" Then
            nReturn = objMyForm.setField("cbD3workerreceivehealthcarewhere emergency dept", "Emergency Dept.")  'works now
        ElseIf rsHS("EC_FAPROVIDED") = "AHOS" Then
            nReturn = objMyForm.setField("cbD3workerreceivehealthcarewhere hospital", "Admitted to hospital")   'worked
        ElseIf rsHS("EC_FAPROVIDED") = "HPOF" Then
            nReturn = objMyForm.setField("cbD3workerreceivehealthcarewhere HP office", "Health Professional Office") 'works now
        ElseIf rsHS("EC_FAPROVIDED") = "CLIN" Then
            nReturn = objMyForm.setField("cbD3workerreceivehealthcarewhere clinic", "Clinic")   'worked
        Else
            nReturn = objMyForm.setField("cbD3workerreceivehealthcarewhere other", "Other") 'worked
            nReturn = objMyForm.setField("cbtxtD3workerreceivehealthcarewhere other description", GetTABLDesc("ECFF", rsHS("EC_FAPROVIDED")))
        End If
    End If
    
    If Not IsNull(rsHS("EC_PHYSNM")) Then
        nReturn = objMyForm.setField("cbtxtD3workerreceivehealthcarewhere namelocation1", rsHS("EC_PHYSNM"))
    End If
    If Not IsNull(rsHS("EC_PHYSADDR")) And Not IsNull(rsHS("EC_DOCPHONE")) Then
        nReturn = objMyForm.setField("cbtxtD3workerreceivehealthcarewhere namelocation2", rsHS("EC_PHYSADDR") & ". Phone: " & rsHS("EC_DOCPHONE"))
    Else
        If IsNull(rsHS("EC_PHYSADDR")) And Not IsNull(rsHS("EC_DOCPHONE")) Then
            nReturn = objMyForm.setField("cbtxtD3workerreceivehealthcarewhere namelocation2", "Phone: " & rsHS("EC_DOCPHONE"))
        ElseIf Not IsNull(rsHS("EC_PHYSADDR")) And IsNull(rsHS("EC_DOCPHONE")) Then
            nReturn = objMyForm.setField("cbtxtD3workerreceivehealthcarewhere namelocation2", rsHS("EC_PHYSADDR") & ".")
        End If
    End If
    
    MDIMain.panHelp(0).FloodPercent = 35
    
    
    'E. Lost Time - No Lost Time
    '1. After the day of accident/awareness of illness, this worker:
    If Not IsNull(rsForm7Sec("F7_RETURNED_TO")) Then
        If rsForm7Sec("F7_RETURNED_TO") = "R" Then
            nReturn = objMyForm.setField("cbE1", "Reg work no lost time")
        ElseIf rsForm7Sec("F7_RETURNED_TO") = "M" Then
            nReturn = objMyForm.setField("cbE1", "Modified work no lost time")  'Mod work no lost time
        ElseIf rsForm7Sec("F7_RETURNED_TO") = "L" Then
            nReturn = objMyForm.setField("cbE1", "Lost time")
        End If
    End If

    If Not IsNull(rsForm7Sec("F7_LOST_DATE")) Then
        nReturn = objMyForm.setField("txtE1workerdatelosttimeDay", Day(rsForm7Sec("F7_LOST_DATE")))
        nReturn = objMyForm.setField("txtE1workerdatelosttimeMonth", month(rsForm7Sec("F7_LOST_DATE")))
        nReturn = objMyForm.setField("txtE1workerdatelosttimeYear", Right(Year(rsForm7Sec("F7_LOST_DATE")), 2))
    End If

    If Not IsNull(rsForm7Sec("F7_RETURN_DATE")) Then
        nReturn = objMyForm.setField("txtE1DatertwDay", Day(rsForm7Sec("F7_RETURN_DATE")))
        nReturn = objMyForm.setField("txtE1DatertwMonth", month(rsForm7Sec("F7_RETURN_DATE")))
        nReturn = objMyForm.setField("txtE1DatertwDayYear", Right(Year(rsForm7Sec("F7_RETURN_DATE")), 2))
    End If

    If Not IsNull(rsForm7Sec("F7_RETURN_REG_MOD")) Then
        If rsForm7Sec("F7_RETURN_REG_MOD") = "R" Then
            nReturn = objMyForm.setField("cbE1rtwtypeofwork", "Regular Work")
        Else
            nReturn = objMyForm.setField("cbE1rtwtypeofwork", "Modified Work")
        End If
    End If

    '2. This LostTime - No LostTime - Modified Work information was confirmed by:
    If Not IsNull(rsForm7Sec("F7_CONFIRM_BY")) Then
        If rsForm7Sec("F7_CONFIRM_BY") = "M" Then
            nReturn = objMyForm.setField("cbE2Modworkconfirmby", "Myself")
        Else
            nReturn = objMyForm.setField("cbE2Modworkconfirmby", "Other")
        End If
    End If

    If Not IsNull(rsForm7Sec("F7_CONFIRM_NAME")) Then
        nReturn = objMyForm.setField("txtE2Modworkconfirmbyname", rsForm7Sec("F7_CONFIRM_NAME"))
    End If

    If Not IsNull(rsForm7Sec("F7_CONFIRM_PHONE")) Then
        nReturn = objMyForm.setField("txtE2Modworkconfirmbytelephoneareacode", Mid(Replace(Replace(Replace(Replace(rsForm7Sec("F7_CONFIRM_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 1, 3))
        nReturn = objMyForm.setField("txtE2Modworkconfirmbytelephone", Mid(Replace(Replace(Replace(Replace(rsForm7Sec("F7_CONFIRM_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 4, 3) & " " & Mid(Replace(Replace(Replace(Replace(rsForm7Sec("F7_CONFIRM_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 7, 4))
    End If
    If Not IsNull(rsForm7Sec("F7_CONFIRM_PHONE_EXT")) Then
        nReturn = objMyForm.setField("txtE2ModworkconfirmbytelephoneEXT", rsForm7Sec("F7_CONFIRM_PHONE_EXT"))
    End If


    'F. Return to Work
    '1.
    If Not IsNull(rsForm7Sec("F7_LIMITATION")) Then
        If rsForm7Sec("F7_LIMITATION") <> 0 Then
            nReturn = objMyForm.setField("CBF1providedwithlimitations", "Yes")
        Else
            nReturn = objMyForm.setField("CBF1providedwithlimitations", "No")
        End If
    End If
    
    '2.
    If Not IsNull(rsForm7Sec("F7_DISCUSSED")) Then
        If rsForm7Sec("F7_DISCUSSED") <> 0 Then
            nReturn = objMyForm.setField("CBF2modifiedworkdiscussed", "Yes")
        Else
            nReturn = objMyForm.setField("CBF2modifiedworkdiscussed", "No")
        End If
    End If
    
    '3.
    If Not IsNull(rsForm7Sec("F7_OFFERED")) Then
        If rsForm7Sec("F7_OFFERED") <> 0 Then
            nReturn = objMyForm.setField("CBF3modifiedworkoffered", "Yes")
        Else
            nReturn = objMyForm.setField("CBF3modifiedworkoffered", "No")
        End If
    End If
    
    If Not IsNull(rsForm7Sec("F7_ACCEPT_DECLINE")) Then
        If rsForm7Sec("F7_ACCEPT_DECLINE") = "A" Then
            nReturn = objMyForm.setField("CBF3modifiedworkofferedwasit", "Accepted")
        Else
            nReturn = objMyForm.setField("CBF3modifiedworkofferedwasit", "Declined")
        End If
    End If

    If Not IsNull(rsForm7Sec("F7_DECLINE_ATTACHED")) Then
        If rsForm7Sec("F7_DECLINE_ATTACHED") <> 0 Then
            nReturn = objMyForm.setField("CBF3modifiedworkdeclinedattachcopy", "Yes")
        Else
            nReturn = objMyForm.setField("CBF3modifiedworkdeclinedattachcopy", "no")
        End If
    End If
    
    '4. Who is responsible for arranging worker's return to work
    If Not IsNull(rsForm7Sec("F7_RESPONSIBLE")) Then
        If rsForm7Sec("F7_RESPONSIBLE") = "M" Then
            nReturn = objMyForm.setField("cbF4whoarrangedrtw", "Myself")
        Else
            nReturn = objMyForm.setField("cbF4whoarrangedrtw", "other")
        End If
    End If

    If Not IsNull(rsForm7Sec("F7_RESPONS_NAME")) Then
        nReturn = objMyForm.setField("txtF4whoarrangedrtwname", rsForm7Sec("F7_RESPONS_NAME"))
    End If

    If Not IsNull(rsForm7Sec("F7_RESPONS_PHONE")) Then
        nReturn = objMyForm.setField("txtF4whoarrangedrtwtelephoneareacode", Mid(Replace(Replace(Replace(Replace(rsForm7Sec("F7_RESPONS_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 1, 3))
        nReturn = objMyForm.setField("txtF4whoarrangedrtwtelephone", Mid(Replace(Replace(Replace(Replace(rsForm7Sec("F7_RESPONS_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 4, 3) & " " & Mid(Replace(Replace(Replace(Replace(rsForm7Sec("F7_RESPONS_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 7, 4))
    End If
    If Not IsNull(rsForm7Sec("F7_RESPONS_PHONE_EXT")) Then
        nReturn = objMyForm.setField("txtF4whoarrangedrtwtelephoneext", rsForm7Sec("F7_RESPONS_PHONE_EXT"))
    End If
    
    
    'G. Regular Wage/Employment Information
    If rsStatCat("SC_WORKER_TYPE") = "PFT" Then
        nReturn = objMyForm.setField("cbG1is this worker pft", "Permanent full time")
    ElseIf rsStatCat("SC_WORKER_TYPE") = "PPT" Then
        nReturn = objMyForm.setField("cbG1is this worker ppt", "Permanent part time")
    ElseIf rsStatCat("SC_WORKER_TYPE") = "TFT" Then
        nReturn = objMyForm.setField("cbG1is this worker tft", "Temporary full time")
    ElseIf rsStatCat("SC_WORKER_TYPE") = "TPT" Then
        nReturn = objMyForm.setField("cbG1is this worker tpt", "Temporary part time")
    ElseIf rsStatCat("SC_WORKER_TYPE") = "CI" Then
        nReturn = objMyForm.setField("cbG1is this worker ci", "casual/irregular")   'worked
    ElseIf rsStatCat("SC_WORKER_TYPE") = "SEAS" Then
        nReturn = objMyForm.setField("cbG1is this worker seas", "Seasonal") 'works
    ElseIf rsStatCat("SC_WORKER_TYPE") = "CONT" Then
        nReturn = objMyForm.setField("cbG1is this worker contract", "Contract") 'works
    ElseIf rsStatCat("SC_WORKER_TYPE") = "STUD" Then
        nReturn = objMyForm.setField("cbG1is this worker student", "Student")   'works
    ElseIf rsStatCat("SC_WORKER_TYPE") = "UT" Then
        nReturn = objMyForm.setField("cbG1is this worker ut", "Unpaid trainee")
    ElseIf rsStatCat("SC_WORKER_TYPE") = "RA" Then
        nReturn = objMyForm.setField("cbG1is this worker ra", "registered apprentice") 'worked
    ElseIf rsStatCat("SC_WORKER_TYPE") = "OI" Then
        nReturn = objMyForm.setField("cbG1is this worker  oi", "Optional Insurance")  'worked
    ElseIf rsStatCat("SC_WORKER_TYPE") = "OOSC" Then
        nReturn = objMyForm.setField("cbG1is this worker  oosc", "Owner operator (sub) contractor")
    ElseIf rsStatCat("SC_WORKER_TYPE") = "OTHR" Then
        nReturn = objMyForm.setField("cbG1is this worker other", "Other")   'worked
        nReturn = objMyForm.setField("txtG1is this worker other desc", rsStatCat("SC_OTHER_DESC"))
    End If
    
    
    'Get the Hourly Rate
    If rsSal("SH_SALCD") = "H" Then
        xHourlyRate = Round2DEC(rsSal("SH_SALARY"))
        
    ElseIf rsSal("SH_SALCD") = "M" Then
        If rsSal("SH_WHRS") = 0 Then
            xHourlyRate = 0
        Else
            xHourlyRate = Round2DEC(((rsSal("SH_SALARY") * 12) / rsSal("SH_WHRS")) / 52)
        End If
    ElseIf rsSal("SH_SALCD") = "A" Then
        If rsSal("SH_WHRS") = 0 Then
            xHourlyRate = 0
        Else
            xHourlyRate = Round2DEC((rsSal("SH_SALARY") / rsSal("SH_WHRS")) / 52)
        End If
        
    ElseIf rsSal("SH_SALCD") = "D" Then
        'woodbridge get's Daily salary
        If rsSal("SH_WHRS") = 0 Then
            xHourlyRate = 0
        Else
            If GetLeapYear(Year(Date)) Then
                xHourlyRate = Round2DEC(((rsSal("SH_SALARY") * 366) / rsSal("SH_WHRS")) / 52)
            Else
                xHourlyRate = Round2DEC(((rsSal("SH_SALARY") * 365) / rsSal("SH_WHRS")) / 52)
            End If
            
            'Ticket #17654 - formula correction
            If Not IsNull(rsJOB("JH_DHRS")) Then
                If rsJOB("JH_DHRS") <> 0 Then
                    xHourlyRate = Round2DEC((rsSal("SH_SALARY") / rsJOB("JH_DHRS")))
                Else
                    xHourlyRate = 0
                End If
            Else
                xHourlyRate = 0
            End If
        End If
    End If
    nReturn = objMyForm.setField("txtG2regrateof pay", xHourlyRate)
    nReturn = objMyForm.setField("cbG2rateofpayper", "hour") 'hour or yes?
    
    MDIMain.panHelp(0).FloodPercent = 40
    
    'H - Additional Wage Information
    '1. Net Claim Code or Amount
    'If Not IsNull(rsEMP("ED_TD1DOL")) Then
    If Not IsNull(rsForm7Sec("F7_FED_AMT")) Then
        'nReturn = objMyForm.setField("txtH1nextclaimcodefederal", rsEMP("ED_TD1DOL"))
        nReturn = objMyForm.setField("txtH1nextclaimcodefederal", rsForm7Sec("F7_FED_AMT"))
    End If
    'If Not IsNull(rsEMP("ED_PROVAMT")) Then
    If Not IsNull(rsForm7Sec("F7_PROV_AMT")) Then
        'nReturn = objMyForm.setField("txtH1nextclaimcodeprovincial", rsEMP("ED_PROVAMT"))
        nReturn = objMyForm.setField("txtH1nextclaimcodeprovincial", rsForm7Sec("F7_PROV_AMT"))
    End If
    
    '2. Vacation Pay - on each cheque?
    If Not IsNull(rsForm7Sec("F7_VAC_PAY")) Then
        If rsForm7Sec("F7_VAC_PAY") <> 0 Then
            nReturn = objMyForm.setField("cbH2vacationpaypercheque", "Yes")
        Else
            nReturn = objMyForm.setField("cbH2vacationpaypercheque", "No")
        End If
    End If
    
    'Provide percentage
    If Not IsNull(rsForm7Sec("F7_VACPC")) Then
        nReturn = objMyForm.setField("txtH2vacationpayperchequepercent", rsForm7Sec("F7_VACPC"))
    End If

    '3. Date and hour last worked
    If Not IsNull(rsForm7Sec("F7_LAST_WORK_DATE")) Then
        nReturn = objMyForm.setField("txtH3datelastworkedDay", Day(rsForm7Sec("F7_LAST_WORK_DATE")))
        nReturn = objMyForm.setField("txtH3datelastworkedMonth", month(rsForm7Sec("F7_LAST_WORK_DATE")))
        nReturn = objMyForm.setField("txtH3datelastworkedYear", Right(Year(rsForm7Sec("F7_LAST_WORK_DATE")), 2))
    End If

    If Not IsNull(rsForm7Sec("F7_LAST_WORK_TIME")) Then
        'If Val(Left(rsForm7Sec("F7_LAST_WORK_TIME"), 2)) < 12 Then
        If Not IsNull(rsForm7Sec("F7_LAST_WORK_AMPM")) Then
            If rsForm7Sec("F7_LAST_WORK_AMPM") = "A" Then
                nReturn = objMyForm.setField("txtH3timelastworkedAM", rsForm7Sec("F7_LAST_WORK_TIME"))
                nReturn = objMyForm.setField("cbtimelastworked", "AM")
            ElseIf rsForm7Sec("F7_LAST_WORK_AMPM") = "P" Then
                nReturn = objMyForm.setField("txtH3timelastworkedPM", rsForm7Sec("F7_LAST_WORK_TIME"))
                nReturn = objMyForm.setField("cbtimelastworked", "PM")
            End If
        End If
    End If

    '4. Normal working hours on last day worked
    If Not IsNull(rsForm7Sec("F7_LAST_DAY_WORK_FTIME")) Then
        'If Val(Left(rsForm7sec("F7_LAST_DAY_WORK_FTIME"), 2)) < 12 Then
        If Not IsNull(rsForm7Sec("F7_LAST_DAY_WORK_FAMPM")) Then
            If rsForm7Sec("F7_LAST_DAY_WORK_FAMPM") = "A" Then
                nReturn = objMyForm.setField("txtH4timelastworkedfromAM", rsForm7Sec("F7_LAST_DAY_WORK_FTIME"))
                nReturn = objMyForm.setField("cbh4timelastworkedfrom", "AM")
            ElseIf rsForm7Sec("F7_LAST_DAY_WORK_FAMPM") = "P" Then
                nReturn = objMyForm.setField("txtH4timelastworkedfromPM", rsForm7Sec("F7_LAST_DAY_WORK_FTIME"))
                nReturn = objMyForm.setField("cbh4timelastworkedfrom", "PM")
            End If
        End If
    End If

    If Not IsNull(rsForm7Sec("F7_LAST_DAY_WORK_TTIME")) Then
        'If Val(Left(rsForm7sec("F7_LAST_DAY_WORK_TTIME"), 2)) < 12 Then
        If Not IsNull(rsForm7Sec("F7_LAST_DAY_WORK_TAMPM")) Then
            If rsForm7Sec("F7_LAST_DAY_WORK_TAMPM") = "A" Then
                nReturn = objMyForm.setField("txtH4timelastworkedtoAM", rsForm7Sec("F7_LAST_DAY_WORK_TTIME"))
                nReturn = objMyForm.setField("cbh4timelastworkedto", "AM")
            ElseIf rsForm7Sec("F7_LAST_DAY_WORK_TAMPM") = "P" Then
                nReturn = objMyForm.setField("txtH4timelastworkedtoPM", rsForm7Sec("F7_LAST_DAY_WORK_TTIME"))
                nReturn = objMyForm.setField("cbh4timelastworkedto", "PM")
            End If
        End If
    End If

    '5. Actual earnings for last day worked
    If Not IsNull(rsForm7Sec("F7_LAST_DAY_ACT_EARN")) Then
        nReturn = objMyForm.setField("txtH5actualearninglastday", rsForm7Sec("F7_LAST_DAY_ACT_EARN"))
    End If

    '6. Normal earnings for last day worked
    If Not IsNull(rsForm7Sec("F7_LAST_DAY_NORM_EARN")) Then
        nReturn = objMyForm.setField("txtH5normalearninglastday", rsForm7Sec("F7_LAST_DAY_NORM_EARN"))
    End If

    '7. Advances on wages
    If Not IsNull(rsForm7Sec("FY_WORKER_PAID")) Then
        If rsForm7Sec("FY_WORKER_PAID") <> 0 Then
            nReturn = objMyForm.setField("cbH7advanceearnings", "Yes")
        Else
            nReturn = objMyForm.setField("cbH7advanceearnings", "No")
        End If
    End If
    
    ' - Full/Regular or Other
    If Not IsNull(rsForm7Sec("F7_WORKER_FTREGOTHR")) Then
        If rsForm7Sec("F7_WORKER_FTREGOTHR") = "F" Then
            nReturn = objMyForm.setField("cbH7advanceearningsamount", "Full Regular")
        Else
            nReturn = objMyForm.setField("cbH7advanceearningsamount", "Other")
            nReturn = objMyForm.setField("txtH7advanceearningsamount other desc", rsForm7Sec("F7_WORKER_OTHER"))
        End If
    End If
        
    '8. Other Earnings (Not Regular Wages)
    'Week is Sun - Sat.
    If IsDate(rsHS("EC_OCCDATE")) Then
                
        'Get which day of the week it is and get the # of days before the last Saturday
        Select Case Weekday(rsHS("EC_OCCDATE"))
            Case vbMonday
                xSat = -2
            Case vbTuesday
                xSat = -3
            Case vbWednesday
                xSat = -4
            Case vbThursday
                xSat = -5
            Case vbFriday
                xSat = -6
            Case vbSaturday
                xSat = -7
            Case vbSunday
                xSat = -8
        End Select
        
        'Week 1 - From
        'Compute the date of the last Saturday
        'xSatDate = DateAdd("d", xSat, rsHS("EC_OCCDATE"))
        'Date prior to Incident Date
        xSatDate = DateAdd("d", -1, CVDate(rsHS("EC_OCCDATE")))
        
        'Compute the date of last Sunday - 1 week prior
        xSunDate = DateAdd("d", -6, xSatDate)
        
        If IsDate(rsForm7Sec("F7_OTH_EARN_FROM_WK1")) Then
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek1fromDAY", Day(xSunDate))
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek1fromMONTH", month(xSunDate))
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek1fromYEAR", Right(Year(xSunDate), 2))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek1fromDAY", Day(rsForm7Sec("F7_OTH_EARN_FROM_WK1")))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek1fromMONTH", month(rsForm7Sec("F7_OTH_EARN_FROM_WK1")))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek1fromYEAR", Right(Year(rsForm7Sec("F7_OTH_EARN_FROM_WK1")), 2))
        End If
        
        'Week 1 - To
        If IsDate(rsForm7Sec("F7_OTH_EARN_TO_WK1")) Then
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek1toDAY", Day(xSatDate))
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek1toMONTH", month(xSatDate))
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek1toYEAR", Right(Year(xSatDate), 2))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek1toDAY", Day(rsForm7Sec("F7_OTH_EARN_TO_WK1")))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek1toMONTH", month(rsForm7Sec("F7_OTH_EARN_TO_WK1")))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek1toYEAR", Right(Year(rsForm7Sec("F7_OTH_EARN_TO_WK1")), 2))
        End If
        
        'Week 2 - From
        xSatDate = DateAdd("d", -1, xSunDate)
        xSunDate = DateAdd("d", -6, xSatDate)
        If IsDate(rsForm7Sec("F7_OTH_EARN_FROM_WK2")) Then
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek2fromDAY", Day(xSunDate))
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek2fromMONTH", month(xSunDate))
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek2fromYEAR", Right(Year(xSunDate), 2))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek2fromDAY", Day(rsForm7Sec("F7_OTH_EARN_FROM_WK2")))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek2fromMONTH", month(rsForm7Sec("F7_OTH_EARN_FROM_WK2")))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek2fromYEAR", Right(Year(rsForm7Sec("F7_OTH_EARN_FROM_WK2")), 2))
        End If
        'Week 2 - To
        If IsDate(rsForm7Sec("F7_OTH_EARN_TO_WK2")) Then
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek2toDAY", Day(xSatDate))
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek2toMONTH", month(xSatDate))
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek2toYEAR", Right(Year(xSatDate), 2))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek2toDAY", Day(rsForm7Sec("F7_OTH_EARN_TO_WK2")))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek2toMONTH", month(rsForm7Sec("F7_OTH_EARN_TO_WK2")))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek2toYEAR", Right(Year(rsForm7Sec("F7_OTH_EARN_TO_WK2")), 2))
        End If
        
        'Week 3 - From
        xSatDate = DateAdd("d", -1, xSunDate)
        xSunDate = DateAdd("d", -6, xSatDate)
        If IsDate(rsForm7Sec("F7_OTH_EARN_FROM_WK3")) Then
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek3fromDAY", Day(xSunDate))
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek3fromMONTH", month(xSunDate))
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek3fromYEAR", Right(Year(xSunDate), 2))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek3fromDAY", Day(rsForm7Sec("F7_OTH_EARN_FROM_WK3")))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek3fromMONTH", month(rsForm7Sec("F7_OTH_EARN_FROM_WK3")))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek3fromYEAR", Right(Year(rsForm7Sec("F7_OTH_EARN_FROM_WK3")), 2))
        End If
        'Week 3 - To
        If IsDate(rsForm7Sec("F7_OTH_EARN_TO_WK3")) Then
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek3toDAY", Day(xSatDate))
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek3toMONTH", month(xSatDate))
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek3toYEAR", Right(Year(xSatDate), 2))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek3toDAY", Day(rsForm7Sec("F7_OTH_EARN_TO_WK3")))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek3toMONTH", month(rsForm7Sec("F7_OTH_EARN_TO_WK3")))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek3toYEAR", Right(Year(rsForm7Sec("F7_OTH_EARN_TO_WK3")), 2))
        End If
        
        'Week 4 - From
        xSatDate = DateAdd("d", -1, xSunDate)
        xSunDate = DateAdd("d", -6, xSatDate)
        If IsDate(rsForm7Sec("F7_OTH_EARN_FROM_WK4")) Then
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek4fromDAY", Day(xSunDate))
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek4fromMONTH", month(xSunDate))
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek4fromYEAR", Right(Year(xSunDate), 2))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek4fromDAY", Day(rsForm7Sec("F7_OTH_EARN_FROM_WK4")))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek4fromMONTH", month(rsForm7Sec("F7_OTH_EARN_FROM_WK4")))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek4fromYEAR", Right(Year(rsForm7Sec("F7_OTH_EARN_FROM_WK4")), 2))
        End If
        'Week 4 - To
        If IsDate(rsForm7Sec("F7_OTH_EARN_TO_WK4")) Then
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek4toDAY", Day(xSatDate))
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek4toMONTH", month(xSatDate))
            'nReturn = objMyForm.setField("txtH8otherearningNRWWeek4toYEAR", Right(Year(xSatDate), 2))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek4toDAY", Day(rsForm7Sec("F7_OTH_EARN_TO_WK4")))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek4toMONTH", month(rsForm7Sec("F7_OTH_EARN_TO_WK4")))
            nReturn = objMyForm.setField("txtH8otherearningNRWWeek4toYEAR", Right(Year(rsForm7Sec("F7_OTH_EARN_TO_WK4")), 2))
        End If
    End If
    
    'Mandatory Overtime Pay
    If Not IsNull(rsForm7Sec("F7_MAND_OVT_PAY_WK1")) Then
        nReturn = objMyForm.setField("txtH8Week1mopay", rsForm7Sec("F7_MAND_OVT_PAY_WK1"))
    End If
    If Not IsNull(rsForm7Sec("F7_MAND_OVT_PAY_WK2")) Then
        nReturn = objMyForm.setField("txtH8Week2mopay", rsForm7Sec("F7_MAND_OVT_PAY_WK2"))
    End If
    If Not IsNull(rsForm7Sec("F7_MAND_OVT_PAY_WK3")) Then
        nReturn = objMyForm.setField("txtH8Week3mopay", rsForm7Sec("F7_MAND_OVT_PAY_WK3"))
    End If
    If Not IsNull(rsForm7Sec("F7_MAND_OVT_PAY_WK4")) Then
        nReturn = objMyForm.setField("txtH8Week4mopay", rsForm7Sec("F7_MAND_OVT_PAY_WK4"))
    End If
    
    'Voluntary Overtime Pay
    If Not IsNull(rsForm7Sec("F7_VOL_OVT_PAY_WK1")) Then
        nReturn = objMyForm.setField("txtH8Week1vopay", rsForm7Sec("F7_VOL_OVT_PAY_WK1"))
    End If
    If Not IsNull(rsForm7Sec("F7_VOL_OVT_PAY_WK2")) Then
        nReturn = objMyForm.setField("txtH8Week2vopay", rsForm7Sec("F7_VOL_OVT_PAY_WK2"))
    End If
    If Not IsNull(rsForm7Sec("F7_VOL_OVT_PAY_WK3")) Then
        nReturn = objMyForm.setField("txtH8Week3vopay", rsForm7Sec("F7_VOL_OVT_PAY_WK3"))
    End If
    If Not IsNull(rsForm7Sec("F7_VOL_OVT_PAY_WK4")) Then
        nReturn = objMyForm.setField("txtH8Week4vopay", rsForm7Sec("F7_VOL_OVT_PAY_WK4"))
    End If
        
    'Other Earnings Week 1
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_1")) Then
        nReturn = objMyForm.setField("HEADINGS1", rsForm7Sec("F7_OTH_EARN_1"))
    Else
        nReturn = objMyForm.setField("HEADINGS1", "")
    End If
    
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_1_WK1")) Then
        nReturn = objMyForm.setField("txtH8Week1other1", rsForm7Sec("F7_OTH_EARN_1_WK1"))
    End If
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_2_WK1")) Then
        nReturn = objMyForm.setField("txtH8Week1other2", rsForm7Sec("F7_OTH_EARN_2_WK1"))
    End If
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_3_WK1")) Then
        nReturn = objMyForm.setField("txtH8Week1other3", rsForm7Sec("F7_OTH_EARN_3_WK1"))
    End If
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_4_WK1")) Then
        nReturn = objMyForm.setField("txtH8Week1other4", rsForm7Sec("F7_OTH_EARN_4_WK1"))
    End If
    
    'Other Earnings Week 2
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_2")) Then
        nReturn = objMyForm.setField("HEADINGS2", rsForm7Sec("F7_OTH_EARN_2"))
    Else
        nReturn = objMyForm.setField("HEADINGS2", "")
    End If
    
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_1_WK2")) Then
        nReturn = objMyForm.setField("txtH8Week2other1", rsForm7Sec("F7_OTH_EARN_1_WK2"))
    End If
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_2_WK2")) Then
        nReturn = objMyForm.setField("txtH8Week2other2", rsForm7Sec("F7_OTH_EARN_2_WK2"))
    End If
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_3_WK2")) Then
        nReturn = objMyForm.setField("txtH8Week2other3", rsForm7Sec("F7_OTH_EARN_3_WK2"))
    End If
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_4_WK2")) Then
        nReturn = objMyForm.setField("txtH8Week2other4", rsForm7Sec("F7_OTH_EARN_4_WK2"))
    End If
    
    'Other Earnings Week 3
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_3")) Then
        nReturn = objMyForm.setField("HEADINGS3", rsForm7Sec("F7_OTH_EARN_3"))
    Else
        nReturn = objMyForm.setField("HEADINGS3", "")
    End If
    
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_1_WK3")) Then
        nReturn = objMyForm.setField("txtH8Week3other1", rsForm7Sec("F7_OTH_EARN_1_WK3"))
    End If
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_2_WK3")) Then
        nReturn = objMyForm.setField("txtH8Week3other2", rsForm7Sec("F7_OTH_EARN_2_WK3"))
    End If
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_3_WK3")) Then
        nReturn = objMyForm.setField("txtH8Week3other3", rsForm7Sec("F7_OTH_EARN_3_WK3"))
    End If
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_4_WK3")) Then
        nReturn = objMyForm.setField("txtH8Week3other4", rsForm7Sec("F7_OTH_EARN_4_WK3"))
    End If
    
    'Other Earnings Week 4
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_4")) Then
        nReturn = objMyForm.setField("HEADINGS4", rsForm7Sec("F7_OTH_EARN_4"))
    Else
        nReturn = objMyForm.setField("HEADINGS4", "")
    End If
    
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_1_WK4")) Then
        nReturn = objMyForm.setField("txtH8Week4other1", rsForm7Sec("F7_OTH_EARN_1_WK4"))
    End If
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_2_WK4")) Then
        nReturn = objMyForm.setField("txtH8Week4other2", rsForm7Sec("F7_OTH_EARN_2_WK4"))
    End If
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_3_WK4")) Then
        nReturn = objMyForm.setField("txtH8Week4other3", rsForm7Sec("F7_OTH_EARN_3_WK4"))
    End If
    If Not IsNull(rsForm7Sec("F7_OTH_EARN_4_WK4")) Then
        nReturn = objMyForm.setField("txtH8Week4other4", rsForm7Sec("F7_OTH_EARN_4_WK4"))
    End If
    
    
    MDIMain.panHelp(0).FloodPercent = 50
    
    'I - Work Schedule
    If Not IsNull(rsForm7Sec("F7_WORKSCH")) Then
        If rsForm7Sec("F7_WORKSCH") = "N" Then
            nReturn = objMyForm.setField("cbIschedule", "A")
            If Not IsNull(rsForm7Sec("F7_REG_SCHD_SUN")) Then
                nReturn = objMyForm.setField("txtIAsunday", rsForm7Sec("F7_REG_SCHD_SUN"))
            End If
            If Not IsNull(rsForm7Sec("F7_REG_SCHD_MON")) Then
                nReturn = objMyForm.setField("txtIAmonday", rsForm7Sec("F7_REG_SCHD_MON"))
            End If
            If Not IsNull(rsForm7Sec("F7_REG_SCHD_TUE")) Then
                nReturn = objMyForm.setField("txtIAtuesday", rsForm7Sec("F7_REG_SCHD_TUE"))
            End If
            If Not IsNull(rsForm7Sec("F7_REG_SCHD_WED")) Then
                nReturn = objMyForm.setField("txtIAwednesday", rsForm7Sec("F7_REG_SCHD_WED"))
            End If
            If Not IsNull(rsForm7Sec("F7_REG_SCHD_THU")) Then
                nReturn = objMyForm.setField("txtIAthursday", rsForm7Sec("F7_REG_SCHD_THU"))
            End If
            If Not IsNull(rsForm7Sec("F7_REG_SCHD_FRI")) Then
                nReturn = objMyForm.setField("txtIAfriday", rsForm7Sec("F7_REG_SCHD_FRI"))
            End If
            If Not IsNull(rsForm7Sec("F7_REG_SCHD_SAT")) Then
                nReturn = objMyForm.setField("txtIAsaturday", rsForm7Sec("F7_REG_SCHD_SAT"))
            End If
        ElseIf rsForm7Sec("F7_WORKSCH") = "R" Then
            nReturn = objMyForm.setField("cbIschedule", "B")
            If Not IsNull(rsForm7Sec("F7_NUM_DAYS_ON")) Then
                nReturn = objMyForm.setField("txtIBnoofdayson", rsForm7Sec("F7_NUM_DAYS_ON"))
            End If
            If Not IsNull(rsForm7Sec("F7_NUM_DAYS_OFF")) Then
                nReturn = objMyForm.setField("txtIBnoofdaysoff", rsForm7Sec("F7_NUM_DAYS_OFF"))
            End If
            If Not IsNull(rsForm7Sec("F7_HRS_SHIFT")) Then
                nReturn = objMyForm.setField("txtIBhourspershiftnoofdayson", rsForm7Sec("F7_HRS_SHIFT"))
            End If
            If Not IsNull(rsForm7Sec("F7_NUM_WKS_CYCLE")) Then
                nReturn = objMyForm.setField("txtIBnoofweeksincycle", rsForm7Sec("F7_NUM_WKS_CYCLE"))
            End If
        Else
            nReturn = objMyForm.setField("cbIschedule", "C")
            
            'Week 1
            If Not IsNull(rsForm7Sec("F7_FWEEK1")) Then
                nReturn = objMyForm.setField("txtICvwsWeek1FROM day", Day(rsForm7Sec("F7_FWEEK1")))
                nReturn = objMyForm.setField("txtICvwsWeek1FROM month", month(rsForm7Sec("F7_FWEEK1")))
                nReturn = objMyForm.setField("txtICvwsWeek1FROM year", Right(Year(rsForm7Sec("F7_FWEEK1")), 2))
            End If
            If Not IsNull(rsForm7Sec("F7_TWEEK1")) Then
                nReturn = objMyForm.setField("txtICvwsWeek1TO day", Day(rsForm7Sec("F7_TWEEK1")))
                nReturn = objMyForm.setField("txtICvwsWeek1TO month", month(rsForm7Sec("F7_TWEEK1")))
                nReturn = objMyForm.setField("txtICvwsWeek1TO year", Right(Year(rsForm7Sec("F7_TWEEK1")), 2))
            End If
            If Not IsNull(rsForm7Sec("F7_TOT_HRS_WEEK1")) Then
                nReturn = objMyForm.setField("txtICtotalhoursworkedWeek1", rsForm7Sec("F7_TOT_HRS_WEEK1"))
            End If
            If Not IsNull(rsForm7Sec("F7_TOT_SHIFT_WEEK1")) Then
                nReturn = objMyForm.setField("txtICtotalshiftsworkedWeek1", rsForm7Sec("F7_TOT_SHIFT_WEEK1"))
            End If
            
            
            'Week 2
            If Not IsNull(rsForm7Sec("F7_FWEEK2")) Then
                nReturn = objMyForm.setField("txtICvwsWeek2FROM day", Day(rsForm7Sec("F7_FWEEK2")))
                nReturn = objMyForm.setField("txtICvwsWeek2FROM month", month(rsForm7Sec("F7_FWEEK2")))
                nReturn = objMyForm.setField("txtICvwsWeek2FROM year", Right(Year(rsForm7Sec("F7_FWEEK2")), 2))
            End If
            If Not IsNull(rsForm7Sec("F7_TWEEK2")) Then
                nReturn = objMyForm.setField("txtICvwsWeek2TO day", Day(rsForm7Sec("F7_TWEEK2")))
                nReturn = objMyForm.setField("txtICvwsWeek2TO month", month(rsForm7Sec("F7_TWEEK2")))
                nReturn = objMyForm.setField("txtICvwsWeek2TO year", Right(Year(rsForm7Sec("F7_TWEEK2")), 2))
            End If
            If Not IsNull(rsForm7Sec("F7_TOT_HRS_WEEK2")) Then
                nReturn = objMyForm.setField("txtICtotalhoursworkedWeek2", rsForm7Sec("F7_TOT_HRS_WEEK2"))
            End If
            If Not IsNull(rsForm7Sec("F7_TOT_SHIFT_WEEK2")) Then
                nReturn = objMyForm.setField("txtICtotalshiftsworkedWeek2", rsForm7Sec("F7_TOT_SHIFT_WEEK2"))
            End If
        
            'Week 3
            If Not IsNull(rsForm7Sec("F7_FWEEK3")) Then
                nReturn = objMyForm.setField("txtICvwsWeek3FROM day", Day(rsForm7Sec("F7_FWEEK3")))
                nReturn = objMyForm.setField("txtICvwsWeek3FROM month", month(rsForm7Sec("F7_FWEEK3")))
                nReturn = objMyForm.setField("txtICvwsWeek3FROM year", Right(Year(rsForm7Sec("F7_FWEEK3")), 2))
            End If
            If Not IsNull(rsForm7Sec("F7_TWEEK3")) Then
                nReturn = objMyForm.setField("txtICvwsWeek3TO day", Day(rsForm7Sec("F7_TWEEK3")))
                nReturn = objMyForm.setField("txtICvwsWeek3TO month", month(rsForm7Sec("F7_TWEEK3")))
                nReturn = objMyForm.setField("txtICvwsWeek3TO year", Right(Year(rsForm7Sec("F7_TWEEK3")), 2))
            End If
            If Not IsNull(rsForm7Sec("F7_TOT_HRS_WEEK3")) Then
                nReturn = objMyForm.setField("txtICtotalhoursworkedWeek3", rsForm7Sec("F7_TOT_HRS_WEEK3"))
            End If
            If Not IsNull(rsForm7Sec("F7_TOT_SHIFT_WEEK3")) Then
                nReturn = objMyForm.setField("txtICtotalshiftsworkedWeek3", rsForm7Sec("F7_TOT_SHIFT_WEEK3"))
            End If
            
            'Week
            If Not IsNull(rsForm7Sec("F7_FWEEK4")) Then
                nReturn = objMyForm.setField("txtICvwsWeek4FROM day", Day(rsForm7Sec("F7_FWEEK4")))
                nReturn = objMyForm.setField("txtICvwsWeek4FROM month", month(rsForm7Sec("F7_FWEEK4")))
                nReturn = objMyForm.setField("txtICvwsWeek4FROM year", Right(Year(rsForm7Sec("F7_FWEEK4")), 2))
            End If
            If Not IsNull(rsForm7Sec("F7_TWEEK4")) Then
                nReturn = objMyForm.setField("txtICvwsWeek4TO day", Day(rsForm7Sec("F7_TWEEK4")))
                nReturn = objMyForm.setField("txtICvwsWeek4TO month", month(rsForm7Sec("F7_TWEEK4")))
                nReturn = objMyForm.setField("txtICvwsWeek4TO year", Right(Year(rsForm7Sec("F7_TWEEK4")), 2))
            End If
            If Not IsNull(rsForm7Sec("F7_TOT_HRS_WEEK4")) Then
                nReturn = objMyForm.setField("txtICtotalhoursworkedWeek4", rsForm7Sec("F7_TOT_HRS_WEEK4"))
            End If
            If Not IsNull(rsForm7Sec("F7_TOT_SHIFT_WEEK4")) Then
                nReturn = objMyForm.setField("txtICtotalshiftsworkedWeek4", rsForm7Sec("F7_TOT_SHIFT_WEEK4"))
            End If
        End If
    End If


    'J. Filled By - Ticket #22682 - Release 8.0
    If Not IsNull(rsForm7Sec("F7_NAME")) Then
        nReturn = objMyForm.setField("txtIjname", rsForm7Sec("F7_NAME"))
    End If
    If Not IsNull(rsForm7Sec("F7_TITLE")) Then
        nReturn = objMyForm.setField("txtJtitle", rsForm7Sec("F7_TITLE"))
    End If
    If Not IsNull(rsForm7Sec("F7_PHONE")) Then
        nReturn = objMyForm.setField("txtJtelephoneareacode", Mid(Replace(Replace(Replace(Replace(rsForm7Sec("F7_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 1, 3))
        nReturn = objMyForm.setField("txtJtelephone", Mid(Replace(Replace(Replace(Replace(rsForm7Sec("F7_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 4, 3) & " " & Mid(Replace(Replace(Replace(Replace(rsForm7Sec("F7_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 7, 4))
        nReturn = objMyForm.setField("txtJtelephoneextension", Mid(Replace(Replace(Replace(Replace(rsForm7Sec("F7_PHONE"), ")", ""), "(", ""), " ", ""), "-", ""), 11))
    End If
    
    'K. Additional Information
    If Not IsNull(rsForm7Sec("F7_ADDITIONAL_INFO")) Then
        nReturn = objMyForm.setField("TXTKADDITIONAL TEXT", rsForm7Sec("F7_ADDITIONAL_INFO"))
    End If
    
    
    'Get the location to save the file in
    xPathToSaveIn = GetComPreferEmail("WSIBFORM7PATH")
    If Len(xPathToSaveIn) = 0 Then
        xPathToSaveIn = glbIHRREPORTS
    End If
    
    If Right(xPathToSaveIn, 1) <> "\" Then xPathToSaveIn = xPathToSaveIn & "\"
    
    'Save completed form file into a new PDF file
    If rsHS("EC_ANY_CONCERNS") <> 0 Or rsForm7Sec("F7_DECLINE_ATTACHED") <> 0 Then
        'Concerns PDF and/or Written Offer pdf found to merge
        
        'Save the Form 7 generated first with the above values filled in the form.
        objMyForm.saveFile (xPathToSaveIn & glbLEE_ID & "_Form7.pdf")
        
        MDIMain.panHelp(0).FloodPercent = 60
        
        'Open the attached Concerned Document, so it can be merged
        'glbIHRREPORTS & glbLEE_ID & "_Concerned.pdf"
        
        'Open the Concern's PDF
        If rsHS("EC_ANY_CONCERNS") <> 0 Then
            'Set the attached document keys
            glbDocName = "INJURYWF7"
            glbDocKey = IIf(IsNull(rsHS("EC_DOCKEY")), "", rsHS("EC_DOCKEY"))
            glbJob = rsHS("EC_CASE")
            SQLQ = getSQL("frmEHSINJURYWF7")
            
            'Open the Concerns document
            xConcDocFound = OpenDocument(SQLQ, "_Concerned", glbLEE_ID)
        End If
        
        'Open the Written Offer PDF
        If rsForm7Sec("F7_DECLINE_ATTACHED") <> 0 Then
            'Set the attached document keys
            glbDocName = "INJURYWF7_WRITTENOFR"
            glbDocKey = IIf(IsNull(rsForm7Sec("F7_DOCKEY")), "", rsForm7Sec("F7_DOCKEY"))
            glbJob = rsForm7Sec("F7_CASE")
            SQLQ = getSQL("frmEInjF7Sections")
            
            'Open hte Writter Offer document
            xWrtnDocFound = OpenDocument(SQLQ, "_WrittenOffer", glbLEE_ID)
        End If
        
        'Merge both the attached documents if both have been attached before merging with Form 7
        'If only one document attached then simply merge that with Form 7 document.
        If (rsHS("EC_ANY_CONCERNS") <> 0 And xConcDocFound) And (rsForm7Sec("F7_DECLINE_ATTACHED") <> 0 And xWrtnDocFound) Then
            If (Dir(xPathToSaveIn & glbLEE_ID & "_Concerned.pdf")) <> "" And (Dir(xPathToSaveIn & glbLEE_ID & "_WrittenOffer.pdf")) <> "" Then
                'Merge the two PDF documents and save the merged PDF file in _AttachedPDF.pdf.
                nReturn = objMyForm.mergePDF(xPathToSaveIn & glbLEE_ID & "_Concerned.pdf", xPathToSaveIn & glbLEE_ID & "_WrittenOffer.pdf", xPathToSaveIn & glbLEE_ID & "_AttachedPDF.pdf")
                
                'Now merge _Form7.pdf with the above merged attached document or if only one attached document
                'then merge with that.
                nReturn = objMyForm.mergePDF(xPathToSaveIn & glbLEE_ID & "_Form7.pdf", xPathToSaveIn & glbLEE_ID & "_AttachedPDF.pdf", xPathToSaveIn & glbLEE_ID & "_WSIBForm7.pdf")
            End If
        Else
            'Only one document attached. Merge that with Form7 document.
            If (rsHS("EC_ANY_CONCERNS") <> 0 And xConcDocFound) Then
                If (Dir(xPathToSaveIn & glbLEE_ID & "_Concerned.pdf")) <> "" Then
                    'Merge the two PDF documents and save the merged PDF file in _WSIBForm7.pdf.
                    nReturn = objMyForm.mergePDF(xPathToSaveIn & glbLEE_ID & "_Form7.pdf", xPathToSaveIn & glbLEE_ID & "_Concerned.pdf", xPathToSaveIn & glbLEE_ID & "_WSIBForm7.pdf")
                End If
            ElseIf (rsForm7Sec("F7_DECLINE_ATTACHED") <> 0 And xWrtnDocFound) Then
                If (Dir(xPathToSaveIn & glbLEE_ID & "_WrittenOffer.pdf")) <> "" Then
                    'Merge the two PDF documents and save the merged PDF file in _WSIBForm7.pdf.
                    nReturn = objMyForm.mergePDF(xPathToSaveIn & glbLEE_ID & "_Form7.pdf", xPathToSaveIn & glbLEE_ID & "_WrittenOffer.pdf", xPathToSaveIn & glbLEE_ID & "_WSIBForm7.pdf")
                End If
            End If
        End If
        
        
'        MDIMain.panHelp(0).FloodPercent = 65
        
'        If (Dir(xPathToSaveIn & glbLEE_ID & "_Concerned.pdf")) <> "" Then
'            'Concatenate two PDF documents and save the merged PDF file.
'            nReturn = objMyForm.mergePDF(xPathToSaveIn & glbLEE_ID & "_Form7.pdf", xPathToSaveIn & glbLEE_ID & "_Concerned.pdf", xPathToSaveIn & glbLEE_ID & "_WSIBForm7.pdf")
'        End If
    
        MDIMain.panHelp(0).FloodPercent = 80

        'Clean Up folder
        'Delete the employee's Concerns, Written Offer and Form 7 temporarily created
        'Delete _Form 7.pdf if Concerns or Written Offer documents were found else rename/copy _Form7.pdf to
        '_WSIBForm7.pdf and then delete _Form7.pdf
        If Not xConcDocFound And Not xWrtnDocFound Then
            'Copy the _Form7.pdf to _WSIBForm7.pdf
            If (Dir(xPathToSaveIn & glbLEE_ID & "_Form7.pdf")) <> "" Then
                FileCopy xPathToSaveIn & glbLEE_ID & "_Form7.pdf", xPathToSaveIn & glbLEE_ID & "_WSIBForm7.pdf"
            End If
            
            'Delete _Form7.pdf
            If (Dir(xPathToSaveIn & glbLEE_ID & "_Form7.pdf")) <> "" Then Kill xPathToSaveIn & glbLEE_ID & "_Form7.pdf"
        Else
            If (Dir(xPathToSaveIn & glbLEE_ID & "_Form7.pdf")) <> "" Then Kill xPathToSaveIn & glbLEE_ID & "_Form7.pdf"
        End If
        
        'Deleting _Concerned.pdf
        If (Dir(xPathToSaveIn & glbLEE_ID & "_Concerned.pdf")) <> "" Then
            Call FileAttributeForm7(xPathToSaveIn & glbLEE_ID & "_Concerned.pdf", "-r", xPathToSaveIn)
            Call Pause(5)
    
            Kill xPathToSaveIn & glbLEE_ID & "_Concerned.pdf"
        End If
        
        'Deleting _WrittenOffer.pdf
        If (Dir(xPathToSaveIn & glbLEE_ID & "_WrittenOffer.pdf")) <> "" Then
            Call FileAttributeForm7(xPathToSaveIn & glbLEE_ID & "_WrittenOffer.pdf", "-r", xPathToSaveIn)
            Call Pause(5)
    
            Kill xPathToSaveIn & glbLEE_ID & "_WrittenOffer.pdf"
        End If
        
        MDIMain.panHelp(0).FloodPercent = 90
    Else
        'Set the attached document keys
        If Not IsNull(rsHS("EC_DOCKEY")) Then
            glbDocKey = rsHS("EC_DOCKEY")
        Else
            glbDocKey = 0
        End If
        glbJob = rsHS("EC_CASE")
        
        'No Concerns PDF document to merge
        objMyForm.saveFile (xPathToSaveIn & glbLEE_ID & "_WSIBForm7.pdf")
        
        MDIMain.panHelp(0).FloodPercent = 80
    End If
    
    MDIMain.panHelp(0).FloodPercent = 90
    
    'If above is successfull, i.e. nReturn = 1, then save the completed Merged pdf into the Incident
    'Document Attachment table in the infoHR_DOC database as part of other incident documents.
    If nReturn = 1 Then
        'Save the document in the Incident Attachment
        glbDocName = "INCIDENT"
        xFileName = xPathToSaveIn & glbLEE_ID & "_WSIBForm7.pdf"
        xFileExtension = GetFileExtension(xFileName)
    
        'Create Incident Attachment Record to get the DOCNO - glbDocTmp
        'Add in HRDOC_HEALTH_SAFETY/TERM_HRDOC_HEALTH_SAFETY
        glbDocTmp = Add_Incident_Attachment_Record(glbJob, rsHS("EC_OCCDATE"))
        
        'Add in HRDOC_HEALTH_SAFETY2/TERM_HRDOC_HEALTH_SAFETY2
        'Add teh FRM9 code first if not existing
        Call CheckHRTABLCode("DOCT", "FRM7", "Form 7")
        'Release 8.0 - Grant permission to this Form 7 Document Type code for this user as well so the user can see the
        'Document of this Document Type
        Call Grant_DocumentTypeCode_Security(glbUserID, "FRM7", "Form 7")
        
        Call AppendIncident(glbLEE_ID, xFileName, xFileExtension, "FRM7", "FRM7 - " & Format(Now, "mm/dd/yyyy"))
    
        MDIMain.panHelp(0).FloodPercent = 95
    
        'Delete the WSIB Form 7 file now that as it has been saved into the Incident Attachment record
        If (Dir(xFileName)) <> "" Then
            'Call FileAttributeForm7(xFileName, "-r", xPathToSaveIn)
            'Call Pause(5)
            Kill xFileName
        End If
                
        'Close all the recordsets
        rsEMP.Close
        Set rsEMP = Nothing
        rsCompMst.Close
        Set rsCompMst = Nothing
        rsHS.Close
        Set rsHS = Nothing
        rsJOB.Close
        Set rsJOB = Nothing
        rsSal.Close
        Set rsSal = Nothing
        rsStatCat.Close
        Set rsStatCat = Nothing
        rsForm7Sec.Close
        Set rsForm7Sec = Nothing
        
        MDIMain.panHelp(0).FloodPercent = 100
        
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = ""
        Screen.MousePointer = DEFAULT
        
        MsgBox "WSIB Form 7 generated successfully and saved under 'Incident Documents' screen." & vbCrLf & vbCrLf & "Please go to the 'Incident Documents' screen to view or print Form 7.", vbInformation, "WSIB - Form 7 Generation"
    Else
        'Close all the recordsets
        rsEMP.Close
        Set rsEMP = Nothing
        rsCompMst.Close
        Set rsCompMst = Nothing
        rsHS.Close
        Set rsHS = Nothing
        rsJOB.Close
        Set rsJOB = Nothing
        rsSal.Close
        Set rsSal = Nothing
        rsStatCat.Close
        Set rsStatCat = Nothing
        rsForm7Sec.Close
        Set rsForm7Sec = Nothing
        
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = ""
        Screen.MousePointer = DEFAULT
        
        MsgBox "WSIB Form 7 cannot be generated.", vbCritical, "WSIB - Form 7 Generation"
    End If
    
End Sub

Private Function Round2DEC(tmpNUM) 'laura nov 10, 1997
Dim strNUM As String, X%

If glbCompDecHR <> 2 And glbCompDecHR <> 3 And glbCompDecHR <> 4 Then
    glbCompDecHR = 2  'THIS SHOULD NOT HAPPEN BUT IS A VALID DEFAULT
End If
Round2DEC = Round(tmpNUM, glbCompDecHR)

End Function

Private Function OpenDocument(zSQLQ, zName, xEmpNo) As Boolean ' As Long)
    
    On Error GoTo ErrHandler
    
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
    Dim FileExt As String
    Dim SQLQ
    Dim errPoint As String
    
    
    OpenDocument = False
    
    'Get the location to save the file in
    'TempDir = glbIHRREPORTS
    TempDir = GetComPreferEmail("WSIBFORM7PATH")
    If Len(Trim(TempDir)) = 0 Then
        TempDir = glbIHRREPORTS
    End If

    'Retrieve the document
    SQLQ = zSQLQ
    rsPHOTO.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic, adLockOptimistic
    
    'No document found
    If rsPHOTO.EOF Then Exit Function
    
    'Document found
    'Set the Extension of the document
    If IsNull(rsPHOTO("FILEEXT")) Then
        FileExt = ""
        TempFile = Replace(Trim(Replace(TempDir, Chr(0), "")) & "\" & xEmpNo & zName & ".tmp", "\\", "\")
    Else
        FileExt = rsPHOTO("FILEEXT")
        TempFile = Replace(Trim(Replace(TempDir, Chr(0), "")) & "\" & xEmpNo & zName & "." & FileExt & "", "\\", "\")
    End If
    
    FileNumber = FreeFile
        
    'Set the file attributes
    If (Dir(TempFile)) <> "" Then
        Call FileAttributeForm7(TempFile, "-r", TempDir)
        Call Pause(5)
        
        Kill TempFile
    End If
    
    'Open the document into a temp. file
    Open TempFile For Binary Access Write As FileNumber
    
    ReDim byteChunk(rsPHOTO("DOC").ActualSize)
    
    'Ticket #18176 - Since the .docx and xlsx are actually a zip file, when doing the following adds an
    'extra byte causing the file not to be opened in Word or Excel. By triming that extra byte seems to
    'resolve the issue. This trimmed byte does not cause any issue for normal doc, xls, or pdf file so
    'no condition has been added that it should trim for .docx and .xlsx only for now.
    'byteChunk() = rsPHOTO("DOC").GetChunk(rsPHOTO("DOC").ActualSize)
    byteChunk() = rsPHOTO("DOC").GetChunk(rsPHOTO("DOC").ActualSize - 1)
    
    Put FileNumber, , byteChunk()

    Close FileNumber
    'Kill (TempFile)
    rsPHOTO.Close
        
    'Read only
    Call FileAttributeForm7(TempFile, "+r", TempDir)
    
    'Open the attachment
    'Shell "cmd /c " & GetShortName(TempFile)

    OpenDocument = True
    
    Exit Function
    
ErrHandler:
    MsgBox Err.Description & " - " & FileNumber & ": " & TempFile & " - " & errPoint, , "Error"
    OpenDocument = False
End Function

Private Sub FileAttributeForm7(xFileName, xAttribute, xTempDir)
Dim TempFile2
Dim errPoint As String

    On Error GoTo ErrHandler_FileAttrib
            
    TempFile2 = Replace(Trim(Replace(xTempDir, Chr(0), "")) & "\IhrDoc.Bat", "\\", "\")
    
    Open TempFile2 For Output As #5
    
    Print #5, "attrib " & xAttribute & " " & GetShortName(xFileName)
    
    Close #5
    
    Shell "cmd /c " & GetShortName(TempFile2)
    

Exit Sub
    
ErrHandler_FileAttrib:
    MsgBox Err.Description & " - Attribute: " & xAttribute & " - " & xFileName & ": " & TempFile2 & " - " & errPoint, , "Error"
End Sub

Private Function Add_Incident_Attachment_Record(xCaseNo, xIncDate) As Integer
    Dim rsHSDoc As New ADODB.Recordset
    Dim SQLQ As String
    Dim xDocNo As Integer
    Dim xDocDesc As String
    Dim xFieldList1, xFieldList2 As String
    
    xFieldList1 = "DE_ID,DE_COMPNO,DE_EMPNBR,DE_CASE,DE_OCCDATE,DE_DOCNO,DE_DOCDESC,DE_FILEEXT,DE_TYPE,DE_LDATE,DE_LTIME,DE_LUSER"
    xFieldList2 = "DE_ID,DE_COMPNO,DE_EMPNBR,DE_CASE,DE_OCCDATE,DE_DOCNO,DE_DOCDESC,DE_FILEEXT,DE_TYPE,DE_LDATE,DE_LTIME,DE_LUSER,TERM_SEQ"
        
    If glbtermopen Then
        SQLQ = "SELECT " & xFieldList2 & " FROM Term_HRDOC_HEALTH_SAFETY"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " AND DE_CASE = " & xCaseNo
        SQLQ = SQLQ & " ORDER BY DE_DOCNO DESC"
        rsHSDoc.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "SELECT " & xFieldList1 & " FROM HRDOC_HEALTH_SAFETY"
        SQLQ = SQLQ & " WHERE DE_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND DE_CASE = " & xCaseNo
        SQLQ = SQLQ & " ORDER BY DE_DOCNO DESC"
        rsHSDoc.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    End If
    
    If rsHSDoc.EOF Then
        'No documents attached for the Case #
        xDocNo = 1
    Else
        'Other documents found for this Case #
        'Compute the Doc No
        rsHSDoc.MoveFirst
        xDocNo = rsHSDoc("DE_DOCNO") + 1
    End If
    
    'Add the new Incident Attachment record and so the next step of document attachment can be done using
    'the DE_DOCNO
    rsHSDoc.AddNew
    rsHSDoc("DE_COMPNO") = "001"
    rsHSDoc("DE_EMPNBR") = glbLEE_ID
    rsHSDoc("DE_CASE") = xCaseNo
    rsHSDoc("DE_DOCNO") = xDocNo
    If glbtermopen Then
        rsHSDoc("TERM_SEQ") = glbTERM_Seq
    End If
    
    If IsDate(xIncDate) Then
        rsHSDoc("DE_OCCDATE") = CVDate(Format(xIncDate, "mm/dd/yyyy"))
    End If
    
    xDocDesc = Left("Form 7 -" & Now, 30)
    rsHSDoc("DE_DOCDESC") = xDocDesc
    'rsHSDoc("DE_FILEEXT")
    rsHSDoc("DE_TYPE") = "INCIDENT"
    rsHSDoc("DE_LDATE") = Date
    rsHSDoc("DE_LTIME") = Time$
    rsHSDoc("DE_LUSER") = glbUserID
    rsHSDoc.Update
        
    rsHSDoc.Close
    Set rsHSDoc = Nothing
    
    Add_Incident_Attachment_Record = xDocNo
End Function

Private Sub Update_Form9_Fields(xOldClaimNo)
    Dim rsF9 As New ADODB.Recordset
    Dim SQLQ As String
    
    'Update the old Claim # with the new one in the Form 9 table
    If glbtermopen Then
        SQLQ = "SELECT F9_EMPNBR,F9_CASE,F9_WCBNBR,F9_WCBFDTE,F9_LDATE,F9_LTIME,F9_LUSER"
        SQLQ = SQLQ & " FROM Term_OHS_FORM9"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " AND F9_CASE = " & Data1.Recordset!EC_CASE
        SQLQ = SQLQ & " AND F9_WCBNBR = '" & xOldClaimNo & "'"
        rsF9.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "SELECT F9_EMPNBR,F9_CASE,F9_WCBNBR,F9_WCBFDTE,F9_LDATE,F9_LTIME,F9_LUSER"
        SQLQ = SQLQ & " FROM HR_OHS_FORM9"
        SQLQ = SQLQ & " WHERE F9_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND F9_CASE = " & Data1.Recordset!EC_CASE
        SQLQ = SQLQ & " AND F9_WCBNBR = '" & xOldClaimNo & "'"
        rsF9.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    
    If Not rsF9.EOF Then
        rsF9("F9_WCBNBR") = medClaimNo
        If IsDate(dlpDate(0).Text) Then
            rsF9("F9_WCBFDTE") = CVDate(dlpDate(0))
        End If
        
        rsF9("F9_LDATE") = Date
        rsF9("F9_LTIME") = Time$
        rsF9("F9_LUSER") = glbUserID
        
        rsF9.Update
    End If
    
    rsF9.Close
    Set rsF9 = Nothing
End Sub
