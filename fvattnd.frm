VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmVATTEND 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Attendance"
   ClientHeight    =   9780
   ClientLeft      =   165
   ClientTop       =   -1755
   ClientWidth     =   11805
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
   ScaleHeight     =   9780
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   5295
      LargeChange     =   315
      Left            =   11400
      Max             =   100
      SmallChange     =   315
      TabIndex        =   98
      Top             =   2280
      Width           =   300
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   39
      Top             =   9360
      Width           =   11805
      _Version        =   65536
      _ExtentX        =   20823
      _ExtentY        =   741
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
      Begin VB.CommandButton cmdSalaryChange 
         Appearance      =   0  'Flat
         Caption         =   "Salary Change"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5520
         TabIndex        =   116
         Top             =   0
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CommandButton cmdRecalPoints 
         Appearance      =   0  'Flat
         Caption         =   "Recalculate"
         Height          =   375
         Left            =   3630
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CommandButton cmdViewDiscip 
         Appearance      =   0  'Flat
         Caption         =   "Discipline Report"
         Height          =   375
         Left            =   1710
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CommandButton cmdAnother 
         Appearance      =   0  'Flat
         Caption         =   "+ &Another"
         Height          =   375
         Left            =   480
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8340
         Top             =   180
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
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fvattnd.frx":0000
      Height          =   1755
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "fvattnd.frx":0014
      TabIndex        =   0
      Top             =   450
      Width           =   11505
   End
   Begin Threed.SSPanel panEEName 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   11805
      _Version        =   65536
      _ExtentX        =   20823
      _ExtentY        =   767
      _StockProps     =   15
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Font3D          =   1
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
         Left            =   8160
         TabIndex        =   106
         Top             =   90
         Width           =   1305
      End
      Begin VB.Label lblEENum 
         Caption         =   "Label2"
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
         Left            =   1320
         TabIndex        =   42
         Top             =   90
         Width           =   1335
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
         Height          =   300
         Left            =   3060
         TabIndex        =   41
         Top             =   67
         Width           =   1740
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EEId"
         DataField       =   "AD_EMPNBR"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3120
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblEmpID 
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
         Height          =   240
         Left            =   5760
         TabIndex        =   37
         Top             =   120
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   36
         Top             =   105
         Width           =   1065
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6600
      Top             =   10320
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   2
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
   Begin VB.Frame fraDetail 
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   120
      TabIndex        =   48
      Top             =   2280
      Width           =   11535
      Begin VB.CommandButton cmdWFCHide 
         Caption         =   "Hide REG && OT"
         Height          =   375
         Left            =   9480
         TabIndex        =   115
         Top             =   80
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   270
         Left            =   10260
         TabIndex        =   112
         Top             =   2910
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtLEPoint 
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
         Left            =   7080
         MaxLength       =   20
         TabIndex        =   16
         Tag             =   "00-L/LE Point"
         Top             =   5670
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame frmMulti 
         Caption         =   "Position/Salary Information"
         Height          =   2745
         Left            =   5160
         TabIndex        =   55
         Top             =   0
         Width           =   4275
         Begin VB.CommandButton cmdPostion 
            Caption         =   "P&ositions"
            Height          =   255
            Left            =   90
            TabIndex        =   56
            Tag             =   "Postions"
            Top             =   300
            Width           =   975
         End
         Begin VB.TextBox txtDHRS 
            Appearance      =   0  'Flat
            DataField       =   "AD_DHRS"
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
            Left            =   1935
            TabIndex        =   22
            Tag             =   "00-Usual working hours per day"
            Top             =   1500
            Width           =   855
         End
         Begin VB.TextBox txtWHRS 
            Appearance      =   0  'Flat
            DataField       =   "AD_WHRS"
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
            Left            =   1935
            TabIndex        =   23
            Tag             =   "00- Number of hours in work week"
            Top             =   1800
            Width           =   975
         End
         Begin VB.ComboBox comPayPer 
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
            Left            =   1935
            TabIndex        =   21
            Tag             =   "Choose annum or hour"
            Top             =   1170
            Width           =   1215
         End
         Begin VB.TextBox txtPayrollID 
            Appearance      =   0  'Flat
            DataField       =   "AD_PAYROLL_ID"
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
            Left            =   1930
            MaxLength       =   25
            TabIndex        =   25
            Tag             =   "00-Payroll ID"
            Top             =   2400
            Width           =   1815
         End
         Begin INFOHR_Controls.CodeLookup clpJob 
            DataField       =   "AD_JOB"
            Height          =   285
            Left            =   1620
            TabIndex        =   18
            Tag             =   "01-Job Code"
            Top             =   270
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "n/a"
            MaxLength       =   25
            LookupType      =   5
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "AD_ORG"
            Height          =   285
            Index           =   0
            Left            =   1620
            TabIndex        =   19
            Tag             =   "00-Union Code"
            Top             =   570
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDOR"
         End
         Begin MSMask.MaskEdBox medSalary 
            DataField       =   "AD_SALARY"
            Height          =   285
            Left            =   1935
            TabIndex        =   20
            Tag             =   "00-Usual working Salary"
            Top             =   870
            Width           =   1305
            _ExtentX        =   2302
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
         Begin INFOHR_Controls.CodeLookup clpGLNo 
            DataField       =   "AD_GLNO"
            Height          =   285
            Left            =   1620
            TabIndex        =   24
            Tag             =   "00-General Ledger - Code"
            Top             =   2100
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   25
            LookupType      =   3
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Per"
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
            Left            =   180
            TabIndex        =   65
            Top             =   1230
            Width           =   300
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Salary"
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
            Left            =   180
            TabIndex        =   64
            Top             =   930
            Width           =   540
         End
         Begin VB.Label lblHrsWeek 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hours/Week"
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
            Left            =   180
            TabIndex        =   63
            Top             =   1830
            Width           =   930
         End
         Begin VB.Label lblHrsDay 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hours/Day"
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
            Left            =   180
            TabIndex        =   62
            Top             =   1530
            Width           =   780
         End
         Begin VB.Label lblSalCode 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "H/A"
            DataField       =   "AD_SALCD"
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
            Height          =   195
            Left            =   2790
            TabIndex        =   61
            Top             =   1260
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Index           =   9
            Left            =   180
            TabIndex        =   60
            Top             =   630
            Width           =   660
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
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
            Left            =   180
            TabIndex        =   59
            Top             =   330
            Width           =   780
         End
         Begin VB.Label lblPayID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payroll ID"
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
            Left            =   180
            TabIndex        =   58
            Top             =   2400
            Width           =   675
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "G/L #"
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
            Left            =   180
            TabIndex        =   57
            Top             =   2130
            Width           =   435
         End
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         DataField       =   "AD_SHIFT"
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
         Left            =   1545
         MaxLength       =   4
         TabIndex        =   8
         Tag             =   "00-Shift code"
         Top             =   1830
         Width           =   800
      End
      Begin VB.TextBox txtWSIB 
         Appearance      =   0  'Flat
         DataField       =   "AD_WCBNBR"
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
         Left            =   1545
         MaxLength       =   20
         TabIndex        =   11
         Tag             =   "00-Claim Number"
         Top             =   2130
         Width           =   1515
      End
      Begin VB.TextBox memComments 
         Appearance      =   0  'Flat
         DataField       =   "AD_COMM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Tag             =   "00-Comments"
         Top             =   3990
         Width           =   7755
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "AD_LDATE"
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
         Left            =   6540
         MaxLength       =   25
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   6420
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "AD_LTIME"
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
         Left            =   6990
         MaxLength       =   25
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   6390
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "AD_LUSER"
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
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   6420
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CheckBox chkEmeryTabl 
         Caption         =   "chkEmeryTabl"
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
         Left            =   7710
         TabIndex        =   51
         Top             =   6210
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox txtChrgCode 
         Appearance      =   0  'Flat
         DataField       =   "AD_CHRGCODE"
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
         Left            =   1545
         MaxLength       =   15
         TabIndex        =   6
         Tag             =   "00-Enter Charge Code"
         Top             =   1530
         Width           =   1275
      End
      Begin VB.TextBox txtPoint 
         Appearance      =   0  'Flat
         DataField       =   "AD_POINT"
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
         Left            =   1545
         MaxLength       =   20
         TabIndex        =   13
         Tag             =   "00-Point"
         Top             =   2430
         Width           =   1215
      End
      Begin VB.CheckBox chkUpload 
         Caption         =   "Upload Flag"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7620
         TabIndex        =   33
         Top             =   3660
         Width           =   1275
      End
      Begin VB.ComboBox comShiftType 
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
         Left            =   1560
         TabIndex        =   9
         Tag             =   "Choose Shift Type"
         Top             =   1830
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         DataField       =   "AD_DISCIPLINE"
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
         Left            =   1545
         MaxLength       =   4
         TabIndex        =   7
         Tag             =   "01-Counselling Type- Code"
         Top             =   5550
         Visible         =   0   'False
         Width           =   900
      End
      Begin Threed.SSCheck chkBackDated 
         Height          =   225
         Left            =   3750
         TabIndex        =   49
         Tag             =   "Hours to be added to employee's seniority."
         Top             =   30
         Visible         =   0   'False
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Back Dated "
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
      Begin INFOHR_Controls.CodeLookup clpChrgCode 
         Height          =   285
         Left            =   1230
         TabIndex        =   50
         Tag             =   "00-Enter Charge Code"
         Top             =   1530
         Visible         =   0   'False
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   5
      End
      Begin INFOHR_Controls.EmployeeLookup elpSupShow 
         Height          =   285
         Left            =   1230
         TabIndex        =   4
         Tag             =   "10-Employee Number of individual's supervisor"
         Top             =   930
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.DateLookup dlpToDate 
         Height          =   285
         Left            =   1230
         TabIndex        =   2
         Tag             =   "41-Date of Attendance"
         Top             =   330
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpReviewDate 
         DataField       =   "AD_DOA"
         Height          =   285
         Left            =   1230
         TabIndex        =   1
         Tag             =   "41-Date of Attendance"
         Top             =   30
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   503
         MultiSelect     =   -1  'True
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "AD_REASON"
         Height          =   285
         Index           =   1
         Left            =   1230
         TabIndex        =   3
         Tag             =   "01-Attendance Reason"
         Top             =   630
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ADRE"
      End
      Begin Threed.SSCheck chkIncident 
         DataField       =   "AD_INCID"
         Height          =   225
         Left            =   1560
         TabIndex        =   28
         Tag             =   "Is this a new incidence of illness?"
         Top             =   3660
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Incident"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCheck ChkInc 
         DataField       =   "AD_INDICATOR"
         Height          =   225
         Left            =   2640
         TabIndex        =   29
         Tag             =   "Incentive -  Attendance Management"
         Top             =   3660
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Incentive"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkSeniority 
         DataField       =   "AD_SEN"
         Height          =   225
         Left            =   3840
         TabIndex        =   30
         Tag             =   "Hours to be added to employee's seniority."
         Top             =   3660
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Seniority"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSCheck ChkFMLA 
         DataField       =   "AD_FMLA"
         Height          =   225
         Left            =   6660
         TabIndex        =   32
         Tag             =   "USA changes only"
         Top             =   3660
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "FMLA"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "AD_HRS"
         Height          =   285
         Left            =   1545
         TabIndex        =   5
         Tag             =   "11-Hours for this reason "
         Top             =   1230
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
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
      Begin Threed.SSCheck ChkEMELEA 
         DataField       =   "AD_EMELEA"
         Height          =   225
         Left            =   4860
         TabIndex        =   31
         Tag             =   "Emergency Leave"
         Top             =   3660
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Emergency Leave"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   1230
         TabIndex        =   10
         Tag             =   "00-Fund"
         Top             =   1830
         Visible         =   0   'False
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
      End
      Begin INFOHR_Controls.CodeLookup clpGLNum 
         Height          =   285
         Left            =   1230
         TabIndex        =   12
         Tag             =   "00-General Ledger - Code"
         Top             =   2130
         Visible         =   0   'False
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   3
      End
      Begin INFOHR_Controls.DateLookup dlpPayEndDate 
         Height          =   285
         Left            =   6630
         TabIndex        =   67
         Tag             =   "41-Date of Attendance"
         Top             =   0
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin MSMask.MaskEdBox medMachineRate 
         Height          =   285
         Left            =   7080
         TabIndex        =   27
         Tag             =   "00-Machine Rate"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "AD_MACHINE_NUM"
         Height          =   285
         Index           =   4
         Left            =   1230
         TabIndex        =   15
         Tag             =   "00-Machine #"
         Top             =   3030
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   15
         LookupType      =   12
      End
      Begin MSMask.MaskEdBox medMachineHours 
         Height          =   285
         Left            =   7080
         TabIndex        =   26
         Tag             =   "00-Machine Hours"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "AD_PROJECT_CODE"
         Height          =   285
         Index           =   3
         Left            =   1230
         TabIndex        =   14
         Tag             =   "00-Account Code"
         Top             =   2730
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   11
         LookupType      =   11
      End
      Begin VB.TextBox txtSup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         DataField       =   "AD_SUPER"
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
         Left            =   1790
         MaxLength       =   12
         TabIndex        =   66
         Tag             =   "10-Employee Number of individual's supervisor"
         Top             =   930
         Visible         =   0   'False
         Width           =   1275
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "AD_REGION"
         Height          =   285
         Index           =   5
         Left            =   1230
         TabIndex        =   17
         Tag             =   "00-Region"
         Top             =   3330
         Visible         =   0   'False
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin VB.Label lblRegion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
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
         TabIndex        =   114
         Top             =   3375
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label lblImport 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Attendance"
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
         Left            =   8715
         TabIndex        =   113
         Top             =   2925
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Image imgNoSec 
         Height          =   240
         Left            =   9840
         Picture         =   "fvattnd.frx":704B
         Top             =   2925
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comp. Time Outstanding"
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
         Left            =   3450
         TabIndex        =   111
         Top             =   5010
         Width           =   1845
      End
      Begin VB.Label lblCompTimeOS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   5070
         TabIndex        =   110
         Top             =   4995
         Width           =   750
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ASL Outstanding"
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
         Left            =   3450
         TabIndex        =   109
         Top             =   5250
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sick Outstanding"
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
         Left            =   3450
         TabIndex        =   108
         Top             =   4770
         Width           =   1380
      End
      Begin VB.Label lbCompTimeOSday 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
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
         Left            =   5940
         TabIndex        =   107
         Top             =   4995
         Width           =   285
      End
      Begin VB.Label lbldays 
         AutoSize        =   -1  'True
         Caption         =   "Day"
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
         Left            =   9750
         TabIndex        =   105
         Top             =   4980
         Width           =   285
      End
      Begin VB.Label lblEMLOSV 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   8940
         TabIndex        =   104
         Top             =   4980
         Width           =   675
      End
      Begin VB.Label lblEMLOS 
         AutoSize        =   -1  'True
         Caption         =   "Emergency Leave Outstanding"
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
         Left            =   6600
         TabIndex        =   103
         Top             =   4980
         Width           =   2190
      End
      Begin VB.Label lblMachineRate 
         BackStyle       =   0  'Transparent
         Caption         =   "Machine Rate"
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
         Left            =   5340
         TabIndex        =   102
         Top             =   3090
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblMachineHours 
         BackStyle       =   0  'Transparent
         Caption         =   "Machine Hours"
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
         Left            =   5340
         TabIndex        =   101
         Top             =   2820
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Code"
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
         Left            =   0
         TabIndex        =   100
         Top             =   2730
         Width           =   1020
      End
      Begin VB.Label lblMachine 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Machine #"
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
         TabIndex        =   99
         Top             =   3030
         Width           =   765
      End
      Begin VB.Label lblCNum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comp"
         DataField       =   "AD_COMPNO"
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
         Left            =   8310
         TabIndex        =   97
         Top             =   5520
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   96
         Top             =   1230
         Width           =   510
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Charge Code"
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
         TabIndex        =   95
         Top             =   1530
         Width           =   930
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
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
         TabIndex        =   94
         Top             =   1830
         Width           =   315
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Claim #"
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
         Index           =   24
         Left            =   0
         TabIndex        =   93
         Top             =   2130
         Width           =   525
      End
      Begin VB.Label lblIncident 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Incident"
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
         Height          =   300
         Left            =   2670
         TabIndex        =   92
         Top             =   1260
         Width           =   2295
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
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
         TabIndex        =   91
         Top             =   3990
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "AttSupervisor"
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
         TabIndex        =   90
         Top             =   930
         Width           =   945
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reason"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   89
         Top             =   630
         Width           =   660
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   88
         Top             =   30
         Width           =   885
      End
      Begin VB.Label lblOvertime 
         BackStyle       =   0  'Transparent
         Caption         =   "Overtime Bank"
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
         Left            =   30
         TabIndex        =   87
         Top             =   5010
         Width           =   1215
      End
      Begin VB.Label txtOvertime 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "x"
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
         Left            =   2160
         TabIndex        =   86
         Top             =   5010
         Width           =   495
      End
      Begin VB.Label lblOvertimeDays 
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
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
         Left            =   2760
         TabIndex        =   85
         Top             =   5010
         Width           =   375
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
         Index           =   13
         Left            =   0
         TabIndex        =   84
         Top             =   330
         Width           =   585
      End
      Begin VB.Label lblVACOSday 
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
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
         Left            =   2760
         TabIndex        =   82
         Top             =   4770
         Width           =   375
      End
      Begin VB.Label lblSICKOS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   5070
         TabIndex        =   81
         Top             =   4770
         Width           =   750
      End
      Begin VB.Label lblVACOS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   1890
         TabIndex        =   80
         Top             =   4770
         Width           =   750
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vacation Outstanding"
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
         Left            =   30
         TabIndex        =   79
         Top             =   4770
         Width           =   1620
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Point"
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
         Index           =   15
         Left            =   0
         TabIndex        =   78
         Top             =   2430
         Width           =   360
      End
      Begin VB.Label lbASLOSday 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
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
         Left            =   5940
         TabIndex        =   77
         Top             =   5235
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label lblASLOS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   5190
         TabIndex        =   76
         Top             =   5235
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblUpload 
         DataField       =   "AD_UPLOAD"
         Height          =   225
         Left            =   8820
         TabIndex        =   75
         Top             =   3060
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Disciplinary"
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
         Left            =   0
         TabIndex        =   74
         Top             =   5550
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
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
         Index           =   0
         Left            =   2730
         TabIndex        =   73
         Top             =   5580
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Emergency Leave Taken"
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
         Left            =   6600
         TabIndex        =   72
         Top             =   4770
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Label lblEMLDay 
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
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
         Left            =   9750
         TabIndex        =   71
         Top             =   4770
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblEMLTaken 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   8940
         TabIndex        =   70
         Top             =   4770
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "L/LE Point"
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
         Left            =   5340
         TabIndex        =   69
         Top             =   5700
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay End Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   22
         Left            =   5250
         TabIndex        =   68
         Top             =   30
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Image imgSec 
         Height          =   240
         Left            =   9840
         Picture         =   "fvattnd.frx":7195
         Top             =   2925
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblSICKOSday 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
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
         Left            =   5940
         TabIndex        =   83
         Top             =   4770
         Width           =   285
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   300
      Left            =   1980
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Day"
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
      Left            =   2760
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
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
      Left            =   8040
      TabIndex        =   43
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmVATTEND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbEEID&
Dim fglbSDate As Variant
Dim fUPMode As Integer ', fglbEmptyNew As Integer
Dim locEmpnbr, locDate, locReason, locHours
Dim locSupShow, locChrgCode, locShift, Answer
Dim locIncident, locSen, locInc, locFmla, locEMELEA, locUpload
Dim locWsib, locPoint
Dim locJob, locUnion, locSalary, locSalCode, locDHrs, locWHrs
Dim locComment
Dim locBackDated, locPayEndDate

Dim oldDate, oldReason, oldHours
Dim oldJob, oldSalary, oldSalCode, oldWhrs
Dim oldComment

Dim SavEML, SavVac, SavSick, AddChg, cntSick, savIncid, SavOutE, SavOutV, SavOutS, SaveHours, SavOutOT, SavOvt, SavOutCT, SavCT, SavMaxBank, SavOTBank
Dim SavVCOBank 'Ticket #14635
Dim SaveOTEmail As String
Dim Fdate, Tdate, fdateS, tdateS, OTFdate, OTTdate
Dim SavEnt, savEnt1, xAD, oldHrs
Dim savEntDate
Dim savReas
Dim SavSup, SavShift, SavDeac
Dim fglbRetry, xmedHours, xAnother
Dim xOldReason, xNewReason
Dim ChgOldNew, xDispODay, xDispODay01, HourGlb
Dim ReaOld, ReaNew, HoursOld, HoursNew, HoursBase 'For County of Brant
Dim AskWeekend, SkipWeekend, AskHoliday, AskHolidayAns, SkipHoliday
Dim fglbEMELEA As Integer
Dim fglbPoint
Dim fglbINC As Integer
Dim fglbSen As Integer
Dim fglbJobList As String
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim whsccExceedFlag As Boolean, whsccExceedOrgNum, whsccExceedRemNum
Dim whsccASLFlag As Boolean, whsccAnotherFlag As Boolean, glbxID
Dim RsATTwhscc As New ADODB.Recordset
Dim fglbNew As Integer
Dim xInciFlag As Boolean
Dim xCurSalary, xSalCD, xJob, xDHrs, xWHrs, xORG
Dim xDeptno, xAdminBy, xGLNO, xDiv
Dim xSDate
Dim xPayrollID
Dim xDiscipFlag As Boolean, xOccuAmount
Dim fglbDeleteble, fglbUpdateble, fglbAddable
Dim fglbSendOTEmail As Boolean 'Use a form global instead of a function
Dim fgotholidays As Boolean
Dim fglbESQLQ, fglbWSQLQ
Dim hlist As String
Dim xHideWFCAttCodes As Boolean

Private Enum ReadWrite
    RW_READ = 0
    RW_WRITE = 1
End Enum

Private Function ReCalcOT(xReaOld, xReaNew, xHoursOld, xHoursNew)
If glbCompSerial = "S/N - 2187W" Then  'City of Port Colborne
    Exit Function
End If
Dim xVal, xBase, xBaseTot
    xBase = HoursBase
    If xReaOld <> xReaNew Then
        xBase = HoursBase
        If UCase(xReaOld) = "OT" Or UCase(xReaOld) = "OT15" Or UCase(xReaOld) = "OT20" Or (glbCompSerial = "S/N - 2375W" And (UCase(xReaOld) = "OT05" Or UCase(xReaOld) = "OT25")) Then     'City of Timmins - Ticket #16168
            If Not (UCase(xReaNew) = "OT" Or UCase(xReaNew) = "OT15" Or UCase(xReaNew) = "OT20") And Not (glbCompSerial = "S/N - 2375W" And (UCase(xReaNew) = "OT05" Or UCase(xReaNew) = "OT25")) Then  'City of Timmins - Ticket #16168
                    xBase = HoursBase
            Else
                If UCase(xReaNew) = "OT" Then
                    xBase = HoursBase
                End If
                If UCase(xReaNew) = "OT15" Then
                    xBase = HoursBase * 1.5
                End If
                If UCase(xReaNew) = "OT20" Then
                    xBase = HoursBase * 2
                End If
            
                'City of Timmins - Ticket #16168
                If glbCompSerial = "S/N - 2375W" Then
                    If UCase(xReaNew) = "OT05" Then xBase = HoursBase * 0.5
                    If UCase(xReaNew) = "OT25" Then xBase = HoursBase * 2.5
                End If
            End If
        Else
            If (UCase(xReaNew) = "OT" Or UCase(xReaNew) = "OT15" Or UCase(xReaNew) = "OT20") Or (glbCompSerial = "S/N - 2375W" And (UCase(xReaNew) = "OT05" Or UCase(xReaNew) = "OT25")) Then   'City of Timmins - Ticket #16168
                If UCase(xReaNew) = "OT" Then
                    xBase = HoursBase
                End If
                If UCase(xReaNew) = "OT15" Then
                    xBase = HoursBase * 1.5
                End If
                If UCase(xReaNew) = "OT20" Then
                    xBase = HoursBase * 2
                End If
            
                'City of Timmins - Ticket #16168
                If glbCompSerial = "S/N - 2375W" Then
                    If UCase(xReaNew) = "OT05" Then xBase = HoursBase * 0.5
                    If UCase(xReaNew) = "OT25" Then xBase = HoursBase * 2.5
                End If
            End If
        End If
        If xBase > 0 Then
            medHours = xBase
        End If
        Call Set_Para_ForBrant
        Exit Function
    End If
    If xHoursOld <> xHoursNew Then
        If UCase(xReaNew) = "OT" Or UCase(xReaNew) = "OT15" Or UCase(xReaNew) = "OT20" Or (glbCompSerial = "S/N - 2375W" And (UCase(xReaNew) = "OT05" Or UCase(xReaNew) = "OT25")) Then 'City of Timmins - Ticket #16168
                If UCase(xReaNew) = "OT" Then
                    xBase = xHoursNew
                End If
                If UCase(xReaNew) = "OT15" Then
                    xBase = xHoursNew * 1.5
                End If
                If UCase(xReaNew) = "OT20" Then
                    xBase = xHoursNew * 2
                End If
                
                'City of Timmins - Ticket #16168
                If glbCompSerial = "S/N - 2375W" Then
                    If UCase(xReaNew) = "OT05" Then xBase = xHoursNew * 0.5
                    If UCase(xReaNew) = "OT25" Then xBase = xHoursNew * 2.5
                End If
                
                medHours = xBase
        End If
        Call Set_Para_ForBrant
        Exit Function
    
    End If
End Function

Private Sub Set_Para_ForBrant()
        ReaOld = clpCode(1).Text
        ReaNew = ReaOld
        HoursOld = medHours
        HoursNew = HoursOld
        If Len(Trim(HoursOld)) > 0 Then
            HoursBase = Val(HoursOld)
        Else
            HoursBase = 0
        End If
        'If UCase(ReaOld) = "OT" Then
        '    HoursBase = HoursOld
        'End If
        If UCase(ReaOld) = "OT15" Then
            HoursBase = Round((HoursBase / 1.5), 2)
        End If
        If UCase(ReaOld) = "OT20" Then
            HoursBase = Round((HoursBase / 2), 2)
        End If

    'City of Timmins - Ticket #16168
    If glbCompSerial = "S/N - 2375W" Then
        If UCase(ReaOld) = "OT05" Then HoursBase = Round((HoursBase / 0.5), 2)
        If UCase(ReaOld) = "OT25" Then HoursBase = Round((HoursBase / 2.5), 2)
    End If
End Sub

Private Function ValReasonCha(xOld, xNew)
Dim xVal, xBase, xBaseTot
    xVal = 0
    If xOld <> xNew Then
        If UCase(xOld) = "OT" Or UCase(xOld) = "OT10" Or UCase(xOld) = "OT15" Or UCase(xOld) = "OT20" Or (glbCompSerial = "S/N - 2375W" And (UCase(xOld) = "OT05" Or UCase(xOld) = "OT25")) Then 'City of Timmins - Ticket #16168
            If Not (UCase(xNew) = "OT" Or UCase(xNew) = "OT10" Or UCase(xNew) = "OT15" Or UCase(xNew) = "OT20") And Not (glbCompSerial = "S/N - 2375W" And (UCase(xNew) = "OT05" Or UCase(xNew) = "OT25")) Then  'City of Timmins - Ticket #16168
                ChgOldNew = 1
                xBaseTot = Val(medHours)
                If UCase(xOld) = "OT" Or UCase(xOld) = "OT10" Then
                    xBase = xBaseTot
                End If
                If UCase(xOld) = "OT15" Then
                    xBase = xBaseTot / 1.5
                End If
                If UCase(xOld) = "OT20" Then
                    xBase = xBaseTot / 2
                End If
                
                'City of Timmins - Ticket #16168
                If glbCompSerial = "S/N - 2375W" Then
                    If UCase(xOld) = "OT05" Then xBase = xBaseTot / 0.5
                    If UCase(xOld) = "OT25" Then xBase = xBaseTot / 2.5
                End If
                
                xDispODay = xBase - xBaseTot
                xDispODay01 = xBase - xBaseTot
                xVal = -xBaseTot
            Else
                ChgOldNew = 1
                xBaseTot = Val(medHours)
                If UCase(xOld) = "OT" Or UCase(xOld) = "OT10" Then
                    xBase = xBaseTot
                End If
                If UCase(xOld) = "OT15" Then
                    xBase = xBaseTot / 1.5
                End If
                If UCase(xOld) = "OT20" Then
                    xBase = xBaseTot / 2
                End If
    
                'City of Timmins - Ticket #16168
                If glbCompSerial = "S/N - 2375W" Then
                    If UCase(xOld) = "OT05" Then xBase = xBaseTot / 0.5
                    If UCase(xOld) = "OT25" Then xBase = xBaseTot / 2.5
                End If
    
                If UCase(xNew) = "OT" Or UCase(xNew) = "OT10" Then
                    xVal = xBase
                End If
                If UCase(xNew) = "OT15" Then
                    xVal = xBase * 1.5
                End If
                If UCase(xNew) = "OT20" Then
                    xVal = xBase * 2
                End If
                
                'City of Timmins - Ticket #16168
                If glbCompSerial = "S/N - 2375W" Then
                    If UCase(xNew) = "OT05" Then xVal = xBase * 0.5
                    If UCase(xNew) = "OT25" Then xVal = xBase * 2.5
                End If
                
                xVal = xVal - xBaseTot
                xDispODay01 = xVal 'xBase - xBaseTot
                xDispODay = xVal
            End If
        Else
            If (UCase(xNew) = "OT" Or UCase(xNew) = "OT10" Or UCase(xNew) = "OT15" Or UCase(xNew) = "OT20") Or (glbCompSerial = "S/N - 2375W" And (UCase(xNew) = "OT05" Or UCase(xNew) = "OT25")) Then  'City of Timmins - Ticket #16168
                ChgOldNew = 1
                xBaseTot = Val(medHours)
                
                If UCase(xNew) = "OT" Or UCase(xNew) = "OT10" Then
                    xBase = xBaseTot
                End If
                If UCase(xNew) = "OT15" Then
                    xBase = xBaseTot * 1.5
                End If
                If UCase(xNew) = "OT20" Then
                    xBase = xBaseTot * 2
                End If
                
                'City of Timmins - Ticket #16168
                If glbCompSerial = "S/N - 2375W" Then
                    If UCase(xNew) = "OT05" Then xBase = xBaseTot * 0.5
                    If UCase(xNew) = "OT25" Then xBase = xBaseTot * 2.5
                End If
                
                xDispODay = xBase
                xDispODay01 = xBase - xBaseTot
                xVal = xBase
            Else

            End If
       End If
    Else
        If UCase(clpCode(1).Text) = "OT" Or UCase(clpCode(1).Text) = "OT10" Or UCase(clpCode(1).Text) = "OT15" Or (glbCompSerial = "S/N - 2375W" And (UCase(clpCode(1).Text) = "OT05" Or UCase(clpCode(1).Text) = "OT25")) Then 'City of Timmins - Ticket #16168
            xBaseTot = Val(medHours)
            xVal = xBaseTot - HourGlb
        End If
    End If
    ValReasonCha = xVal
End Function

Private Function chkAttendance()
Dim SQLQ As String, Msg$, dd&, Response%
Dim DgDef As Variant, Title$, DCurPDate As Variant

chkAttendance = False

'''On Error GoTo chkPerH_Err

lblSalCode = Left(comPayPer, 1)

'Ticket #25268: Check first if this Attendance record that was created by ESS as Approved request
If AddChg = "C" Then
    If Is_ESSApproved_Record(rsDATA!AD_ATT_ID) Then
        'Check if Date, Reason or Hours have changed. Do not allow these three fields to change
        If oldDate <> dlpReviewDate Or oldReason <> clpCode(1) Or oldHours <> medHours Then
            'MsgBox "This record was added from ESS - Request Approval. The 'From Date', 'Reason' or 'Hours' cannot be changed from info:HR.", vbExclamation, "ESS Approved Attendance Record"
            'dlpReviewDate.SetFocus
            'Exit Function
            Msg$ = "This record was added from ESS - Request Approval." & vbCrLf & vbCrLf & "Are you sure you want to change 'From Date', 'Reason' or 'Hours' from info:HR?"
            Response% = MsgBox(Msg, 36, "Confirm change: ESS Approved Attendance Record")
            If Response% <> 6 Then Exit Function
        End If
    End If
End If

'Ticket #26576 - WDGPHU - Cannot add FX* codes from Attendance
If glbCompSerial = "S/N - 2411W" And AddChg = "A" And UCase(Left(clpCode(1), 2)) = "FX" And gsFLEX_LOGIC Then
    If clpCode(1).Text = "FX+Y" Then 'Ticket #27771 Franks 12/01/2015 - allow user to enter FX+Y manually
        'ok with this code
    Else
        MsgBox "You cannot add Flex Time Attendance from here. Please use ESS Module.", vbExclamation, "info:HR - Flex Time entry restricted"
        Exit Function
    End If
End If

'Ticket #30305 - Disable Compensatory Time Entries
If gsDISABLE_COMPTIME Then
    If Left(clpCode(1).Text, 2) = "OT" Or Left(clpCode(1).Text, 2) = "CT" Then
        MsgBox "You cannot maintain Compensatory Time Attendance records from here. Please use ESS Module.", vbExclamation, "info:HR - Compensatory Time entry restricted"
        Exit Function
    End If
End If

If elpSupShow.Text = "0" Then elpSupShow.Text = ""
If Len(elpSupShow.Text) > 0 Then
    If elpSupShow.Caption = "Unassigned" Then
        elpSupShow.SetFocus
        MsgBox "Invalid Supervisor Code"
        Exit Function
    End If
End If

If Len(dlpReviewDate.Text) < 1 Then
    Msg$ = "Date is required"
    dlpReviewDate.SetFocus
    MsgBox Msg$
    Exit Function
Else
    If Not IsDate(dlpReviewDate.Text) Then
        Msg$ = "Not a Valid Date"
        dlpReviewDate.SetFocus
        MsgBox Msg$
        Exit Function
    End If
End If
If Len(dlpToDate.Text) > 0 And dlpToDate.Enabled Then
    If Not IsDate(dlpToDate.Text) Then
        Msg$ = "Not a Valid Date"
        dlpToDate.SetFocus
        MsgBox Msg$
        Exit Function
    End If
End If
If IsDate(dlpToDate.Text) And IsDate(dlpReviewDate.Text) Then
    If DateDiff("d", dlpToDate.Text, dlpReviewDate.Text) > 0 Then
        Msg$ = "From date must be earlier than To Date"
        dlpReviewDate.SetFocus
        MsgBox Msg$
        Exit Function
    End If
End If
If chkBackDated Then
    If Len(dlpPayEndDate) = 0 Then
        MsgBox "Pay End Date must be entered if Back Dated is checked"
        dlpPayEndDate.SetFocus
        Exit Function
    End If
    If Not IsDate(dlpPayEndDate) Then
        MsgBox "Invalid Pay End Date"
        dlpPayEndDate.SetFocus
        Exit Function
    End If
End If
If Len(clpCode(1)) > 0 Then
    If clpCode(1).Caption = "Unassigned" Then
        MsgBox "Invalid Reason code"
        clpCode(1).SetFocus
        Exit Function
    End If
Else
    MsgBox "Reason code is required"
    clpCode(1).SetFocus
    Exit Function
End If
clpCode(1).Text = Trim(clpCode(1).Text)

If Len(medHours) <= 0 Then
    MsgBox "Hours is required"
    medHours.SetFocus
    Exit Function
Else
    If Not IsNumeric(medHours) Then
        MsgBox "Hours is invalid"
        medHours.SetFocus
        Exit Function
    End If
    If Val(medHours) > 99999.9999 Then
        MsgBox "Hours is Invalid"
        medHours.SetFocus
        Exit Function
    End If
End If

If glbMulti Or glbCompSerial = "S/N - 2192W" Or glbCompSerial = "S/N - 2366W" Then
    If Len(Trim(txtDHRS)) = 0 Then txtDHRS = 0
    If Len(Trim(txtWHRS)) = 0 Then txtWHRS = 0
    If Len(Trim(medsalary)) = 0 Then medsalary = 0
    If Not IsNumeric(medsalary) Then
        MsgBox "Salary invalid"
        medsalary.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtDHRS) Then
        MsgBox "Hours/Day is invalid"
        txtDHRS.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtWHRS) Then
        MsgBox "Hours/Week is invalid"
        txtWHRS.SetFocus
        Exit Function
    End If
    If glbVadim Then
        If Val(medsalary) = 0 Then
            MsgBox "Salary is required field"
            medsalary.SetFocus
            Exit Function
        End If
        If lblSalCode = "A" And Val(txtWHRS) = 0 Then
            MsgBox "Hours/Week is required field if Annual Salary is entered"
            txtWHRS.SetFocus
            Exit Function
        End If
    End If
    'Franks Sep 11,02 for WHSCC
    If Len(clpCode(0)) > 0 Then
        If clpCode(0).Caption = "Unassigned" Then
            MsgBox "Invalid Union code"
            clpCode(0).SetFocus
            Exit Function
        End If
    Else
        If glbWHSCC And clpCode(1) = "USB" Then
            MsgBox "Union code is required if Reason is 'USB'"
            clpCode(0).SetFocus
            Exit Function
        End If
    End If
    'Franks Sep 11,02 for WHSCC
End If

'Casey House
If glbCompSerial = "S/N - 2214W" Then
    If UCase(clpCode(1).Text) = "D2" And Len(Trim(clpChrgCode.Text)) = 0 Then
        MsgBox "If Reason Code is 'D2', Department must be entered"
        clpChrgCode.SetFocus
        Exit Function
    End If
    If Len(clpChrgCode) > 0 Then
        If clpChrgCode.Caption = "Unassigned" Then
            MsgBox "If Department is entered it must be valid"
            clpChrgCode.SetFocus
            Exit Function
        End If
    End If
    If Len(clpCode(2)) > 0 Then
        If clpCode(2).Caption = "Unassigned" Then
            MsgBox "Invalid Fund code"
            clpCode(2).SetFocus
            Exit Function
        End If
    End If
    If Len(clpGLNum) > 0 Then
        If clpGLNum.Caption = "Unassigned" Then
            MsgBox lStr("If G/L Number is entered it must be valid")
            clpGLNum.SetFocus
            Exit Function
        End If
    End If
End If

'Ticket #17323 - Oshawa CHC
If glbCompSerial = "S/N - 2396W" Then
    If Len(clpChrgCode) > 0 Then
        If clpChrgCode.Caption = "Unassigned" Then
            MsgBox "If G/L # is entered it must be valid"
            clpChrgCode.SetFocus
            Exit Function
        End If
    End If
End If

If glbCompSerial = "S/N - 2411W" Then 'WDGPHU - Ticket #24655
    If Len(clpCode(5)) > 0 Then
        If clpCode(5).Caption = "Unassigned" Then
            MsgBox lStr("Invalid Region code")
            clpCode(5).SetFocus
            Exit Function
        End If
    End If
End If

'Ticket #27771 Franks 12/01/2015
If glbCompSerial = "S/N - 2411W" And (AddChg = "A" Or AddChg = "C") And clpCode(1).Text = "FX+Y" Then
    If Not WDGPH_Check_FX(glbLEE_ID, "FX-X", CVDate(dlpReviewDate.Text), medHours.Text) Then   '(xEmpNo, xCode, xDate, xHours)
        Msg$ = "No 'FX-X' Attendance record found with the same From Date and Hours. " & Chr(10) & "Cannot enter 'FX+Y' Attendance. "
        MsgBox Msg$
        Exit Function
    End If
    
    'Ticket #29306 - For FX-X -ve hours, there should be FX+Y +ve hours
    If medHours < 0 Then
        MsgBox "Invalid 'FX+Y' Hours. It must be Positive Hours"
        medHours.SetFocus
        Exit Function
    End If
    
End If

chkAttendance = True
Call Payroll_Integration

Exit Function

chkPerH_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkPerf", "HR Attendance", "edit/Add")
Call RollBack '28July99 js

End Function
Private Sub Payroll_Integration(Optional UptType)
'NOT IN USE IN VB CODE, HAS TRIGGER IN THE DATABASE
'Dim HRAtts As New Collection
'Dim UpdPayroll As Boolean
'
'If xAttendance = "Attendance_History" Or glbtermopen Then Exit Sub
'
'If IsMissing(UptType) Then
'    If isChanged_Attendance(HRAtts, oldDate, dlpReviewDate) Then UpdPayroll = True
'    If isChanged_Attendance(HRAtts, oldReason, clpCode(1)) Then UpdPayroll = True
'    If isChanged_Attendance(HRAtts, oldHours, medHours) Then UpdPayroll = True
'    If isChanged_Attendance(HRAtts, oldComment, memComments) Then UpdPayroll = True
'    If isChanged_Attendance(HRAtts, oldJob, clpJob) Then UpdPayroll = True
'    If isChanged_Attendance(HRAtts, oldSalary, medSalary) Then UpdPayroll = True
'    If isChanged_Attendance(HRAtts, oldSalCode, lblSalCode) Then UpdPayroll = True
'    If isChanged_Attendance(HRAtts, oldWhrs, txtWHRS) Then UpdPayroll = True
'    If fglbNew Then
'        UptType = "A"
'    Else
'        UptType = "M"
'    End If
'ElseIf UptType = "D" Then
'    If isChanged_Attendance(HRAtts, "", Data1.Recordset("AD_REASON")) Then UpdPayroll = True
'    If isChanged_Attendance(HRAtts, 0, medHours) Then UpdPayroll = True
'    If isChanged_Attendance(HRAtts, "", dlpReviewDate) Then UpdPayroll = True
'    UpdPayroll = True
'End If
'If UpdPayroll Then
'    Call Passing_Attendance_Changes(HRAtts, UptType, glbLEE_ID)
'End If

End Sub

Private Function ChkTermStatus()
Dim SQLQ, Logx, Msg$, SavReviewDate, Result
Dim rsTB As New ADODB.Recordset
Dim xFLAG As Boolean, DgDef As Variant
    xFLAG = True
    If glbtermopen Then
        ChkTermStatus = xFLAG
        Exit Function
    End If
    SQLQ = "SELECT ED_EMPNBR,ED_EMP FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
    If rsTB.State <> 0 Then rsTB.Close
    rsTB.Open SQLQ, gdbAdoIhr001
    If Not rsTB.EOF Then
        If rsTB("ED_EMP") = "TERM" Then
            xFLAG = False
        End If
    End If
    rsTB.Close
    If Not xFLAG Then
        DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
        Msg$ = "You are entering Attendance for a terminated employee " & Chr(10)
        Msg$ = Msg$ & "Do you wish to continue? " & Chr(10)
        Msg$ = Msg$ & "      Click Yes to accept this Attendance record." & Chr(10)
        Msg$ = Msg$ & "      Click No to edit it again." & Chr(10)
        Result = MsgBox(Msg$, DgDef, "Terminated Employee")
        If Not (Result = IDNO) Then
            xFLAG = True
        End If
    End If
    ChkTermStatus = xFLAG
End Function
Private Function ChkDup()
Dim SQLQ, Logx, Msg$, SavReviewDate
Dim rsTB As New ADODB.Recordset

ChkDup = False

Logx = False
SavReviewDate = dlpReviewDate.Text
If glbtermopen Then
    SQLQ = "SELECT AD_EMPNBR FROM Term_ATTENDANCE WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & glbLEE_ID
End If
SQLQ = SQLQ & " AND AD_REASON = '" & clpCode(1).Text & "'"
SQLQ = SQLQ & " AND AD_DOA = " & Date_SQL(dlpReviewDate.Text)
SQLQ = SQLQ & " AND AD_HRS = " & medHours   'Hemu - Jerry said it should check with hours as well.
If AddChg <> "A" Then
    SQLQ = SQLQ & " AND AD_ATT_ID <> " & Data1.Recordset("AD_ATT_ID")
End If
If xAD <> "AD" Then
    SQLQ = Replace(SQLQ, "AD_", "AH_")
    SQLQ = Replace(SQLQ, "HR_ATTENDANCE", "HR_ATTENDANCE_HISTORY")
End If
If glbtermopen Then
    rsTB.Open SQLQ, gdbAdoIhr001X
Else
    rsTB.Open SQLQ, gdbAdoIhr001
End If
If Not rsTB.EOF Then Logx = True
rsTB.Close

'Franks Aug 08,02 T#2544
If glbWFC Then
    Dim xtot
    xtot = Val(medHours)
    If glbtermopen Then
        SQLQ = "SELECT AD_EMPNBR,AD_HRS FROM Term_ATTENDANCE WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = "SELECT AD_EMPNBR,AD_HRS FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & glbLEE_ID
    End If
    SQLQ = SQLQ & " AND AD_DOA = " & Date_SQL(dlpReviewDate.Text) & " "
    If AddChg <> "A" Then
        SQLQ = SQLQ & " AND AD_ATT_ID <> " & Data1.Recordset("AD_ATT_ID")
    End If
    If xAD <> "AD" Then
        SQLQ = Replace(SQLQ, "AD_", "AH_")
        SQLQ = Replace(SQLQ, "HR_ATTENDANCE", "HR_ATTENDANCE_HISTORY")
    End If
    If glbtermopen Then
        rsTB.Open SQLQ, gdbAdoIhr001X
    Else
        rsTB.Open SQLQ, gdbAdoIhr001
    End If

    If Not rsTB.EOF Then
        Do While Not rsTB.EOF
            If xAD = "AD" Then
                xtot = xtot + rsTB("AD_HRS")
            Else
                xtot = xtot + rsTB("AH_HRS")
            End If
            rsTB.MoveNext
        Loop
        rsTB.MoveFirst
    End If
    If xtot > SaveHours Then
        Msg$ = "Warning: The total number of hours entered for " & SavReviewDate & "  exceeds " & Chr(10)
        Msg$ = Msg$ & "the employee's hours per day as defined on their Position History screen" & Chr(10)
        Msg$ = Msg$ & "Click OK to accept this record. Otherwise, Click on Cancel to edit it again"
        Answer = MsgBox(Msg$, 1)
        If Answer = IDCANCEL Then
            dlpToDate.Text = ""
            Exit Function
        Else
            ChkDup = True
        End If
    End If
End If
'Franks Aug 08,02 T#2544

If Logx = True Then
    Msg$ = "Duplicate exist. OK to proceed"
    Answer = MsgBox(Msg$, 1)
    If Answer = IDCANCEL Then
''        dlpReviewDate.Text = ""
''        dlpReviewDate.SetFocus
''        'vbxTrueGrid(0).SetFocus    'Hemu - 07/02/2003 Commented - giving error and its not required
''        dlpToDate.Text = ""
        Call cmdCancel_Click
        Exit Function
    Else
        ChkDup = True
    End If
Else
ChkDup = True
End If

End Function

Private Sub chkBackDated_Click(Value As Integer)
If chkBackDated And chkBackDated.Enabled Then
    lblTitle(22).Visible = True
    dlpPayEndDate.Visible = True
Else
    lblTitle(22).Visible = False
    dlpPayEndDate.Visible = False
End If
End Sub

Private Sub ChkEMELEA_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub ChkFMLA_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub ChkInc_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkIncident_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkSeniority_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpChrgCode_LostFocus()
txtChrgCode = clpChrgCode
End Sub

Private Sub clpCode_Change(Index As Integer)
Call ATTCode_Desc(Index)
End Sub

Private Sub clpCode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub
Private Sub clpGLNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub
Private Sub clpGLNum_LostFocus()
    If glbCompSerial = "S/N - 2214W" Then
        txtWSIB = clpGLNum
    End If
End Sub

Private Sub clpJob_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub

Private Sub cmdAnother_Click()
Dim X%, Msg$, xSup
Dim DgDef As Double
Dim SQLQ, xID


Dim rsATT As New ADODB.Recordset
'''On Error GoTo Add_Err
cmdAnother.Enabled = False
If Not chkAttendance() Then GoTo flgEXIT

'Hemu - 06/23/2003 Begin - Skip AskWeekend Ticket # 4360
If glbCompSerial = "S/N - 2190W" Then
    AskWeekend = False
    SkipWeekend = False
End If
'Hemu - 06/23/2003 End

'Added by Bryan 31/Oct/05 Ticket#9648
'When doing "Another" the old hours should not be what they started with but what the
'last value was.
HoursOld = medHours
'end Bryan

'Ticket #18188

 
'If glbCompSerial = "S/N - 2418W" Then
    'If AskHoliday Then
       
        If fgotholidays = False Then
            Dim d2 As Date
            
            If IsDate(dlpToDate.Text) Then
                d2 = dlpToDate.Text
            Else
                d2 = dlpReviewDate.Text
            End If
            hlist = IsSTATHoliday(dlpReviewDate.Text, d2)
            fgotholidays = True
        End If
        If InStr(hlist, Date_SQL(dlpReviewDate.Text)) > 0 Then
            If AskHoliday Then
                Msg$ = "Do you want exclude STAT holidays?"
                AskHoliday = False
                SkipHoliday = False
                X% = MsgBox(Msg$, 36)
                AskHolidayAns = X%
            End If
            If AskHolidayAns = 6 Then
                SkipHoliday = True
                dlpReviewDate.Text = DateAdd("d", 1, dlpReviewDate.Text)
                Exit Sub
            End If
        End If
    'End If
'End If

If Weekday(dlpReviewDate.Text) = 7 Or Weekday(dlpReviewDate.Text) = 1 Then
    If AskWeekend Then
        Msg$ = "Do you want exclude Saturday/Sunday?"
        AskWeekend = False
        SkipWeekend = False
        X% = MsgBox(Msg$, 36)
        If X% = 6 Then
            SkipWeekend = True
            dlpReviewDate.Text = DateAdd("d", 1, dlpReviewDate.Text)
            Exit Sub
        End If
    Else
        If SkipWeekend Then
            dlpReviewDate.Text = DateAdd("d", 1, dlpReviewDate.Text)
            Exit Sub
        End If
    End If
End If

'Ticket #25500 - just for EntEccess() to account for OT15 and OT20 - moved from down up here
' For City of Niagara Fulls
If xAnother = 1 Then
    xmedHours = Val(medHours)
End If
If glbNiagaraFulls And (Not glbtermopen) Then
    Select Case UCase(clpCode(1).Text)
        Case "OT15"
            medHours = xmedHours * 1.5
        Case "OT20"
            medHours = xmedHours * 2  'Val(medHours) * 2
    End Select
End If

If Not EntEccess() Then GoTo flgEXIT
'Hemu - EML
'If glbSQL Or glbOracle Then If Not EmlEntEccess() Then GoTo flgEXIT
If Not EmlEntEccess() Then GoTo flgEXIT
'Hemu

If glbWHSCC Then
    whsccAnotherFlag = True
    If Not EntStatEccess(glbLEE_ID, CVDate(dlpReviewDate)) Then Exit Sub
    If clpCode(1) = "ASL" Then
        If Not EntASLEccess(glbLEE_ID, CVDate(dlpReviewDate), medHours) Then Exit Sub
    End If
    If clpCode(1) = "USB" Then
        If Not EntUSBEccess(glbLEE_ID, CVDate(dlpReviewDate), medHours) Then Exit Sub
    End If
End If

'Surrey Place to check if the Employee Status is "TERM"
If glbCompSerial = "S/N - 2347W" And AddChg = "A" Then
    If Not ChkTermStatus Then GoTo flgEXIT
End If
If AddChg = "A" Then
    If Not ChkDup() Then GoTo flgEXIT
End If

'Ticket #26604 - Moved Up
'For City of Niagara Fulls
'If xAnother = 1 Then
'    xmedHours = Val(medHours)
'End If
'If glbNiagaraFulls And (Not glbtermopen) Then
'    Select Case UCase(clpCode(1).Text)
'        Case "OT15"
'            medHours = xmedHours * 1.5
'        Case "OT20"
'            medHours = xmedHours * 2  'Val(medHours) * 2
'    End Select
'End If

'V7.6
'Current Salary to AD_SALARY FOR Casey House
If glbCompSerial = "S/N - 2214W" Then
    If Len(txtChrgCode) = 0 Then
        If Len(xDeptno) > 0 Then
            clpChrgCode = xDeptno: txtChrgCode = xDeptno
        End If
    End If
    If Len(txtShift) = 0 Then
        If Len(xAdminBy) > 0 Then
            clpCode(2) = xAdminBy: txtShift = xAdminBy
        End If
    End If
    If Len(txtWSIB) = 0 Then
        If Len(xGLNO) > 0 Then
            clpGLNum = xGLNO: txtWSIB = xGLNO
        End If
    End If
End If
    
    If Not IsNumeric(medsalary) Then
        medsalary = xCurSalary
    Else
        If Val(medsalary) = 0 Then
        medsalary = xCurSalary
        End If
    End If
    If Len(lblSalCode) = 0 Then
        If Len(xSalCD) > 0 Then lblSalCode = xSalCD
    End If
    If Len(clpJob) = 0 Then
        If Len(xJob) > 0 Then clpJob = xJob
    End If
    If Len(clpCode(0)) = 0 Then
        If Len(xORG) > 0 Then clpCode(0) = xORG
    End If
    If Len(txtDHRS) = 0 Then
        If xDHrs > 0 Then txtDHRS = xDHrs
    End If
    If Len(txtWHRS) = 0 Then
        If xWHrs > 0 Then txtWHRS = xWHrs
    End If
    


locSupShow = elpSupShow.Text
locShift = txtShift
If Trim(elpSupShow.Text) = "" And Len(SavSup) > 0 Then elpSupShow.Text = ShowEmpnbr(SavSup)
If glbCompSerial = "S/N - 2394W" Then ' St. John's - Ticket #15053
    If Trim(txtChrgCode) = "" And Len(SavShift) > 0 Then txtChrgCode = SavShift
Else
    If Trim(txtShift) = "" And Len(SavShift) > 0 Then txtShift = SavShift
End If

'If chkIncident = True Then cntSick = cntSick + 1: UpdateIncCount RW_WRITE

'lblIncident.Caption = "Total # of Incidents = " & Str(cntSick)

If Left(clpCode(1).Text, 3) = "VAC" Then
    If DateValue(dlpReviewDate.Text) >= Fdate And DateValue(dlpReviewDate.Text) <= Tdate Then
        SavVac = SavVac + medHours
    End If
    SavOutV = SavOutV - SavVac
End If
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #14635
    If Left(clpCode(1).Text, 3) = "VCO" Then
        If DateValue(dlpReviewDate.Text) >= Fdate And DateValue(dlpReviewDate.Text) <= Tdate Then
            SavVCOBank = SavVCOBank + medHours
        End If
    End If
End If
If Left(clpCode(1).Text, 3) = "SIC" Then
    If DateValue(dlpReviewDate.Text) >= fdateS And DateValue(dlpReviewDate.Text) <= tdateS Then
        SavSick = SavSick + medHours
    End If
    SavOutS = SavOutS - SavSick
End If

'Town of Aurora
'If glbCompSerial = "S/N - 2378W" Then
    If Left(clpCode(1).Text, 2) = "OT" Then
        If DateValue(dlpReviewDate.Text) >= OTFdate And DateValue(dlpReviewDate.Text) <= OTTdate Then
            SavOvt = SavOvt + medHours
        End If
    ElseIf Left(clpCode(1).Text, 2) = "CT" Then
        If DateValue(dlpReviewDate.Text) >= OTFdate And DateValue(dlpReviewDate.Text) <= OTTdate Then
            SavCT = SavCT + medHours
        End If
    End If
'End If

'Hemu - EML
'If glbSQL Or glbOracle Then
'Hemu
    If chkEMELEA Then
        If DateValue(dlpReviewDate.Text) >= CVDate(GetMonth("Jan") & " 1," & Year(Date)) And DateValue(dlpReviewDate.Text) <= CVDate(GetMonth("Dec") & " 31," & Year(Date)) Then
        'linamar stuff added by Bryan 13/10/05 Ticket#9264
            If Not glbLinamar Then
                SavEML = SavEML + medHours
            Else
                SavEML = SavEML + 1
            End If
        End If
        SavOutE = SavOutE - SavEML
    End If
'End If

AddChg = "NO"

If glbCompSerial = "S/N - 2173W" Then
    Call Check_Overtime_Bank(True)
End If

UPDVACSICK True
'Hemu - EML
'If glbSQL Or glbOracle Then UPDEML True
Call Calculate_EML_Taken

'Calculate Comp Time Outstanding - Ticket #17345
Call Calculate_Outstanding_CompTime

'Hemu
UPDHRENTIT clpCode(1).Text, medHours

locEmpnbr = lblEEID
locDate = dlpReviewDate.Text
locReason = clpCode(1).Text
locHours = medHours
locChrgCode = clpChrgCode.Text
locComment = memComments

locIncident = chkIncident
locSen = chkSeniority
locEMELEA = chkEMELEA
locInc = ChkInc
locFmla = ChkFMLA
locWsib = txtWSIB
locPoint = txtPoint
locUpload = lblUpload

locJob = clpJob.Text
locUnion = clpCode(0).Text
locSalary = medsalary
locSalCode = lblSalCode
locDHrs = txtDHRS
locWHrs = txtWHRS
locBackDated = chkBackDated
locPayEndDate = dlpPayEndDate

Screen.MousePointer = HOURGLASS

If glbtermopen Then
   rsDATA!TERM_SEQ = glbTERM_Seq
End If

Call UpdUStats(Me)

Call UpdCodes(xAD) 'Ticket #28846 Franks 08/16/2016

Call Set_Control("U", Me, rsDATA)
xSup = getEmpnbr(elpSupShow.Text)
If Val(xSup) = 0 Then
    rsDATA!AD_SUPER = Null
Else
    rsDATA!AD_SUPER = xSup
End If
xDiscipFlag = False
'Disciplinary action for Whitby, Check how many Incident flags
If glbWFC And glbPlantCode = "WHBY" And xAD = "AD" Then 'And AddChg = "A" Then
    Call WhitbyGetIncidentFlags(lblEmpID, CVDate(dlpReviewDate), clpCode(1))
End If
gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans
xID = rsDATA("AD_ATT_ID")

'If glbWFC Then 'Ticket #24124 Franks 07/24/2013
If glbWFC Or glbMitchellPlastics Then 'Ticket #24112 Franks 07/30/2013
    Call WFC_Attend_To_AT(glbLEE_ID, "M", rsDATA("AD_DOA"), rsDATA("AD_REASON"), xID)
End If

'Hemu - EML
Call Calculate_EML_Taken
'Hemu

'Calculate Comp Time Outstanding - Ticket #17345
Call Calculate_Outstanding_CompTime

'Ticket #28207 - The exceeding prompt was not coming up when a date range of data entry is done for OT.
'The recalculate of the OT will keep track of the exceeding hours and prompt accordingly.
Call UPDOVERTIME

If glbWHSCC Then
    If whsccExceedFlag And Not glbtermopen Then
        If IsDate(dlpToDate) Then
            glbxID = xID
            Call whsccASLatt
            xID = glbxID
            locHours = whsccExceedOrgNum
        Else
            locReason = "ASL"
            locHours = whsccExceedRemNum
            glbxID = xID
            Call whsccASLatt
            xID = glbxID
            whsccExceedFlag = False
        End If
    End If
    
    If clpCode(1) = "ASL" Then
        Call UpdateASL(lblEmpID, CVDate(dlpReviewDate), medHours, "U")
    End If
    If clpCode(1) = "USB" Then
        Call ReCalcUSB("AD_EMPNBR = " & lblEmpID, "EMP")
    End If
End If

'Create Disciplinary action for Whitby
If xDiscipFlag Then
    If glbWFC And glbPlantCode = "WHBY" And xAD = "AD" Then
        Call WhitbyUpdateDisciplinary(lblEmpID, CVDate(dlpReviewDate), clpCode(1)) ', medHours, "U")
    End If
End If

Data1.Refresh
DoEvents

If Not IsNull(xID) Then
    Data1.Recordset.Find "AD_ATT_ID=" & xID
End If
If Not Data1.Recordset.EOF Then     'Otherwise getting EOF or BOF error
    Dim xKey
    xKey = Data1.Recordset("AD_EMPNBR")
    xKey = xKey & "|" & Format(Data1.Recordset("AD_DOA"), "dd-mmm-yyyy")
    xKey = xKey & "|" & Data1.Recordset("AD_REASON")
    Call Attendance_Master_Integration(xKey, xID)
End If
DoEvents

'Hemu - Testing ASL
If glbWHSCC Then
    'Data1.Refresh
    If whsccASLFlag Then
        Call modSTUPD(False)
        AddChg = " "
        Screen.MousePointer = DEFAULT
        whsccASLFlag = False
        Exit Sub
    End If
End If
'Hemu
'-------------------------------------
AddChg = "A"

rsDATA.AddNew

'Me.vbxTrueGrid(0).Enabled = True
lblCNum.Caption = "001"
lblEEID = locEmpnbr
dlpReviewDate.Text = DateAdd("d", 1, locDate)
clpCode(1).Text = locReason
medHours = locHours
elpSupShow.Text = locSupShow
clpChrgCode.Text = locChrgCode
txtShift = locShift
chkBackDated = locBackDated
dlpPayEndDate = locPayEndDate

' dkostka - 11/16/2001 - Incident flag should shut off when using +Another.
'2454W - Showa Ticket #25250 Franks 04/07/2014
If glbCompSerial = "S/N - 2350W" Or glbCompSerial = "S/N - 2454W" Then
    chkIncident = locIncident
Else
    chkIncident = False ' locIncident
End If
chkSeniority = locSen
chkEMELEA = locEMELEA
ChkInc = locInc
ChkFMLA = locFmla
txtWSIB = locWsib
txtPoint = locPoint
lblUpload = locUpload

memComments = locComment
clpJob.Text = locJob
clpCode(0).Text = locUnion
medsalary = locSalary
lblSalCode = locSalCode
txtDHRS = locDHrs
txtWHRS = locWHrs

SavEML = 0
SavVac = 0
SavVCOBank = 0 'Vitalaire
SavSick = 0

'Town of Aurora
'If glbCompSerial = "S/N - 2378W" Then
    SavOvt = 0
    SavCT = 0
'End If

flgEXIT:
Screen.MousePointer = DEFAULT
xAnother = 2
cmdAnother.Enabled = True
Exit Sub



Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAnother", "ATTEND", "Update")
Call RollBack '28July99 js
End Sub

Sub cmdCancel_Click()
'''On Error GoTo Can_Err

SavEML = 0
SavVac = 0
SavVCOBank = 0
SavSick = 0

'Town of Aurora
'If glbCompSerial = "S/N - 2378W" Then
    SavOvt = 0
    SavCT = 0
'End If

If savEnt1 <> 0 Then
'    UPDHRENTIT Data1.Recordset("AD_REASON"), savEnt1
    savEnt1 = 0
    savEntDate = ""
    savReas = ""
End If
fglbNew = False
cmdAnother.Visible = False
Screen.MousePointer = HOURGLASS

rsDATA.CancelUpdate
Call Display_Value

Call txtSUP_Change
Screen.MousePointer = DEFAULT
dlpToDate.Text = ""

Call modSTUPD(True)  ' reset screen's attributes

Me.vbxTrueGrid(0).SetFocus


Exit Sub

Can_Err:
Data1.Refresh
Resume Next
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
Call RollBack '28July99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMVATTEND" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, X
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, rc%, DtTm As Variant, xID
Dim xEmpNo, xDOA, xReason 'For glbWHSCC


If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
 '   cmdModify.Enabled = False   'laura 03/10/98
    Exit Sub
End If

'''On Error GoTo Del_Err

'Ticket #25268: Check first if this Attendance record that was created by ESS as Approved request.
If vbxTrueGrid(0).SelBookmarks.count = 0 Then vbxTrueGrid(0).SelBookmarks.Add Data1.Recordset.Bookmark
For X = 0 To vbxTrueGrid(0).SelBookmarks.count - 1
    Data1.Recordset.Bookmark = vbxTrueGrid(0).SelBookmarks(X)
    xID = Data1.Recordset("AD_ATT_ID")

    If Is_ESSApproved_Record(xID) Then
        If X = 0 Then
            Msg = "This record was added from ESS - Request Approval." & vbCrLf & vbCrLf
            'MsgBox "This record was added from ESS - Request Approval. It cannot be deleted from info:HR.", vbExclamation, "ESS Approved Attendance Record"
            'Exit Sub
        ElseIf X > 0 Then
            Msg = "There are records in this selection that were added from ESS - Request Approval." & vbCrLf & vbCrLf
            Exit For
        End If
    End If
Next

Msg = Msg & "Are You Sure You Want To Delete "
If vbxTrueGrid(0).SelBookmarks.count > 1 Then
    Msg = Msg & "These Records?"
Else
    Msg = Msg & "This Record?"
End If
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

DtTm = Now
Dim xKey
AddChg = "D"
xInciFlag = False

If vbxTrueGrid(0).SelBookmarks.count = 0 Then vbxTrueGrid(0).SelBookmarks.Add Data1.Recordset.Bookmark
For X = 0 To vbxTrueGrid(0).SelBookmarks.count - 1
    Call Payroll_Integration("D")
    Data1.Recordset.Bookmark = vbxTrueGrid(0).SelBookmarks(X)
    xID = Data1.Recordset("AD_ATT_ID")
    
    'If glbWFC Then 'Ticket #24124 Franks 07/25/2013
    If glbWFC Or glbMitchellPlastics Then 'Ticket #24112 Franks 07/30/2013
        Call WFC_Attend_To_AT(glbLEE_ID, "D", Data1.Recordset("AD_DOA"), Data1.Recordset("AD_REASON"), xID)
    End If
    
    xKey = Data1.Recordset("AD_EMPNBR")
    xKey = xKey & "|" & Format(Data1.Recordset("AD_DOA"), "dd-mmm-yyyy")
    xKey = xKey & "|" & Data1.Recordset("AD_REASON")
    Call Attendance_Master_Integration(xKey, , True)
    
    'Hemu - The procedure was being called with AD_REASON as the parameter to be compareed against EF_FREAS
    'but this will not work because for Vacation and Sick the Followup Reason is VAC and SICK - so made
    'the change accordingly. Also, the procedure is called with AD_DOA as the parameter to be compared
    'against the EF_FDATE but this will not work either because EF_FDATE stores system date - so made
    'the change accordingly.
    If Left(Data1.Recordset("AD_REASON"), 3) = "VAC" Then
        'Call VacSickHourlyFollowUp("ED_EMPNBR=" & Data1.Recordset("AD_EMPNBR"), "Delete", "VAC", Data1.Recordset("AD_DOA"))
        Call VacSickHourlyFollowUp("ED_EMPNBR=" & Data1.Recordset("AD_EMPNBR"), "Delete", "VAC", Data1.Recordset("AD_LDATE"))
    ElseIf Left(Data1.Recordset("AD_REASON"), 3) = "SIC" Then
        'Call VacSickHourlyFollowUp("ED_EMPNBR=" & Data1.Recordset("AD_EMPNBR"), "Delete", "SICK", Data1.Recordset("AD_DOA"))
        Call VacSickHourlyFollowUp("ED_EMPNBR=" & Data1.Recordset("AD_EMPNBR"), "Delete", "SICK", Data1.Recordset("AD_LDATE"))
    Else
        'Call VacSickHourlyFollowUp("ED_EMPNBR=" & Data1.Recordset("AD_EMPNBR"), "Delete", Data1.Recordset("AD_REASON"), Data1.Recordset("AD_DOA"))
        Call VacSickHourlyFollowUp("ED_EMPNBR=" & Data1.Recordset("AD_EMPNBR"), "Delete", Data1.Recordset("AD_REASON"), Data1.Recordset("AD_LDATE"))
    End If
    
    If glbWHSCC Then
        xReason = Data1.Recordset("AD_REASON")
        xEmpNo = Data1.Recordset("AD_EMPNBR")
        If xReason = "ASL" Then
            xDOA = Data1.Recordset("AD_DOA")
            Call UpdateASL(xEmpNo, xDOA, 0, "D")
        End If
    End If
    
    If chkIncident = True Then   'laura 02/24/98
        xInciFlag = True
        'cntSick = cntSick - 1
        'UpdateIncCount RW_WRITE
        'lblIncident.Caption = "Total # of Incidents = " & Str(cntSick)
    End If
    
    If Left(Data1.Recordset("AD_REASON"), 3) = "VAC" Then
        If Data1.Recordset("AD_DOA") >= Fdate And Data1.Recordset("AD_DOA") <= Tdate Then
            SavVac = 0 - Data1.Recordset("AD_HRS")
        End If
        SavOutV = SavOutV - SavVac
    End If
    
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #14635
        If Left(Data1.Recordset("AD_REASON"), 3) = "VCO" Then
            If Data1.Recordset("AD_DOA") >= Fdate And Data1.Recordset("AD_DOA") <= Tdate Then
                SavVCOBank = 0 - Data1.Recordset("AD_HRS")
            End If
        End If
    End If

    If Left(Data1.Recordset("AD_REASON"), 3) = "SIC" Then
        If Data1.Recordset("AD_DOA") >= fdateS And Data1.Recordset("AD_DOA") <= tdateS Then
            SavSick = 0 - Data1.Recordset("AD_HRS")
        End If
        SavOutS = SavOutS - SavSick
    End If
    
    'Town of Aurora
    'If glbCompSerial = "S/N - 2378W" Then
        If Left(Data1.Recordset("AD_REASON"), 2) = "OT" Then
            If Data1.Recordset("AD_DOA") >= OTFdate And Data1.Recordset("AD_DOA") <= OTTdate Then
                SavOvt = 0 - Data1.Recordset("AD_HRS")
            End If
        ElseIf Left(Data1.Recordset("AD_REASON"), 2) = "CT" Then
            If Data1.Recordset("AD_DOA") >= OTFdate And Data1.Recordset("AD_DOA") <= OTTdate Then
                SavCT = 0 - Data1.Recordset("AD_HRS")
            End If
        End If
    'End If
    
    'Hemu - EML
    'If glbSQL Or glbOracle Then
    'Hemu
        If chkEMELEA Then
            If Data1.Recordset("AD_DOA") >= CVDate(GetMonth("Jan") & " 1," & Year(Date)) And Data1.Recordset("AD_DOA") <= CVDate(GetMonth("Dec") & " 31," & Year(Date)) Then
                'linamar stuff added by Bryan 13/10/05 Ticket#9264
                If Not glbLinamar Then
                    SavEML = 0 - Data1.Recordset("AD_HRS")
                Else
                    SavEML = -1
                End If
            End If
            SavOutE = SavOutE - SavEML
        End If
    'End If
    
    If glbCompSerial = "S/N - 2173W" Then
        Call Check_Overtime_Bank(True)
    End If
    
    UPDVACSICK True
    
    'Hemu - EML
    'If glbSQL Or glbOracle Then UPDEML True
    Calculate_EML_Taken
    'Hemu

    'Calculate Comp Time Outstanding - Ticket #17345
    Call Calculate_Outstanding_CompTime
    
    UPDHRENTIT Data1.Recordset("AD_REASON"), 0
    
    If glbtermopen Then
        gdbAdoIhr001X.BeginTrans
        gdbAdoIhr001X.Execute "DELETE FROM Term_ATTENDANCE WHERE AD_ATT_ID=" & xID
        gdbAdoIhr001X.CommitTrans
        
        If gsAttachment_DB Then
            gdbAdoIhr001_DOC.Execute "Delete from Term_HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " AND AD_REASON='" & clpCode(1).Text & "' AND AD_DOA=" & Date_SQL(dlpReviewDate.Text) & " and AD_DOCKEY=" & glbDocKey & " "
        End If
    Else
        If xAD = "AD" Then
            If glbWFC And glbPlantCode = "WHBY" Then
                Call WhitbyDisciplineDelete(xID)
            End If
            If glbBurlTech Then
                Call BurlTechDisciplineDelete(xID)
            End If
            gdbAdoIhr001.BeginTrans
            gdbAdoIhr001.Execute "DELETE FROM HR_ATTENDANCE WHERE AD_ATT_ID=" & xID
            gdbAdoIhr001.CommitTrans
            
            If gsAttachment_DB Then
                gdbAdoIhr001_DOC.Execute "delete from HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(glbDocName) & "' AND AD_EMPNBR = " & glbLEE_ID & " AND AD_REASON='" & clpCode(1).Text & "' AND AD_DOA=" & Date_SQL(dlpReviewDate.Text) & " and AD_DOCKEY=" & glbDocKey & " "
            End If
            
            If glbWFC And glbPlantCode = "WHBY" Then
                Call Whitby60daysRule(lblEmpID, "D")
            End If
        Else
            gdbAdoIhr001.BeginTrans
            gdbAdoIhr001.Execute "DELETE FROM HR_ATTENDANCE_HISTORY WHERE AH_ATT_ID=" & xID
            gdbAdoIhr001.CommitTrans
            
            '7.9 Enhancement
            If gsAttachment_DB Then
                gdbAdoIhr001_DOC.Execute "delete from HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(glbDocName) & "' AND AD_EMPNBR = " & glbLEE_ID & " AND AD_REASON='" & clpCode(1).Text & "' AND AD_DOA=" & Date_SQL(dlpReviewDate.Text) & " and AD_DOCKEY=" & glbDocKey & " "
            End If
            
        End If
    End If
    If glbWHSCC Then
        If xReason = "USB" Then
            Call ReCalcUSB("AD_EMPNBR = " & lblEmpID, "EMP")
        End If
    End If
    DoEvents
Next

'If glbBurlTech Then 'BTI Points Recalculate
'    Call BTIPoint(lblEEID)
'End If

Data1.Refresh

'Hemu - EML
Call Calculate_EML_Taken
'Hemu

'Calculate Comp Time Outstanding - Ticket #17345
Call Calculate_Outstanding_CompTime

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If
fglbNew = False
'Call modSTUPD(True)
Call SET_UP_MODE

If xInciFlag Then
    UpdateIncCount RW_WRITE
End If

If glbWHSCC Then
    'Hemu - Testing ASL
    SQLQ = "ED_EMPNBR = " & glbLEE_ID
    Call EntReCalc(SQLQ)
    'Hemu
    'Franks 05/22/2003 Ticket# 4103 Show ASL Balance in this year on Attendance screen
    Call EntASLBalance(glbLEE_ID)
End If

'Town of Aurora
'If glbCompSerial = "S/N - 2378W" Then
    Call UPDOVERTIME
'End If

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "ATTEND", "Delete")
Call RollBack '28July99 js

End Sub

Private Sub BurlTechDisciplineDelete(xAD_ID)
'If Disciplinary action was deleted, the matched Attendance record should be delete
Dim rsTemp As New ADODB.Recordset
Dim rsTem2 As New ADODB.Recordset
Dim SQLQ, xDiscipStep, xNextStepPlus
Dim xEmpNo, xType, xIncDate, xATTReason
Dim xNextStepVal
    ''Disable it until Whitby is ready
    'Exit Sub
    
    SQLQ = "SELECT AD_EMPNBR,AD_DOA,AD_REASON,AD_DISCIPLINE FROM HR_ATTENDANCE WHERE AD_ATT_ID = " & xAD_ID
    rsTem2.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsTem2.EOF Then Exit Sub
    If IsNull(rsTem2("AD_DISCIPLINE")) Then Exit Sub 'No Disciplinary action
    If Len(Trim((rsTem2("AD_DISCIPLINE")))) = 0 Then Exit Sub 'No Disciplinary action
    xEmpNo = rsTem2("AD_EMPNBR")
    xType = rsTem2("AD_DISCIPLINE")
    xIncDate = rsTem2("AD_DOA")
    xATTReason = rsTem2("AD_REASON")
    rsTem2.Close

    'Delete Discipline from HR_COUNSEL
    SQLQ = "DELETE FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND CL_ATTREASON = '" & xATTReason & "' "
    SQLQ = SQLQ & "AND CL_ATTDATE = " & Date_SQL(xIncDate) & " "
    gdbAdoIhr001.Execute SQLQ
    
    Unload frmECounsel
End Sub
Private Sub WhitbyDisciplineDelete(xAD_ID)
'If Disciplinary action was deleted, the matched Attendance record should be delete
'Also check ED_DISCIPLINENEXT in HREMP, if this Disciplinary action is the current action
'then ED_DISCIPLINENEXT should be ED_DISCIPLINENEXT -1
Dim rsTemp As New ADODB.Recordset
Dim rsTem2 As New ADODB.Recordset
Dim SQLQ, xDiscipStep, xNextStepPlus
Dim xEmpNo, xType, xIncDate, xATTReason
Dim xNextStepVal
    'Enable this function Ticket# 6656
    ''Disable it until Whitby is ready
    'Exit Sub
    
    SQLQ = "SELECT AD_EMPNBR,AD_DOA,AD_REASON,AD_DISCIPLINE FROM HR_ATTENDANCE WHERE AD_ATT_ID = " & xAD_ID
    rsTem2.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsTem2.EOF Then Exit Sub
    If IsNull(rsTem2("AD_DISCIPLINE")) Then Exit Sub 'No Disciplinary action
    If Len(Trim((rsTem2("AD_DISCIPLINE")))) = 0 Then Exit Sub 'No Disciplinary action
    xEmpNo = rsTem2("AD_EMPNBR")
    xType = rsTem2("AD_DISCIPLINE")
    xIncDate = rsTem2("AD_DOA")
    xATTReason = rsTem2("AD_REASON")
    rsTem2.Close

    'Check if xType in HR_DISCIPLINE_STEPS table
    SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS WHERE DS_DISCIPLINE = '" & xType & "' "
    rsTem2.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsTem2.EOF Then
        Exit Sub
    Else
        xDiscipStep = rsTem2("DS_STEPNO")
    End If
    rsTem2.Close
    
    'Delete Discipline from HR_COUNSEL
    SQLQ = "DELETE FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND CL_TYPE = '" & xType & "' "
    SQLQ = SQLQ & "AND CL_ATTDATE = " & Date_SQL(xIncDate) & " "
    gdbAdoIhr001.Execute SQLQ
    
    Unload frmECounsel
    
''
''    ''Delete Discipline from HR_COUNSEL
''    'SQLQ = "DELETE FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
''    'SQLQ = SQLQ & "AND CL_TYPE = '" & xType & "' "
''    'SQLQ = SQLQ & "AND CL_INCDATE = " & Date_SQL(xIncDate) & " "
''    'gdbAdoIhr001.Execute SQLQ
''
''    'Check if this Disciplinary action is the current action
''    SQLQ = "SELECT ED_EMPNBR, ED_DISCIPLINENEXT FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
''    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
''    xNextStepVal = 1
''    If Not rsTemp.EOF Then
''        If Not IsNull(rsTemp("ED_DISCIPLINENEXT")) Then
''            xNextStepVal = rsTemp("ED_DISCIPLINENEXT")
''            'If (rsTemp("ED_DISCIPLINENEXT") = xDiscipStep + 1) And xDiscipStep >= 1 Then
''            '
''            '    rsTemp("ED_DISCIPLINENEXT") = xDiscipStep
''            '    rsTemp.Update
''            'End If
''        End If
''    End If
''    rsTemp.Close
''
''    Unload frmECounsel
''
''    If Not (xDiscipStep + 1 = xNextStepVal) Then
''        'IF this attendance record is not current Disciplinary step
''        'Delete Discipline from HR_COUNSEL
''        SQLQ = "DELETE FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
''        SQLQ = SQLQ & "AND CL_TYPE = '" & xType & "' "
''        SQLQ = SQLQ & "AND CL_INCDATE = " & Date_SQL(xIncDate) & " "
''        gdbAdoIhr001.Execute SQLQ
''
''        Exit Sub
''    Else
''        'Check if this attendance record is current Disciplinary step, but there are multi current steps
''        'if this attendance record is not the latest record, don't change ED_DISCIPLINENEXT
''        SQLQ = "SELECT AD_EMPNBR,AD_DOA,AD_REASON,AD_DISCIPLINE FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & xEmpNo & " "
''        SQLQ = SQLQ & "AND AD_DISCIPLINE = '" & xType & "' "
''        SQLQ = SQLQ & "ORDER BY AD_DOA DESC"
''        If rsTem2.State <> 0 Then rsTem2.Close
''        rsTem2.Open SQLQ, gdbAdoIhr001, adOpenStatic
''        If Not rsTem2.EOF Then
''            If Not (CVDate(rsTem2("AD_DOA")) = CVDate(xIncDate)) Then
''                'Delete Discipline from HR_COUNSEL
''                SQLQ = "DELETE FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
''                SQLQ = SQLQ & "AND CL_TYPE = '" & xType & "' "
''                SQLQ = SQLQ & "AND CL_INCDATE = " & Date_SQL(xIncDate) & " "
''                gdbAdoIhr001.Execute SQLQ
''                rsTem2.Close
''                Exit Sub
''            End If
''        End If
''        rsTem2.Close
''    End If
''
''    SQLQ = "UPDATE HREMP SET ED_DISCIPLINENEXT = " & WhitbyPreStep(xEmpNo, xDiscipStep) & " WHERE ED_EMPNBR = " & xEmpNo
''    gdbAdoIhr001.Execute SQLQ
''
''    'Delete Discipline from HR_COUNSEL
''    SQLQ = "DELETE FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
''    SQLQ = SQLQ & "AND CL_TYPE = '" & xType & "' "
''    SQLQ = SQLQ & "AND CL_INCDATE = " & Date_SQL(xIncDate) & " "
''    gdbAdoIhr001.Execute SQLQ
''

    
End Sub
'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
'Function WhitbyPreStep(xEmpNo, xStep)
'Dim rsDisci As New ADODB.Recordset
'Dim rsCounsel As New ADODB.Recordset
'Dim SQLQ, xPreStep, I, xMum
''Dim xArray(10, 2)
'    xPreStep = xStep
'    SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS WHERE DS_STEPNO <= " & xStep & " ORDER BY DS_STEPNO DESC"
'    rsDisci.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    If rsDisci.EOF Then
'        rsDisci.Close
'        GoTo End_line
'    End If
'    xPreStep = 1
'    Do While Not rsDisci.EOF
'        SQLQ = "SELECT CL_EMPNBR FROM HR_COUNSEL WHERE CL_TYPE = '" & rsDisci("DS_DISCIPLINE") & "' "
'        SQLQ = SQLQ & "AND  CL_EMPNBR = " & xEmpNo
'        If rsCounsel.State <> 0 Then rsCounsel.Close
'        rsCounsel.Open SQLQ, gdbAdoIhr001, adOpenStatic
'        If Not rsCounsel.EOF Then
'            xPreStep = rsDisci("DS_STEPNO")
'            rsCounsel.Close
'            rsDisci.Close
'            GoTo End_line
'        End If
'        rsCounsel.Close
'        rsDisci.MoveNext
'    Loop
'
'End_line:
'    WhitbyPreStep = xPreStep
'End Function

Sub cmdModify_Click()
    Dim Skll As String, Skllvl As String, SklDte As String
    Dim SQLQ As String
    
    '''On Error GoTo Mod_Err
    
    xOldReason = clpCode(1).Text
    SavEML = 0
    SavVac = 0
    SavVCOBank = 0
    SavSick = 0
    savEnt1 = 0
    HourGlb = Val(medHours)
    savEntDate = ""
    savReas = ""
    oldHrs = 0
    
    'Town of Aurora
    'If glbCompSerial = "S/N - 2378W" Then
        SavOvt = 0
        SavCT = 0
    'End If
    
    oldDate = dlpReviewDate
    oldReason = clpCode(1)
    oldHours = medHours
    oldJob = clpJob
    oldSalary = medsalary
    oldSalCode = lblSalCode
    oldWhrs = txtWHRS
    oldComment = memComments
    oldHrs = medHours
    
    'Ticket #27910 - Option to change Salary information
    If glbCompSerial = "S/N - 2411W" Then
        cmdSalaryChange.Enabled = False
    End If
    
    If Data1.Recordset.EOF Then Exit Sub
    
    If Left(Data1.Recordset("AD_REASON"), 3) = "VAC" Then
        If Data1.Recordset("AD_DOA") >= Fdate And Data1.Recordset("AD_DOA") <= Tdate Then
            SavVac = 0 - Data1.Recordset("AD_HRS")
        End If
    End If
    
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #14635
        If Left(Data1.Recordset("AD_REASON"), 3) = "VCO" Then
            If Data1.Recordset("AD_DOA") >= Fdate And Data1.Recordset("AD_DOA") <= Tdate Then
                SavVCOBank = 0 - Data1.Recordset("AD_HRS")
            End If
        End If
    End If
    
    If Left(Data1.Recordset("AD_REASON"), 3) = "SIC" Then
        If Data1.Recordset("AD_DOA") >= fdateS And Data1.Recordset("AD_DOA") <= tdateS Then
            SavSick = 0 - Data1.Recordset("AD_HRS")
        End If
    End If
    
    'Town of Aurora
    'If glbCompSerial = "S/N - 2378W" Then
        If Left(Data1.Recordset("AD_REASON"), 2) = "OT" Then
            If Data1.Recordset("AD_DOA") >= OTFdate And Data1.Recordset("AD_DOA") <= OTTdate Then
                SavOvt = 0 - Data1.Recordset("AD_HRS")
            End If
        ElseIf Left(Data1.Recordset("AD_REASON"), 2) = "CT" Then
            If Data1.Recordset("AD_DOA") >= OTFdate And Data1.Recordset("AD_DOA") <= OTTdate Then
                SavCT = 0 - Data1.Recordset("AD_HRS")
            End If
        End If
    'End If
    
    'Hemu - EML
    'If glbSQL Or glbOracle Then
    'Hemu
        If chkEMELEA Then
            If Data1.Recordset("AD_DOA") >= CVDate(GetMonth("Jan") & " 1," & Year(Date)) And Data1.Recordset("AD_DOA") <= CVDate(GetMonth("Dec") & " 31," & Year(Date)) Then
                'Linamar stuff added by Bryan 13/10/05 Ticket#9264
                If Not glbLinamar Then
                    SavEML = 0 - Data1.Recordset("AD_HRS")
                Else
                    SavEML = -1
                End If
            End If
        End If
    'End If
    
    If chkIncident = True Then    'laura 02/24/98
        'cntSick = cntSick - 1
        'UpdateIncCount RW_WRITE
    End If
    
    savEntDate = Data1.Recordset("AD_DOA")
    savReas = Data1.Recordset("AD_REASON")
    'UPDHRENTIT Data1.Recordset("AD_REASON"), 0 - Data1.Recordset("AD_HRS")
    savEnt1 = Data1.Recordset("AD_HRS")
    AddChg = "C"
    
    'Ticket #27910 - Option to change Salary information
    If glbCompSerial = "S/N - 2411W" Then
        cmdSalaryChange.Enabled = True
    End If
    
    Call Set_Para_ForBrant
    
    Exit Sub
    
Mod_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
    Call RollBack '28July99 js
    Resume Next
End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
    Dim SQLQ As String, Msg$, X%
    
    '''On Error GoTo AddN_Err
    
    
    fglbNew = True
    
    'Call modSTUPD(True)
    
    AddChg = "A"
    Call SET_UP_MODE
    
    '7.9 Enhancement
    If gsAttachment_DB Then 'And xAD = "AD" Then
        lblImport.Visible = True 'False
        imgSec.Visible = False
        imgNoSec.Visible = True 'False
        cmdImport.Visible = True 'False
    End If
    
    Call Set_Control("B", Me)
    If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
        locEmpnbr = lblEEID
        locDate = dlpReviewDate.Text
        locReason = clpCode(1).Text
        locHours = medHours
        locSupShow = elpSupShow.Text
        locChrgCode = clpChrgCode.Text
        locShift = txtShift
    End If
    xmedHours = 0
    'data1.Recordset.AddNew
    ''' Sam add July 2002 * Remove Binding Control
    rsDATA.AddNew
    
    'Defaulting the Shift value for Brant County Health Unit
    If glbCompSerial = "S/N - 2226W" Then
        comShiftType.ListIndex = 0
        txtShift.Text = "FD"
    End If
    
    oldDate = ""
    oldReason = ""
    oldHours = 0
    oldJob = ""
    oldSalary = 0
    oldSalCode = ""
    oldWhrs = 0
    oldComment = ""
    
    SavEML = 0
    SavVac = 0
    SavVCOBank = 0
    SavSick = 0
    savEnt1 = 0
    lblEEID.Caption = glbLEE_ID
    lblCNum.Caption = "001"
    savEntDate = ""
    savReas = ""
    chkBackDated.Value = False
    
    'Town of Aurora
    'If glbCompSerial = "S/N - 2378W" Then
        SavOvt = 0
        SavCT = 0
    'End If
    
    If Trim(elpSupShow.Text) = "" And Len(SavSup) > 0 Then elpSupShow.Text = ShowEmpnbr(SavSup)
    If glbCompSerial = "S/N - 2394W" Then ' St. John's - Ticket #15053
        If Trim(txtChrgCode) = "" And Len(SavShift) > 0 Then txtChrgCode = SavShift
    Else
        If Trim(txtShift) = "" And Len(SavShift) > 0 Then txtShift = SavShift
    End If
    If glbLinamar Then txtDHRS = 8
    
    If glbCompSerial = "S/N - 2214W" Then
        If Len(xDeptno) > 0 Then
            clpChrgCode = xDeptno: txtChrgCode = xDeptno
        End If
        If Len(xAdminBy) > 0 Then
            clpCode(2) = xAdminBy: txtShift = xAdminBy
        End If
        If Len(xGLNO) > 0 Then
            clpGLNum = xGLNO: txtWSIB = xGLNO
        End If
    Else
        clpGLNo = xGLNO
        txtPayrollID = xPayrollID
        medsalary = xCurSalary
        lblSalCode = xSalCD
        clpJob = xJob
        
        If Not IsNull(xDHrs) Then
            txtDHRS = xDHrs
        End If
        If Not IsNull(xWHrs) Then
            txtWHRS = xWHrs
        End If
        If Not IsNull(xORG) Then
            clpCode(0) = xORG
        End If
    End If
    If glbCompSerial = "S/N - 2376W" Then ' AFN Ticket #16251
        If Len(xDiv) > 0 Then
            clpChrgCode = xDiv: txtChrgCode = xDiv
        End If
    End If
    
    'Ticket #27910 - Option to change Salary information
    If glbCompSerial = "S/N - 2411W" Then
        cmdSalaryChange.Enabled = False
    End If

    dlpToDate.Enabled = True
    cmdAnother.Visible = True
    AskWeekend = True
    'Ticket #18188
    'If glbCompSerial = "S/N - 2418W" Then
        AskHoliday = True
    'Else
    '    AskHoliday = False
    'End If
    
    'If glbCBrant Then
        Call Set_Para_ForBrant
    'End If
    dlpReviewDate.SetFocus
    
    xAnother = 1
    
    'Ticket #7500 - Town of Ajax
    'tkt#10423 Jerry said remove serial#control for add_security
    'If glbCompSerial = "S/N - 2173W" Then
        MDIMain.MainToolBar.ButtonS("save").Enabled = True
        MDIMain.MainToolBar.ButtonS("cancel").Enabled = True
    'End If
    
    Exit Sub
    
AddN_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "ATTEND", "Add")
    Call RollBack '28July99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim rsATT As New ADODB.Recordset
Dim X%, Msg$, Answer, SQLQ, Logx, xWrk
Dim DgDef As Double
Dim xID, xSup
Dim xKey

'''On Error GoTo cmdOK_Err

fgotholidays = False
hlist = ""

If xAD = "AH" And Not gSec_Upd_Attendance_History Then
    MsgBox "You do not have authority to make changes"
    Call cmdCancel_Click
    Exit Sub
End If

If glbCompSerial = "S/N - 2214W" Then
    'To let LostFocus of clpCode(2) work, for Casey House, AD_SHIFT lookup on Admin By
    dlpReviewDate.SetFocus
    DoEvents
End If

'Jerry said remove Serial#Control for Add_Attendance Security tkt#10423
'If glbCompSerial = "S/N - 2173W" Then
    If (gSec_Add_Attendance And Not gSec_Upd_Attendance) And AddChg = "C" Then
        MsgBox "You do not have authority to make changes"
        Call cmdCancel_Click
        Exit Sub
    End If
'End If

If Not chkAttendance() Then Exit Sub

If fglbNew Then
    xKey = 0
    xKey = xKey & "|" & Format(Date, "dd-mmm-yyyy")
    xKey = xKey & "|ANY"
Else
    xKey = Data1.Recordset("AD_EMPNBR")
    xKey = xKey & "|" & Format(Data1.Recordset("AD_DOA"), "dd-mmm-yyyy")
    xKey = xKey & "|" & Data1.Recordset("AD_REASON")
End If

lblUpload = IIf(chkUpload.Value, "Y", "N")
xNewReason = clpCode(1).Text
xmedHours = Val(medHours)

If IsDate(dlpToDate.Text) Then
    For X = 1 To DateDiff("d", dlpReviewDate.Text, dlpToDate.Text)
        Call cmdAnother_Click
        If dlpToDate.Text = "" Or fglbRetry Then Exit For
    Next
    If Not fglbRetry Then
        dlpToDate.Text = ""
'        Call modSTUPD(True)
        Call SET_UP_MODE
    End If
    If Weekday(dlpReviewDate.Text) = 7 Or Weekday(dlpReviewDate.Text) = 1 Then
        If SkipWeekend Then
            fglbNew = False
            rsDATA.CancelUpdate
            Data1.Refresh
            Call Display_Value
            Exit Sub
        End If
    End If
End If

If glbCompSerial = "S/N - 2173W" Then
    If Not Check_Overtime_Bank(True) Then Exit Sub
End If

'Ticket #28207 - Added back in as only AddChg = "A" condition will fix the issue in the Ticket #27181
'Ticket #27181 - Commenting this out as the procedure EntEccess is already computing the * 1.5 and * 2.0
''Ticket #25500 - just for EntEccess() to account for OT15 and OT20 hours exceeding
''For City of Niagara Fulls
If glbNiagaraFulls And (Not glbtermopen) And AddChg = "A" Then
    Select Case UCase(clpCode(1).Text)
        Case "OT15"
            medHours = xmedHours * 1.5
        Case "OT20"
            medHours = xmedHours * 2  'Val(medHours) * 2
    End Select
End If

If Not EntEccess() Then Exit Sub
'Hemu - EML
'If glbSQL Or glbOracle Then If Not EmlEntEccess() Then Exit Sub
If Not EmlEntEccess() Then Exit Sub
'Hemu

If glbWHSCC Then
    whsccAnotherFlag = False
    If Not EntStatEccess(glbLEE_ID, CVDate(dlpReviewDate)) Then Exit Sub
    If clpCode(1) = "ASL" Then
        If Not EntASLEccess(glbLEE_ID, CVDate(dlpReviewDate), medHours) Then Exit Sub
    End If
    If clpCode(1) = "USB" Then
        If Not EntUSBEccess(glbLEE_ID, CVDate(dlpReviewDate), medHours) Then Exit Sub
    End If
End If


'Surrey Place to check if the Employee Status is "TERM"
If glbCompSerial = "S/N - 2347W" And AddChg = "A" Then
    If Not ChkTermStatus Then Exit Sub
End If

If Data1.Recordset.RecordCount > 0 Then If Not ChkDup() Then Exit Sub

If chkIncident = True Then cntSick = cntSick + 1

lblIncident.Caption = "Total # of Incidents = " & Str(cntSick)
If Left(clpCode(1).Text, 3) = "VAC" Then
    If DateValue(dlpReviewDate.Text) >= Fdate And DateValue(dlpReviewDate.Text) <= Tdate Then
        SavVac = SavVac + medHours
    End If
    SavOutV = SavOutV - SavVac
End If

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #14635
    If Left(clpCode(1).Text, 3) = "VCO" Then
        If DateValue(dlpReviewDate.Text) >= Fdate And DateValue(dlpReviewDate.Text) <= Tdate Then
            SavVCOBank = SavVCOBank + medHours
        End If
    End If
End If

If Left(clpCode(1).Text, 3) = "SIC" Then
    If DateValue(dlpReviewDate.Text) >= fdateS And DateValue(dlpReviewDate.Text) <= tdateS Then
        SavSick = SavSick + medHours
    End If
    SavOutS = SavOutS - SavSick
End If

'Town of Aurora
'If glbCompSerial = "S/N - 2378W" Then
    If Left(clpCode(1).Text, 2) = "OT" Then
        If DateValue(dlpReviewDate.Text) >= OTFdate And DateValue(dlpReviewDate.Text) <= OTTdate Then
            SavOvt = SavOvt + medHours
        End If
    ElseIf Left(clpCode(1).Text, 2) = "CT" Then
        If DateValue(dlpReviewDate.Text) >= OTFdate And DateValue(dlpReviewDate.Text) <= OTTdate Then
            SavCT = SavCT + medHours
        End If
    End If
'End If

'Hemu - EML
'If glbSQL Or glbOracle Then
'Hemu
    If chkEMELEA Then
        If DateValue(dlpReviewDate.Text) >= CVDate(GetMonth("Jan") & " 1," & Year(Date)) And DateValue(dlpReviewDate.Text) <= CVDate(GetMonth("Dec") & " 31," & Year(Date)) Then
            'linamar added by Bryan 3/Oct/05 Ticket#9264
            If Not glbLinamar Then
                SavEML = SavEML + medHours
            Else
                SavEML = SavEML + 1
            End If
        End If
        SavOutE = SavOutE - SavEML
    End If
'End If

UPDVACSICK True

'Hemu - EML
'If glbSQL Or glbOracle Then UPDEML True
Calculate_EML_Taken
'Hemu

'Calculate Comp Time Outstanding - Ticket #17345
Call Calculate_Outstanding_CompTime

UPDHRENTIT clpCode(1).Text, medHours
'UpdateIncCount RW_WRITE

'V7.6
'Current Salary to AD_SALARY FOR Casey House
'If glbCompSerial = "S/N - 2214W" Then
    If Not IsNumeric(medsalary) Then
        medsalary = xCurSalary
    Else
        If Val(medsalary) = 0 Then
        medsalary = xCurSalary
        End If
    End If
    If Len(lblSalCode) = 0 Then
        If Len(xSalCD) > 0 Then lblSalCode = xSalCD
    End If
    If Len(clpJob) = 0 Then
        If Len(xJob) > 0 Then clpJob = xJob
    End If
    If Len(clpCode(0)) = 0 Then
        If Len(xORG) > 0 Then clpCode(0) = xORG
    End If
    If Len(txtDHRS) = 0 Then
        If xDHrs > 0 Then txtDHRS = xDHrs
    End If
    If Len(txtWHRS) = 0 Then
        If xWHrs > 0 Then txtWHRS = xWHrs
    End If

'Casey House
If glbCompSerial = "S/N - 2214W" Then
    If Len(txtShift) = 0 Then
        If Len(xAdminBy) > 0 Then txtShift = xAdminBy
    End If
    If Len(txtChrgCode) = 0 Then
        If Len(xDeptno) > 0 Then txtChrgCode = xDeptno
    End If
    If Len(txtWSIB) = 0 Then
        If Len(xGLNO) > 0 Then txtWSIB = xGLNO
    End If
End If

'Call VacSickHourlyFollowUp(clpCode(1).Text, dlpReviewDate)
Call VacSickHourlyFollowUp("ED_EMPNBR=" & lblEmpID, , clpCode(1).Text, dlpReviewDate)

glbflgFU = False
'Calculate_EML_Taken

Call UpdUStats(Me)

Call UpdCodes(xAD) 'Ticket #28846 Franks 08/16/2016

Call Set_Control("U", Me, rsDATA)

xSup = getEmpnbr(elpSupShow.Text)
If Val(xSup) = 0 Then
    rsDATA!AD_SUPER = Null
Else
    rsDATA!AD_SUPER = xSup
End If

'Ticket #18668 - 7.9 Enhancement
'Ticket #30369 - Update _SOURCE for Edits as well excluding WFC as they only need to be updated on New Records only as mentioned by Frank in his comments below
If glbWFC Then
    If fglbNew Then 'Ticket #28373 Franks 03/23/2016 - the AD_SOURCE can't be overriten since ESS to AT need it, it only can be setup on new record
        If xAD = "AD" Then
            rsDATA("AD_SOURCE") = "IHRATT"
        Else
            rsDATA("AH_SOURCE") = "IHRATH"
        End If
    End If
Else
    If xAD = "AD" Then
        rsDATA("AD_SOURCE") = "IHRATT"
    Else
        rsDATA("AH_SOURCE") = "IHRATH"
    End If
End If

xDiscipFlag = False: xOccuAmount = 0
If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    rsDATA.Resync

    If gsAttachment_DB Then
        'If Not fglbNew Then
            gdbAdoIhr001_DOC.Execute "Update Term_HRDOC_ATTENDANCE SET AD_REASON='" & rsDATA("AD_REASON") & "',AD_DOA=" & Date_SQL(rsDATA("AD_DOA")) & " WHERE AD_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " AND AD_REASON='" & oldReason & "' AND AD_DOA=" & Date_SQL(oldDate) & " AND AD_DOCKEY = " & glbDocKey
        'End If
    End If

    xID = rsDATA!AD_ATT_ID
Else
    'Disciplinary action for Whitby, Check how many Incident flags
    If glbWFC And glbPlantCode = "WHBY" And xAD = "AD" And AddChg = "A" Then
        Call WhitbyGetIncidentFlags(lblEmpID, CVDate(dlpReviewDate), clpCode(1))
    End If
    
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
        
    If glbVadim Then
        rsDATA.Requery
    Else
        rsDATA.Resync
    End If
        
    If gsAttachment_DB Then
        'If Not fglbNew Then
            gdbAdoIhr001_DOC.Execute "Update HRDOC_ATTENDANCE SET AD_REASON='" & rsDATA("AD_REASON") & "',AD_DOA=" & Date_SQL(rsDATA("AD_DOA")) & " WHERE AD_TYPE='" & UCase(glbDocName) & "' AND AD_EMPNBR = " & glbLEE_ID & " AND AD_REASON='" & oldReason & "' AND AD_DOA=" & Date_SQL(oldDate) & " AND AD_DOCKEY = " & glbDocKey
            'gdbAdoIhr001_DOC.Execute "Update HRDOC_ATTENDANCE SET AD_REASON='" & rsDATA("AD_REASON") & "',AD_DOA=" & Date_SQL(rsDATA("AD_DOA")) & " WHERE AD_TYPE='" & UCase(glbDocName) & "' AND AD_EMPNBR = " & glbLEE_ID & " AND AD_DOCKEY = " & glbDocKey
        'End If
    End If
        
    xID = rsDATA!AD_ATT_ID
    
    'If glbWFC Then 'Ticket #24124 Franks 07/24/2013
    If glbWFC Or glbMitchellPlastics Then 'Ticket #24112 Franks 07/30/2013
        If (Not fglbNew) And (Not oldDate = dlpReviewDate.Text) Then 'Ticket #24155 Franks 07/30/2013
            Call WFC_Attend_To_AT(glbLEE_ID, "M", rsDATA("AD_DOA"), rsDATA("AD_REASON"), xID, oldDate, oldReason)
        Else
            Call WFC_Attend_To_AT(glbLEE_ID, "M", rsDATA("AD_DOA"), rsDATA("AD_REASON"), xID, , oldReason)
        End If
    End If
End If

If gsAttachment_DB Then
    If glbDocNewRecord Then 'New Record only
        If Len(glbDocImpFile) > 0 Then
            glbDocKey = xID
            glbAttDOA = dlpReviewDate.Text
            glbAttReason = clpCode(1).Text
            If glbtermopen Then
                Call AttachmentAdd(glbTERM_ID, glbDocImpFile, glbDocType, glbDocDesc)
            Else
                Call AttachmentAdd(glbLEE_ID, glbDocImpFile, glbDocType, glbDocDesc)
            End If
        End If
    End If
    glbDocImpFile = ""
End If

'Leeds and Grenville - Ticket #19441 - Just incase they add a new OT records from info:HR Attendance
'instead of ESS Request/Approval - I am updating with Expiry Date so that in the Weekly Adjustments
'this record is included.
'For N.B.P.H.U update the OT expiry date to 6 Months ticket # 15985
If glbCompSerial = "S/N - 2345W" Or (glbCompSerial = "S/N - 2233W" And fglbNew) Then
    Call UpdateOTExpiry(xID)
End If

'Hemu - EML
Call Calculate_EML_Taken
'Hemu

'Calculate Comp Time Outstanding - Ticket #17345
Call Calculate_Outstanding_CompTime

'Franks Aug 26 for T#2593
If glbWHSCC Then
    If clpCode(1) = "ASL" Then
        Call UpdateASL(lblEmpID, CVDate(dlpReviewDate), medHours, "U")
    End If
    If clpCode(1) = "USB" Then
        Call ReCalcUSB("AD_EMPNBR = " & lblEmpID, "EMP")
    End If
    If whsccExceedFlag And Not glbtermopen Then
        glbxID = xID
        Call whsccASLatt
        xID = glbxID
    End If
End If
'Franks Aug 26 for T#2593

'Create Disciplinary action for Whitby
If xDiscipFlag Then
    If glbWFC And glbPlantCode = "WHBY" And xAD = "AD" Then
        Call WhitbyUpdateDisciplinary(lblEmpID, CVDate(dlpReviewDate), clpCode(1)) ', medHours, "U")
    End If
End If

If glbBurlTech Then 'BTI Points Recalculate
    Call BTIPoint(lblEEID)
End If

If glbWFC And fglbNew Then 'Ticket #25148 Frank 03/05/2014
    If Left(cmdWFCHide.Caption, 4) = "Unhi" Then
        xHideWFCAttCodes = False
        cmdWFCHide.Caption = "Hide REG && OT"
        'the program must show all att records for this employee, otherwise it causes error for new REG and OT record
        Call EERetrieve
    End If
End If

Data1.Refresh

DoEvents

Data1.Recordset.Find "AD_ATT_ID=" & rsDATA!AD_ATT_ID
'Call modSTUPD(True)

UpdateIncCount RW_WRITE

fglbNew = False
'Call SET_UP_MODE
AddChg = " "

'Franks 05/22/2003 Ticket# 4103 Show ASL Balance in this year on Attendance screen
If glbWHSCC Then
    Call EntASLBalance(glbLEE_ID)
End If


'Check 60days rule of Disciplinary action for Whitby
If glbWFC And glbPlantCode = "WHBY" And xAD = "AD" Then
    Call Whitby60daysRule(lblEmpID, "")
    Data1.Refresh
    Data1.Recordset.Find "AD_ATT_ID=" & xID
Else
    Data1.Refresh
    Data1.Recordset.Find "AD_ATT_ID=" & xID
End If

'Frank 06/14/2004 Ticket# 6345 Close Vacation/Sick Entitlement and Hourly Entitlement forms
Unload frmVACSICK
Unload frmVACSICKO
Unload frmHrEnt
'Frank 06/14/2004 Ticket# 6345

'Town of Aurora
'If glbCompSerial = "S/N - 2378W" Then
    Call UPDOVERTIME
'End If

'Town of Ajax
If glbCompSerial = "S/N - 2173W" Then
    Call Recalculate_OTBANK
End If

'Jaddy Jan 24, 2005 for the integration
xKey = Data1.Recordset("AD_EMPNBR")
xKey = xKey & "|" & Format(Data1.Recordset("AD_DOA"), "dd-mmm-yyyy")
xKey = xKey & "|" & Data1.Recordset("AD_REASON")

Call Attendance_Master_Integration(xKey, Data1.Recordset("AD_ATT_ID"))
'Jaddy Jan 24, 2005-end

cmdAnother.Visible = False

Screen.MousePointer = DEFAULT

If NextFormIF("Attendance") Then
    Call cmdNew_Click
End If

'Ticket #12012
'Delete is disabled after add new on Att screen
Call SET_UP_MODE

Exit Sub

cmdOK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "ATTEND", "Update")
Call RollBack '28July99 js
Resume Next
End Sub

Private Sub cmdImport_Click()
    glbDocNewRecord = fglbNew
    glbDocName = "Attendance"
    If fglbNew Then
        glbDocKey = 0
        
        If Len(dlpReviewDate.Text) = 0 Or Len(clpCode(1).Text) = 0 Then
            MsgBox "'" & lStr("From Date") & "' and '" & lStr("Reason") & "' must be entered before attaching a document", vbExclamation
            Exit Sub
        Else
            glbAttReason = clpCode(1).Text
            glbAttDOA = dlpReviewDate.Text
        End If
    Else
        '8.0 - Ticket #22682 - Noticed an issue when testing for _DOCTYPE and _USRDESC
        If Not IsNull(rsDATA("AD_DOCKEY")) Then
            glbDocKey = rsDATA("AD_DOCKEY")
        Else
            glbDocKey = rsDATA("AD_ATT_ID") 'Ticket #16018
        End If
        glbAttReason = rsDATA("AD_REASON")
        glbAttDOA = rsDATA("AD_DOA")
    End If
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmVATTEND")
End Sub

Private Sub cmdSalaryChange_Click()
    frmMsgSalUpd.Show 1
    rsDATA.Requery
    Data1.Refresh
    DoEvents
End Sub

Private Sub cmdWFCHide_Click()
    If Left(cmdWFCHide.Caption, 4) = "Hide" Then
        xHideWFCAttCodes = True
        cmdWFCHide.Caption = "Unhide REG && OT" 'Hide REG && OT
    Else
        If Left(cmdWFCHide.Caption, 4) = "Unhi" Then
            xHideWFCAttCodes = False
            cmdWFCHide.Caption = "Hide REG && OT"
        End If
    End If
    Call EERetrieve
End Sub

Private Sub imgSec_Click()
    Dim SQLQ
    SQLQ = getSQL("frmVATTEND")
    Call FillMemoFile(SQLQ, "Attendance")
End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdPostion_Click()
Dim oJob As String, OJobD As String

oJob = clpJob.Text
OJobD = clpJob.Caption

Load frmJOBS
frmJOBS.Show 1

'If Len(glbJob) < 1 Then
If Len(glbPos) < 1 Then
    clpJob.Text = oJob
    clpJob.Caption = OJobD
Else
    clpJob.Text = glbPos
    clpJob.Caption = glbPosDesc
    
    'clpJob.Text = glbJob
    'clpJob.Caption = glbJobDesc
End If
End Sub

Sub cmdPrint_Click()
Dim RHeading As String, dscGroup$, xReport, X%
RHeading = lblEEName.Caption & "'s Attendance Information"

If glbtermopen Then
'    cmdPrint.Enabled = False
    Me.vbxCrystal.ReportSource = crptTrueDBGrid
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
Else
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
    
    ' dkostka - 02/22/2002 - Added print button support for SQL.
If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 2
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next
    End If
    If xAD = "AD" Then
        xReport = glbIHRREPORTS & "rgridatt.rpt"  'Attendance
        Me.vbxCrystal.ReportFileName = xReport
        Me.vbxCrystal.SelectionFormula = "{HR_ATTENDANCE.AD_EMPNBR}=" & glbLEE_ID & " "
    Else
        RHeading = lblEEName.Caption & "'s Attendance History Information"
        xReport = glbIHRREPORTS & "rgridath.rpt"  'Attendance history
        Me.vbxCrystal.ReportFileName = xReport
        Me.vbxCrystal.SelectionFormula = "{HR_ATTENDANCE_HISTORY.AH_EMPNBR}=" & glbLEE_ID & " "
    End If
End If
    Me.vbxCrystal.Destination = 1
    Me.vbxCrystal.Action = 1
    Me.vbxCrystal.Reset
'    cmdPrint.Enabled = True

End Sub
Sub cmdView_Click()
Dim RHeading As String, dscGroup$, xReport, X%
RHeading = lblEEName.Caption & "'s Attendance Information"

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

If glbtermopen Then
'    cmdPrint.Enabled = False
    Me.vbxCrystal.ReportSource = crptTrueDBGrid
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
Else
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
    
    ' dkostka - 02/22/2002 - Added print button support for SQL.
If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 2
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next
    End If
    If xAD = "AD" Then
        xReport = glbIHRREPORTS & "rgridatt.rpt"  'Attendance
        Me.vbxCrystal.ReportFileName = xReport
        If glbWFC And xHideWFCAttCodes Then  'Ticket #25148 Franks 03/04/2014
            Me.vbxCrystal.SelectionFormula = "{HR_ATTENDANCE.AD_EMPNBR}=" & glbLEE_ID & " AND NOT ({HR_ATTENDANCE.AD_REASON} = 'REG' OR {HR_ATTENDANCE.AD_REASON} = 'OT') "
        Else
            Me.vbxCrystal.SelectionFormula = "{HR_ATTENDANCE.AD_EMPNBR}=" & glbLEE_ID & " "
        End If
    Else
        RHeading = lblEEName.Caption & "'s Attendance History Information"
        xReport = glbIHRREPORTS & "rgridath.rpt"  'Attendance history
        Me.vbxCrystal.ReportFileName = xReport
        If glbWFC And xHideWFCAttCodes Then  'Ticket #25148 Franks 03/04/2014
            Me.vbxCrystal.SelectionFormula = "{HR_ATTENDANCE_HISTORY.AH_EMPNBR}=" & glbLEE_ID & " AND NOT ({HR_ATTENDANCE_HISTORY.AH_REASON} = 'REG' OR {HR_ATTENDANCE_HISTORY.AH_REASON} = 'OT') "
        Else
            Me.vbxCrystal.SelectionFormula = "{HR_ATTENDANCE_HISTORY.AH_EMPNBR}=" & glbLEE_ID & " "
        End If
    End If
End If
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.Action = 1
    Me.vbxCrystal.Reset
'    cmdPrint.Enabled = True

End Sub
'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdRecalPoints_Click()
Dim Msg, a%
    Msg = "This function will check Unexcused, Emergency Leave Flag," & Chr(10)
    Msg = Msg & "also reset the Absence Point" & Chr(10)
    Msg = Msg & "Are You Sure You Want To Do This? "
    a% = MsgBox(Msg, 36, "Confirm ")
    If a% <> 6 Then Exit Sub
    Call BTIPoint(lblEEID, True)
    'Data1.Refresh
    MsgBox " Recalculate Completed."
    Unload Me
End Sub


Private Sub cmdViewDiscip_Click()
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzdiscip.rpt"
    Me.vbxCrystal.Connect = RptODBC_SQL
    Me.vbxCrystal.SubreportToChange = "AttDetail"
    Me.vbxCrystal.Connect = RptODBC_SQL
    Me.vbxCrystal.SubreportToChange = ""
    Me.vbxCrystal.SelectionFormula = "{HRATTWRK.AD_WRKEMP}='" & glbUserID & "'"
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.WindowTitle = "Disciplinary Report"
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
End Sub

Private Sub comPayPer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub

Private Sub comShiftType_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comShiftType_LostFocus()
    If comShiftType.ListIndex = 0 Then txtShift.Text = "FD"
    If comShiftType.ListIndex = 1 Then txtShift.Text = "AM"
    If comShiftType.ListIndex = 2 Then txtShift.Text = "PM"
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "ATTEND", "SELECT")

End Sub



Private Function EENames()
'Dim SQLQ As String
'Dim countr   As Integer  ' EENames_Snap is definded at form level
''''On Error GoTo EENames_Err
'EENames = False         ' if not found - no depts
'Screen.MousePointer = HOURGLASS
'SQLQ = "Select * from qry_EE_Names "
'SQLQ = SQLQ & " Where " & glbSeleDeptUn
'SQLQ = SQLQ & " ORDER BY ED_EMPNBR"
'If fsnapEENames.State <> 0 Then fsnapEENames.Close
'fsnapEENames.Open SQLQ, gdbAdoIhr001, adOpenStatic
'Screen.MousePointer = DEFAULT
'EENames = True
'Exit Function
'EENames_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EENames", "HREMP", "Select")
'Call RollBack '28July99 js

End Function

Sub UpdateOTExpiry(ByVal ID As String)
    
    Dim SQLQ As String
    Dim tbName As String
    tbName = "HR_ATTENDANCE"
    If xAD = "AH" Then
        tbName = "HR_ATTENDANCE_HISTORY"
    End If
    SQLQ = "SELECT * FROM " & tbName & " WHERE " & xAD & "_ATT_ID = " & ID
    Dim rs As ADODB.Recordset
    Set rs = gdbAdoIhr001.Execute(SQLQ)
    
    
    If rs.EOF Then
         Exit Sub
    End If
    If rs.RecordCount > 0 Then
         If Left(clpCode(1).Text, 2) = "OT" Then
            'Leeds and Grenville - Ticket #19441 - Just incase they add OT records from info:HR Attendance
            'instead of ESS Request/Approval - I am updating with Expiry Date so that in the Weekly Adjustments
            'this record is included.
            If glbCompSerial = "S/N - 2233W" Then
                If tbName = "HR_ATTENDANCE" Then
                    SQLQ = "UPDATE " & tbName & " SET " & xAD & "_bankhrs_exp=" & Date_SQL(DateAdd("d", 30, CDate(dlpReviewDate.Text))) & " WHERE " & xAD & "_ATT_ID= " & ID
                    gdbAdoIhr001.Execute SQLQ
                End If
            Else
                SQLQ = "UPDATE " & tbName & " SET " & xAD & "_bankhrs_exp=" & Date_SQL(DateAdd("m", 6, CDate(dlpReviewDate.Text))) & " WHERE " & xAD & "_ATT_ID= " & ID
                gdbAdoIhr001.Execute SQLQ
            End If
         End If
    End If
    
End Sub

Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

'''On Error GoTo EERError
    If glbLinamar Then
        SQLQ = " CASE WHEN AD_SUPER IS NOT NULL AND LEN(AD_SUPER)>2 "
        SQLQ = SQLQ & " THEN RIGHT(AD_SUPER,3)+'-'+"
        SQLQ = SQLQ & " LEFT(AD_SUPER,LEN(AD_SUPER)-3) "
        SQLQ = SQLQ & " ELSE STR(AD_SUPER) END "
        SQLQ = SQLQ & " AS SUPER "
        'Ticket #28846 Franks 08/16/2016
        SQLQ = SQLQ & ", CASE WHEN AD_SHIFT IS NOT NULL AND LEN(AD_SHIFT)>3 "
        SQLQ = SQLQ & " THEN SUBSTRING(AD_SHIFT,4,2) "
        SQLQ = SQLQ & " ELSE AD_SHIFT END "
        SQLQ = SQLQ & " AS SHIFT"
    Else
        If glbOracle Then
            SQLQ = " AD_SUPER AS SUPER "
        Else
            SQLQ = " STR(AD_SUPER) AS SUPER "
        End If
        'Ticket #28846 Franks 08/16/2016
        SQLQ = "AD_SHIFT AS SHIFT "
    End If
Screen.MousePointer = HOURGLASS
If glbtermopen Then
    SQLQ = "SELECT Term_ATTENDANCE.*, " & SQLQ
    SQLQ = SQLQ & " FROM Term_ATTENDANCE "
    SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
    If glbWFC Then 'Ticket #25148 Franks 03/04/2014
        If xHideWFCAttCodes Then
            SQLQ = SQLQ & " AND NOT (AD_REASON = 'REG' OR AD_REASON = 'OT') "
        End If
    End If
    SQLQ = SQLQ & " ORDER BY AD_DOA DESC, AD_EMPNBR"
ElseIf xAD = "AD" Then
    SQLQ = "SELECT HR_ATTENDANCE.*, " & SQLQ
    
    SQLQ = SQLQ & " FROM HR_ATTENDANCE "
    SQLQ = SQLQ & " WHERE AD_EMPNBR = " & glbLEE_ID
    If glbWFC Then 'Ticket #25148 Franks 03/04/2014
        If xHideWFCAttCodes Then
            SQLQ = SQLQ & " AND NOT (AD_REASON = 'REG' OR AD_REASON = 'OT') "
        End If
    End If
    SQLQ = SQLQ & " ORDER BY AD_DOA DESC, AD_EMPNBR"
Else
    SQLQ = "SELECT " & Replace(SQLQ, "AD_", "AH_") & ","
    SQLQ = SQLQ & "AH_COMPNO AS AD_COMPNO,"
    SQLQ = SQLQ & "AH_EMPNBR AS AD_EMPNBR,"
    SQLQ = SQLQ & "AH_DOA AS AD_DOA,"
    SQLQ = SQLQ & "AH_HRS AS AD_HRS,"
    SQLQ = SQLQ & "AH_REASON AS AD_REASON,"
    SQLQ = SQLQ & "AH_CHRGCODE AS AD_CHRGCODE,"
    SQLQ = SQLQ & "AH_PROJECT_CODE AS AD_PROJECT_CODE,"
    SQLQ = SQLQ & "AH_SUPER AS AD_SUPER,"
    SQLQ = SQLQ & "AH_SHIFT AS AD_SHIFT,"
    SQLQ = SQLQ & "AH_WCBNBR AS AD_WCBNBR,"
    SQLQ = SQLQ & "AH_INCID AS AD_INCID,"
    SQLQ = SQLQ & "AH_FMLA AS AD_FMLA,"
    SQLQ = SQLQ & "AH_INDICATOR AS AD_INDICATOR,"
    SQLQ = SQLQ & "AH_SEN AS AD_SEN,"
    SQLQ = SQLQ & "AH_EMELEA AS AD_EMELEA,"
    SQLQ = SQLQ & "AH_COMM AS AD_COMM,"
    SQLQ = SQLQ & "AH_ATT_ID AS AD_ATT_ID,"
    
    SQLQ = SQLQ & "AH_JOB AS AD_JOB,"
    SQLQ = SQLQ & "AH_ORG AS AD_ORG,"
    SQLQ = SQLQ & "AH_SALARY AS AD_SALARY,"
    SQLQ = SQLQ & "AH_SALCD AS AD_SALCD,"
    SQLQ = SQLQ & "AH_DHRS AS AD_DHRS,"
    SQLQ = SQLQ & "AH_WHRS AS AD_WHRS,"
    SQLQ = SQLQ & "AH_POINT AS AD_POINT,"
    SQLQ = SQLQ & "AH_UPLOAD AS AD_UPLOAD,"
    SQLQ = SQLQ & "AH_CALCHRS AS AD_CALCHRS,"
    SQLQ = SQLQ & "AH_PAYROLL_ID AS AD_PAYROLL_ID,"
    SQLQ = SQLQ & "AH_GLNO AS AD_GLNO,"
   
    SQLQ = SQLQ & "AH_LDATE AS AD_LDATE,"
    SQLQ = SQLQ & "AH_LTIME AS AD_LTIME,"
    SQLQ = SQLQ & "AH_LUSER AS AD_LUSER,"
    SQLQ = SQLQ & "AH_MACHINE_RATE AS AD_MACHINE_RATE,"  'ticket 8332, error 3265
    SQLQ = SQLQ & "AH_MACHINE_HRS AS AD_MACHINE_HRS," 'ticket 8332, error 3265
    SQLQ = SQLQ & "AH_MACHINE_NUM AS AD_MACHINE_NUM," 'ticket 8332, error 3265
    'Hemu - already added above - duplicate entry
    'SQLQ = SQLQ & "AH_PROJECT_CODE AS AD_PROJECT_CODE," 'ticket 8332, error 3265
    'Hemu
    
    'Hemu - 06/08/2004 Begin - Ticket #6306
    SQLQ = SQLQ & "AH_DISCIPLINE AS AD_DISCIPLINE"
    'Hemu - 06/08/2004 End
    If glbBurlTech Then
        SQLQ = SQLQ & ",AH_LEPOINT AS AD_LEPOINT"
    End If
    If glbCompSerial = "S/N - 2242W" Then  'ccac london
        SQLQ = SQLQ & ",AH_PAYENDDATE AS AD_PAYENDDATE"
    End If
    
    SQLQ = SQLQ & ",AH_REGION AS AD_REGION"

    SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY "
    SQLQ = SQLQ & " WHERE AH_EMPNBR = " & glbLEE_ID
    If glbWFC Then 'Ticket #25148 Franks 03/04/2014
        If xHideWFCAttCodes Then
            SQLQ = SQLQ & " AND NOT (AH_REASON = 'REG' OR AH_REASON = 'OT') "
        End If
    End If
    SQLQ = SQLQ & " ORDER BY AH_DOA DESC, AH_EMPNBR"
End If

Data1.RecordSource = SQLQ
Data1.Refresh

EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Attendance", "HR Attendance", "SELECT")
Call RollBack '28July99 js

Exit Function

End Function

Private Function EntEccess()
Dim X%, Msg$, Answer, SQLQ, Logx, xWrk, SavWork0
Dim DgDef As Double, Response%, SavWork, Title$, Msg1$

EntEccess = False
Msg1$ = ". Select" & Chr(10)
Msg1$ = Msg1$ & "'Abort' to terminate this transaction," & Chr(10)
Msg1$ = Msg1$ & "'Retry' to modify the transaction or" & Chr(10)
Msg1$ = Msg1$ & "'Ignore' to update Attendance Master"
DgDef = MB_ABORTRETRYIGNORE + MB_ICONSTOP + MB_DEFBUTTON2
Title$ = "info:HR - ATTENDANCE ENTRY"
fglbRetry = False
If Left(clpCode(1).Text, 3) = "VAC" Then
    UPDVACSICK False
    SavWork = SavOutV - SavVac
    If DateValue(dlpReviewDate.Text) >= Fdate And DateValue(dlpReviewDate.Text) <= Tdate Then
        SavWork = SavWork - medHours
    End If
    If SavWork < 0 Then
        'Hemu   - Town of Ajax - Remove "Ignore" button
        If glbCompSerial = "S/N - 2173W" Then
            If (gSec_Add_Attendance And Not gSec_Upd_Attendance) Or (Not gSec_Add_Attendance And gSec_Upd_Attendance) Then
                Msg1$ = ". Select" & Chr(10)
                Msg1$ = Msg1$ & "'Cancel' to terminate this transaction or" & Chr(10)
                Msg1$ = Msg1$ & "'Retry' to modify the transaction" & Chr(10)
                DgDef = MB_RETRYCANCEL + MB_ICONSTOP + MB_DEFBUTTON2
            End If
        End If
        'Hemu
        
        SavWork0 = 0 - SavWork
        Msg$ = "Warning: VACATION Entitlement has" & Chr(10)
        Msg$ = Msg$ & "been exceeded by " & Format(SavWork0, "Fixed") & " hours" & Msg1$
        Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
        If Response% = IDABORT Or Response% = IDCANCEL Then   'Hemu
            If IsDate(dlpToDate.Text) Then
                dlpToDate.Text = ""
            Else
                Call cmdCancel_Click
            End If
        End If
        'If Response% = IDRETRY Or Response% = IDABORT Then GoTo EntEcc
        If Response% = IDRETRY Or Response% = IDABORT Or Response% = IDCANCEL Then GoTo EntEcc
    End If
End If

If Left(clpCode(1).Text, 3) = "SIC" Then
    UPDVACSICK False
    SavWork = SavOutS - SavSick
    If DateValue(dlpReviewDate.Text) >= fdateS And DateValue(dlpReviewDate.Text) <= tdateS Then
        SavWork = SavWork - medHours
    End If
    If SavWork < 0 Then
        If glbWHSCC Then 'Franks Aug 1,02 WHSCC
            Msg1$ = ". Select" & Chr(10)
            Msg1$ = Msg1$ & "'Yes' to assign ASL to remainder," & Chr(10)
            Msg1$ = Msg1$ & "'No' to edit this transaction"
            DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON1
            SavWork0 = 0 - SavWork
            Msg$ = "Warning: SICK TIME Entitlement has" & Chr(10)
            Msg$ = Msg$ & "been exceeded by " & Format(SavWork0, "Fixed") & " hours" & Msg1$
            Response% = MsgBox(Msg, DgDef, Title)
            If Response% = IDNO Then
                Response% = IDRETRY
                GoTo EntEcc
            Else
                whsccExceedOrgNum = Val(medHours)
                whsccExceedRemNum = SavWork0
                If Not EntASLEccess(glbLEE_ID, CVDate(dlpReviewDate), whsccExceedRemNum) Then
                    'Exit Function
                    Response% = IDRETRY
                    GoTo EntEcc
                End If
                medHours = Val(medHours) - SavWork0
                whsccExceedFlag = True
            End If
        Else
        
            'Hemu   - Town of Ajax - Remove "Ignore" button
            If glbCompSerial = "S/N - 2173W" Then
                If (gSec_Add_Attendance And Not gSec_Upd_Attendance) Or (Not gSec_Add_Attendance And gSec_Upd_Attendance) Then
                    Msg1$ = ". Select" & Chr(10)
                    Msg1$ = Msg1$ & "'Cancel' to terminate this transaction or" & Chr(10)
                    Msg1$ = Msg1$ & "'Retry' to modify the transaction" & Chr(10)
                    DgDef = MB_RETRYCANCEL + MB_ICONSTOP + MB_DEFBUTTON2
                End If
            End If
            'Hemu
        
            SavWork0 = 0 - SavWork
            Msg$ = "Warning: SICK TIME Entitlement has" & Chr(10)
            Msg$ = Msg$ & "been exceeded by " & Format(SavWork0, "Fixed") & " hours" & Msg1$
            Response% = MsgBox(Msg, DgDef, Title)
            If Response% = IDABORT Or Response% = IDCANCEL Then   'Hemu
                If IsDate(dlpToDate.Text) Then
                    dlpToDate.Text = ""
                Else
                    Call cmdCancel_Click
                End If
            End If
            'If Response% = IDRETRY Or Response% = IDABORT Then GoTo EntEcc
            If Response% = IDRETRY Or Response% = IDABORT Or Response% = IDCANCEL Then GoTo EntEcc
        End If
    End If
End If

'Town of Aurora
'If glbCompSerial = "S/N - 2378W" Then
    If Left(clpCode(1).Text, 2) = "CT" Or Left(clpCode(1).Text, 2) = "OT" Then
        UPDVACSICK False
        If Left(clpCode(1).Text, 2) = "CT" Then
            SavWork = SavOutOT - SavCT
            If DateValue(dlpReviewDate.Text) >= OTFdate And DateValue(dlpReviewDate.Text) <= OTTdate Then
                SavWork = SavWork - medHours
            End If
            If SavWork < 0 Then
                SavWork0 = 0 - SavWork
                Msg$ = "Warning: OVERTIME BANKED taken has" & Chr(10)
                Msg$ = Msg$ & "been exceeded by " & Format(SavWork0, "Fixed") & " hours" & Msg1$
                Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                If Response% = IDABORT Or Response% = IDCANCEL Then   'Hemu
                    If IsDate(dlpToDate.Text) Then
                        dlpToDate.Text = ""
                    Else
                        Call cmdCancel_Click
                    End If
                End If
                'If Response% = IDRETRY Or Response% = IDABORT Then GoTo EntEcc
                If Response% = IDRETRY Or Response% = IDABORT Or Response% = IDCANCEL Then GoTo EntEcc
                
                'Exceeded in using the Overtime Bank time - send an email
                If SaveOTEmail <> "" Then
                    Call Send_Email_Overtime_Exceeded(SaveOTEmail, SavWork0, "CT")
                    If fglbSendOTEmail = False Then
                        MsgBox "Outstanding Overtime Bank exceeded email cannot be sent." & vbCrLf & "Please check the Setup -> Overtime Master screen for the correct Email address.", vbOKOnly, "Outstanding Overtime Bank Exceeded Email"
                        'commented by Bryan Ticket#11837 Oct 5
                        'Just because they can't send email doesn't mean the transaction should be cancelled.
                        'Call cmdCancel_Click
                    End If
                End If
            End If
        ElseIf Left(clpCode(1).Text, 2) = "OT" Then
            If SavMaxBank <> "" Then
                SavWork = SavMaxBank - SavOTBank 'SavOvt
                If DateValue(dlpReviewDate.Text) >= OTFdate And DateValue(dlpReviewDate.Text) <= OTTdate Then
                    'Ticket 11837, Oct 5
                    'Bryan removed the stuff in brackets... we've already subtracted SavOTBank, why subtract it again??
                    SavWork = SavWork - medHours ' (SavOTBank + medHours)    'SavWork - medHours
                End If
                If SavWork < 0 Then
                    SavWork0 = 0 - SavWork
                    Msg$ = "Warning: OVERTIME BANK has" & Chr(10)
                    Msg$ = Msg$ & "exceeded the Maximum Bank by " & Format(SavWork0, "Fixed") & " hours" & Msg1$
                    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                    If Response% = IDABORT Or Response% = IDCANCEL Then   'Hemu
                        If IsDate(dlpToDate.Text) Then
                            dlpToDate.Text = ""
                        Else
                            Call cmdCancel_Click
                        End If
                    End If
                    'If Response% = IDRETRY Or Response% = IDABORT Then GoTo EntEcc
                    If Response% = IDRETRY Or Response% = IDABORT Or Response% = IDCANCEL Then GoTo EntEcc
                    
                    'The ignore button has been pressed
                    EntEccess = True
                    
                    'Exceeded in using the Overtime Bank time - send an email
                    If SaveOTEmail <> "" Then
                        Call Send_Email_Overtime_Exceeded(SaveOTEmail, SavWork0, "MAX")
                        If fglbSendOTEmail = False Then
                            MsgBox "Overtime Bank exceeded Maximum Bank email cannot be sent." & vbCrLf & "Please check the Setup -> Overtime Master screen for the correct Email address.", vbOKOnly, "Outstanding Overtime Bank Exceeded Email"
                            'Commented by Bryan, Just because they can't send the email they've already hit the ignore buton on the transaction
                            'Call cmdCancel_Click
                        Else
                        End If
                    End If
                End If
            End If
        End If
    End If
'End If

xWrk = HRENTITLE(clpCode(1).Text, medHours)
If xWrk < 0 Then
    'possible that this has been set to true from OT and should be false here
    EntEccess = False
    'Hemu   - Town of Ajax - Remove "Ignore" button
    If glbCompSerial = "S/N - 2173W" Then
        If (gSec_Add_Attendance And Not gSec_Upd_Attendance) Or (Not gSec_Add_Attendance And gSec_Upd_Attendance) Then
            Msg1$ = ". Select" & Chr(10)
            Msg1$ = Msg1$ & "'Cancel' to terminate this transaction or" & Chr(10)
            Msg1$ = Msg1$ & "'Retry' to modify the transaction" & Chr(10)
            DgDef = MB_RETRYCANCEL + MB_ICONSTOP + MB_DEFBUTTON2
        End If
    End If
    'Hemu

    SavWork0 = 0 - xWrk
    'Msg$ = "Warning: [" & clpCode(1).Text & "] Entitlement has" & Chr(10)
    'Ticket #28279 Franks 03/04/2016 - add "Hourly " in front of Entitlement
    Msg$ = "Warning: [" & clpCode(1).Text & "] Hourly Entitlement has" & Chr(10)
    Msg$ = Msg$ & "been exceeded by " & Format(SavWork0, "Fixed") & " hours" & Msg1$
    Response% = MsgBox(Msg, DgDef, Title)
    'If Response% = IDABORT Then
    If Response% = IDABORT Or Response% = IDCANCEL Then   'Hemu
            If IsDate(dlpToDate.Text) Then
                dlpToDate.Text = ""
            Else
                Call cmdCancel_Click
            End If
    End If
    'If Response% = IDRETRY Or Response% = IDABORT Then GoTo EntEcc
    If Response% = IDRETRY Or Response% = IDABORT Or Response% = IDCANCEL Then GoTo EntEcc
End If

EntEccess = True

Exit Function
EntEcc:
If Response% = IDRETRY Then
    fglbRetry = True
'    clpCode(1).SetFocus
End If
End Function

Private Function EmlEntEccess()
Dim X%, Msg$, Answer, SQLQ, Logx, xWrk, SavWork0
Dim DgDef As Double, Response%, SavWork, Title$, Msg1$

EmlEntEccess = False
Msg1$ = ". Select" & Chr(10)
Msg1$ = Msg1$ & "'Abort' to terminate this transaction," & Chr(10)
Msg1$ = Msg1$ & "'Retry' to modify the transaction or" & Chr(10)
Msg1$ = Msg1$ & "'Ignore' to update Attendance Master"
DgDef = MB_ABORTRETRYIGNORE + MB_ICONSTOP + MB_DEFBUTTON2
Title$ = "info:HR - ATTENDANCE ENTRY"
fglbRetry = False

If chkEMELEA Then
    Calculate_EML_Taken
    SavWork = SavOutE - SavEML
    If DateValue(dlpReviewDate.Text) >= CVDate(GetMonth("Jan") & " 1," & Year(Date)) And DateValue(dlpReviewDate.Text) <= CVDate(GetMonth("Dec") & " 31," & Year(Date)) Then
        'Linamar added by Bryan 13/Oct/05 Tiocket#9264
        If Not glbLinamar Then
            SavWork = SavWork - medHours
        Else
            SavWork = SavWork - 1
        End If
    End If
    If SavWork < 0 Then
        
        'Hemu   - Town of Ajax - Remove "Ignore" button
        If glbCompSerial = "S/N - 2173W" Then
            If (gSec_Add_Attendance And Not gSec_Upd_Attendance) Or (Not gSec_Add_Attendance And gSec_Upd_Attendance) Then
                Msg1$ = ". Select" & Chr(10)
                Msg1$ = Msg1$ & "'Cancel' to terminate this transaction or" & Chr(10)
                Msg1$ = Msg1$ & "'Retry' to modify the transaction" & Chr(10)
                DgDef = MB_RETRYCANCEL + MB_ICONSTOP + MB_DEFBUTTON2
            End If
        End If
        'Hemu
        
        SavWork0 = 0 - SavWork
        Msg$ = "Warning: Emergency Leave Entitlement has" & Chr(10)
        Msg$ = Msg$ & "been exceeded by " & Format(SavWork0, "Fixed")
        'Linamar added by Bryan 13/Oct/05 Tiocket#9264
        If Not glbLinamar Then
            Msg$ = Msg$ & " hours" & Msg1$
        Else
            Msg$ = Msg$ & " incidents" & Msg1$
        End If
        Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
        If Response% = IDABORT Or Response% = IDCANCEL Then   'Hemu
            If IsDate(dlpToDate.Text) Then
                dlpToDate.Text = ""
            Else
                Call cmdCancel_Click
            End If
        End If
        If Response% = IDRETRY Or Response% = IDABORT Or Response% = IDCANCEL Then GoTo EntEcc1
    End If
End If


EmlEntEccess = True

Exit Function
EntEcc1:
If Response% = IDRETRY Then
    fglbRetry = True
'    clpCode(1).SetFocus
End If
End Function
Private Sub Fd_CurrJob()
Dim SQLQ As String
Dim rsJOB As New ADODB.Recordset
If glbtermopen Then
    SQLQ = "Select JH_REPTAU,JH_SHIFT,JH_DHRS FROM Term_JOB_HISTORY"
    SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
    SQLQ = SQLQ & " and JH_CURRENT <>0 "
    rsJOB.Open SQLQ, gdbAdoIhr001X, adOpenStatic
Else
    SQLQ = "Select JH_REPTAU,JH_SHIFT,JH_DHRS FROM HR_JOB_HISTORY"
    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " and JH_CURRENT <>0 "
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
End If

If rsJOB.BOF And rsJOB.EOF Then
    SavSup = ""
    SavShift = ""
    SaveHours = 0
Else
    SavSup = IIf(IsNull(rsJOB("JH_REPTAU")), "", rsJOB("JH_REPTAU"))
    SavShift = IIf(IsNull(rsJOB("JH_SHIFT")), "", rsJOB("JH_SHIFT"))
    SaveHours = IIf(IsNull(rsJOB("JH_DHRS")), 0, rsJOB("JH_DHRS"))
End If

End Sub

Private Sub dlpReviewDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub
Private Sub dlpReviewDate_MultiSelect(FromDate As Date, ToDate As Date)
    dlpReviewDate = FromDate
    dlpToDate = ToDate
End Sub

Private Sub dlpToDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub
Private Sub elpSupShow_Change()
'txtSup.Text = getEmpnbr(elpSupShow.Text)
End Sub

Private Sub elpSupShow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub

Private Sub Form_Activate()

glbOnTop = "FRMVATTEND"
Call SET_UP_MODE

End Sub

Private Sub Form_Deactivate()
    SavDeac = True
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMVATTEND"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub

Public Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim X%
Dim xtime, SQLQ As String, numInc%

xtime = Time
glbOnTop = "FRMVATTEND"
SavDeac = False
SavOutOT = 0

clpJob.TextBoxWidth = 1215 'Ticket #26726 Franks 06/15/2015

If (glbNiagaraFulls Or glbCompSerial = "S/N - 2173W") And (Not glbtermopen) Then
    lblOvertime.Visible = True
    txtOvertime.Visible = True
    lblOvertimeDays.Visible = True
Else
    lblOvertime.Visible = False
    txtOvertime.Visible = False
    lblOvertimeDays.Visible = False
End If

If glbCompSerial = "S/N - 2242W" Then  'ccac london
    chkBackDated.Visible = True
'    lblTitle(22).Visible = True
'    dlpPayEndDate.Visible = True
    dlpPayEndDate.DataField = "AD_PAYENDDATE"
End If

'Hemu - EML
'If glbCompSerial = "S/N - 2288W" Or glbCompSerial = "S/N - 2350W" Then
    lblEMLTaken.Visible = True
    lblEMLDay.Visible = True
    lblTitle(19).Visible = True
'End If
'Hemu

If glbVadim Then
    lblTitle(10).FontBold = True
    lblTitle(11).FontBold = True
    clpCode(4).isVadim = True
End If

If glbCompSerial = "S/N - 2375W" Then   'City of Timmins
    'Hide G/L No field.
    clpGLNo.Visible = False
    lblTitle(20).Visible = False
End If

If glbCompSerial = "S/N - 2279W" Then   'Friesens Corporation - Ticket #16189
    'Hide Machine Code fields
    clpCode(4).Visible = False
    lblMachine.Visible = False
End If

'WDGPHU - Ticket #26576
'Leeds and Grenville - Ticket #19441
If glbCompSerial = "S/N - 2233W" Or (glbCompSerial = "S/N - 2411W" And gsFLEX_LOGIC) Then
    ChkInc.Caption = "Frozen"
    ChkInc.Tag = "Frozen"
    vbxTrueGrid(0).Columns(4).Caption = "Frozen"
    ChkInc.Enabled = False
End If

If xAttendance = "Attendance_History" Then
    xAD = "AH"

    'Ticket #18668
    lblTitle(25).Visible = False
    lblCompTimeOS.Visible = False
    lbCompTimeOSday.Visible = False
    
    '7.9 - Changes
    lblTitle(16).Visible = False
    lblVACOS.Visible = False
    lblVACOSday.Visible = False
    
    lblOvertime.Visible = False
    txtOvertime.Visible = False
    lblOvertimeDays.Visible = False
    
    lblTitle(17).Visible = False
    lblSICKOS.Visible = False
    lblSICKOSday.Visible = False
    
    lblTitle(25).Visible = False
    lblCompTimeOS.Visible = False
    lbCompTimeOSday.Visible = False
    
    lblTitle(18).Visible = False
    lblASLOS.Visible = False
    lbASLOSday.Visible = False
    
    lblTitle(19).Visible = False
    lblEMLTaken.Visible = False
    lblEMLDay.Visible = False
    
    lblEMLOS.Visible = False
    lblEMLOSV.Visible = False
    lblDays.Visible = False
Else
    xAD = "AD"

    'Ticket #18668
    lblTitle(25).Visible = True
    lblCompTimeOS.Visible = True
    lbCompTimeOSday.Visible = True

    '7.9 - Changes
    lblTitle(16).Visible = True
    lblVACOS.Visible = True
    lblVACOSday.Visible = True
        
    lblTitle(17).Visible = True
    lblSICKOS.Visible = True
    lblSICKOSday.Visible = True
            
    lblTitle(19).Visible = True
    lblEMLTaken.Visible = True
    lblEMLDay.Visible = True
    
    lblEMLOS.Visible = True
    lblEMLOSV.Visible = True
    lblDays.Visible = True
End If

chkEMELEA.Enabled = False

If xAttendance <> "Attendance_History" Then
    lblTitle(16).Visible = gSec_Inq_Entitlements
    lblVACOS.Visible = gSec_Inq_Entitlements
    lblVACOSday.Visible = gSec_Inq_Entitlements
End If

'wellington duffrine ticket ##17736
'S.U.C.C.E.S.S - Ticket #19099
If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
    lblVACOSday.Caption = "Hour"
    lblSICKOSday.Caption = "Hour"
End If

If xAttendance <> "Attendance_History" Then
    lblTitle(17).Visible = gSec_Inq_Entitlements
    lblSICKOS.Visible = gSec_Inq_Entitlements
    lblSICKOSday.Visible = gSec_Inq_Entitlements
End If

Call setCaption(lblTitle(1))
Call setCaption(lblTitle(13))
Call setCaption(lblTitle(2))
Call setCaption(lblTitle(3))
Call setCaption(lblTitle(4))
Call setCaption(lblTitle(5))
Call setCaption(lblTitle(6))
Call setCaption(lblTitle(24))
Call setCaption(lblTitle(15))
Call setCaption(lblTitle(20))
Call setCaption(lblTitle(23))
Call setCaption(lblMachine)

lblTitle(4).Caption = lStr("Hours")

If glbCompSerial = "S/N - 2411W" Then 'WDGPHU - Ticket #24655
    lblRegion.Visible = True
    clpCode(5).Visible = True
    Call setCaption(lblRegion)
    
    'Ticket #27910 - Option to change Salary information
    cmdSalaryChange.Left = cmdViewDiscip.Left
    cmdSalaryChange.Visible = True
Else
    lblRegion.Visible = False
    clpCode(5).Visible = False
    cmdSalaryChange.Visible = False
End If

'Ticket #23537 and Release 8.0
lblHrsDay.Caption = lStr("Hours/Day")
lblHrsWeek.Caption = lStr("Hours/Week")


If glbCompSerial = "S/N - 2430W" Then 'Ticket #21167 Franks 11/07/2011
    lblTitle(23).Caption = "Program"
End If

If lblTitle(3).Caption = "AttSupervisor" Then lblTitle(3).Caption = "Supervisor"

'If glbCompSerial = "S/N - 2351W" Then ' For Burlington Tech.
'    clpChrgCode.Visible = True
'    clpChrgCode.LookupType = ChargeCode
'    clpChrgCode.MaxLength = 7
'    txtChrgCode.Visible = False
If glbCompSerial = "S/N - 2217W" Or glbCompSerial = "S/N - 2214W" Or glbCompSerial = "S/N - 2241W" Then
    clpChrgCode.Visible = True
    clpChrgCode.LookupType = Department
    clpChrgCode.MaxLength = 7
    txtChrgCode.Visible = False
ElseIf glbCompSerial = "S/N - 2376W" Then ' AFN Ticket #16251
    clpChrgCode.Visible = True
    clpChrgCode.LookupType = Division
    clpChrgCode.MaxLength = 4
    txtChrgCode.Visible = False
    lblTitle(5).Caption = lStr("Division")
    lblTitle(5).Tag = "00-Enter " & lStr("Division") & ""
    vbxTrueGrid(0).Columns(7).Caption = lStr("Division")
ElseIf glbCompSerial = "S/N - 2192W" Then  ' county essex
    'charge code link to gl account
    lblTitle(5) = lStr("G/L #")
    txtChrgCode.Visible = False
    clpChrgCode.Visible = True
    clpChrgCode.LookupType = GL
    clpChrgCode.MaxLength = 25
    
    'use multi frame
    frmMulti.Visible = True
    'do not show job
    cmdPostion.Visible = False
    lblTitle(8).Visible = False
    clpJob.Visible = False
    ' union
    lblTitle(9).Visible = False
    clpCode(0).Visible = False
    
    lblTitle(10).Top = 270
    medsalary.Top = 270
    
    lblTitle(11).Top = 570
    comPayPer.Top = 570
    
    lblHrsDay.Visible = False
    txtDHRS.Visible = False
    lblHrsWeek.Visible = False
    txtWHRS.Visible = False
    lblTitle(20).Visible = False
    clpGLNo.Visible = False
    lblPayID.Visible = False
    txtPayrollID.Visible = False
    
    lblMachineRate.Visible = True
    medMachineRate.Visible = True
    medMachineRate.DataField = "AD_MACHINE_RATE"
    
    
    lblMachineHours.Visible = True
    medMachineHours.Visible = True
    medMachineHours.DataField = "AD_MACHINE_HRS"
    
    frmMulti.Height = 1000
ElseIf glbCompSerial = "S/N - 2366W" Then 'FYC Muskoka
    'do not show job
    cmdPostion.Visible = False
    lblTitle(8).Visible = False
    clpJob.Visible = False
    ' union
    lblTitle(9).Visible = False
    clpCode(0).Visible = False
    ' Hours/day
    lblHrsDay.Visible = False
    txtDHRS.Visible = False
    lblHrsWeek.Visible = False
    txtWHRS.Visible = False
    'GLNo
    lblTitle(20).Visible = False
    clpGLNo.Visible = False
    lblPayID.Visible = False
    txtPayrollID.Visible = False
    
ElseIf glbCBrant Then
    clpChrgCode.Visible = True
    clpChrgCode.LookupType = Job
    clpChrgCode.MaxLength = 6
    txtChrgCode.Visible = False
End If

If glbWHSCC Then
    lblTitle(18).Visible = True
    lblASLOS.Visible = True
    lbASLOSday.Visible = True
End If

If (glbWFC And glbPlantCode = "WHBY") Or glbBurlTech Then
    lblTitle(5).Visible = False
    txtChrgCode.Visible = False
    lblTitle(14).Visible = True
    txtCode(0).Visible = True
    lblTitle(14).Top = lblTitle(5).Top
    txtCode(0).Top = txtChrgCode.Top
    lblCodeDesc(0).Top = txtChrgCode.Top
    vbxTrueGrid(0).Columns(7).Caption = "Disciplinary"
    vbxTrueGrid(0).Columns(7).DataField = "AD_DISCIPLINE"
    If (glbWFC And glbPlantCode = "WHBY") Then
        cmdViewDiscip.Visible = True
    End If
End If
If glbWFC Then 'Ticket #25148 Franks 03/04/2014
    cmdWFCHide.Visible = True
    xHideWFCAttCodes = False
End If
If glbVadim Then
    lblMachine = "Equipment"
    clpCode(4).TABLTitle = "Equipment"
    clpCode(4).Tag = "Equipment #"
End If
If glbCompSerial = "S/N - 2350W" Then
    ChkInc.Visible = False
    chkSeniority.Visible = False
    lblTitle(17).Visible = False
    lblSICKOS.Visible = False
    lblSICKOSday.Visible = False
End If
If glbCompSerial = "S/N - 2454W" Then 'Showa Ticket #25250 Franks 04/07/2014
    ChkInc.Visible = False
End If

If glbCompSerial = "S/N - 2347W" Then 'Surrey Place
    ChkInc.Caption = "LTD"
    ChkInc.Tag = "LTD"
    vbxTrueGrid(0).Columns(4).Caption = "LTD"
End If

If glbCompSerial = "S/N - 2214W" Then 'Casey House  - Ticket #15276
    ChkInc.Caption = "HOOPP"
    ChkInc.Tag = "HOOPP"
    vbxTrueGrid(0).Columns(4).Caption = "HOOPP"
End If

If glbCompSerial = "S/N - 2388W" Then 'DNSSAB Ticket #14260
    ChkInc.Caption = "No Sick Ent"
    ChkInc.Tag = "No Sick Ent"
    vbxTrueGrid(0).Columns(4).Caption = "No Sick Ent"
End If

If glbCompSerial = "S/N - 2376W" Then   'Assembly of First Nations - Ticket #16181
    clpCode(1).MaxLength = 10
End If

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If
Call setCaption(lblTitle(9))
Screen.MousePointer = HOURGLASS

If glbMulti Or glbCompSerial = "S/N - 2192W" Or glbCompSerial = "S/N - 2366W" Or glbCompSerial = "S/N - 2363W" Then
    comPayPer.Clear
    comPayPer.AddItem "Annum"
    comPayPer.AddItem "Hour"
    comPayPer.AddItem "Monthly"
    'If glbCompSerial = "S/N - 2282W" Then  - It's for everyone
        comPayPer.AddItem "Daily "
    'End If
    frmMulti.Visible = True
Else
    comPayPer.Clear
    comPayPer.AddItem "Annum"
    comPayPer.AddItem "Hour"
    comPayPer.AddItem "Monthly"
    'If glbCompSerial = "S/N - 2282W" Then 'Ticket #12019   - It's for everyone
        comPayPer.AddItem "Daily "
    'End If
End If

'Hemu - 02/16/2004 Begin - Brant County Health Unit - Ticket # 5600
If glbCompSerial = "S/N - 2226W" Then
    comShiftType.Clear
    comShiftType.AddItem "FD"
    comShiftType.AddItem "AM"
    comShiftType.AddItem "PM"
    comShiftType.Visible = True
Else
    comShiftType.Visible = False
End If
'Hemu - 02/16/2004 End

If glbBurlTech Then
    vbxTrueGrid(0).Columns(3).Visible = False
    vbxTrueGrid(0).Columns(4).Caption = "Unexcused"
    vbxTrueGrid(0).Columns(5).Caption = "Excused"
    vbxTrueGrid(0).Columns(9).Caption = "Absence Point"
    'chkIncident.Visible = False 'Ticket #13372 make it visible
    ChkInc.Caption = "Unexcused"
    chkSeniority.Caption = "Excused"
    lblTitle(15).Caption = "Absence Point"
    'lblTitle(21).Visible = True
    'txtLEPoint.Visible = True
    'lblTitle(21).Top = lblTitle(15).Top
    'txtLEPoint.Top = txtPoint.Top
    'vbxTrueGrid(0).Columns(10).DataField = "AD_LEPOINT"
    'txtLEPoint.DataField = "AD_LEPOINT"
    cmdRecalPoints.Visible = True
    lblIncident.Visible = False
    vbxTrueGrid(0).Columns(10).Visible = False
    'Burlington doesn't want to use Machine number, it is used by Payweb interface though.
    'Bryan Mar 28, 2007 Ticket#12667
    lblMachine.Visible = False
    clpCode(4).Visible = False
End If

If glbLambton Or glbCompSerial = "S/N - 2394W" Or glbCompSerial = "S/N - 2241W" Then 'or St. John's Rehab - Ticket #14954
    txtShift.MaxLength = 4
End If

X% = EENames()

If glbCompSerial = "S/N - 2235W" Or glbCompSerial = "S/N - 2236W" Then
    lblTitle(4).Caption = "Days"
    vbxTrueGrid(0).Columns(3) = "Days"
End If
Screen.MousePointer = DEFAULT
AddChg = " "

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
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If


Screen.MousePointer = HOURGLASS

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    If xAD = "AD" Then
        Me.Caption = "Attendance - " & Left$(glbLEE_SName, 5)
    Else
        Me.Caption = "Attendance_History - " & Left$(glbLEE_SName, 5)
    End If
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEmpID.Caption = lblEEID

UpdateIncCount RW_READ

Call Display_Value
Call modSTUPD(False)

If Not gSec_Upd_Attendance Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
End If

Call Fd_CurrJob 'Ticket# 7293 moved back from the bottom
Call CR_JobHis_Snap

If glbCompSerial = "S/N - 2173W" Then
    Call Check_Overtime_Bank(False)
End If

' dkostka - 12/20/2001 - Moved incident counting to seperate sub.
''UpdateIncCount RW_READ
UPDVACSICK False

'Hemu - EML
'If glbSQL Or glbOracle Then UPDEML False

'Re-opened it up again for Linamar because of Ticket #14376 - speed issue has been resolved
'with indexing on the HR_ATTENDANCE database. When they upgraded the index was not recreated.
'Ticket #12772, do not call this function on form load, it takes long time for large database (Linamar)
'Ticket #13813, There was no EML taken on the screen if this was commented, but donot show it for Linamar
'If Not glbLinamar Then
    Call Calculate_EML_Taken
'End If
'Hemu

'Calculate Comp Time Outstanding - Ticket #17345
Call Calculate_Outstanding_CompTime

'Hemu - Testing ASL
'Franks 05/22/2003 Ticket# 4103 Show ASL Balance in this year on Attendance screen
If glbWHSCC Then
    Call EntASLBalance(glbLEE_ID)
End If
'Hemu

If SavDeac Then
    SavDeac = False
    Exit Sub
End If
If glbCountry = "U.S.A." Then
    ChkFMLA.Visible = True
Else
    ChkFMLA.Visible = False
End If


'Ticket #30192 - Moved the Job selection above here before Salary selection, and looked for Default Position option if glbMulti and then took the Start Date and Job
'from Job History and looked for corresponding salary in the Salary History.
'Get current Job FOR Casey House
Dim rsSalT As New ADODB.Recordset

SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS,JH_SDATE FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & glbLEE_ID
If glbMulti Then    'Ticket #30192
    SQLQ = SQLQ & " AND JH_POSITION_CONTROL = 'YES'"
End If
xJob = "": xDHrs = 0: xWHrs = 0: xSDate = ""
rsSalT.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsSalT.EOF Then
    xJob = rsSalT("JH_JOB")
    xDHrs = rsSalT("JH_DHRS")
    xWHrs = rsSalT("JH_WHRS")
    xSDate = rsSalT("JH_SDATE")
End If
rsSalT.Close

'Get current salary FOR Casey House
SQLQ = "SELECT SH_EMPNBR,SH_CURRENT,SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & glbLEE_ID
SQLQ = SQLQ & " AND SH_JOB = '" & xJob & "'"    'Ticket #30192
SQLQ = SQLQ & " AND SH_SDATE = " & Date_SQL(xSDate) 'Ticket #30192
rsSalT.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsSalT.EOF Then
    xCurSalary = rsSalT("SH_SALARY")
    xSalCD = rsSalT("SH_SALCD")
Else
    xCurSalary = 0
    xSalCD = ""
End If
rsSalT.Close

SQLQ = "SELECT ED_EMPNBR, ED_DEPTNO,ED_GLNO,ED_ADMINBY,ED_ORG,ED_PAYROLL_ID,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
xORG = "": xDeptno = "": xAdminBy = "": xGLNO = "": xDiv = ""
rsSalT.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsSalT.EOF Then
    If Not IsNull(rsSalT("ED_DEPTNO")) Then
        xDeptno = rsSalT("ED_DEPTNO")
    End If
    If Not IsNull(rsSalT("ED_DIV")) Then 'Ticket #16251
        xDiv = rsSalT("ED_DIV")
    End If
    If Not IsNull(rsSalT("ED_GLNO")) Then
        xGLNO = rsSalT("ED_GLNO")
    End If
    If Not IsNull(rsSalT("ED_ADMINBY")) Then
        xAdminBy = rsSalT("ED_ADMINBY")
    End If
    If Not IsNull(rsSalT("ED_ORG")) Then
        xORG = rsSalT("ED_ORG")
    End If
    If Not IsNull(rsSalT("ED_PAYROLL_ID")) Then
        xPayrollID = rsSalT("ED_PAYROLL_ID")
    End If
End If
rsSalT.Close

If glbCompSerial = "S/N - 2214W" Then '
    lblTitle(5).Caption = "Department"
    lblTitle(6).Caption = lStr("Administered By") '"Fund"
    lblTitle(24).Caption = lStr("G/L #") '"Account #"
    lblTitle(5).Tag = "00-Enter Department"
    lblTitle(6).Tag = "00-Enter " & lStr("Administered By")  'Fund"
    lblTitle(24).Tag = "00-Enter Account #"
    vbxTrueGrid(0).Columns(7).Caption = "Department"
    vbxTrueGrid(0).Columns(8).Caption = lStr("Administered By")  '"Fund"
    'Save or Edit AD_SHIFT using clpCode(2) instead of txtShift
    txtShift.Visible = False
    txtShift.MaxLength = 4
    clpCode(2).Visible = True
    'Save or Edit AD_WCBNBR using clpGLNum instead of txtWSIB
    txtWSIB.Visible = False
    clpGLNum.MaxLength = 20
    clpGLNum.Visible = True
End If

If glbCompSerial = "S/N - 2376W" Then  'Ticket #16589 AFN
    'charge code link to gl account
    lblTitle(5) = lStr("G/L #")
    txtChrgCode.Visible = False
    clpChrgCode.Visible = True
    clpChrgCode.LookupType = GL
    clpChrgCode.MaxLength = 25
End If

If glbCompSerial = "S/N - 2396W" Then  'Ticket #17323 - Oshawa CHC
    'Charge Code link to GL Account
    lblTitle(5).Caption = lStr("G/L #")
    txtChrgCode.Visible = False
    clpChrgCode.Visible = True
    clpChrgCode.LookupType = GL
    clpChrgCode.MaxLength = 25
    clpChrgCode.Tag = lStr("00-Enter G/L #")
End If

If glbCompSerial = "S/N - 2241W" Then
    lblTitle(5).Caption = "Department"
    clpChrgCode.Tag = "00-Enter Department"
    vbxTrueGrid(0).Columns(7).Caption = "Department"
End If

clpCode(3).TextBoxWidth = 1500
clpGLNo.TextBoxWidth = 1500

If glbLinamar Then 'Ticket #28846 Franks 08/16/2016
    Call LinamarSceenSetup
End If

Call INI_Controls(Me)
'Call Fd_CurrJob
'Call CR_JobHis_Snap

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

'Hemu   - Town of Ajax - Remove "Ignore" button
If glbCompSerial = "S/N - 2173W" And ((gSec_Add_Attendance And Not gSec_Upd_Attendance) Or (Not gSec_Add_Attendance And gSec_Upd_Attendance)) Then
    Keepfocus = False
Else
    'Ticket #26576 - WDGPHU - Cannot maintain FX* codes from Attendance
    If glbCompSerial = "S/N - 2411W" And UCase(Left(clpCode(1).Text, 2)) = "FX" And gsFLEX_LOGIC Then
        Call cmdCancel_Click
    ElseIf gsDISABLE_COMPTIME And (UCase(Left(clpCode(1).Text, 2)) = "OT" Or UCase(Left(clpCode(1).Text, 2)) = "CT") Then
        'Ticket #30305 - Disable Compensatory Time Entries
        Call cmdCancel_Click
    Else
        Keepfocus = Not isUpdated(Me)
    End If
End If
'Hemu

Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Resize()

fraDetail.Height = 5535 '5000
If Me.ScaleWidth - 200 > 0 Then vbxTrueGrid(0).Width = Me.ScaleWidth - 200
If Me.Height >= vbxTrueGrid(0).Height + panEEName.Height + fraDetail.Height + panControls.Height Then '+ 230 Then
    scrControl.Value = 0
    fraDetail.Top = vbxTrueGrid(0).Height + panEEName.Height + 240
    scrControl.Visible = False
    Exit Sub
End If
If Me.Height < vbxTrueGrid(0).Height + panEEName.Height + scrControl.Top + panControls.Height + 400 Then Exit Sub
scrControl.Visible = True

scrControl.Max = vbxTrueGrid(0).Height + panEEName.Height + fraDetail.Height + panControls.Height + 250 - Me.Height
scrControl.Left = Me.Width - scrControl.Width - 120
scrControl.Height = Me.Height - scrControl.Top - panControls.Height - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)

    MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
    Set frmVATTEND = Nothing 'carmen may 00
    Call NextForm
End Sub

Private Function HRENTITLE(savKey, SavEnt)
Dim rsTB As New ADODB.Recordset
Dim rsAttend As New ADODB.Recordset
Dim SQLQ
Dim xWrk

'Ticket #17924 - Begin
Dim xOrgSavKey As String
If Right(savKey, 1) = "+" Then Exit Function    'This is Earned code - no warning message on this
If Right(savKey, 1) = "-" Or Mid(savKey, 3, 1) = "-" Then
    'If the code has suffix of "-" then it is hourly entitl. with Flex logic
    'Ticket #1859 - Or the 3rd character is "-" then it is hourly entitlement additional Flex logic
    'Save the original code for Attendance records retrieval
    xOrgSavKey = savKey

    'Replace with "+" so the Hourly Entitlement record is retrieved which is suffixed with "+ "
    If Mid(savKey, 3, 1) = "-" Then
        'Ticket #1859
        'This is multiple taken codes for one bank - only need the first two chars and then concatenate with +
        savKey = Left(savKey, 2) & "+"
    Else
        savKey = Left(savKey, Len(savKey) - 1) & "+"
    End If
End If
'Ticket #17924 - End

HRENTITLE = 0
SQLQ = "SELECT HE_EMPNBR,HE_FDATE,HE_TDATE,HE_TYPE,HE_ENTITLE,HE_TAKEN,HE_PREV FROM HRENTHRS "
SQLQ = SQLQ & "WHERE HE_EMPNBR = " & glbLEE_ID & " AND HE_TYPE = '" & savKey & "'"
'Hemu - Ticket # 9448 - Since the client can maintain history of the entitlement esp. for STD cases
'       we will need to get the latest date range to compare the hours against.
SQLQ = SQLQ & " ORDER BY HE_FDATE DESC,HE_TDATE DESC"
'Hemu
rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If rsTB.EOF Then
    HRENTITLE = 0
    rsTB.Close
    Exit Function
End If
xWrk = 0

'Hemu
rsTB.MoveFirst
'Hemu

If DateValue(dlpReviewDate.Text) >= rsTB("HE_FDATE") And DateValue(dlpReviewDate.Text) <= rsTB("HE_TDATE") Then
'    'Ticket #17924 - Begin
'    'Reset the original code back to get the Attendance records
    'Ticket #18629 the "savKey = xOrgSavKey" caused blank savKey for normal Attendance Code,
    '--such as "FD", so this logic should only be for "-" code
    If Right(savKey, 1) = "+" Or Mid(savKey, 3, 1) = "+" Then
        savKey = xOrgSavKey
    End If
'    'Ticket #17924 - End
    
    'Ticket #1859 - Multiple taken codes for one bank - sum all the taken codes
    If Mid(savKey, 3, 1) = "-" Then
        savKey = Left(savKey, 3) & "%"
    End If
    
    'Calculate most current Taken entitlement
    SQLQ = "SELECT SUM(AD_HRS) AS HRS_TAKEN FROM HR_ATTENDANCE"
    SQLQ = SQLQ & " WHERE HR_ATTENDANCE.AD_DOA>=" & Date_SQL(rsTB("HE_FDATE"))
    SQLQ = SQLQ & " AND (HR_ATTENDANCE.AD_DOA)<=" & Date_SQL(rsTB("HE_TDATE"))
    SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_REASON like '" & savKey & "'"  '"%'"
    SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_EMPNBR=" & glbLEE_ID
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsAttend.EOF Then
        rsTB("HE_TAKEN") = IIf(IsNull(rsAttend("HRS_TAKEN")), 0, rsAttend("HRS_TAKEN"))
        rsTB.Update
    End If
    rsAttend.Close
    
    If Len(SavEnt) = 0 Then SavEnt = 0
    If Not fglbNew Then
        'xWrk = (rsTB("HE_PREV") + rsTB("HE_ENTITLE")) - rsTB("HE_TAKEN") - (SavEnt - oldHrs)
        xWrk = (IIf(IsNull(rsTB("HE_PREV")), 0, rsTB("HE_PREV")) + IIf(IsNull(rsTB("HE_ENTITLE")), 0, rsTB("HE_ENTITLE"))) - IIf(IsNull(rsTB("HE_TAKEN")), 0, rsTB("HE_TAKEN")) - (SavEnt - oldHrs)
    Else
        'xWrk = (rsTB("HE_PREV") + rsTB("HE_ENTITLE")) - rsTB("HE_TAKEN") - SavEnt
        xWrk = (IIf(IsNull(rsTB("HE_PREV")), 0, rsTB("HE_PREV")) + IIf(IsNull(rsTB("HE_ENTITLE")), 0, rsTB("HE_ENTITLE"))) - IIf(IsNull(rsTB("HE_TAKEN")), 0, rsTB("HE_TAKEN")) - SavEnt
    End If
End If

HRENTITLE = xWrk
rsTB.Close
End Function

Private Sub lblEmpID_Change()
lblEENum = ShowEmpnbr(lblEEID)
Call Fd_CurrJob ' Ticket #14555
End Sub

Private Sub lblUpload_Change()
chkUpload.Value = IIf(lblUpload = "Y", 1, 0)
End Sub

Private Sub medHours_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medHours_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub

Private Sub medHours_LostFocus()
    'If glbCBrant Then
    If Not glbNiagaraFulls And Not (glbCompSerial = "S/N - 2241W") Then 'Ticket #15133
        Call ReCalcOT(ReaOld, clpCode(1).Text, HoursOld, medHours)
    End If
    'End If
End Sub

Private Sub medMachineHours_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub



Private Sub medMachineRate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub

Private Sub medSalary_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub

Private Sub memComments_GotFocus()

Call SetPanHelp(ActiveControl)
MDIMain.panHelp(1).Caption = "Unlimited"    'laura jan 05, 1998
MDIMain.panHelp(2).Caption = " "

End Sub

Private Function modSpecificID(EEID&)
Dim Msg$, Def$, Response$, Title$

modSpecificID = 0

Def$ = CStr(EEID&)
Msg$ = "If you desire a lookup for a specific Employee"
Msg$ = Msg$ & " enter their employee number."
Msg$ = Msg$ & "The default value below is the Employee "
Msg$ = Msg$ & "number of the last individual you reviewed."
Msg$ = Msg$ & Chr(10) & Chr(10)
Msg$ = Msg$ & "If you want to all attendance records "
Msg$ = Msg$ & "Enter 0. "
Title$ = "Specific Attendance Records?"
Response$ = InputBox$(Msg$, Title$, Def$)

If Len(Response$) > 0 Then
    If IsNumeric(Response$) Then
        modSpecificID = CLng(Response$)
    End If
End If

End Function

Private Sub modSTUPD(YN)
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
'cmdNew.Enabled = FT     'DWS if ft = true?
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT

'WDGPHU - Ticket #26576
'Leeds and Grenville - Ticket #19441
If glbCompSerial = "S/N - 2233W" Or (glbCompSerial = "S/N - 2411W" And gsFLEX_LOGIC) Then
    ChkInc.Enabled = False
Else
    ChkInc.Enabled = TF
End If
chkIncident.Enabled = TF
chkSeniority.Enabled = TF
chkBackDated.Enabled = TF
ChkFMLA.Enabled = TF            'Jaddy 6/4/99
ChkInc.Font3D = (TF + 1)        '
chkIncident.Font3D = (TF + 1)   '
chkSeniority.Font3D = (TF + 1)  '
chkEMELEA.Font3D = (TF + 1)
ChkFMLA.Font3D = (TF + 1)       '
txtWSIB.Enabled = TF            '
txtPoint.Enabled = TF
clpCode(3).Enabled = TF
clpCode(4).Enabled = TF
medMachineHours.Enabled = TF
medMachineRate.Enabled = TF
'If glbBurlTech Then
'    txtLEPoint.Enabled = TF
'End If
medHours.Enabled = TF
'memComments.Enabled = TF
memComments.Locked = FT
clpChrgCode.Enabled = TF
txtChrgCode.Enabled = TF
clpCode(1).Enabled = TF
dlpReviewDate.Enabled = TF
txtShift.Enabled = TF
elpSupShow.Enabled = TF
dlpToDate.Enabled = TF 'False
'vbxTrueGrid(0).Enabled = FT

If glbWHSCC Then
    'If change the date for ASL, it's too complicated to change the data in ASL table
    'so if the user wants to change the date for ASL, the only way is delete it and
    'then add a new one
    If AddChg = "C" And UCase(clpCode(1)) = "ASL" Then
        dlpReviewDate.Enabled = False
        clpCode(1).Enabled = False
    End If
End If

chkEMELEA.Enabled = TF And chkEmeryTabl <> 0

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    cmdAnother.Visible = False
Else
    cmdAnother.Enabled = True
End If

If glbCompSerial = "S/N - 2192W" Or glbCompSerial = "S/N - 2366W" Then 'county of essex
    frmMulti.Enabled = TF
End If
'If glbMulti Then
    frmMulti.Enabled = TF
    cmdPostion.Visible = TF
'End If

If glbCompSerial = "S/N - 2288W" Then 'Musashi - Ticket #12690
    'Check if the user has access to this employee's salary information
    If Allow_User_To_View("ACTIVE") = False Then
        medsalary.Visible = False
        lblTitle(10).Visible = False
        lblTitle(11).Visible = False
        comPayPer.Visible = False
    Else
        medsalary.Visible = gSec_Inq_Salary
        If medsalary.Visible = True Then medsalary.Enabled = gSec_Upd_Salary
        lblTitle(10).Visible = gSec_Inq_Salary
        lblTitle(11).Visible = True
        comPayPer.Visible = True
    End If
Else
    medsalary.Visible = gSec_Inq_Salary
    If medsalary.Visible = True Then medsalary.Enabled = gSec_Upd_Salary
    lblTitle(10).Visible = gSec_Inq_Salary
    'Ticket #13249 Frank 06/22/2007
    'make the change to '-NON' security to make sure the information is not displayed on the attendance screen
    'If glbWFC Then - Ticket #18435 Frank 04/30/2010, Samuel needs this too. So open it for all if they use "-EXE" or "-NON"
        If glbNoNONE Or glbNoEXEC Then
            If xORG = "NONE" Or xORG = "EXEC" Then
                medsalary.Visible = False
                lblTitle(10).Visible = False
            End If
        End If
    'End If
End If

glbDocName = "Attendance"
'7.9 Enhancement
If gsAttachment_DB Then 'And xAD = "AD" Then
    'rsDATA.Requery
    If Not (rsDATA.BOF And rsDATA.EOF) Then
        If rsDATA.RecordCount > 0 Then
            If Not IsNull(rsDATA("AD_DOCKEY")) Then
                glbDocKey = rsDATA("AD_DOCKEY")
            Else
                glbDocKey = 0
            End If
        Else
            If Not IsNull(Data1.Recordset("AD_DOCKEY")) Then
                glbDocKey = Data1.Recordset("AD_DOCKEY")
            Else
                glbDocKey = 0
            End If
        End If
        
        glbAttReason = clpCode(1).Text
        If IsDate(dlpReviewDate.Text) Then
            glbAttDOA = dlpReviewDate.Text
        End If
    End If
    
    Call DispimgIcon(Me, "frmVATTEND")
    If gSec_Upd_Attendance Then
        If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
            cmdImport.Visible = False
        Else
            cmdImport.Visible = True
        End If
    End If
End If

End Sub

Private Sub clpCode_LostFocus(Index As Integer)
    If Index = 1 Then
        If AddChg = "A" Then
            '2454W - Showa Ticket #25250 Franks 04/07/2014
            If glbCompSerial = "S/N - 2350W" Or glbCompSerial = "S/N - 2454W" Then
                chkIncident = fglbINC
            Else
                ChkInc = fglbINC
            End If
            chkSeniority = fglbSen
        End If
        If IsNumeric(fglbPoint) Then
            txtPoint = fglbPoint
        Else
            If glbBurlTech Then
            Else
            txtPoint = ""
            End If
        End If
        If glbLinamar Then
            If UCase(clpCode(1)) = "VAC" Then medHours = 8
        End If
        If fglbEMELEA <> 0 Then
            If Not chkEMELEA Then    'And cmdOK.Enabled Then 'check the condition
                Dim Msg, a%
                Msg = "Emergency Leave selected." & Chr(10)
                Msg = Msg & "Do the hours deduct from the Employee's Emergency Leave?"
                a% = MsgBox(Msg, 36, "Confirm")
                If a% = 6 Then
                    chkEMELEA = True
                End If
            End If
            chkEMELEA.Enabled = True
        Else
            If chkEMELEA Then 'And cmdOK.Enabled Then  'check the condition
                chkEMELEA = False
            End If
            chkEMELEA.Enabled = False
        End If
    End If
    If Not glbNiagaraFulls And glbCompSerial <> "S/N - 2241W" Then
        Call ReCalcOT(ReaOld, clpCode(1), HoursOld, medHours)
    End If
    'Casey House
    If glbCompSerial = "S/N - 2214W" Then
        If Index = 2 Then
            txtShift = clpCode(2)
        End If
    End If
End Sub


Private Sub clpJob_LostFocus()
Dim xNSalary, xNOrg, xNDHRS, xNWHRS
Dim TE As New ADODB.Recordset
Dim SQLQ

xNSalary = 0
xNOrg = ""
xNDHRS = 0
xNWHRS = 0
If AddChg = "C" Or AddChg = "A" Then
    SQLQ = "SELECT SH_SALCD,SH_SALARY FROM HR_SALARY_HISTORY WHERE SH_EMPNBR=" & lblEmpID & " AND SH_JOB='" & clpJob & "' AND SH_CURRENT<>0 "
    TE.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not TE.EOF Then
        If TE("SH_SALCD") = "A" Then
            comPayPer.ListIndex = 0
        ElseIf TE("SH_SALCD") = "H" Then
            comPayPer.ListIndex = 1
        ElseIf TE("SH_SALCD") = "D" And glbCompSerial = "S/N - 2282W" Then
            comPayPer.ListIndex = 3
        Else
            comPayPer.ListIndex = 2
        End If
        xNSalary = TE("SH_SALARY")
    End If
    TE.Close
    SQLQ = "SELECT JH_ORG,JH_DHRS,JH_WHRS,JH_PAYROLL_ID,JH_GLNO FROM HR_JOB_HISTORY WHERE JH_EMPNBR=" & lblEmpID & " AND JH_JOB='" & clpJob & "' AND JH_CURRENT<>0 "
    TE.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not TE.EOF Then
        xNOrg = IIf(IsNull(TE("JH_ORG")), xORG, TE("JH_ORG"))
        xNDHRS = TE("JH_DHRS")
        xNWHRS = TE("JH_WHRS")
        medsalary = xNSalary
        clpCode(0) = xNOrg
        txtDHRS = IIf(IsNull(xNDHRS), 0, xNDHRS)
        txtWHRS = IIf(IsNull(xNWHRS), 0, xNWHRS)
        clpGLNo = TE("JH_GLNO") & ""
        txtPayrollID = TE("JH_PAYROLL_ID") & ""
    End If
    TE.Close
End If
End Sub


Private Sub scrControl_Change()

fraDetail.Top = 240 + vbxTrueGrid(0).Height + panEEName.Height - scrControl.Value * ((panControls.Height + scrControl.Max) / scrControl.Max)
End Sub


Private Sub txtChrgCode_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtChrgCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub
Private Sub txtCode_Change(Index As Integer)
Dim rsDiscip As New ADODB.Recordset
Dim SQLQ
    lblCodeDesc(0).Caption = ""
    If Len(txtCode(0).Text) > 0 Then
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'CETY' AND TB_KEY = '" & txtCode(0).Text & "'"
        rsDiscip.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsDiscip.EOF Then
            lblCodeDesc(0).Caption = rsDiscip("TB_DESC")
            lblCodeDesc(0).Visible = True
        End If
        rsDiscip.Close
    End If
End Sub

Private Sub txtDHRS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub

Private Sub txtPayrollID_GotFocus()
 Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtPayrollID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub

Private Sub txtPoint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtPoint_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub

Private Sub txtShift_Change()
If Not glbCompSerial = "S/N - 2226W" Then Exit Sub
    If Len(txtShift.Text) > 0 Then
        If txtShift.Text = "FD" Then
            comShiftType.ListIndex = 0
        ElseIf txtShift.Text = "AM" Then
            comShiftType.ListIndex = 1
        Else
            comShiftType.ListIndex = 2
        End If
    Else
        comShiftType = ""
    End If
End Sub

Private Sub txtShift_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub
Private Sub txtSUP_Change()
elpSupShow = ShowEmpnbr(txtSup.Text)
End Sub


Private Sub txtWHRS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub

Private Sub txtWSIB_DblClick()
Dim lastWCB

lastWCB = txtWSIB

If Len(glbWSIB) > 0 Then
    txtWSIB = glbWSIB
Else
    txtWSIB = lastWCB
End If

End Sub

Private Sub txtWSIB_GotFocus()
    Call SetPanHelp(ActiveControl)      'Jaddy 6/4/99
End Sub

Private Sub txtShift_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub UPDOVERTIME()
Dim rsOvtEmp As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim SQLQ As String

'Recalculate the Overtime Bank
SQLQ = "SELECT * FROM HR_OVERTIME_BANK WHERE OT_EMPNBR = " & glbLEE_ID
rsOvtEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not rsOvtEmp.EOF Then
    Call ReCalcOvt("OT_EMPNBR = " & glbLEE_ID)
Else
'Do not create the overtime records for employees if the user has not run the Update entitlement.
'    SQLQ = "SELECT ED_EMPNBR, ED_ORG, ED_OTBANK, ED_PT, ED_EMP FROM HREMP "
'    SQLQ = SQLQ & " WHERE ED_EMPNBR = " & glbLEE_ID
'    SQLQ = SQLQ & " AND ED_ORG IN (SELECT OM_ORG FROM HR_OVERTIME_MASTER)"
'    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'
'    If Not rsEMP.EOF Then
'        'Add record Overtime Bank records
'        rsOvtEmp.AddNew
'        rsOvtEmp("OT_COMPNO") = "001"
'        rsOvtEmp("OT_EMPNBR") = glbLEE_ID
'        rsOvtEmp("OT_PBANK") = 0
'        rsOvtEmp("OT_BANK") = Get_OvertimeBank(glbLEE_ID) * Overtime_Multiplier(rsEMP("ED_ORG"), rsEMP("ED_PT"), rsEMP("ED_EMP"))
'        rsOvtEmp("OT_BANKT") = Get_OvertimeTaken(glbLEE_ID)
'        rsOvtEmp("OT_EFDATE") = Format("1/1/" & Year(Now()), "mm/dd/yyyy")
'        rsOvtEmp("OT_ETDATE") = Format("12/31/" & Year(Now()), "mm/dd/yyyy")
'        rsOvtEmp("OT_LDATE") = Date
'        rsOvtEmp("OT_LTIME") = Time$
'        rsOvtEmp("OT_LUSER") = glbUserID
'        rsOvtEmp.Update
'
'        rsEMP("ED_OTBANK") = rsOvtEmp("OT_BANK") - rsOvtEmp("OT_BANKT")
'        rsEMP.Update
'    End If
'    rsEMP.Close
End If
rsOvtEmp.Close

'Refresh the Overtime Bank value on the Attendance screen value
'Ticket #23655 - For City of Niagara Falls - ED_OTBANK is the total Banked Time.
If glbCompSerial = "S/N - 2276W" Then
    Dim rsHREmp As New ADODB.Recordset
    SQLQ = "SELECT ED_OTBANK FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHREmp.EOF Then
        Call DisOvertime(IIf(IsNull(rsHREmp("ED_OTBANK")), 0, rsHREmp("ED_OTBANK")))
    End If
    rsHREmp.Close
    Set rsHREmp = Nothing
End If

End Sub

Private Sub UPDHRENTIT(savKey, SavEnt)
Dim rsTB As New ADODB.Recordset
Dim rsTBTemp As New ADODB.Recordset 'Hemu
Dim SQLQ, SQLT  'Hemu

If Len(SavEnt) = 0 Then Exit Sub
If Len(savKey) = 0 Then Exit Sub

''Ticket #17924 - Begin
Dim xOrgSavKey As String
Dim xOrgSavReas As String

'Retain original values
xOrgSavKey = savKey
xOrgSavReas = savReas

'If Hourly Earned then HE_ENTITLE needs to be updated
If Right(savKey, 1) = "+" Then GoTo HourlyEarned

'If Hourly Taken then HE_TAKEN needs to be updated
If Right(savKey, 1) = "-" Or Mid(savKey, 3, 1) = "-" Then
    'Change "-" code to "+" so that the Hourly Entitlement record is found
    'ABC- to ABC+
    'Ticket #1859 - Or the 3rd character is "-" then it is hourly entitlement additional Flex logic
    'Replace with "+" so the Hourly Entitlement record is retrieved which is suffixed with "+ "
    If Mid(savKey, 3, 1) = "-" Then
        'Ticket #1859
        'This is multiple taken codes for one bank - only need the first two chars and then concatenate with +
        savKey = Left(savKey, 2) & "+"
    Else
        savKey = Left(savKey, Len(savKey) - 1) & "+"
    End If
End If
''Ticket #17924 - End

If glbtermopen Then
    SQLQ = "SELECT HE_TAKEN,HE_ENTITLE,HE_ID FROM Term_ENTHRS "
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "SELECT HE_TAKEN,HE_ENTITLE,HE_ID FROM HRENTHRS "
    SQLQ = SQLQ & " WHERE HE_EMPNBR = " & glbLEE_ID
End If
SQLT = SQLQ 'Hemu
SQLQ = SQLQ & " AND HE_TYPE = '" & savKey & "'"
SQLQ = SQLQ & " AND HE_FDATE<= " & Date_SQL(dlpReviewDate.Text)
SQLQ = SQLQ & " AND HE_TDATE>= " & Date_SQL(dlpReviewDate.Text)

If glbtermopen Then
    rsTB.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

''Ticket #17924 - Begin
If Right(savReas, 1) = "-" Or Mid(savReas, 3, 1) = "-" Then
    'Change "-" code to "+" so that the Hourly Entitlement record is found
    'ABC- to ABC+
    If Mid(savReas, 3, 1) = "-" Then
        'Ticket #1859 - 'AB-1 to AB-
        'This is multiple taken codes for one bank - only need the first two chars and then concatenate with +
        savReas = Left(savReas, 2) & "+"
    Else
        savReas = Left(savReas, Len(savReas) - 1) & "+"
    End If
End If
''Ticket #17924 - End

If rsTB.EOF Then
    SQLT = SQLT & " AND HE_TYPE = '" & savReas & "'"
    SQLT = SQLT & " AND HE_FDATE<= " & Date_SQL(savEntDate)
    SQLT = SQLT & " AND HE_TDATE>= " & Date_SQL(savEntDate)
    If glbtermopen Then
        rsTBTemp.Open SQLT, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        rsTBTemp.Open SQLT, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    If Not rsTBTemp.EOF Then
        If AddChg = "C" Then
            ''Ticket #17924 - Begin
            If Right(xOrgSavReas, 1) <> "+" And Mid(xOrgSavReas, 3, 1) <> "+" Then
                rsTBTemp("HE_TAKEN") = IIf(IsNull(rsTBTemp("HE_TAKEN")), 0, rsTBTemp("HE_TAKEN")) - savEnt1
            Else
                rsTBTemp("HE_ENTITLE") = IIf(IsNull(rsTBTemp("HE_ENTITLE")), 0, rsTBTemp("HE_ENTITLE")) - savEnt1
            End If
            ''Ticket #17924 - End
            rsTBTemp.Update
        End If
    End If
    rsTBTemp.Close
    
    Exit Sub
Else
    If AddChg = "C" Then
        SQLT = SQLT & " AND HE_TYPE = '" & savReas & "'"
        SQLT = SQLT & " AND HE_FDATE<= " & Date_SQL(savEntDate)
        SQLT = SQLT & " AND HE_TDATE>= " & Date_SQL(savEntDate)
        If glbtermopen Then
            rsTBTemp.Open SQLT, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsTBTemp.Open SQLT, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        If Not rsTBTemp.EOF Then
            ''Ticket #17924 - Begin
            If Right(xOrgSavReas, 1) <> "+" And Mid(xOrgSavReas, 3, 1) <> "+" Then
                rsTBTemp("HE_TAKEN") = IIf(IsNull(rsTBTemp("HE_TAKEN")), 0, rsTBTemp("HE_TAKEN")) - savEnt1
            Else
                rsTBTemp("HE_ENTITLE") = IIf(IsNull(rsTBTemp("HE_ENTITLE")), 0, rsTBTemp("HE_ENTITLE")) - savEnt1
            End If
            ''Ticket #17924 - End
            rsTBTemp.Update
        End If
        rsTBTemp.Close
    End If
End If

rsTB.Requery

If Not rsTB.EOF Then
    rsTB.MoveFirst
End If

Do Until rsTB.EOF
    'Hemu - Begin
    If AddChg = "C" Then
        SQLT = SQLT & " AND HE_FDATE<= " & Date_SQL(savEntDate)
        SQLT = SQLT & " AND HE_TDATE>= " & Date_SQL(savEntDate)
        If glbtermopen Then
            rsTBTemp.Open SQLT, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsTBTemp.Open SQLT, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        If rsTBTemp.EOF Then
            savEnt1 = 0
        End If
        rsTBTemp.Close
    End If
    'Hemu - End
    
    If AddChg = "C" Or AddChg = "D" Then
        If AddChg = "C" Then 'And savReas <> savKey Then
            rsTB("HE_TAKEN") = IIf(IsNull(rsTB("HE_TAKEN")), 0, rsTB("HE_TAKEN")) + SavEnt
        Else
            ''Ticket #17924 - Begin
            If Right(xOrgSavReas, 1) <> "+" And Mid(xOrgSavReas, 3, 1) <> "+" Then
                rsTB("HE_TAKEN") = IIf(IsNull(rsTB("HE_TAKEN")), 0, rsTB("HE_TAKEN")) + SavEnt - savEnt1
            Else
                rsTB("HE_TAKEN") = IIf(IsNull(rsTB("HE_TAKEN")), 0, rsTB("HE_TAKEN")) + SavEnt
            End If
            ''Ticket #17924 - Endi
        End If
    Else
        rsTB("HE_TAKEN") = IIf(IsNull(rsTB("HE_TAKEN")), 0, rsTB("HE_TAKEN")) + SavEnt
    End If
    rsTB.Update
    rsTB.MoveNext
    glbENTScreen = True
Loop
rsTB.Close
Exit Sub

''Ticket #17924 - Begin
HourlyEarned:
If glbtermopen Then
    SQLQ = "SELECT HE_ENTITLE,HE_TAKEN,HE_ID FROM Term_ENTHRS "
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "SELECT HE_ENTITLE,HE_TAKEN,HE_ID FROM HRENTHRS "
    SQLQ = SQLQ & " WHERE HE_EMPNBR = " & glbLEE_ID
End If
SQLT = SQLQ 'Hemu
SQLQ = SQLQ & " AND HE_TYPE = '" & savKey & "'"
SQLQ = SQLQ & " AND HE_FDATE<= " & Date_SQL(dlpReviewDate.Text)
SQLQ = SQLQ & " AND HE_TDATE>= " & Date_SQL(dlpReviewDate.Text)

If glbtermopen Then
    rsTB.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

If Right(savReas, 1) = "-" Or Mid(savReas, 3, 1) = "-" Then
    'Change "-" code to "+" so that the Hourly Entitlement record is found
    'ABC- to ABC+
    'Ticket #1859 - Or the 3rd character is "-" then it is hourly entitlement additional Flex logic
    'Replace with "+" so the Hourly Entitlement record is retrieved which is suffixed with "+ "
    If Mid(savReas, 3, 1) = "-" Then
        'Ticket #1859
        'This is multiple taken codes for one bank - only need the first two chars and then concatenate with +
        savReas = Left(savReas, 2) & "+"
    Else
        savReas = Left(savReas, Len(savReas) - 1) & "+"
    End If
End If

If rsTB.EOF Then
    SQLT = SQLT & " AND HE_TYPE = '" & savReas & "'"
    SQLT = SQLT & " AND HE_FDATE<= " & Date_SQL(savEntDate)
    SQLT = SQLT & " AND HE_TDATE>= " & Date_SQL(savEntDate)
    If glbtermopen Then
        rsTBTemp.Open SQLT, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        rsTBTemp.Open SQLT, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    If Not rsTBTemp.EOF Then
        If AddChg = "C" Then
            If Right(xOrgSavReas, 1) <> "+" And Mid(xOrgSavReas, 3, 1) <> "+" Then
                rsTBTemp("HE_TAKEN") = IIf(IsNull(rsTBTemp("HE_TAKEN")), 0, rsTBTemp("HE_TAKEN")) - savEnt1
            Else
                rsTBTemp("HE_ENTITLE") = IIf(IsNull(rsTBTemp("HE_ENTITLE")), 0, rsTBTemp("HE_ENTITLE")) - savEnt1
            End If

            rsTBTemp.Update
        End If
    End If
    rsTBTemp.Close

    Exit Sub
Else
    If AddChg = "C" Then
        SQLT = SQLT & " AND HE_TYPE = '" & savReas & "'"
        SQLT = SQLT & " AND HE_FDATE<= " & Date_SQL(savEntDate)
        SQLT = SQLT & " AND HE_TDATE>= " & Date_SQL(savEntDate)
        If glbtermopen Then
            rsTBTemp.Open SQLT, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsTBTemp.Open SQLT, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        If Not rsTBTemp.EOF Then
            If Right(xOrgSavReas, 1) <> "+" And Mid(xOrgSavReas, 3, 1) <> "+" Then
                rsTBTemp("HE_TAKEN") = IIf(IsNull(rsTBTemp("HE_TAKEN")), 0, rsTBTemp("HE_TAKEN")) - savEnt1
            Else
                rsTBTemp("HE_ENTITLE") = IIf(IsNull(rsTBTemp("HE_ENTITLE")), 0, rsTBTemp("HE_ENTITLE")) - savEnt1
            End If
            rsTBTemp.Update
        End If
        rsTBTemp.Close
    End If
End If

rsTB.Requery

If Not rsTB.EOF Then
    rsTB.MoveFirst
End If

Do Until rsTB.EOF
    'Hemu - Begin
    If AddChg = "C" Then
        SQLT = SQLT & " AND HE_FDATE<= " & Date_SQL(savEntDate)
        SQLT = SQLT & " AND HE_TDATE>= " & Date_SQL(savEntDate)
        If glbtermopen Then
            rsTBTemp.Open SQLT, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsTBTemp.Open SQLT, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        If rsTBTemp.EOF Then
            savEnt1 = 0
        End If
        rsTBTemp.Close
    End If
    'Hemu - End

    If AddChg = "C" Or AddChg = "D" Then
        If AddChg = "C" Then 'And savReas <> savKey Then
            rsTB("HE_ENTITLE") = rsTB("HE_ENTITLE") + SavEnt
        Else
            If Right(xOrgSavReas, 1) <> "+" And Mid(xOrgSavReas, 3, 1) <> "+" Then
                rsTB("HE_ENTITLE") = rsTB("HE_ENTITLE") + SavEnt - savEnt1
            Else
                rsTB("HE_ENTITLE") = rsTB("HE_ENTITLE") + SavEnt - savEnt1
            End If
        End If
    Else
        rsTB("HE_ENTITLE") = rsTB("HE_ENTITLE") + SavEnt
    End If
    rsTB.Update
    rsTB.MoveNext
    glbENTScreen = True
Loop
rsTB.Close
''Ticket #17924 - End

End Sub

Private Sub UPDVACSICK(xmthr)
Dim rsTB As New ADODB.Recordset
Dim rsOvtBank As New ADODB.Recordset
Dim SQLQ
Dim IfChange As Boolean
Dim xOTBANK, xTemp, xTemp01, xNum
Dim xVCOBankFlag As Boolean
IfChange = False
'If xmthr = True And SavVac = 0 And SavSick = 0 Then Exit Sub

SQLQ = "SELECT ED_VACT,ED_SICKT,ED_INCIDCNT,"
SQLQ = SQLQ & "ED_EFDATE,ED_ETDATE,ED_EFDATES,ED_ETDATES,"
If glbNiagaraFulls And (Not glbtermopen) Then
    SQLQ = SQLQ & "ED_VAC,ED_PVAC,ED_SICK,ED_PSICK,ED_OTBANK "
Else
    SQLQ = SQLQ & "ED_VAC,ED_PVAC,ED_SICK,ED_PSICK "
End If
If glbCompSerial = "S/N - 2380W" Then 'VitalAire
    SQLQ = SQLQ & ",ED_ANNVAC,ED_OTBANK"
End If
If glbtermopen Then
    SQLQ = SQLQ & ",TERM_SEQ "
    SQLQ = SQLQ & " FROM Term_HREMP WHERE TERM_SEQ=" & glbTERM_Seq
    rsTB.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = SQLQ & ",ED_EMPNBR "
    SQLQ = SQLQ & " FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
    rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

If rsTB.EOF Then
    Exit Sub
End If

' For City of Niagara Falls '
If glbNiagaraFulls And (Not glbtermopen) Then
    Call DisOvertime(rsTB("ED_OTBANK"))
End If
If glbNiagaraFulls And xmthr = True And (Not glbtermopen) Then
    If UCase(clpCode(1).Text) = "ASIC" Then
        If AddChg = "A" Then
            rsTB("ED_SICK") = rsTB("ED_SICK") + Val(medHours)
            IfChange = True
        End If
        If AddChg = "C" Then
            rsTB("ED_SICK") = rsTB("ED_SICK") + (Val(medHours) - savEnt1)
            IfChange = True
        End If
        If AddChg = "D" Then
            rsTB("ED_SICK") = rsTB("ED_SICK") - Val(medHours)
            IfChange = True
        End If
    End If
    ChgOldNew = 0
    xDispODay = 0
    'Hemu - 06/07/2004 Begin - Ticket # 6292
    'If UCase(clpCode(1).Text) = "OT" Or UCase(clpCode(1).Text) = "OT10" Or UCase(clpCode(1).Text) = "OT15" Or UCase(clpCode(1).Text) = "OT20" Or (ValReasonCha(xOldReason, xNewReason) <> 0 And AddChg <> "D") Then
    If UCase(clpCode(1).Text) = "OT" Or UCase(clpCode(1).Text) = "OT10" Or UCase(clpCode(1).Text) = "OT15" Or UCase(clpCode(1).Text) = "OT20" Then
    'Hemu - 06/07/2004 End
        'SaveHours
        If AddChg = "A" Then
            xOTBANK = xmedHours 'Val(medHours)
            If xAnother = 2 Then
                If UCase(clpCode(1).Text) = "OT15" Then xOTBANK = xmedHours
                If UCase(clpCode(1).Text) = "OT20" Then xOTBANK = xmedHours
            Else
                If UCase(clpCode(1).Text) = "OT15" Then xOTBANK = xmedHours * 1.5  'Val(medHours) * 1.5
                If UCase(clpCode(1).Text) = "OT20" Then xOTBANK = xmedHours * 2
            End If
            
            'Comment out the custom logic - Ticket #14412
            'If IIf(IsNull(rsTB("ED_OTBANK")), 0, rsTB("ED_OTBANK")) + xOTBANK > SaveHours * 10 Then
            '    xNum = xOTBANK
            '    xOTBANK = SaveHours * 10
            '    If SaveHours = 0 Then
            '        MsgBox "Hours/Day for this employee is zero for their current position "
            '    Else
            '        MsgBox "Overtime Bank maximum has been reach. Maximum Overtime bank is 10 days"
            '    End If

            '    medHours = (SaveHours * 10) - IIf(IsNull(rsTB("ED_OTBANK")), 0, rsTB("ED_OTBANK"))
            'Else
                'xOTBANK = rsTB("ED_OTBANK") + xOTBANK
                xOTBANK = IIf(IsNull(rsTB("ED_OTBANK")), 0, rsTB("ED_OTBANK")) + xOTBANK
                Select Case UCase(clpCode(1).Text)
                    Case "OT15"
                        If xAnother = 2 Then
                            medHours = xmedHours * 1.5  'Ticket #28207
                        Else
                            medHours = xmedHours * 1.5 'Val(medHours) * 1.5
                        End If
                    Case "OT20"
                        If xAnother = 2 Then
                            medHours = xmedHours * 2    'Ticket #28207
                        Else
                            medHours = xmedHours * 2  'Val(medHours) * 2
                        End If
                End Select
            'End If
            
            rsTB("ED_OTBANK") = xOTBANK 'rsTB("ED_OTBANK") + xOTBANK
            If rsTB("ED_OTBANK") < 0 Then rsTB("ED_OTBANK") = 0
 
            Call DisOvertime(rsTB("ED_OTBANK"))

            IfChange = True
        End If
        
        If AddChg = "NO" Then
            xOTBANK = xmedHours
            If UCase(clpCode(1).Text) = "OT15" Then xOTBANK = xmedHours * 1.5  'Val(medHours) * 1.5
            If UCase(clpCode(1).Text) = "OT20" Then xOTBANK = xmedHours * 2

            'Comment out the custom logic - Ticket #14412
            'If IIf(IsNull(rsTB("ED_OTBANK")), 0, rsTB("ED_OTBANK")) + xOTBANK > SaveHours * 10 Then
            '    xOTBANK = SaveHours * 10
            '    If SaveHours = 0 Then
            '        MsgBox "Hours/Day for this employee is zero for their current position "
            '    Else
            '        MsgBox "Overtime Bank maximum has been reach. Maximum Overtime bank is 10 days"
            '    End If
                
            '    medHours = (SaveHours * 10) - IIf(IsNull(rsTB("ED_OTBANK")), 0, rsTB("ED_OTBANK"))
            'Else
                xOTBANK = IIf(IsNull(rsTB("ED_OTBANK")), 0, rsTB("ED_OTBANK")) + xOTBANK
                Select Case UCase(clpCode(1).Text)
                    Case "OT15"
                        medHours = xmedHours * 1.5 'Val(medHours) * 1.5
                    Case "OT20"
                        If xAnother = 2 Then
                            medHours = xmedHours
                        Else
                            medHours = xmedHours * 2  'Val(medHours) * 2
                        End If
                End Select
            'End If
            
            rsTB("ED_OTBANK") = xOTBANK 'rsTB("ED_OTBANK") + xOTBANK
            If rsTB("ED_OTBANK") < 0 Then rsTB("ED_OTBANK") = 0
            Call DisOvertime(rsTB("ED_OTBANK"))
            IfChange = True
        End If
        
        If AddChg = "C" Then
            xTemp = IIf(IsNull(rsTB("ED_OTBANK")), 0, rsTB("ED_OTBANK")) + (Val(medHours) - savEnt1)
            xTemp01 = Val(medHours) 'HourGlb 'Val(medHours)
            xNum = IIf(IsNull(rsTB("ED_OTBANK")), 0, rsTB("ED_OTBANK")) - HourGlb
            rsTB("ED_OTBANK") = IIf(IsNull(rsTB("ED_OTBANK")), 0, rsTB("ED_OTBANK")) + (HourGlb - savEnt1) + ValReasonCha(xOldReason, xNewReason)

            If rsTB("ED_OTBANK") < 0 Then rsTB("ED_OTBANK") = 0
            
            'Comment out the custom logic - Ticket #14412
            'If rsTB("ED_OTBANK") > SaveHours * 10 Then
            '    If SaveHours = 0 Then
            '        MsgBox "Hours/Day for this employee is zero for their current position "
            '    Else
            '        MsgBox "Overtime Bank maximum has been reach. Maximum Overtime bank is 10 days"
            '    End If
            '    rsTB("ED_OTBANK") = SaveHours * 10
            '    medHours = (SaveHours * 10) - (xNum)
            'Else
                If ChgOldNew = 1 Then
                    'medHours = xTemp + xDispODay
                    medHours = xTemp01 + xDispODay01
                End If
            'End If
            Call DisOvertime(rsTB("ED_OTBANK"))
            IfChange = True
        End If
        If AddChg = "D" Then
            rsTB("ED_OTBANK") = IIf(IsNull(rsTB("ED_OTBANK")), 0, rsTB("ED_OTBANK")) - Val(medHours)
            If rsTB("ED_OTBANK") < 0 Then rsTB("ED_OTBANK") = 0
            Call DisOvertime(rsTB("ED_OTBANK"))
            IfChange = True
        End If
    End If
End If

' dkostka - 12/20/2001 - Pulled incident count update code out of this sub, it was causing
'   far too many problems here, moved to its own sub, UpdateIncCount.
xVCOBankFlag = True
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #14635
    If Not SavVCOBank = 0 Then
        xVCOBankFlag = False
    End If
End If

If xmthr = True And SavVac = 0 And SavSick = 0 And xVCOBankFlag Then
    If IfChange Then
        rsTB.Update
    End If
    rsTB.Close
    Exit Sub
End If

If xmthr = True Then

    rsTB("ED_VACT") = rsTB("ED_VACT") + SavVac
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #14635
        If IsNull(rsTB("ED_OTBANK")) Then
            rsTB("ED_OTBANK") = SavVCOBank
        Else
            rsTB("ED_OTBANK") = rsTB("ED_OTBANK") + SavVCOBank
        End If
    End If
    rsTB("ED_SICKT") = rsTB("ED_SICKT") + SavSick
    ' dkostka - 12/20/2001 - Moved incident counting to seperate sub.
    'rsTB("ED_INCIDCNT") = cntSick
    rsTB.Update
    'Franks Nov 25,02 - Displaying Vacation and Sick Outstanding
    If SaveHours > 0 Then
        'wellington duffrine ticket ##17736
        'S.U.C.C.E.S.S - Ticket #19099
        If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
            lblVACOS = Format(((rsTB("ED_VAC") + rsTB("ED_PVAC") - rsTB("ED_VACT"))), "Fixed")
            lblSICKOS = Format(((rsTB("ED_SICK") + rsTB("ED_PSICK") - rsTB("ED_SICKT"))), "Fixed")
        Else
        
            lblVACOS = Format(((rsTB("ED_VAC") + rsTB("ED_PVAC") - rsTB("ED_VACT")) / SaveHours), "Fixed")
            lblSICKOS = Format(((rsTB("ED_SICK") + rsTB("ED_PSICK") - rsTB("ED_SICKT")) / SaveHours), "Fixed")
        End If
    Else
        lblVACOS = Format(0, "Fixed")
        lblSICKOS = Format(0, "Fixed")
    End If
    rsTB.Close
    glbENTScreen = True
Else
    Fdate = rsTB("ED_EFDATE")
    Tdate = rsTB("ED_ETDATE")
    fdateS = rsTB("ED_EFDATES")
    tdateS = rsTB("ED_ETDATES")
    ' dkostka - 12/20/2001 - Moved incident counting to seperate sub.
    'cntSick = rsTB("ED_INCIDCNT")
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire
        SavOutV = rsTB("ED_ANNVAC") - rsTB("ED_VACT")
    Else
        SavOutV = rsTB("ED_VAC") + rsTB("ED_PVAC") - rsTB("ED_VACT")
    End If
    SavOutS = rsTB("ED_SICK") + rsTB("ED_PSICK") - rsTB("ED_SICKT")
    'Franks Nov 25,02 - Displaying Vacation and Sick Outstanding
    If SaveHours > 0 Then
        'wellington duffrine ticket ##17736
        'S.U.C.C.E.S.S - Ticket #19099
        If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
            lblVACOS = Format((SavOutV), "Fixed")
            lblSICKOS = Format((SavOutS), "Fixed")
        Else
            lblVACOS = Format((SavOutV / SaveHours), "Fixed")
            lblSICKOS = Format((SavOutS / SaveHours), "Fixed")
        End If
    Else
        lblVACOS = Format(0, "Fixed")
        lblSICKOS = Format(0, "Fixed")
    End If
    
    'Town of Aurora
    'If glbCompSerial = "S/N - 2378W" Then
        Dim rsOTBank As New ADODB.Recordset
        Dim rsOTMst As New ADODB.Recordset
        
        SQLQ = "SELECT OT_EMPNBR, OT_BANKT, OT_EFDATE, OT_ETDATE,"
        SQLQ = SQLQ & " OT_BANK,OT_PBANK, OT_MBANK "
        SQLQ = SQLQ & " FROM HR_OVERTIME_BANK WHERE OT_EMPNBR=" & glbLEE_ID
        rsOTBank.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
        If Not rsOTBank.EOF Then
            OTFdate = rsOTBank("OT_EFDATE")
            OTTdate = rsOTBank("OT_ETDATE")
            SavOTBank = rsOTBank("OT_BANK")
            SavOutOT = rsOTBank("OT_BANK") + rsOTBank("OT_PBANK") - rsOTBank("OT_BANKT")
            SavMaxBank = rsOTBank("OT_MBANK")
        End If
        rsOTBank.Close
        
        SQLQ = "SELECT OM_ORG,OM_MAX_BANK_HRS,OM_EMAIL "
        SQLQ = SQLQ & " FROM HR_OVERTIME_MASTER WHERE OM_ORG = "
        SQLQ = SQLQ & " (SELECT ED_ORG FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & ")"
        rsOTMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsOTMst.EOF Then
            'SavMaxBank = rsOTMst("OM_MAX_BANK_HRS")
            If Not IsNull(rsOTMst("OM_EMAIL")) Then
                SaveOTEmail = rsOTMst("OM_EMAIL")
            End If
        Else
            SaveOTEmail = ""
        End If
        rsOTMst.Close
    'End If
End If

'SavVac = 0
'SavSick = 0
lblIncident.Caption = "Total # of Incidents = " & Str(cntSick)

End Sub
'Private Sub vbxTrueGrid_GotFocus(Index As Integer)
'    Call SetPanHelp(ActiveControl)
'End Sub

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


Private Sub medSalary_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medSalary_LostFocus()
If Len(Trim(medsalary)) = 0 Then medsalary = 0
End Sub

Private Sub comPayPer_LostFocus()
If comPayPer.ListIndex = 0 Then lblSalCode.Caption = "A"
If comPayPer.ListIndex = 1 Then lblSalCode.Caption = "H"
If comPayPer.ListIndex = 2 Then lblSalCode.Caption = "M"
If comPayPer.ListIndex = 3 Then lblSalCode.Caption = "D"
End Sub


Private Sub comPayPer_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub
Private Sub lblSalCode_Change()
'commented by Bryan Ticket#11702, this stuff was un-serial controlled
'If glbMulti Or glbCompSerial = "S/N - 2192W" Or glbCompSerial = "S/N - 2366W" Or glbCompSerial = "S/N - 2363W" Then
    If Len(lblSalCode) > 0 Then
        If lblSalCode = "A" Then
            comPayPer.ListIndex = 0
        ElseIf lblSalCode = "H" Then
            comPayPer.ListIndex = 1
        ElseIf lblSalCode = "M" Then
            comPayPer.ListIndex = 2
        ElseIf lblSalCode = "D" Then
            comPayPer.ListIndex = 3
        End If
    Else
        comPayPer = ""
    End If
'End If
End Sub
Private Sub txtDHRS_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtDHRS_LostFocus()
If Len(Trim(txtDHRS)) = 0 Then txtDHRS = 0
End Sub
Private Sub txtWHRS_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtWHRS_LostFocus()
If Len(Trim(txtWHRS)) = 0 Then txtWHRS = 0
End Sub

Private Sub Job_Desc()
Dim SQLQ As String
'''On Error GoTo Pos_Err
Dim dynaJobs As New ADODB.Recordset
 clpJob = ""
 clpJob.ShowDescription = False
If Len(clpJob.Text) > 0 Then
     clpJob.Caption = "Unassigned"
     clpJob.ShowDescription = True
    dynaJobs.Open "HRJOB", gdbAdoIhr001, adOpenDynamic
    If dynaJobs.EOF And dynaJobs.BOF Then Exit Sub
    SQLQ = "JB_CODE = '" & clpJob.Text & "'"
    dynaJobs.Find SQLQ
    If Not dynaJobs.EOF Then clpJob.Caption = dynaJobs("JB_DESCR")
End If

Exit Sub

Pos_Err:
If Err = 94 Then
    Err = 0
    Resume Next
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Snap", "JOBS", "SELECT")
Call RollBack '28July99 js

End Sub

Private Sub DisOvertime(xNum) '''
Dim xTemp '

    If SaveHours > 0 Then
        'wellington duffrine ticket ##17736
        'S.U.C.C.E.S.S - Ticket #19099
        If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
             xTemp = Format((xNum), "Fixed")
        Else
             xTemp = Format((xNum / SaveHours), "Fixed")
        End If
    Else
        xTemp = 0
    End If
    If xTemp > 1 Then
        'wellington duffrine ticket ##17736
        'S.U.C.C.E.S.S - Ticket #19099
        If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
            lblOvertimeDays.Caption = "Hours"
        Else
            lblOvertimeDays.Caption = "Days"
        End If
    Else
        'S.U.C.C.E.S.S - Ticket #19099
        If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
            lblOvertimeDays.Caption = "Hour"
        Else
            lblOvertimeDays.Caption = "Day"
        End If
    End If
    If xTemp <> 0 Then
        txtOvertime = xTemp
    Else
        txtOvertime = ""
    End If
    
End Sub

Private Sub Calculate_EML_Taken()
    Dim rsATT As New ADODB.Recordset
    Dim SQLQ As String
    Dim toteml As Double ' This will be in Days
    Dim htoteml As Double 'this will be the stored hours
    Dim xOuts As Double 'Outstanding EML
    Dim rsTB As New ADODB.Recordset
    Dim xATTDHRS As Double
    'Hemu - EML

        If glbtermopen Then
            SQLQ = "SELECT AD_HRS AS EMLTAKEN, AD_EMPNBR,AD_DHRS FROM Term_ATTENDANCE "
            SQLQ = SQLQ & " WHERE AD_EMELEA <> 0 AND AD_EMPNBR = " & glbTERM_Seq
            If glbOracle Then
                SQLQ = SQLQ & " AND TO_CHAR(AD_DOA,'YYYY')  = " & Year(Date)
            Else
                SQLQ = SQLQ & " AND YEAR(AD_DOA) = " & Year(Date)
            End If
            rsATT.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            SQLQ = "SELECT AD_HRS AS EMLTAKEN, AD_EMPNBR,AD_DHRS FROM HR_ATTENDANCE "
            SQLQ = SQLQ & " WHERE AD_EMELEA <> 0 AND AD_EMPNBR = " & glbLEE_ID
            If glbOracle Then
                SQLQ = SQLQ & " AND TO_CHAR(AD_DOA,'YYYY') = " & Year(Date)
            Else
                SQLQ = SQLQ & " AND YEAR(AD_DOA) = " & Year(Date)
            End If

            rsATT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        If Not rsATT.EOF Then
            toteml = 0
            rsATT.MoveFirst
            Do While Not rsATT.EOF
                If SaveHours > 0 Then 'Logic taken from reports - that's how they want it to work
                    If glbLinamar Or glbCompSerial = "S/N - 2288W" Then 'Ticket #13177 Frank 06/07/2007, one record is one EML day
                        If Not IsNull(rsATT("EMLTAKEN")) Then           'Musashi - - Ticket #16786
                            If rsATT("EMLTAKEN") > 0 Then
                                toteml = toteml + 1
                            End If
                        End If
                    Else
                        If Not IsNull(rsATT("EMLTAKEN")) Then
                            'Ticket #29593 - Listowel Technology Inc. - #2350
                            'Ticket #29595 - City of Kenora - #2487
                            'Ticket #26021 - Putting it back to round to 1day for City of Woodstock, Ticket #28102 - KTH Shelburne
                            'Release 8.0 - Ticket #24545: Jerry wants this to be same as the EML Report, i.e.
                            'to show actual # of days taken and outstanding instead of rounding up to 1.
                            If glbCompSerial = "S/N - 2282W" Or glbCompSerial = "S/N - 2183W" Or glbCompSerial = "S/N - 2393W" Or glbCompSerial = "S/N - 2350W" Or glbCompSerial = "S/N - 2487W" Then
                                toteml = toteml + (Int((rsATT("EMLTAKEN") / SaveHours) + 0.9999))
                            Else
                                'Ticket #25893 Franks 08/15/2014 - begin
                                'toteml = toteml + ((rsATT("EMLTAKEN") / SaveHours))
                                xATTDHRS = 0
                                If Not IsNull(rsATT("AD_DHRS")) Then
                                    If rsATT("AD_DHRS") > 0 Then
                                        xATTDHRS = rsATT("AD_DHRS")
                                    End If
                                End If
                                If xATTDHRS > 0 Then
                                    toteml = toteml + ((rsATT("EMLTAKEN") / xATTDHRS))
                                Else
                                    If SaveHours > 0 Then
                                        toteml = toteml + ((rsATT("EMLTAKEN") / SaveHours))
                                    End If
                                End If
                                'Ticket #25893 Franks 08/15/2014 - end
                            End If
                        End If
                    End If
                    If Not IsNull(rsATT("EMLTAKEN")) Then
                        htoteml = htoteml + rsATT("EMLTAKEN")
                    End If
                Else
                    toteml = toteml + 0
                    htoteml = htoteml + 0
                End If
                rsATT.MoveNext
            Loop
            'ticket #17736
            'S.U.C.C.E.S.S - Ticket #19099 -
            If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
                lblEMLTaken.Caption = Format((toteml * SaveHours), "Fixed")
                lblEMLDay.Caption = IIf((toteml * SaveHours) > 1, "Hours", "Hour")
            Else
                lblEMLTaken.Caption = Format(toteml, "Fixed")
                lblEMLDay.Caption = IIf(toteml > 1, "Days", "Day")
            End If
    'added by Bryan ticket#11705 Sep 19, 2006
            If glbtermopen Then
                SQLQ = "SELECT ED_EML,ED_EMLT,ED_DHRS,TERM_SEQ"
                SQLQ = SQLQ & " FROM Term_HREMP WHERE TERM_SEQ=" & glbTERM_Seq
                rsTB.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
            Else
                SQLQ = "SELECT ED_EML,ED_EMLT,ED_DHRS,ED_EMPNBR"
                SQLQ = SQLQ & " FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
                rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            End If
            
            rsTB("ED_EMLT") = htoteml
            rsTB.Update
            
            'uncomment in order to allow EML outstanding to be measured in parts of days
            'If gbCompSerial = "S/N - 2288W" Then 'musashi
            If Not IsNull(rsTB("ED_EML")) Then
                xOuts = rsTB("ED_EML") - toteml
            Else
                xOuts = 0 - toteml
            End If
            
            'Hemu
            SavOutE = xOuts * SaveHours
            'Hemu
            
           'Else
           '     xOuts = (rsTB("ED_EML") * SaveHours) - htoteml
           ' End If
           
           
           'wellington duffrine ticket ##17736
           'S.U.C.C.E.S.S - Ticket #19099
            If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
                lblEMLOSV.Caption = Format((xOuts * SaveHours), "Fixed")
                If CDbl(lblEMLOSV.Caption) > 1 Then lblDays.Caption = "Hours" Else lblDays.Caption = "Hour"
            Else
                lblEMLOSV.Caption = Format(xOuts, "Fixed")
                If CDbl(lblEMLOSV.Caption) > 1 Then lblDays.Caption = "Days" Else lblDays.Caption = "Day"
            End If
           
'end bryan
        Else
            If glbtermopen Then
                SQLQ = "SELECT ED_EML,ED_EMLT,ED_DHRS,TERM_SEQ"
                SQLQ = SQLQ & " FROM Term_HREMP WHERE TERM_SEQ=" & glbTERM_Seq
                rsTB.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
            Else
                SQLQ = "SELECT ED_EML,ED_EMLT,ED_DHRS,ED_EMPNBR"
                SQLQ = SQLQ & " FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
                rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            End If
            
            'wellington duffrine ticket ##17736
            'S.U.C.C.E.S.S - Ticket #19099
            If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
                lblEMLDay.Caption = "Hour"
            Else
                lblEMLDay.Caption = "Day"
            End If
            lblEMLTaken.Caption = "0.00"
            
            
            If Not IsNull(rsTB("ED_EML")) Then
                lblEMLOSV.Caption = Format(rsTB("ED_EML"), "Fixed")
                'wellington duffrine ticket ##17736
                'S.U.C.C.E.S.S - Ticket #19099
                If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
                    If IsNumeric(rsTB("ED_EML")) Then
                      lblEMLOSV.Caption = Format((rsTB("ED_EML") * SaveHours), "Fixed")
                    End If
                    If CDbl(lblEMLOSV.Caption) > 1 Then lblDays.Caption = "Hours" Else lblDays.Caption = "Hour"
                Else
                    lblEMLOSV.Caption = Format((rsTB("ED_EML")), "Fixed")
                    If CDbl(lblEMLOSV.Caption) > 1 Then lblDays.Caption = "Days" Else lblDays.Caption = "Day"
                End If
            Else
                lblEMLOSV.Caption = Format(0, "Fixed")
                If CDbl(lblEMLOSV.Caption) > 1 Then lblDays.Caption = "Days" Else lblDays.Caption = "Day"
                'wellington duffrine ticket ##17736
                'S.U.C.C.E.S.S - Ticket #19099
                If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
                    lblDays.Caption = "Hour"
                End If
            End If
            SavOutE = CDbl(lblEMLOSV.Caption) * SaveHours
            rsTB.Close
            
        End If
        rsATT.Close

    'Hemu
    
End Sub

Private Sub UPDEML(xmthr)
Dim rsTB As New ADODB.Recordset
Dim SQLQ
Dim IfChange As Boolean
Dim xOTBANK, xTemp, xTemp01, xNum, xEMLOSV
IfChange = False

If glbtermopen Then
    SQLQ = "SELECT ED_EML,ED_EMLT,ED_DHRS,TERM_SEQ"
    SQLQ = SQLQ & " FROM Term_HREMP WHERE TERM_SEQ=" & glbTERM_Seq
    rsTB.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "SELECT ED_EML,ED_EMLT,ED_DHRS,ED_EMPNBR"
    SQLQ = SQLQ & " FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
    rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

If rsTB.EOF Then
    Exit Sub
End If

If xmthr = True Then
    rsTB("ED_EMLT") = rsTB("ED_EMLT") + SavEML
    rsTB.Update
    glbENTScreen = True
    
    'wellington duffrine ticket ##17736
    'S.U.C.C.E.S.S - Ticket #19099
    If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
        If Not IsNull(rsTB("ED_EML")) And Not IsNull(rsTB("ED_EMLT")) Then
            lblEMLOSV.Caption = Format(((rsTB("ED_EML") - rsTB("ED_EMLT")) * SaveHours), "Fixed")
            If (rsTB("ED_EML") * SaveHours) > 1 Then lblDays.Caption = "Hours" Else lblDays.Caption = "Hour"
        End If
    Else
        lblEMLOSV.Caption = Format(rsTB("ED_EML") - rsTB("ED_EMLT"), "Fixed")
          If rsTB("ED_EML") > 1 Then lblDays.Caption = "Days" Else lblDays.Caption = "Day"
    End If
  
Else
    'Linamar added by Bryan 13/Oct/05 Ticket#9264
    
    If Not glbLinamar Then
        'EML Taken is stored as hours, EML is stored as days
        'For most it will be any EML will be a day
        If glbCompSerial = "S/N - 2288W" Then   'musashi
            
        Else
            SavOutE = (rsTB("ED_EML") * SaveHours) - rsTB("ED_EMLT")
        End If
    ' Added by Sam to show outstanding EML on Jerry's request 06/23/2006
        If SaveHours > 0 Then
            'wellington duffrine ticket ##17736
            'S.U.C.C.E.S.S - Ticket #19099
            If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
                lblEMLOSV.Caption = Format(SavOutE, "Fixed")
                If SavOutE > 1 Then lblDays.Caption = "Hours" Else lblDays.Caption = "Hour"
            Else
                lblEMLOSV.Caption = Format(SavOutE / SaveHours, "Fixed")
                If rsTB("ED_EML") > 1 Then lblDays.Caption = "Days" Else lblDays.Caption = "Day"
            End If
        Else
            'S.U.C.C.E.S.S - Ticket #19099
            If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
                lblEMLOSV.Caption = "0.00"
                lblDays.Caption = "Hour"
            Else
                lblEMLOSV.Caption = "0.00"
                lblDays.Caption = "Day"
            End If
        End If
    Else
        SavOutE = rsTB("ED_EML") - rsTB("ED_EMLT")
    End If
End If

rsTB.Close
End Sub

' dkostka - 12/20/2001 - Fixed incident count updating, moved to its own sub here.
' Frank - 10/23/03 - ticket #4787
' The old way of Incident calculation didn't work in v7.1, get the incident number from DB after Insert, Edit and Delete
Private Sub UpdateIncCount(ReadWriteValue As ReadWrite)
    Dim rsTB As New ADODB.Recordset
    Dim rsTC As New ADODB.Recordset
    Dim rsTD As New ADODB.Recordset
    Dim SQLQ, XCNTIND
    
    If ReadWriteValue = RW_WRITE Then
        If Not glbtermopen Then
            SQLQ = "UPDATE HREMP SET ED_INCIDCNT=0 WHERE ED_EMPNBR = " & glbLEE_ID
            gdbAdoIhr001.Execute SQLQ
            If glbOracle Then
                SQLQ = "UPDATE HREMP SET HREMP.ED_INCIDCNT=( select qry_INCID.INCIDNBR "
                SQLQ = SQLQ & " FROM qry_INCID where  HREMP.ED_EMPNBR = qry_INCID.EMPNBR )"
                SQLQ = SQLQ & " WHERE ED_EMPNBR = " & glbLEE_ID
                gdbAdoIhr001.Execute SQLQ
            ElseIf glbSQL Then
            'Incident Number for Musashi May 27,2002
                If glbCompSerial = "S/N - 2288W" Then
                    SQLQ = "SELECT * FROM HREMP "
                    SQLQ = SQLQ & " WHERE  ED_EMPNBR = " & glbLEE_ID
                    If rsTC.State <> 0 Then rsTC.Close
                    rsTC.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
                    If Not rsTC.EOF Then
                        XCNTIND = 0
                        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsTC("ED_EMPNBR") & ""
                        SQLQ = SQLQ & " AND AD_INCID<>0"
                        If rsTD.State <> 0 Then rsTD.Close
                        rsTD.Open SQLQ, gdbAdoIhr001, adOpenStatic
                        Do While Not rsTD.EOF
                            If IsDate(rsTD("AD_DOA")) Then
                                If DateDiff("d", CVDate(rsTD("AD_DOA")), DateAdd("m", -6, Now)) <= 0 Then
                                    If rsTD("AD_INCID") Then
                                        XCNTIND = XCNTIND + 1
                                    End If
                                End If
                            End If
                            rsTD.MoveNext
                        Loop
                        rsTD.Close
                        rsTC("ED_INCIDCNT") = XCNTIND
                        rsTC.Update
                        rsTC.MoveNext
                    End If
                    rsTC.Close
                Else
                    SQLQ = "UPDATE HREMP SET HREMP.ED_INCIDCNT=qry_INCID.INCIDNBR "
                    SQLQ = SQLQ & " FROM HREMP INNER JOIN qry_INCID ON HREMP.ED_EMPNBR = qry_INCID.EMPNBR"
                    SQLQ = SQLQ & " WHERE ED_EMPNBR = " & glbLEE_ID
                    gdbAdoIhr001.Execute SQLQ
                End If
            Else
                SQLQ = "UPDATE HREMP RIGHT JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR "
                SQLQ = SQLQ & " SET HREMP.ED_INCIDCNT = HREMP.ED_INCIDCNT-HR_ATTENDANCE.AD_INCID "
                SQLQ = SQLQ & " WHERE AD_INCID<>0"
                SQLQ = SQLQ & " AND ED_EMPNBR = " & glbLEE_ID
                gdbAdoIhr001.Execute SQLQ
                
            End If
        End If
    End If
    
    ' We're updating variables from the database, read the count from HREMP
    If glbtermopen Then
        rsTB.Open "SELECT ED_INCIDCNT,TERM_SEQ FROM Term_HREMP WHERE TERM_SEQ=" & glbTERM_Seq, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        rsTB.Open "SELECT ED_INCIDCNT,ED_EMPNBR FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    If IsNull(rsTB("ED_INCIDCNT")) Then
        cntSick = 0
    Else
        cntSick = rsTB("ED_INCIDCNT")
    End If
    lblIncident.Caption = "Total # of Incidents = " & Str(cntSick)
    rsTB.Close

End Sub

''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
Dim SQLQ
If glbtermopen Then
    SQLQ = "SELECT *"
    SQLQ = SQLQ & " FROM Term_ATTENDANCE "
ElseIf xAD = "AD" Then
    SQLQ = "SELECT * "
    SQLQ = SQLQ & " FROM HR_ATTENDANCE "
Else
    SQLQ = "SELECT "
    SQLQ = SQLQ & "AH_COMPNO AS AD_COMPNO,"
    SQLQ = SQLQ & "AH_EMPNBR AS AD_EMPNBR,"
    SQLQ = SQLQ & "AH_DOA AS AD_DOA,"
    SQLQ = SQLQ & "AH_HRS AS AD_HRS,"
    SQLQ = SQLQ & "AH_REASON AS AD_REASON,"
    SQLQ = SQLQ & "AH_CHRGCODE AS AD_CHRGCODE,"
    SQLQ = SQLQ & "AH_PROJECT_CODE AS AD_PROJECT_CODE,"
    SQLQ = SQLQ & "AH_MACHINE_NUM AS AD_MACHINE_NUM,"
    SQLQ = SQLQ & "AH_SUPER AS AD_SUPER,"
    SQLQ = SQLQ & "AH_SHIFT AS AD_SHIFT,"
    SQLQ = SQLQ & "AH_WCBNBR AS AD_WCBNBR,"
    SQLQ = SQLQ & "AH_INCID AS AD_INCID,"
    SQLQ = SQLQ & "AH_FMLA AS AD_FMLA,"
    SQLQ = SQLQ & "AH_INDICATOR AS AD_INDICATOR,"
    SQLQ = SQLQ & "AH_SEN AS AD_SEN,"
    SQLQ = SQLQ & "AH_EMELEA AS AD_EMELEA,"
    SQLQ = SQLQ & "AH_COMM AS AD_COMM,"
    SQLQ = SQLQ & "AH_ATT_ID AS AD_ATT_ID,"
    
    SQLQ = SQLQ & "AH_JOB AS AD_JOB,"
    SQLQ = SQLQ & "AH_ORG AS AD_ORG,"
    SQLQ = SQLQ & "AH_SALARY AS AD_SALARY,"
    SQLQ = SQLQ & "AH_SALCD AS AD_SALCD,"
    SQLQ = SQLQ & "AH_DHRS AS AD_DHRS,"
    SQLQ = SQLQ & "AH_WHRS AS AD_WHRS,"
    SQLQ = SQLQ & "AH_POINT AS AD_POINT,"
    SQLQ = SQLQ & "AH_UPLOAD AS AD_UPLOAD,"
    SQLQ = SQLQ & "AH_CALCHRS AS AD_CALCHRS,"
    SQLQ = SQLQ & "AH_PAYROLL_ID AS AD_PAYROLL_ID,"
    SQLQ = SQLQ & "AH_GLNO AS AD_GLNO,"
   
    SQLQ = SQLQ & "AH_LDATE AS AD_LDATE,"
    SQLQ = SQLQ & "AH_LTIME AS AD_LTIME,"
    SQLQ = SQLQ & "AH_LUSER AS AD_LUSER,"
    
    '7.9 Enhancement
    SQLQ = SQLQ & "AH_DOCKEY AS AD_DOCKEY,"
    
    'Hemu - 06/08/2004 Begin - Ticket #6306
    SQLQ = SQLQ & "AH_DISCIPLINE AS AD_DISCIPLINE"
    'Hemu - 06/08/2004 End
    If glbBurlTech Then
        SQLQ = SQLQ & ",AH_LEPOINT AS AD_LEPOINT"
    End If
    If glbCompSerial = "S/N - 2242W" Then  'ccac london
        SQLQ = SQLQ & ",AH_PAYENDDATE AS AD_PAYENDDATE"
    End If
    
    'Ticket #18668 - 7.9 Enhancement
    SQLQ = SQLQ & ",AH_SOURCE"
    
    SQLQ = SQLQ & ",AH_REGION AS AD_REGION"
    
    SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY "
End If
clpChrgCode = ""
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    If glbtermopen Then
        SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockPessimistic
    ElseIf xAD = "AD" Then
        SQLQ = SQLQ & " WHERE AD_EMPNBR = " & glbLEE_ID
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
    Else
        SQLQ = SQLQ & " WHERE AH_EMPNBR = " & glbLEE_ID
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
    End If
Else
    If xAD = "AD" Then
        SQLQ = SQLQ & " WHERE AD_ATT_ID=" & Data1.Recordset!AD_ATT_ID
    Else
        SQLQ = SQLQ & " WHERE AH_ATT_ID = " & Data1.Recordset!AD_ATT_ID
    End If
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    If glbtermopen Then
          SQLQ = Replace(SQLQ, "AH_ATT_ID", "AD_ATT_ID")
          rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
          rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    Call SET_UP_MODE
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    elpSupShow.Text = "" 'Ticket #14555
    Call Set_Control("R", Me, rsDATA)

    clpChrgCode = txtChrgCode
    clpCode(2) = txtShift
    clpGLNum = txtWSIB
    
    If Not rsDATA.EOF Then Call getCodes(xAD) 'Ticket #28846 Franks 08/16/2016
End If
chkBackDated.Value = IsDate(dlpPayEndDate)

'lblTitle(22).Visible = IsDate(dlpPayEndDate)
'dlpPayEndDate.Visible = IsDate(dlpPayEndDate)


Call SET_UP_MODE
Call cmdModify_Click
 End Sub

Private Sub txtWSIB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
    End If
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Index As Integer, Cancel As Integer)
If glbCompSerial = "S/N - 2173W" And ((gSec_Add_Attendance And Not gSec_Upd_Attendance) Or (Not gSec_Add_Attendance And gSec_Upd_Attendance)) Then
    Cancel = False
Else
    'Ticket #26576 - WDGPHU - Cannot maintain FX* codes from Attendance
    If glbCompSerial = "S/N - 2411W" And UCase(Left(clpCode(1).Text, 2)) = "FX" And gsFLEX_LOGIC Then
        Call cmdCancel_Click
    ElseIf gsDISABLE_COMPTIME And (UCase(Left(clpCode(1).Text, 2)) = "OT" Or UCase(Left(clpCode(1).Text, 2)) = "CT") Then
        'Ticket #30305 - Disable Compensatory Time Entries
        Call cmdCancel_Click
    Else
        Cancel = Not isUpdated(Me)
    End If
End If
End Sub

Private Sub vbxTrueGrid_HeadClick(Index As Integer, ByVal ColIndex As Integer)
        Dim SQLQ As String
    
        If vbxTrueGrid(Index).Tag = "ASC" Then
            vbxTrueGrid(Index).Tag = "DESC"
        Else
            vbxTrueGrid(Index).Tag = "ASC"
        End If
        
    If glbLinamar Then
        SQLQ = " CASE WHEN AD_SUPER IS NOT NULL AND LEN(AD_SUPER)>2 "
        SQLQ = SQLQ & " THEN RIGHT(AD_SUPER,3)+'-'+"
        SQLQ = SQLQ & " LEFT(AD_SUPER,LEN(AD_SUPER)-3) "
        SQLQ = SQLQ & " ELSE STR(AD_SUPER) END "
        SQLQ = SQLQ & " AS SUPER "
    Else
        If glbOracle Then
            SQLQ = " AD_SUPER AS SUPER "
        Else
            SQLQ = " STR(AD_SUPER) AS SUPER "
        End If
    End If

    If glbtermopen Then
        SQLQ = "SELECT Term_ATTENDANCE.*, " & SQLQ
        SQLQ = SQLQ & " FROM Term_ATTENDANCE "
        SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
        If glbWFC Then 'Ticket #25148 Franks 03/04/2014
            If xHideWFCAttCodes Then
                SQLQ = SQLQ & " AND NOT (AD_REASON = 'REG' OR AD_REASON = 'OT') "
            End If
        End If
    ElseIf xAD = "AD" Then
        SQLQ = "SELECT HR_ATTENDANCE.*, " & SQLQ
        SQLQ = SQLQ & " FROM HR_ATTENDANCE "
        SQLQ = SQLQ & " WHERE AD_EMPNBR = " & glbLEE_ID
        If glbWFC Then 'Ticket #25148 Franks 03/04/2014
            If xHideWFCAttCodes Then
                SQLQ = SQLQ & " AND NOT (AD_REASON = 'REG' OR AD_REASON = 'OT') "
            End If
        End If
    Else
        SQLQ = "SELECT " & Replace(SQLQ, "AD_", "AH_") & ","
        SQLQ = SQLQ & "AH_COMPNO AS AD_COMPNO,"
        SQLQ = SQLQ & "AH_EMPNBR AS AD_EMPNBR,"
        SQLQ = SQLQ & "AH_DOA AS AD_DOA,"
        SQLQ = SQLQ & "AH_HRS AS AD_HRS,"
        SQLQ = SQLQ & "AH_REASON AS AD_REASON,"
        SQLQ = SQLQ & "AH_CHRGCODE AS AD_CHRGCODE,"
        SQLQ = SQLQ & "AH_PROJECT_CODE AS AD_PROJECT_CODE,"
        SQLQ = SQLQ & "AH_SUPER AS AD_SUPER,"
        SQLQ = SQLQ & "AH_SHIFT AS AD_SHIFT,"
        SQLQ = SQLQ & "AH_WCBNBR AS AD_WCBNBR,"
        SQLQ = SQLQ & "AH_INCID AS AD_INCID,"
        SQLQ = SQLQ & "AH_FMLA AS AD_FMLA,"
        SQLQ = SQLQ & "AH_INDICATOR AS AD_INDICATOR,"
        SQLQ = SQLQ & "AH_SEN AS AD_SEN,"
        SQLQ = SQLQ & "AH_EMELEA AS AD_EMELEA,"
        SQLQ = SQLQ & "AH_COMM AS AD_COMM,"
        SQLQ = SQLQ & "AH_ATT_ID AS AD_ATT_ID,"
        
        SQLQ = SQLQ & "AH_JOB AS AD_JOB,"
        SQLQ = SQLQ & "AH_ORG AS AD_ORG,"
        SQLQ = SQLQ & "AH_SALARY AS AD_SALARY,"
        SQLQ = SQLQ & "AH_SALCD AS AD_SALCD,"
        SQLQ = SQLQ & "AH_DHRS AS AD_DHRS,"
        SQLQ = SQLQ & "AH_WHRS AS AD_WHRS,"
        SQLQ = SQLQ & "AH_POINT AS AD_POINT,"
        SQLQ = SQLQ & "AH_UPLOAD AS AD_UPLOAD,"
        SQLQ = SQLQ & "AH_CALCHRS AS AD_CALCHRS,"
        SQLQ = SQLQ & "AH_PAYROLL_ID AS AD_PAYROLL_ID,"
        SQLQ = SQLQ & "AH_GLNO AS AD_GLNO,"
       
        SQLQ = SQLQ & "AH_LDATE AS AD_LDATE,"
        SQLQ = SQLQ & "AH_LTIME AS AD_LTIME,"
        SQLQ = SQLQ & "AH_LUSER AS AD_LUSER,"
        SQLQ = SQLQ & "AH_MACHINE_RATE AS AD_MACHINE_RATE,"  'ticket 8332, error 3265
        SQLQ = SQLQ & "AH_MACHINE_HRS AS AD_MACHINE_HRS," 'ticket 8332, error 3265
        SQLQ = SQLQ & "AH_MACHINE_NUM AS AD_MACHINE_NUM," 'ticket 8332, error 3265
        'Hemu - 06/08/2004 Begin - Ticket #6306
        SQLQ = SQLQ & "AH_DISCIPLINE AS AD_DISCIPLINE"
        'Hemu - 06/08/2004 End
        If glbBurlTech Then
            SQLQ = SQLQ & ",AH_LEPOINT AS AD_LEPOINT"
        End If
        If glbCompSerial = "S/N - 2242W" Then  'ccac london
            SQLQ = SQLQ & ",AH_PAYENDDATE AS AD_PAYENDDATE"
        End If
    
        SQLQ = SQLQ & ",AH_REGION AS AD_REGION"

        SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY "
        SQLQ = SQLQ & " WHERE AH_EMPNBR = " & glbLEE_ID
        If glbWFC Then 'Ticket #25148 Franks 03/04/2014
            If xHideWFCAttCodes Then
                SQLQ = SQLQ & " AND NOT (AH_REASON = 'REG' OR AH_REASON = 'OT') "
            End If
        End If
    End If

    If Len(vbxTrueGrid(Index).Columns(ColIndex).DataField) > 0 Then
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid(Index).Columns(ColIndex).DataField & " " & vbxTrueGrid(Index).Tag
    End If
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True

End Sub

Private Sub vbxTrueGrid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        memComments.SetFocus
End If
End Sub

Private Sub vbxTrueGrid_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)

Call Display_Value
'Casey House
If glbCompSerial = "S/N - 2214W" Then
    If Not glbtermopen Then
        If Not Data1.Recordset.EOF Then
            If IsNull(Data1.Recordset("AD_SHIFT")) Then clpCode(2).Text = "" Else clpCode(2).Text = Data1.Recordset("AD_SHIFT")
        Else
            clpCode(2).Text = ""
        End If
    End If
End If
End Sub

Private Sub ATTCode_Desc(Indx As Integer)
Dim SQLQ As String
Dim rsCode As New ADODB.Recordset
'''On Error GoTo ATTCode_Err

If Indx = 1 Then
    fglbEMELEA = 0
    fglbINC = 0
    fglbSen = 0
    fglbPoint = Null
    If Len(clpCode(Indx).Text) > 0 Then
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME='ADRE' "
        SQLQ = SQLQ & " AND TB_KEY='" & clpCode(Indx).Text & "'"
        
        rsCode.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        chkEMELEA.Enabled = False
        If Not rsCode.EOF Then
            If rsCode("TB_USR3") <> 0 Then
                chkEMELEA.Enabled = True
            End If
            fglbINC = rsCode("TB_INDICATOR")
            fglbSen = rsCode("TB_SEN")
            If Not IsNull(rsCode("TB_USR3")) Then
                fglbEMELEA = rsCode("TB_USR3")
            Else
                fglbEMELEA = 0
            End If
            fglbPoint = rsCode("TB_USR2")
        End If
    End If
ElseIf Indx = 4 Then
    If Len(clpCode(Indx).Text) > 0 Then
        SQLQ = "SELECT * FROM HR_MACHINE_NUM WHERE MACHINE_NUM='" & clpCode(Indx).Text & "'"
        rsCode.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsCode.EOF Then
            If IsNumeric(rsCode("RATE")) Then
                medMachineRate = rsCode("RATE")
            End If
        End If
    End If
End If
Exit Sub

ATTCode_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ATT Code Snap", "Key", "SELECT")
Resume Next

End Sub

Private Sub CR_JobHis_Snap()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String
Dim dynaJobHIS As New ADODB.Recordset
'''On Error GoTo JobHis_Err
fglbJobList = ""
Screen.MousePointer = HOURGLASS
If glbtermopen Then
    SQLQ = "Select * from Term_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY JH_CURRENT " & IIf(glbSQL, "DESC", "") & ",JH_SDATE DESC"

    If dynaJobHIS.State <> 0 Then dynaJobHIS.Close
    dynaJobHIS.Open SQLQ, gdbAdoIhr001X, adOpenStatic
Else
    SQLQ = "Select * from HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE JH_EMPNBR=" & glbLEE_ID
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Hours per Week", "HR_JOB_History", "SELECT")
Screen.MousePointer = DEFAULT
Resume Next

End Sub
Private Function EntStatEccess(xEmpNo, xDOA)
Dim X%, Msg$, Msg1$, Answer, SQLQ
Dim DgDef As Double, Response%, SavWork, Title$
Dim rsSTAT As New ADODB.Recordset
Dim FlagStat As Boolean
    EntStatEccess = False
    
    SQLQ = "SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE AD_DOA = ('" & Format(xDOA, "mmm dd,yyyy") & "') "
    SQLQ = SQLQ & "AND AD_REASON = 'STAT' "
    SQLQ = SQLQ & "AND AD_EMPNBR = " & xEmpNo & " "
    If AddChg <> "A" Then
        SQLQ = SQLQ & " AND AD_ATT_ID <> " & Data1.Recordset("AD_ATT_ID")
    End If
    rsSTAT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    FlagStat = False
    If Not rsSTAT.EOF Then
        FlagStat = True
    End If
    rsSTAT.Close
    If FlagStat Then
        Msg1$ = Chr(10) & "You can't enter this Attendance record" & Chr(10)
        'DgDef = MB_RETRYCANCEL + MB_ICONSTOP + MB_DEFBUTTON1
    
        Msg$ = "Warning: There is STAT attendance record of " & Format(xDOA, "mmm dd,yyyy") & " " & Msg1$
        MsgBox Msg$

        If IsDate(dlpToDate) Then
            dlpReviewDate = DateAdd("d", 1, dlpReviewDate)
        Else
            Call cmdCancel_Click
        End If

        Exit Function
    End If

EntStatEccess = True

Exit Function
EntEcc1:
If Response% = IDRETRY Then
    fglbRetry = True
End If

End Function

Private Function EntUSBEccess(xEmpNo, xDOA, xHrs)
Dim X%, Msg$, Msg1$, Answer, SQLQ, Logx, xWrk, SavWork0
Dim DgDef As Double, Response%, SavWork, Title$
Dim rsEmp As New ADODB.Recordset
Dim rsUSBA As New ADODB.Recordset
Dim rsUSBT As New ADODB.Recordset
Dim xID, xUnion
Dim xUSBA, xUSBT, xUSBO
Dim xFDate, xTDate
    EntUSBEccess = False
    
    
    'Get Union Code from Attendance table
    xUnion = clpCode(0) '""
    
    'Get the Amount of USB Bank
    xUSBA = -999
    SQLQ = "SELECT * FROM WHSCC_USB WHERE WU_ORG = '" & xUnion & "' "
    SQLQ = SQLQ & "AND WU_EFDATE <= ('" & Format(xDOA, "mmm dd,yyyy") & "') "
    SQLQ = SQLQ & "AND WU_ETDATE >= ('" & Format(xDOA, "mmm dd,yyyy") & "') "
    rsUSBA.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsUSBA.EOF Then
        xFDate = rsUSBA("WU_EFDATE")
        xTDate = rsUSBA("WU_ETDATE")
        xUSBA = rsUSBA("WU_USB")
    End If
    rsUSBA.Close
    If xUSBA = -999 Then
        'EntUSBEccess = True
        Msg1$ = ". Select" & Chr(10)
        Msg1$ = Msg1$ & "'Retry' to modify the transaction," & Chr(10)
        Msg1$ = Msg1$ & "'Cancel' to terminate this transaction"
        DgDef = MB_RETRYCANCEL + MB_ICONSTOP + MB_DEFBUTTON1
    
        Msg$ = "Warning: Union Sick Bank of " & xUnion & " has not been setup that the" & Chr(10)
        Msg$ = Msg$ & "Sick Entitlement Date range includes " & Format(xDOA, "mmm dd,yyyy") & " " & Msg1$
        Response% = MsgBox(Msg, DgDef, Title)
        If Response% = IDCANCEL Then
            If IsDate(dlpToDate) Then
                dlpToDate = ""
            Else
                Call cmdCancel_Click
            End If
        End If
        Exit Function
    End If
    
    'Get the Amount of USB Taken
    If Not IsEmpty(Data1.Recordset("AD_ATT_ID")) Then
        xID = Data1.Recordset("AD_ATT_ID")
    Else
        xID = -999
    End If
    xUSBT = 0
'    SQLQ = "SELECT SUM(AD_HRS) AS USBTAKEN FROM HR_ATTENDANCE LEFT JOIN HREMP "
'    SQLQ = SQLQ & "ON HR_ATTENDANCE.AD_EMPNBR = HREMP.ED_EMPNBR "
'    SQLQ = SQLQ & "WHERE ED_ORG = '" & xUnion & "' "
    SQLQ = "SELECT SUM(AD_HRS) AS USBTAKEN FROM HR_ATTENDANCE "
    SQLQ = SQLQ & "WHERE AD_ORG = '" & xUnion & "' "
    SQLQ = SQLQ & "AND AD_REASON = 'USB' "
    SQLQ = SQLQ & "AND AD_DOA >= ('" & Format(xFDate, "mmm dd,yyyy") & "') "
    SQLQ = SQLQ & "AND AD_DOA <= ('" & Format(xTDate, "mmm dd,yyyy") & "') "
    SQLQ = SQLQ & "AND AD_ATT_ID <> " & xID & " "
    SQLQ = SQLQ & "GROUP BY AD_ORG "
    rsUSBT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsUSBT.EOF Then
        If rsUSBT("USBTAKEN") > 0 Then
            xUSBT = rsUSBT("USBTAKEN")
        End If
    End If
    rsUSBT.Close
    xUSBT = xUSBT + xHrs
    xUSBO = xUSBA - xUSBT
    If xUSBO < 0 Then
        Msg1$ = ". Select" & Chr(10)
        Msg1$ = Msg1$ & "'Retry' to modify the transaction," & Chr(10)
        Msg1$ = Msg1$ & "'Cancel' to terminate this transaction"
        DgDef = MB_RETRYCANCEL + MB_ICONSTOP + MB_DEFBUTTON1
    
        Msg$ = "Warning: Union Sick Bank of " & xUnion & " has" & Chr(10)
        Msg$ = Msg$ & "been exceeded by " & Format(-xUSBO, "Fixed") & " hours " & Msg1$
        Response% = MsgBox(Msg, DgDef, Title)
        If Response% = IDCANCEL Then
            If IsDate(dlpToDate) Then
                dlpToDate = ""
            Else
                Call cmdCancel_Click
            End If
        End If
        Exit Function
    End If


EntUSBEccess = True

Exit Function
EntEcc1:
If Response% = IDRETRY Then
    fglbRetry = True
End If

End Function
Private Sub WhitbyGetIncidentFlags(xEmpNo, xDOA, xReason)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ, xVPoint, xNextDiscipStep
Dim xCodeFlag As Boolean
Dim xIncidentAmt, xTmpDate1, xTmpDate2
Dim xIncidentVal, xDayAmt, I, xDayDiff
    'Check Attendance Code if it has Absend checked and has Points
    xCodeFlag = True
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_KEY = '" & xReason & "' "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsTemp.EOF Then 'No this code
        xCodeFlag = False
    Else
        'If Not rsTemp("TB_ABSENCE") Then xCodeFlag = False 'Absence was unchecked
        If IsNull(rsTemp("TB_USR2")) Then xCodeFlag = False 'Point is null
        If rsTemp("TB_USR2") = 0 Then xCodeFlag = False 'Point is 0
    End If
    If IsNull(rsTemp("TB_USR2")) Then
        xVPoint = 0
    Else
        xVPoint = rsTemp("TB_USR2")
    End If
    rsTemp.Close
    If Not xCodeFlag Then
        'rsDATA("AD_INCID") = False
        Exit Sub
    End If
    
    'If xNextDiscipStep = 0  then Check the first three occurences, else go to next Disciplinary Step
    SQLQ = "SELECT AD_EMPNBR,AD_DOA,AD_REASON,AD_INCID,TB_ABSENCE,TB_USR2 FROM HR_ATTENDANCE "
    SQLQ = SQLQ & "LEFT JOIN HRTABL ON (HR_ATTENDANCE.AD_REASON = HRTABL.TB_KEY) AND (HR_ATTENDANCE.AD_REASON_TABL = HRTABL.TB_NAME) "
    SQLQ = SQLQ & "WHERE AD_EMPNBR =" & xEmpNo & " "
    SQLQ = SQLQ & "AND AD_DOA >= " & Date_SQL(CVDate(glbDiscipStartDate)) & " "
    SQLQ = SQLQ & "ORDER BY AD_DOA "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xIncidentAmt = 0
    xTmpDate1 = CVDate(glbDiscipStartDate)
    xTmpDate2 = CVDate(glbDiscipStartDate)
    Do While Not rsTemp.EOF
        If rsTemp("AD_INCID") Then
            xIncidentAmt = xIncidentAmt + 1
        End If
        'Frank 05/03/2004 Ticket# 6105, only check Point
        'If rsTemp("TB_ABSENCE") Then
            If Not IsNull(rsTemp("TB_USR2")) Then
                If rsTemp("TB_USR2") > 0 Then
                    If rsTemp("AD_REASON") = xReason Then 'Check the same reason code
                        xTmpDate1 = rsTemp("AD_DOA")
                    End If
                End If
            End If
        'End If
        xTmpDate2 = rsTemp("AD_DOA")
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    'xIncidentVal
    'If xTmpDate1 = xTmpDate2 And (xTmpDate1 <> CVDate(glbDiscipStartDate)) Then
    '    'it means the last Attendance record was Absent,
    '    'don't turn on the Absent flag for the new record
    '    xIncidentVal = False
    'Else
    '    xIncidentVal = True
    '    xIncidentAmt = xIncidentAmt + 1
    'End If
    xDayAmt = DateDiff("d", xTmpDate1, xDOA): xDayDiff = 0
    If xDayAmt > 10 Then
        xDayDiff = 10
    Else
        For I = 1 To xDayAmt
            xTmpDate1 = DateAdd("d", 1, xTmpDate1)
            If Not (Weekday(xTmpDate1) = 1 Or Weekday(xTmpDate1) = 7) Then
                xDayDiff = xDayDiff + 1
            End If
        Next I
    End If
    If xDayDiff > 1 Then
        xIncidentVal = True
        xIncidentAmt = xIncidentAmt + 1
    Else
        xIncidentVal = False
    End If
    rsDATA("AD_INCID") = xIncidentVal
    xOccuAmount = xIncidentAmt
    xDiscipFlag = True
End Sub

Private Sub WhitbyUpdateDisciplinary(xEmpNo, xDOA, xReason) ', xHrs, xFlag)
Dim rsTemp As New ADODB.Recordset
Dim rsTem2 As New ADODB.Recordset
Dim SQLQ, xVPoint, xNextDiscipStep, xNextStepPlus
Dim xCodeFlag As Boolean
Dim CurDiscip, NextDiscip, xREPTAU1
    'Enable this function Ticket# 6656
    ''Disable it until Whitby is ready
    'Exit Sub
    
    'Check what is the Next Disciplinary Step
    xNextDiscipStep = 1
    SQLQ = "SELECT ED_EMPNBR,ED_DISCIPLINENEXT FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("ED_DISCIPLINENEXT")) Then
            xNextDiscipStep = rsTemp("ED_DISCIPLINENEXT")
        End If
    End If
    rsTemp.Close
    
    If xNextDiscipStep = 1 And xOccuAmount <= 3 Then
        'if less than 3 Incident occurences, no Disciplianry action
        Exit Sub
    End If
    
    'Find the Disciplinary Code
    CurDiscip = "***": NextDiscip = "***"
    SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS ORDER BY DS_STEPNO " ' WHERE DS_STEPNO = " & xNextDiscipStep
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsTemp.EOF Then
        rsTemp.Close
        Exit Sub
    Else
        Do While Not rsTemp.EOF
            If rsTemp("DS_STEPNO") = xNextDiscipStep - 1 Then
                CurDiscip = rsTemp("DS_DISCIPLINE")
            End If
            If rsTemp("DS_STEPNO") = xNextDiscipStep Then
                NextDiscip = rsTemp("DS_DISCIPLINE")
            End If
            rsTemp.MoveNext
        Loop
    End If
    rsTemp.Close
    
    'If Next Disciplinary action doesn't exist, exit sub
    If NextDiscip = "***" Then
        Exit Sub
    End If
    'Check if the Current Disciplianry Action exists
    '   If it doesn't exist, create a new one using next Step
    '   If it exists, check if the Counselling Date has beed entered
    '       if not entered, exit sub, don't do anything
    '       if entered, it means the Disciplinary Action was done by HR person, create a new action
    SQLQ = "SELECT * FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND CL_TYPE = '" & CurDiscip & "' "
    SQLQ = SQLQ & "AND CL_LDATE >= " & Date_SQL(CVDate(glbDiscipStartDate)) & " "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        If Not IsDate(rsTemp("CL_COUDATE")) Then
            'if not entered, exit sub, don't do anything
            rsTemp.Close
            Exit Sub
        End If
    End If
    rsTemp.Close
    
    'Reset the current Disciplinary to False before creating a new current
    SQLQ = "UPDATE HR_COUNSEL SET CL_COMPLETED = 0 WHERE CL_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND CL_COMPLETED <> 0 "
    gdbAdoIhr001.Execute SQLQ
    
    xNextStepPlus = xNextDiscipStep + 1
    'Create Next Disciplinary Action
    SQLQ = "SELECT * FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND CL_TYPE = '" & NextDiscip & "' "
    SQLQ = SQLQ & "AND CL_LDATE >= " & Date_SQL(CVDate(glbDiscipStartDate)) & " "
    SQLQ = SQLQ & "AND CL_REASON = 'ATT' "
    SQLQ = SQLQ & "AND CL_INCDATE= " & Date_SQL(xDOA) & " "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsTemp.EOF Then
        
        'Get Next Step Number
        
        'Get Report #1 from current position
        xREPTAU1 = ""
        SQLQ = "SELECT JH_REPTAU FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR = " & xEmpNo & " "
        If rsTem2.State <> 0 Then rsTem2.Close
        rsTem2.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTem2.EOF Then
            If Not IsNull(rsTem2("JH_REPTAU")) Then
                xREPTAU1 = Trim(rsTem2("JH_REPTAU"))
            End If
        End If
        rsTem2.Close
        
        ''Put Disciplinary Action in Attendance table
        'SQLQ = "UPDATE HR_ATTENDANCE SET AD_DISCIPLINE = '" & NextDiscip & "' "
        'SQLQ = SQLQ & "WHERE AD_EMPNBR = " & xEmpNo & " "
        'SQLQ = SQLQ & "AND AD_DOA = " & Date_SQL(xDOA) & " "
        'SQLQ = SQLQ & "AND AD_REASON = '" & xReason & "' "
        'gdbAdoIhr001.Execute SQLQ
        
        rsTemp.AddNew
        rsTemp("CL_COMPNO") = "001"
        rsTemp("CL_EMPNBR") = xEmpNo
        rsTemp("CL_INCDATE") = xDOA
        rsTemp("CL_TYPE") = NextDiscip
        If Len(xREPTAU1) > 0 Then rsTemp("CL_COUBY") = xREPTAU1
        rsTemp("CL_LDATE") = Date
        rsTemp("CL_LTIME") = Time$
        rsTemp("CL_LUSER") = glbUserID
        rsTemp("CL_ATTDATE") = xDOA
        rsTemp("CL_ATTREASON") = xReason
        rsTemp("CL_COMPLETED") = -1
    Else
        rsTemp("CL_COUDATE") = Null
        rsTemp("CL_INCDATE") = xDOA
        rsTemp("CL_LDATE") = Date
        rsTemp("CL_LTIME") = Time$
        rsTemp("CL_LUSER") = glbUserID
        rsTemp("CL_ATTDATE") = xDOA
        rsTemp("CL_ATTREASON") = xReason
        rsTemp("CL_COMPLETED") = -1
    End If
    rsTemp.Update
    
    'Put Disciplinary Action in Attendance table
    SQLQ = "UPDATE HR_ATTENDANCE SET AD_DISCIPLINE = '" & NextDiscip & "' "
    SQLQ = SQLQ & "WHERE AD_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND AD_DOA = " & Date_SQL(xDOA) & " "
    SQLQ = SQLQ & "AND AD_REASON = '" & xReason & "' "
    gdbAdoIhr001.Execute SQLQ
        
    'Create a report records
    SQLQ = "DELETE FROM HRATTWRK WHERE AD_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "SELECT * FROM HRATTWRK WHERE AD_WRKEMP = '" & glbUserID & "' "
    If rsTem2.State <> 0 Then rsTem2.Close
    rsTem2.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsTem2.AddNew
    rsTem2("AD_COMPNO") = "001"
    rsTem2("AD_EMPNBR") = xEmpNo
    rsTem2("AD_DOA") = xDOA
    rsTem2("AD_REASON") = xReason
    rsTem2("AD_DISCIPLINE") = NextDiscip
    'rsTem2("AD_POINT") = ""
    rsTem2("AD_LDATE") = Date
    rsTem2("AD_LTIME") = Time$
    rsTem2("AD_WRKEMP") = glbUserID
    rsTem2.Update
    rsTem2.Close
    Call cmdViewDiscip_Click
    Unload frmECounsel
    rsTemp.Close
    
    If xNextStepPlus > xNextDiscipStep Then
        SQLQ = "UPDATE HREMP SET ED_DISCIPLINENEXT = " & xNextStepPlus & " "
        SQLQ = SQLQ & "WHERE ED_EMPNBR = " & xEmpNo
        gdbAdoIhr001.Execute SQLQ
    End If
    
End Sub
Private Sub UpdateASL(xEmpNo, xDOA, xHrs, xFLAG)
Dim rsASL As New ADODB.Recordset
Dim rsENT As New ADODB.Recordset
Dim rsASLOuts As New ADODB.Recordset
Dim SQLQ, xTaken, xRepaid, xOutStand, xHrsDiff

    If xFLAG = "U" Then 'Add & Edit
        SQLQ = "SELECT ED_EMPNBR,ED_EFDATES,ED_ETDATES FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
        rsENT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsENT.EOF Then
            If IsDate(rsENT("ED_EFDATES")) And IsDate(rsENT("ED_ETDATES")) Then
                'No Date Range check for ASL T#3304
                'If rsENT("ED_EFDATES") <= CVDate(xDOA) And rsENT("ED_ETDATES") >= CVDate(xDOA) Then
                
                    SQLQ = "SELECT * FROM WHSCC_ASL WHERE AS_EMPNBR = " & xEmpNo & " "
                    SQLQ = SQLQ & "AND AS_DOA = ('" & Format(xDOA, "mmm dd,yyyy") & "') "
                    'SQLQ = SQLQ & "AND AS_EFDATES = ('" & Format(rsENT("ED_EFDATES"), "mmm dd,yyyy") & "') "
                    'SQLQ = SQLQ & "AND AS_ETDATES = ('" & Format(rsENT("ED_ETDATES"), "mmm dd,yyyy") & "') "
                    SQLQ = SQLQ & "AND AS_CODE = 'TAKE' "
                    rsASL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If rsASL.EOF Then
                        rsASL.AddNew
                        rsASL("AS_HRSREP") = 0
                    End If
                    rsASL("AS_COMPNO") = "001"
                    rsASL("AS_EMPNBR") = xEmpNo
                    rsASL("AS_DOA") = xDOA
                    rsASL("AS_CODE") = "TAKE"
                    rsASL("AS_HRSTAK") = xHrs
                    'rsASL("AS_EFDATES") = rsENT("ED_EFDATES")
                    'rsASL("AS_ETDATES") = rsENT("ED_ETDATES")
                    rsASL("AS_LDATE") = Format(Now, "SHORT DATE")
                    rsASL("AS_LTIME") = Time$
                    rsASL("AS_LUSER") = glbUserID
                    rsASL.Update
                    rsASL.Close
                    Call Pause(0.2)
                    Call ReCalcASL(xEmpNo, "")
                'End If
            End If
        End If
        rsENT.Close
    End If
    If xFLAG = "D" Then 'Delete
        SQLQ = "DELETE FROM WHSCC_ASL WHERE AS_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND AS_DOA = ('" & Format(xDOA, "mmm dd,yyyy") & "') "
        SQLQ = SQLQ & "AND AS_CODE = 'TAKE' "
        gdbAdoIhr001.Execute SQLQ
        
        Call ReCalcASL(xEmpNo, "")
    End If
End Sub
Private Function EntASLBalance(xEmpNo)
Dim rsENT As New ADODB.Recordset
Dim rsASLT As New ADODB.Recordset
Dim xID, xUnion, SQLQ
Dim xASLA, xASLT, xASLO
Dim xfdateS, xtdateS, Uflag As Boolean
    xASLT = 0
    
    SQLQ = "SELECT ED_EMPNBR,ED_EFDATES,ED_ETDATES FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    rsENT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsENT.EOF Then
        If IsDate(rsENT("ED_EFDATES")) And IsDate(rsENT("ED_ETDATES")) Then
                xfdateS = rsENT("ED_EFDATES")
                xtdateS = rsENT("ED_ETDATES")
                Uflag = True
        End If
    End If
    rsENT.Close
    
    If Not Uflag Then
        lblASLOS.Caption = xASLT
        Exit Function
    End If
    
    'Get the Amount of ASL Taken
    xASLT = 0
    'SQLQ = "SELECT SUM(AD_HRS) AS ASLTAKEN FROM HR_ATTENDANCE "
    'SQLQ = SQLQ & "WHERE AD_REASON = 'ASL' "
    'SQLQ = SQLQ & "AND AD_EMPNBR = " & xEmpNo & " "
    'SQLQ = SQLQ & "AND AD_DOA >= ('" & Format(xfdateS, "mmm dd,yyyy") & "') "
    'SQLQ = SQLQ & "AND AD_DOA <= ('" & Format(xtdateS, "mmm dd,yyyy") & "') "
    'SQLQ = SQLQ & "GROUP BY AD_EMPNBR "
    'Franks 09/02/2003 ticket #4673 ASL was paid out,
    'but there still was amount showing on the Attendance screen
    SQLQ = "SELECT SUM(AS_HRSTAK) AS ASLTAKEN,SUM(AS_HRSREP) AS ASLREP FROM WHSCC_ASL "
    SQLQ = SQLQ & "WHERE AS_EMPNBR = " & xEmpNo & " "
    
    rsASLT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsASLT.EOF Then
        If rsASLT("ASLTAKEN") > 0 Then
            If IsNull(rsASLT("ASLTAKEN")) Then
                xASLT = rsASLT("ASLTAKEN")
            Else
                xASLT = rsASLT("ASLTAKEN") - rsASLT("ASLREP")
            End If
        End If
    End If
    rsASLT.Close
    If SaveHours > 0 Then
    
        'wellington duffrine ticket ##17736
        'S.U.C.C.E.S.S - Ticket #19099
        If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
            lblASLOS.Caption = Format((xASLT), "Fixed")
            EntASLBalance = Format((xASLT / SaveHours), "Fixed")
        Else
            EntASLBalance = Format((xASLT / SaveHours), "Fixed")
            lblASLOS.Caption = EntASLBalance
        End If
        
    End If
End Function

Private Function EntASLEccess(xEmpNo, xDOA, xHrs) 'Maximum ASL is 15 days
Dim X%, Msg$, Msg1$, Answer, SQLQ, Logx, xWrk, SavWork0
Dim DgDef As Double, Response%, Title$
Dim rsENT As New ADODB.Recordset
Dim rsASLT As New ADODB.Recordset
Dim xID, xUnion
Dim xASLA, xASLT, xASLO
Dim xfdateS, xtdateS, Uflag As Boolean

    'Franks 05/22/2003 ticket# 4103
    'Remove the 15 days max
    'Show ASL Balance in this year on Attendance screen
    EntASLEccess = True
    Exit Function
    'Franks 05/22/2003 ticket# 4103
    
    EntASLEccess = False
    If Not SaveHours > 0 Then
        EntASLEccess = True 'if Hour per day is not greater than aero, skip the checking
        Exit Function
    Else
        xASLA = SaveHours * 15
    End If
    If glbtermopen Then
        EntASLEccess = True
        Exit Function
    End If

        
    'No Date Range check for ASL T#3304
    'Check if xDOA is in the date range of Sick Entitlement date
    'SQLQ = "SELECT ED_EMPNBR,ED_EFDATES,ED_ETDATES FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    'If rsENT.State <> 0 Then rsENT.Close
    'rsENT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    'Uflag = False
    'If Not rsENT.EOF Then
    '    If IsDate(rsENT("ED_EFDATES")) And IsDate(rsENT("ED_ETDATES")) Then
    '        If rsENT("ED_EFDATES") <= CVDate(xDOA) And rsENT("ED_ETDATES") >= CVDate(xDOA) Then
    '            xFDates = rsENT("ED_EFDATES")
    '            xTDates = rsENT("ED_ETDATES")
    '            Uflag = True
    '        End If
    '    End If
    'End If
    'rsENT.Close
    'If Not Uflag Then
    '    EntASLEccess = True
    '    Exit Function
    'End If
     
    SQLQ = "SELECT ED_EMPNBR,ED_EFDATES,ED_ETDATES FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    rsENT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsENT.EOF Then
        If IsDate(rsENT("ED_EFDATES")) And IsDate(rsENT("ED_ETDATES")) Then
            If rsENT("ED_EFDATES") <= CVDate(xDOA) And rsENT("ED_ETDATES") >= CVDate(xDOA) Then
                xfdateS = rsENT("ED_EFDATES")
                xtdateS = rsENT("ED_ETDATES")
                Uflag = True
            End If
        End If
    End If
    rsENT.Close
    If Not Uflag Then
        EntASLEccess = True
        Exit Function
    End If
    

    'Get the Amount of ASL Taken
    If Not IsEmpty(rsDATA("AD_ATT_ID")) Then
        xID = rsDATA("AD_ATT_ID")
    Else
        xID = -999
    End If
    xASLT = 0

    SQLQ = "SELECT SUM(AD_HRS) AS ASLTAKEN FROM HR_ATTENDANCE "
    SQLQ = SQLQ & "WHERE AD_REASON = 'ASL' "
    SQLQ = SQLQ & "AND AD_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND AD_DOA >= ('" & Format(xfdateS, "mmm dd,yyyy") & "') "
    SQLQ = SQLQ & "AND AD_DOA <= ('" & Format(xtdateS, "mmm dd,yyyy") & "') "
    SQLQ = SQLQ & "AND AD_ATT_ID <> " & xID & " "
    SQLQ = SQLQ & "GROUP BY AD_EMPNBR "
    rsASLT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsASLT.EOF Then
        If rsASLT("ASLTAKEN") > 0 Then
            xASLT = rsASLT("ASLTAKEN")
        End If
    End If
    rsASLT.Close
    
    xASLT = xASLT + xHrs
    xASLO = xASLA - xASLT

    If xASLO < 0 Then
        If xHrs = -xASLO Then
            Msg$ = "Warning: ASL Bank of Employee " & xEmpNo & " has" & Chr(10)
            Msg$ = Msg$ & "been exceeded by " & Format(-xASLO, "Fixed") & " hours " & Chr(10)
            Msg$ = Msg$ & "This record can't be entered" & Chr(10)
            Msg$ = Msg$ & "The maximum ASL hours are " & Format(xASLA, "Fixed") & " hours (15 days)" '& Msg1$
            'Per Year maximum is 15 days
            MsgBox Msg

            If IsDate(dlpToDate) Then
                dlpToDate = ""
            Else
                Call cmdCancel_Click
            End If

            Exit Function
        Else
            Msg1$ = Chr(10) & "Click OK to accept the " & (xHrs + xASLO) & " hours to reach the top of ASL bank "
            Msg1$ = Msg1$ & Chr(10) & "Otherwise, Click on Cancel to terminate it"
            DgDef = MB_RETRYCANCEL + MB_ICONSTOP + MB_DEFBUTTON1
        
            Msg$ = "Warning: ASL Bank of Employee " & xEmpNo & " has" & Chr(10)
            Msg$ = Msg$ & "been exceeded by " & Format(-xASLO, "Fixed") & " hours " & Chr(10)
            Msg$ = Msg$ & "The maximum ASL hours are " & Format(xASLA, "Fixed") & " hours (15 days)" & Msg1$
            Response% = MsgBox(Msg$, 1)
            whsccASLFlag = False
            If Response% = IDCANCEL Then
                If IsDate(dlpToDate) Then
                    dlpToDate = ""
                Else
                    Call cmdCancel_Click
                End If
                Exit Function
            Else
                If clpCode(1) = "SICK" Then
                    If whsccAnotherFlag Then
                        medHours = xHrs + xASLO
                    Else
                        whsccExceedRemNum = xHrs + xASLO
                    End If
    
                    EntASLEccess = True
                    whsccASLFlag = True
                    Exit Function
                End If
                
                If clpCode(1) = "ASL" Then
                    medHours = xHrs + xASLO
                End If
                
            End If
            
        End If
    End If


EntASLEccess = True

Exit Function
EntEcc1:
If Response% = IDRETRY Then
    fglbRetry = True
End If

End Function

Private Sub whsccASLatt()
Dim SQLQ
            locEmpnbr = lblEEID
            locDate = dlpReviewDate
            locReason = "ASL"
            locHours = whsccExceedRemNum  'medHours
            locChrgCode = txtChrgCode
            locComment = memComments
            locSupShow = elpSupShow
            locIncident = chkIncident
            locSen = chkSeniority
            locEMELEA = chkEMELEA
            locPoint = txtPoint
            locInc = ChkInc
            locFmla = ChkFMLA
            locWsib = txtWSIB
            locJob = clpJob
            locShift = txtShift
            locUnion = clpCode(0)
            locSalary = Val(medsalary)
            locSalCode = lblSalCode
            locDHrs = txtDHRS
            locWHrs = txtWHRS
                                
            SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & locEmpnbr & " "
            If RsATTwhscc.State <> 0 Then RsATTwhscc.Close
            RsATTwhscc.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            RsATTwhscc.AddNew
            RsATTwhscc("AD_COMPNO") = "001"
            RsATTwhscc("AD_EMPNBR") = locEmpnbr
            RsATTwhscc("AD_DOA") = locDate
            RsATTwhscc("AD_REASON_TABL") = "ADRE"
            RsATTwhscc("AD_REASON") = locReason
            RsATTwhscc("AD_HRS") = locHours
            RsATTwhscc("AD_COMM") = locComment
            RsATTwhscc("AD_CHRGCODE") = locChrgCode
            RsATTwhscc("AD_SHIFT") = locShift
            If Val(locSupShow) = 0 Then
                RsATTwhscc("AD_SUPER") = Null
            Else
                RsATTwhscc("AD_SUPER") = locSupShow
            End If
            RsATTwhscc("AD_SEN") = locSen
            RsATTwhscc("AD_INCID") = chkIncident
            RsATTwhscc("AD_INDICATOR") = locInc
            RsATTwhscc("AD_FMLA") = locFmla
            RsATTwhscc("AD_EMELEA") = locEMELEA
            RsATTwhscc("AD_POINT") = locPoint
            RsATTwhscc("AD_JOB") = locJob
            RsATTwhscc("AD_SALARY") = locSalary
            RsATTwhscc("AD_SALCD") = locSalCode
            RsATTwhscc("AD_DHRS") = locDHrs
            RsATTwhscc("AD_WHRS") = locWHrs
            RsATTwhscc("AD_ORG_TABL") = "EDOR"
            RsATTwhscc("AD_ORG") = locUnion
            RsATTwhscc("AD_LDATE") = Format(Date, "Short Date")
            RsATTwhscc("AD_LTIME") = Time$
            RsATTwhscc("AD_LUSER") = glbUserID
            'Call UpdUStats(Me)
            RsATTwhscc.Update
            glbxID = RsATTwhscc("AD_ATT_ID")
            RsATTwhscc.Close
            Call UpdateASL(locEmpnbr, CVDate(locDate), locHours, "U")
            SQLQ = "DELETE FROM HR_ATTENDANCE WHERE AD_HRS = 0  AND AD_EMPNBR = " & locEmpnbr & ""
            gdbAdoIhr001.Execute SQLQ
            whsccExceedFlag = False
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
'UpdateRight = gSec_Upd_Attendance
If glbCompSerial = "S/N - 2343W" Then  'ottawa ccac
    If glbEmpNbr = glbLEE_ID Then
        UpdateRight = False
    End If
End If
'Ticket #7500 - Town of Ajax
'tkt310423 Jerry said remove serial#control for Add_Attendance security
'If glbCompSerial = "S/N - 2173W" Then
If xAttendance = "Attendance_History" Then
    UpdateRight = gSec_Upd_Attendance_History
Else
    UpdateRight = fglbAddable Or gSec_Upd_Attendance
End If
'End If
End Property

Public Property Get Addable() As Boolean
'Ticket #7500 - Town of Ajax
'tkt310423 Jerry said remove serial#control for Add_Attendance security
'If glbCompSerial = "S/N - 2173W" Then
If xAttendance = "Attendance_History" Then
    Addable = gSec_Upd_Attendance_History
Else
    Addable = fglbAddable
End If
'Else
 '   Addable = True
'End If
End Property

Public Property Get Updateble() As Boolean
'Updateble = True ' fglbUpdateble
'Ticket #7500 - Town of Ajax
'tkt310423 Jerry said remove serial#control for Add_Attendance security
'If glbCompSerial = "S/N - 2173W" Then
If xAttendance = "Attendance_History" Then
    'Ticket #26576 - WDGPHU - Cannot maintain FX* codes from Attendance
    If glbCompSerial = "S/N - 2411W" And gsFLEX_LOGIC Then
        If clpCode(1).Text = "FX+Y" Then 'Ticket #27771 Franks 12/01/2015 - allow user to enter FX+Y manually
            Updateble = gSec_Upd_Attendance_History
        Else
            Updateble = IIf(UCase(Left(clpCode(1).Text, 2)) = "FX", False, gSec_Upd_Attendance_History)
        End If
    ElseIf gsDISABLE_COMPTIME And (UCase(Left(clpCode(1).Text, 2)) = "OT" Or UCase(Left(clpCode(1).Text, 2)) = "CT") Then
        'Ticket #30305 - Disable Compensatory Time Entries
        Updateble = False
    Else
        Updateble = gSec_Upd_Attendance_History
    End If
Else
    'Ticket #26576 - WDGPHU - Cannot maintain FX* codes from Attendance
    If glbCompSerial = "S/N - 2411W" And gsFLEX_LOGIC Then
        If clpCode(1).Text = "FX+Y" Then 'Ticket #27771 Franks 12/01/2015 - allow user to enter FX+Y manually
            Updateble = gSec_Upd_Attendance
        Else
            Updateble = IIf(UCase(Left(clpCode(1).Text, 2)) = "FX", False, gSec_Upd_Attendance)
        End If
    ElseIf gsDISABLE_COMPTIME And (UCase(Left(clpCode(1).Text, 2)) = "OT" Or UCase(Left(clpCode(1).Text, 2)) = "CT") Then
        'Ticket #30305 - Disable Compensatory Time Entries
        Updateble = False
    Else
        Updateble = gSec_Upd_Attendance
    End If
End If
'End If

End Property

Public Property Get Deleteble() As Boolean
'Deleteble = fglbDeleteble
'Ticket #7500 - Town of Ajax
'tkt310423 Jerry said remove serial#control for Add_Attendance security
'If glbCompSerial = "S/N - 2173W" Then
If xAttendance = "Attendance_History" Then
    'Ticket #26576 - WDGPHU - Cannot maintain FX* codes from Attendance
    If glbCompSerial = "S/N - 2411W" And gsFLEX_LOGIC Then
        If clpCode(1).Text = "FX+Y" Then 'Ticket #27771 Franks 12/01/2015 - allow user to enter FX+Y manually
            Deleteble = gSec_Upd_Attendance_History
        Else
            Deleteble = IIf(UCase(Left(clpCode(1).Text, 2)) = "FX", False, gSec_Upd_Attendance_History)
        End If
    ElseIf gsDISABLE_COMPTIME And (UCase(Left(clpCode(1).Text, 2)) = "OT" Or UCase(Left(clpCode(1).Text, 2)) = "CT") Then
        'Ticket #30305 - Disable Compensatory Time Entries
        Deleteble = False
    Else
        Deleteble = gSec_Upd_Attendance_History
    End If
Else
    'Ticket #26576 - WDGPHU - Cannot maintain FX* codes from Attendance
    If glbCompSerial = "S/N - 2411W" And gsFLEX_LOGIC Then
        If clpCode(1).Text = "FX+Y" Then 'Ticket #27771 Franks 12/01/2015 - allow user to enter FX+Y manually
            Deleteble = gSec_Upd_Attendance
        Else
            Deleteble = IIf(UCase(Left(clpCode(1).Text, 2)) = "FX", False, gSec_Upd_Attendance)
        End If
    ElseIf gsDISABLE_COMPTIME And (UCase(Left(clpCode(1).Text, 2)) = "OT" Or UCase(Left(clpCode(1).Text, 2)) = "CT") Then
        'Ticket #30305 - Disable Compensatory Time Entries
        Deleteble = False
    Else
        Deleteble = gSec_Upd_Attendance
    End If
End If
'End If

End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Private Sub SET_UP_MODE()
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

fglbDeleteble = True
'fglbUpdateble = True

If glbOttawaCCAC And chkUpload <> 0 Then 'ticket #6148
    fglbDeleteble = False
'    fglbUpdateble = False
End If

'Ticket #7500 - Town of Ajax
'Ticket #10423 Jerry said add security for everybody 05/09/2006
'If glbCompSerial = "S/N - 2173W" Then
    fglbAddable = gSec_Add_Attendance
'End If

If xAD = "AH" Then
    fglbDeleteble = gSec_Upd_Attendance_History
    fglbUpdateble = gSec_Upd_Attendance_History
    fglbAddable = gSec_Upd_Attendance_History
End If

Call set_Buttons(UpdateState)

If Not UpdateRight Then TF = False

Call modSTUPD(TF)

End Sub

Private Sub lblEEID_Change()

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    
 If xAD = "AD" Then
        Me.Caption = "Attendance - " & Left$(glbLEE_SName, 5)
    Else
        Me.Caption = "Attendance_History - " & Left$(glbLEE_SName, 5)
    End If
    
 '   frmVATTEND.Caption = "Attendance - " & Left$(glbLEE_SName, 5)
    
    frmVATTEND.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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

Private Function Check_Overtime_Bank(xCalcFlag) As Boolean
Dim fromCurDate, toCurDate As Date
Dim rsTB As New ADODB.Recordset
Dim IfChange As Boolean
Dim SQLQ  As String
Dim xOTBANK
Dim Msg$
Dim DgDef As Double, Response%, Title$, Msg1$

Check_Overtime_Bank = True

Msg1$ = "Select" & Chr(10)
Msg1$ = Msg1$ & "'Cancel' to terminate this transaction or" & Chr(10)
Msg1$ = Msg1$ & "'Retry' to modify the transaction" & Chr(10)
Title$ = "info:HR - ATTENDANCE ENTRY"
DgDef = MB_RETRYCANCEL + MB_ICONSTOP + MB_DEFBUTTON2
Msg$ = "Warning: Hours exceeding Overtime Banked Time" & Chr(10)
Msg$ = Msg$ & Msg1$

SQLQ = "SELECT ED_VACT,ED_SICKT,ED_INCIDCNT,"
SQLQ = SQLQ & "ED_EFDATE,ED_ETDATE,ED_EFDATES,ED_ETDATES,"
If (Not glbtermopen) Then
    SQLQ = SQLQ & "ED_VAC,ED_PVAC,ED_SICK,ED_PSICK,ED_OTBANK "
Else
    SQLQ = SQLQ & "ED_VAC,ED_PVAC,ED_SICK,ED_PSICK,ED_OTBANK "
End If
If glbtermopen Then
    SQLQ = SQLQ & ",TERM_SEQ "
    SQLQ = SQLQ & " FROM Term_HREMP WHERE TERM_SEQ=" & glbTERM_Seq
    rsTB.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = SQLQ & ",ED_EMPNBR "
    SQLQ = SQLQ & " FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
    rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

If rsTB.EOF Then
    Exit Function
End If

Call DisOvertime(rsTB("ED_OTBANK"))
    
If xCalcFlag = False Then Exit Function

fromCurDate = Format("01/01/" & Year(Now), "mm/dd/yyyy")
toCurDate = Format("12/31/" & Year(Now), "mm/dd/yyyy")

'Hemu - Begin -  Town of Ajax - OT Bank
If glbCompSerial = "S/N - 2173W" Then
    If Left(UCase(clpCode(1).Text), 2) = "OT" Then
        If CVDate(dlpReviewDate) >= CVDate(fromCurDate) And CVDate(dlpReviewDate) <= CVDate(toCurDate) Then
            If AddChg = "A" Then
                rsTB("ED_OTBANK") = IIf(IsNull(rsTB("ED_OTBANK")), 0, rsTB("ED_OTBANK")) + Val(medHours)
                IfChange = True
            End If
            If AddChg = "NO" Then
                xOTBANK = xmedHours
                rsTB("ED_OTBANK") = IIf(IsNull(rsTB("ED_OTBANK")), 0, rsTB("ED_OTBANK")) + Val(medHours)
                IfChange = True
            End If
            If AddChg = "C" Then
                If (CVDate(savEntDate) >= CVDate(fromCurDate) And CVDate(savEntDate) <= CVDate(toCurDate)) _
                    And Left(UCase(xOldReason), 2) = "OT" Then
                    rsTB("ED_OTBANK") = rsTB("ED_OTBANK") + (Val(medHours) - savEnt1)
                ElseIf (CVDate(savEntDate) >= CVDate(fromCurDate) And CVDate(savEntDate) <= CVDate(toCurDate)) _
                    And Left(UCase(xOldReason), 2) = "CT" Then
                    rsTB("ED_OTBANK") = rsTB("ED_OTBANK") + oldHrs
                    rsTB("ED_OTBANK") = rsTB("ED_OTBANK") + Val(medHours)
                Else
                    rsTB("ED_OTBANK") = rsTB("ED_OTBANK") + Val(medHours)
                End If
                IfChange = True
            End If
            If AddChg = "D" Then
                rsTB("ED_OTBANK") = rsTB("ED_OTBANK") - Val(medHours)
                IfChange = True
            End If
        Else
            If AddChg = "C" Then
                If (CVDate(savEntDate) >= CVDate(fromCurDate) And CVDate(savEntDate) <= CVDate(toCurDate)) _
                    And Left(UCase(xOldReason), 2) = "OT" Then
                    rsTB("ED_OTBANK") = rsTB("ED_OTBANK") - oldHrs
                End If
            End If
        End If
    ElseIf Left(UCase(xOldReason), 2) = "OT" And AddChg = "C" Then
        If (CVDate(savEntDate) >= CVDate(fromCurDate) And CVDate(savEntDate) <= CVDate(toCurDate)) Then
            xOTBANK = rsTB("ED_OTBANK") - oldHrs
            
            If Left(UCase(clpCode(1).Text), 2) = "CT" Then
                xOTBANK = xOTBANK - Val(medHours)
                If Val(xOTBANK) < 0 Then
                    Msg$ = "Warning: Hours exceeding Overtime Banked Time: " & rsTB("ED_OTBANK") & Chr(10)
                    Msg$ = Msg$ & Msg1$
                    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                    If Response% = IDABORT Or Response% = IDCANCEL Then   'Hemu
                        If IsDate(dlpToDate.Text) Then
                            dlpToDate.Text = ""
                        Else
                            Check_Overtime_Bank = False
                            IfChange = False
                            Call cmdCancel_Click
                        End If
                    End If
                    If Response% = IDRETRY Or Response% = IDABORT Or Response% = IDCANCEL Then GoTo OvertTime_Bank
                    
                    Check_Overtime_Bank = False
                    IfChange = False
                Else
                    rsTB("ED_OTBANK") = xOTBANK
                    IfChange = True
                End If
            Else
                rsTB("ED_OTBANK") = xOTBANK
                IfChange = True
            End If
        End If
    ElseIf (Left(UCase(xOldReason), 2) = "CT" And Left(UCase(clpCode(1).Text), 2) <> "CT") And AddChg = "C" Then
        If (CVDate(savEntDate) >= CVDate(fromCurDate) And CVDate(savEntDate) <= CVDate(toCurDate)) Then
            rsTB("ED_OTBANK") = rsTB("ED_OTBANK") + oldHrs
            IfChange = True
        End If
    ElseIf Left(UCase(clpCode(1).Text), 2) = "CT" Then
        If CVDate(dlpReviewDate) >= CVDate(fromCurDate) And CVDate(dlpReviewDate) <= CVDate(toCurDate) Then
            If Not IsNull(rsTB("ED_OTBANK")) Then
                If AddChg = "D" Then
                    rsTB("ED_OTBANK") = rsTB("ED_OTBANK") + Val(medHours)
                    IfChange = True
                ElseIf AddChg = "C" Then
                    If (CVDate(savEntDate) >= CVDate(fromCurDate) And CVDate(savEntDate) <= CVDate(toCurDate)) _
                        And Left(UCase(xOldReason), 2) = "CT" Then
                        
                        xOTBANK = (rsTB("ED_OTBANK") + oldHrs) - Val(medHours)
                        If Val(xOTBANK) < 0 Then
                            Msg$ = "Warning: Hours exceeding Overtime Banked Time: " & rsTB("ED_OTBANK") & Chr(10)
                            Msg$ = Msg$ & Msg1$
                            Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                            If Response% = IDABORT Or Response% = IDCANCEL Then   'Hemu
                                If IsDate(dlpToDate.Text) Then
                                    dlpToDate.Text = ""
                                Else
                                    Check_Overtime_Bank = False
                                    IfChange = False
                                    Call cmdCancel_Click
                                End If
                            End If
                            If Response% = IDRETRY Or Response% = IDABORT Or Response% = IDCANCEL Then GoTo OvertTime_Bank
                            
                            Check_Overtime_Bank = False
                            IfChange = False
                        Else
                            rsTB("ED_OTBANK") = xOTBANK
                            IfChange = True
                        End If
                    Else
                        xOTBANK = rsTB("ED_OTBANK") - Val(medHours)
                        If Val(xOTBANK) < 0 Then
                            Msg$ = "Warning: Hours exceeding Overtime Banked Time: " & rsTB("ED_OTBANK") & Chr(10)
                            Msg$ = Msg$ & Msg1$
                            Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                            If Response% = IDABORT Or Response% = IDCANCEL Then   'Hemu
                                If IsDate(dlpToDate.Text) Then
                                    dlpToDate.Text = ""
                                Else
                                    Check_Overtime_Bank = False
                                    IfChange = False
                                    Call cmdCancel_Click
                                End If
                            End If
                            If Response% = IDRETRY Or Response% = IDABORT Or Response% = IDCANCEL Then GoTo OvertTime_Bank
                            
                            Check_Overtime_Bank = False
                            IfChange = False
                        Else
                            rsTB("ED_OTBANK") = xOTBANK
                            IfChange = True
                        End If
                    End If
                End If
                If AddChg = "A" Or AddChg = "NO" Then
                    If rsTB("ED_OTBANK") < Val(medHours) Then
                        'MsgBox "Hours exceeding Overtime banked"
                        Msg$ = "Warning: Hours exceeding Overtime Banked Time: " & rsTB("ED_OTBANK") & Chr(10)
                        Msg$ = Msg$ & Msg1$
                        Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                        If Response% = IDABORT Or Response% = IDCANCEL Then   'Hemu
                            If IsDate(dlpToDate.Text) Then
                                dlpToDate.Text = ""
                            Else
                                Check_Overtime_Bank = False
                                IfChange = False
                                Call cmdCancel_Click
                            End If
                        End If
                        If Response% = IDRETRY Or Response% = IDABORT Or Response% = IDCANCEL Then GoTo OvertTime_Bank
                        
                        Check_Overtime_Bank = False
                        IfChange = False
                    Else
                        rsTB("ED_OTBANK") = rsTB("ED_OTBANK") - Val(medHours)
                        IfChange = True
                    End If
                End If
            Else
                If AddChg <> "D" Then
                    'MsgBox "No any Overtime banked"
                    Msg$ = "Warning: No any Overtime Banked Time" & Chr(10)
                    Msg$ = Msg$ & Msg1$
                    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                    If Response% = IDABORT Or Response% = IDCANCEL Then   'Hemu
                        If IsDate(dlpToDate.Text) Then
                            dlpToDate.Text = ""
                        Else
                            Check_Overtime_Bank = False
                            IfChange = False
                            Call cmdCancel_Click
                        End If
                    End If
                    If Response% = IDRETRY Or Response% = IDABORT Or Response% = IDCANCEL Then GoTo OvertTime_Bank
                End If
                Check_Overtime_Bank = False
                IfChange = False
            End If
        End If
    End If
    Call DisOvertime(rsTB("ED_OTBANK"))
            
End If

If IfChange Then
    rsTB.Update
End If
rsTB.Close

'Hemu - End - Town of Ajax - OT Bank

Exit Function
OvertTime_Bank:
If Response% = IDRETRY Then
    fglbRetry = True
    Check_Overtime_Bank = False
End If
End Function

Private Sub Send_Email_Overtime_Exceeded(xEmail, xHours, xType)

    Dim MailBody As String
    Dim LocCode As String, LocDesc As String
    
    '''On Error GoTo ErrorHandler
    fglbSendOTEmail = False
    Load frmSendEmail
    If Len(frmSendEmail.txtFrom.Text) > 0 Then 'if no email address exists can't send
        If xType = "CT" Then
            frmSendEmail.txtSubject.Text = "info:HR Outstanding Overtime Bank Exceeded"
            MailBody = "The employee below has exceeded using his/her Overtime Banked time" & vbCrLf & vbCrLf
        Else
            frmSendEmail.txtSubject.Text = "info:HR Overtime Bank Exceeded the Maximum Bank"
            MailBody = "The employee below has exceeded his/her Maximum Banked time" & vbCrLf & vbCrLf
        End If
        MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
        MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
        MailBody = MailBody & "Date: " & dlpReviewDate.Text & vbCrLf & vbCrLf
        MailBody = MailBody & "Time Exceeded: " & xHours & " hour(s)" & vbCrLf & vbCrLf
        frmSendEmail.txtBody.Text = MailBody
        
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(0).Caption = "Sending email..."
        'frmSendEmail.txtTo.Text = "hotline@woodbridgegroup.com"
        frmSendEmail.txtTo.Text = xEmail
        frmSendEmail.Tag = ""
        frmSendEmail.cmdSend_Click
        Do
            DoEvents
        Loop Until frmSendEmail.Tag <> ""
        
        If frmSendEmail.Tag = "DONE" Then
            fglbSendOTEmail = True
        End If
        MDIMain.panHelp(0).Caption = ""
        MDIMain.panHelp(0).FloodType = 1
    End If
    
    Unload frmSendEmail
    
Exit Sub
    
ErrorHandler:
    If Err.Number = 364 Then Exit Sub
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number

End Sub

Private Sub Recalculate_OTBANK()
Dim rsEmp As New ADODB.Recordset
Dim rsAttend As New ADODB.Recordset
Dim rsAttendCT As New ADODB.Recordset
Dim SQLQ

'Set ED_OTBANK to zero for the first time otherwise Null will be updated if some Value - Null
SQLQ = "UPDATE HREMP SET ED_OTBANK = 0 WHERE ED_EMPNBR =" & glbLEE_ID
gdbAdoIhr001.Execute SQLQ

SQLQ = "SELECT ED_EMPNBR, ED_OTBANK FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic

If Not rsEmp.EOF Then
    rsEmp.MoveFirst
    
    Do While Not rsEmp.EOF
        
        If glbOracle Then
            SQLQ = "SELECT SUM(AD_HRS) AS OT_SUM FROM HR_ATTENDANCE WHERE substr(AD_REASON,1,2) = 'OT' AND AD_EMPNBR = " & rsEmp("ED_EMPNBR") & " GROUP BY AD_EMPNBR"
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            
            SQLQ = "SELECT SUM(AD_HRS) AS CT_SUM FROM HR_ATTENDANCE WHERE substr(AD_REASON,1,2) = 'CT' AND AD_EMPNBR = " & rsEmp("ED_EMPNBR") & " GROUP BY AD_EMPNBR"
            rsAttendCT.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
        Else
            SQLQ = "SELECT SUM(AD_HRS) AS OT_SUM FROM HR_ATTENDANCE WHERE LEFT(AD_REASON,2) = 'OT' AND AD_EMPNBR = " & rsEmp("ED_EMPNBR") & " GROUP BY AD_EMPNBR"
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            
            SQLQ = "SELECT SUM(AD_HRS) AS CT_SUM FROM HR_ATTENDANCE WHERE LEFT(AD_REASON,2) = 'CT' AND AD_EMPNBR = " & rsEmp("ED_EMPNBR") & " GROUP BY AD_EMPNBR"
            rsAttendCT.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
        End If
        If Not rsAttend.EOF Then
            If Not rsAttendCT.EOF Then
                SQLQ = "UPDATE HREMP SET ED_OTBANK = " & rsAttend("OT_SUM") - rsAttendCT("CT_SUM") & " WHERE ED_EMPNBR = " & rsEmp("ED_EMPNBR")
            Else
                SQLQ = "UPDATE HREMP SET ED_OTBANK = " & rsAttend("OT_SUM") & " WHERE ED_EMPNBR = " & rsEmp("ED_EMPNBR")
            End If
            gdbAdoIhr001.Execute SQLQ
        Else
            If Not rsAttendCT.EOF Then
                SQLQ = "UPDATE HREMP SET ED_OTBANK = " & 0 - rsAttendCT("CT_SUM") & " WHERE ED_EMPNBR = " & rsEmp("ED_EMPNBR")
            Else
                SQLQ = "UPDATE HREMP SET ED_OTBANK = 0 WHERE ED_EMPNBR = " & rsEmp("ED_EMPNBR")
            End If
            gdbAdoIhr001.Execute SQLQ
        End If
        rsAttend.Close
        rsAttendCT.Close
        
        rsEmp.MoveNext
    Loop
End If
rsEmp.Close

End Sub

Private Sub Calculate_Outstanding_CompTime()
    Dim xOutstanding As Double
    
    xOutstanding = (Get_OvertimeBank(glbLEE_ID, "", "") - Get_OvertimeTaken(glbLEE_ID, "", ""))
    
    If glbCompSerial = "S/N - 2433W" Then 'Kerry's Place Ticket #22332 Franks 07/26/2012
        Call Recalculate_KerrysPlaceLieu("ED_EMPNBR = " & glbLEE_ID, xOutstanding)
    End If

    If SaveHours > 0 Then
        'S.U.C.C.E.S.S - Ticket #19099
        If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
            xOutstanding = Format((xOutstanding), "Fixed")
        Else
            xOutstanding = Format((xOutstanding / SaveHours), "Fixed")
        End If
    Else
        xOutstanding = 0
    End If
    
    lblCompTimeOS.Caption = Format(Round(xOutstanding, 2), "Fixed")
    
    If xOutstanding > 1 Or xOutstanding < -1 Then
        'S.U.C.C.E.S.S - Ticket #19099
        If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
            lbCompTimeOSday.Caption = "Hours"
        Else
            lbCompTimeOSday.Caption = "Days"
        End If
    Else
        'S.U.C.C.E.S.S - Ticket #19099
        If glbCompSerial = "S/N - 2411W" Or glbCompSerial = "S/N - 2422W" Then
            lbCompTimeOSday.Caption = "Hour"
        Else
            lbCompTimeOSday.Caption = "Day"
        End If
    End If
End Sub

Private Function Is_ESSApproved_Record(xID) As Boolean
    Dim rsAttend As New ADODB.Recordset
    Dim SQLQ As String
    
    Is_ESSApproved_Record = False
    
    If xAD = "AD" Then
        SQLQ = "SELECT AD_EMPNBR, AD_ATT_ID, AD_SOURCE, AD_REQID FROM HR_ATTENDANCE"
        SQLQ = SQLQ & " WHERE AD_ATT_ID = " & xID
    Else
        SQLQ = "SELECT AH_EMPNBR, AH_ATT_ID, AH_SOURCE, AH_REQID FROM HR_ATTENDANCE_HISTORY"
        SQLQ = SQLQ & " WHERE AH_ATT_ID = " & xID
    End If
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsAttend.EOF Then
        If xAD = "AD" Then
            If rsAttend("AD_SOURCE") = "ESSAP" And Not IsNull(rsAttend("AD_REQID")) Then
                Is_ESSApproved_Record = True
            Else
                Is_ESSApproved_Record = False
            End If
        Else
            If rsAttend("AH_SOURCE") = "ESSAP" And Not IsNull(rsAttend("AH_REQID")) Then
                Is_ESSApproved_Record = True
            Else
                Is_ESSApproved_Record = False
            End If
        End If
    End If
    rsAttend.Close
    Set rsAttend = Nothing
    

End Function

Private Function WDGPH_Check_FX(xEmpNo, xCode, xDATE, xHours) 'Ticket #27771 Franks 12/01/2015
Dim rsTmpAt As New ADODB.Recordset
Dim SQLQ
Dim retval As Boolean
    retval = False
    If IsDate(xDATE) Then
        SQLQ = "SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND AD_DOA = " & Date_SQL(xDATE) & " "
        SQLQ = SQLQ & "AND ABS(AD_HRS) = " & Abs(xHours) & " "    'Ticket #29306 - Added Abs(): For -ve FX-X hours there should be +ve FX+Y hours
        SQLQ = SQLQ & "AND AD_REASON = '" & xCode & "' " 'user enter FX+Y, check if FX-X exist
        rsTmpAt.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
        If Not rsTmpAt.EOF Then
            retval = True
        End If
        rsTmpAt.Close
    End If
    WDGPH_Check_FX = retval
End Function

Private Sub LinamarSceenSetup() 'Ticket #28846 Franks 08/16/2016
    txtShift.Visible = False
    txtShift.DataField = ""
    'clpCode(2).Top = txtShift.Top
    'clpCode(2).Left = clpJob.Left
    clpCode(2).TablName = "SHFT"
    clpCode(2).MaxLength = 10
    clpCode(2).TextBoxWidth = 1000
    clpCode(2).Visible = True
    clpCode(2).TransDiv = Right(glbLEE_ID, 3)
End Sub

Sub getCodes(xAD) 'Ticket #28846 Franks 08/16/2016
If rsDATA.EOF Then Exit Sub
If glbLinamar Then
        If Not IsNull(rsDATA("AD_SHIFT")) Then
            If Len(rsDATA("AD_SHIFT")) > 3 Then clpCode(2).Text = Mid(rsDATA("AD_SHIFT"), 4) Else clpCode(2).Text = rsDATA("AD_SHIFT")
        Else
            clpCode(2).Text = ""
        End If
End If
End Sub
Private Sub UpdCodes(xAD) 'Ticket #28846 Franks 08/16/2016
    If glbLinamar Then
            If Trim(clpCode(2).Text) <> "" Then
                rsDATA("AD_SHIFT") = getShiftCodeforLinamar(clpCode(2).TransDiv & clpCode(2).Text)
            Else
                rsDATA("AD_SHIFT") = ""
            End If
    End If
End Sub
