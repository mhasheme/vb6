VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEHSF9 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "WSIB - Form 9"
   ClientHeight    =   10950
   ClientLeft      =   150
   ClientTop       =   180
   ClientWidth     =   12315
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
   ScaleWidth      =   12315
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   7035
      LargeChange     =   315
      Left            =   12000
      Max             =   100
      SmallChange     =   315
      TabIndex        =   40
      Top             =   2640
      Visible         =   0   'False
      Width           =   300
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   8520
      Top             =   10920
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   10320
      Top             =   10920
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   29
      Top             =   10530
      Width           =   12315
      _Version        =   65536
      _ExtentX        =   21722
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
      Begin VB.CommandButton cmdPrintWF9 
         Appearance      =   0  'Flat
         Caption         =   "Generate Form 9"
         Height          =   375
         Left            =   4883
         TabIndex        =   39
         Tag             =   "Generate Form 7"
         Top             =   0
         Width           =   2295
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6840
         Top             =   240
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
      DataField       =   "F9_LDATE"
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
      Left            =   10920
      MaxLength       =   25
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "F9_LTIME"
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
      Left            =   10920
      MaxLength       =   25
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "F9_LUSER"
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
      Left            =   10920
      MaxLength       =   25
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1000
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   12315
      _Version        =   65536
      _ExtentX        =   21722
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
      Begin VB.Label lblEENumber 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   120
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
         TabIndex        =   35
         Top             =   97
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
         TabIndex        =   34
         Top             =   97
         Width           =   720
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fehsF9.frx":0000
      Height          =   1935
      Left            =   240
      OleObjectBlob   =   "fehsF9.frx":0014
      TabIndex        =   0
      Top             =   480
      Width           =   11415
   End
   Begin VB.Frame ScrFrame 
      BorderStyle     =   0  'None
      Height          =   8175
      Left            =   240
      TabIndex        =   41
      Top             =   2640
      Width           =   11655
      Begin VB.Frame frQ1 
         Height          =   975
         Left            =   0
         TabIndex        =   82
         Top             =   240
         Width           =   11535
         Begin VB.TextBox txtReturnedWork 
            Appearance      =   0  'Flat
            DataField       =   "F9_RETURNED_WORK"
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
            Left            =   4920
            MaxLength       =   50
            TabIndex        =   1
            Tag             =   "01-Area where incident occurred"
            Top             =   175
            Width           =   6435
         End
         Begin VB.TextBox txtRtrnWorkAP 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   11160
            MaxLength       =   5
            TabIndex        =   84
            Top             =   555
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Frame Frame10 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   9600
            TabIndex        =   83
            Top             =   565
            Width           =   1215
            Begin VB.OptionButton optRtrnWorkAP 
               Caption         =   "PM"
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
               Left            =   720
               TabIndex        =   5
               Tag             =   "40-Time returned to work, PM"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton optRtrnWorkAP 
               Caption         =   "AM"
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
               Left            =   0
               TabIndex        =   4
               Tag             =   "40-Time returned to work, AM"
               Top             =   0
               Width           =   615
            End
         End
         Begin INFOHR_Controls.DateLookup dlpRtrnWork 
            Height          =   285
            Left            =   6360
            TabIndex        =   2
            Tag             =   "41-Date Returned to Work"
            Top             =   555
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin MSMask.MaskEdBox medRtrnWorkTime 
            Height          =   285
            Left            =   8595
            TabIndex        =   3
            Tag             =   "10-Time returned to work"
            Top             =   555
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
         Begin VB.Label lblQ1a 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "If so, give date commenced."
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
            Left            =   300
            TabIndex        =   88
            Top             =   595
            Width           =   2010
         End
         Begin VB.Label lblQ1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "1. Has the worker returned to work since the injury?"
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
            TabIndex        =   87
            Top             =   220
            Width           =   3645
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Commenced"
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
            Left            =   4920
            TabIndex        =   86
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
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
            Left            =   8160
            TabIndex        =   85
            Top             =   600
            Width           =   345
         End
      End
      Begin VB.Frame frQ2 
         Height          =   975
         Left            =   0
         TabIndex        =   72
         Tag             =   "v"
         Top             =   1080
         Width           =   11535
         Begin VB.TextBox txtAftLOffFAP 
            Appearance      =   0  'Flat
            DataField       =   "F9_AFT_LOFF_FAMPM"
            Height          =   285
            Left            =   11160
            MaxLength       =   5
            TabIndex        =   76
            Top             =   175
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.TextBox txtAftLOffTAP 
            Appearance      =   0  'Flat
            DataField       =   "F9_AFT_LOFF_TAMPM"
            Height          =   285
            Left            =   11160
            MaxLength       =   5
            TabIndex        =   75
            Top             =   550
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   8760
            TabIndex        =   74
            Top             =   190
            Width           =   1335
            Begin VB.OptionButton optAftLOffFAP 
               Caption         =   "PM"
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
               Left            =   720
               TabIndex        =   9
               Tag             =   "40-Last Worked Time, PM"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton optAftLOffFAP 
               Caption         =   "AM"
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
               Left            =   0
               TabIndex        =   8
               Tag             =   "40-Last Worked From Time, AM"
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   8760
            TabIndex        =   73
            Top             =   565
            Width           =   1335
            Begin VB.OptionButton optAftLOffTAP 
               Caption         =   "PM"
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
               Left            =   720
               TabIndex        =   13
               Tag             =   "40-Last Worked Time, PM"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton optAftLOffTAP 
               Caption         =   "AM"
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
               Left            =   0
               TabIndex        =   12
               Tag             =   "40-Last Worked To Time, AM"
               Top             =   0
               Width           =   615
            End
         End
         Begin INFOHR_Controls.DateLookup dlpAftLOffFDate 
            DataField       =   "F9_AFT_LOFF_FDATE"
            Height          =   285
            Left            =   5400
            TabIndex        =   6
            Tag             =   "41-From Period worker worked after first Layoff"
            Top             =   180
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin INFOHR_Controls.DateLookup dlpAftLOffTDate 
            DataField       =   "F9_AFT_LOFF_TDATE"
            Height          =   285
            Left            =   5400
            TabIndex        =   10
            Tag             =   "41-To Period worker worked after first Layoff"
            Top             =   550
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin MSMask.MaskEdBox medAftLOffFTime 
            DataField       =   "F9_AFT_LOFF_FTM"
            Height          =   285
            Left            =   7755
            TabIndex        =   7
            Tag             =   "10-From Time worker worked after first layoff"
            Top             =   175
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
         Begin MSMask.MaskEdBox medAftLOffTTime 
            DataField       =   "F9_AFT_LOFF_TTM"
            Height          =   285
            Left            =   7755
            TabIndex        =   11
            Tag             =   "10-To Time worker worked after first layoff"
            Top             =   550
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
         Begin VB.Label lblQ2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "2. If the worker worked after the first layoff, please enter dates."
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
            TabIndex        =   81
            Top             =   220
            Width           =   4410
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "From"
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
            Left            =   4920
            TabIndex        =   80
            Top             =   225
            Width           =   345
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "To"
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
            Left            =   4920
            TabIndex        =   79
            Top             =   595
            Width           =   195
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
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
            Left            =   7320
            TabIndex        =   78
            Top             =   225
            Width           =   345
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
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
            Left            =   7320
            TabIndex        =   77
            Top             =   595
            Width           =   345
         End
      End
      Begin VB.Frame frQ3 
         Height          =   615
         Left            =   0
         TabIndex        =   68
         Top             =   1920
         Width           =   11535
         Begin MSMask.MaskEdBox medTotShiftLost 
            DataField       =   "F9_TOT_SHIFT_LOST"
            Height          =   285
            Left            =   6960
            TabIndex        =   14
            Tag             =   "20-Total number of shifts lost"
            Top             =   175
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            MaxLength       =   4
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
         Begin MSMask.MaskEdBox medNoPayHrsShift 
            DataField       =   "F9_PAY_HRS_SHIFT"
            Height          =   285
            Left            =   10440
            TabIndex        =   15
            Tag             =   "20-Number of pay hours per shifts"
            Top             =   175
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            MaxLength       =   2
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
         Begin VB.Label lblQ3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "3. For Rotating Shift Workers Only, please complete the following:"
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
            TabIndex        =   71
            Top             =   220
            Width           =   4635
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total number of shifts lost:"
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
            Left            =   4920
            TabIndex        =   70
            Top             =   220
            Width           =   1845
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Number of pay hours per shift:"
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
            Left            =   8160
            TabIndex        =   69
            Top             =   220
            Width           =   2115
         End
      End
      Begin VB.Frame frQ4 
         Height          =   1095
         Left            =   0
         TabIndex        =   64
         Top             =   2400
         Width           =   11535
         Begin VB.TextBox txtWrkRtnOpinion 
            Appearance      =   0  'Flat
            DataField       =   "F9_WORKER_RTN_OPINION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   4920
            MaxLength       =   450
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Tag             =   "00-Your opinion on worker returning to work"
            Top             =   240
            Width           =   6525
         End
         Begin VB.Label lblQ4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "4. Did worker return as soon as able? (Give your opinion)"
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
            TabIndex        =   67
            Top             =   220
            Width           =   4005
         End
         Begin VB.Label lblQ4a 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "If not, give date and time you consider worker was able. "
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
            Left            =   300
            TabIndex        =   66
            Top             =   460
            Width           =   3990
         End
         Begin VB.Label lblQ4b 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "On what do you base your opinion?"
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
            Left            =   300
            TabIndex        =   65
            Top             =   700
            Width           =   2505
         End
      End
      Begin VB.Frame frQ5 
         Height          =   1815
         Left            =   0
         TabIndex        =   58
         Top             =   3360
         Width           =   11535
         Begin VB.TextBox txtKindofWork 
            Appearance      =   0  'Flat
            DataField       =   "F9_KIND_OF_WORK"
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
            Left            =   4920
            MaxLength       =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Tag             =   "00-Kind of work worker able to do"
            Top             =   240
            Width           =   6525
         End
         Begin VB.TextBox txtWhenFormerWork 
            Appearance      =   0  'Flat
            DataField       =   "F9_WHEN_FORMER_WORK"
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
            Left            =   4920
            MaxLength       =   70
            TabIndex        =   20
            Tag             =   "01-When will the worker be able to do former work"
            Top             =   1320
            Width           =   6435
         End
         Begin VB.TextBox txtServWorth 
            Appearance      =   0  'Flat
            DataField       =   "F9_SERV_WORTH"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4920
            MaxLength       =   76
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Tag             =   "00-Service Worth"
            Top             =   840
            Width           =   4125
         End
         Begin MSMask.MaskEdBox medServWorth 
            DataField       =   "F9_SERV_WORTH_PCT"
            Height          =   285
            Left            =   10650
            TabIndex        =   19
            Tag             =   "11-Service Worth in Percentage"
            Top             =   885
            Width           =   555
            _ExtentX        =   979
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label lblQ5b 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "When, if ever, will worker in your opinion be able to do former work?"
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
            Height          =   375
            Left            =   240
            TabIndex        =   63
            Top             =   1320
            Width           =   4305
         End
         Begin VB.Label lblQ5a 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "If only able to do other than former work, what do you consider services worth?"
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
            Height          =   435
            Left            =   240
            TabIndex        =   62
            Top             =   810
            Width           =   4485
         End
         Begin VB.Label lblQ5 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "5. If unable to do former work, what kind of work is worker doing or able to do?"
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
            Height          =   435
            Left            =   120
            TabIndex        =   61
            Top             =   220
            Width           =   4575
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Please express in terms of percentage"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   9120
            TabIndex        =   60
            Top             =   810
            Width           =   1485
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   11250
            TabIndex        =   59
            Top             =   930
            Width           =   120
         End
      End
      Begin VB.Frame frQ6 
         Height          =   900
         Left            =   0
         TabIndex        =   51
         Top             =   5040
         Width           =   11535
         Begin VB.Frame Frame9 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   6240
            TabIndex        =   53
            Top             =   600
            Width           =   1335
            Begin VB.OptionButton optEarningReducedYN 
               Caption         =   "No"
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
               Left            =   720
               TabIndex        =   23
               Tag             =   "40-Earning Reduced? No"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton optEarningReducedYN 
               Caption         =   "Yes"
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
               Left            =   0
               TabIndex        =   22
               Tag             =   "40-Earning Reduced? Yes"
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.TextBox txtEarningReducedYN 
            Appearance      =   0  'Flat
            DataField       =   "F9_REDUCED_YN"
            Height          =   195
            Left            =   11040
            MaxLength       =   5
            TabIndex        =   52
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin MSMask.MaskEdBox medAvgWeekEarn 
            DataField       =   "F9_AVG_GROSS_WEEKLY"
            Height          =   285
            Left            =   8880
            TabIndex        =   21
            Tag             =   "20-Average Weekly Gross Earnings"
            Top             =   180
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
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
            Format          =   "#,##0.0000;(#,##0.0000)"
            PromptChar      =   "_"
         End
         Begin VB.Label lblQ6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "6. Provide the worker's average gross weekly earnings since returning to work."
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
            TabIndex        =   57
            Top             =   220
            Width           =   5550
         End
         Begin VB.Label lblQ6a 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Are these earnings reduced in any way?"
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
            Left            =   300
            TabIndex        =   56
            Top             =   610
            Width           =   2835
         End
         Begin VB.Line Line1 
            X1              =   300
            X2              =   11445
            Y1              =   530
            Y2              =   530
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "$"
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
            Left            =   8640
            TabIndex        =   55
            Top             =   225
            Width           =   90
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Average weekly gross earnings"
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
            Left            =   6240
            TabIndex        =   54
            Top             =   225
            Width           =   2325
         End
      End
      Begin VB.Frame frQ7 
         Height          =   1335
         Left            =   0
         TabIndex        =   44
         Top             =   5880
         Width           =   11535
         Begin VB.TextBox txtInsurance 
            Appearance      =   0  'Flat
            DataField       =   "F9_INSURANCE"
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
            Left            =   4440
            MaxLength       =   50
            TabIndex        =   27
            Tag             =   "01-Name of Insurance Company"
            Top             =   900
            Width           =   6915
         End
         Begin MSMask.MaskEdBox medGrossTot 
            DataField       =   "F9_GROSS_TOT"
            Height          =   285
            Left            =   6240
            TabIndex        =   24
            Tag             =   "20-Gross Total Payment"
            Top             =   180
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin INFOHR_Controls.DateLookup dlpCoveredFrom 
            DataField       =   "F9_DATE_COV_FDATE"
            Height          =   285
            Left            =   7920
            TabIndex        =   25
            Tag             =   "41-Benefit/Payments Period From"
            Top             =   175
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin INFOHR_Controls.DateLookup dlpCoveredTo 
            DataField       =   "F9_DATE_COV_TDATE"
            Height          =   285
            Left            =   9840
            TabIndex        =   26
            Tag             =   "41-Benefit/Payments Period To"
            Top             =   175
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin VB.Label lblQ7 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   $"fehsF9.frx":540C
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
            Height          =   780
            Left            =   120
            TabIndex        =   50
            Top             =   220
            Width           =   3765
         End
         Begin VB.Line Line2 
            X1              =   4440
            X2              =   11420
            Y1              =   580
            Y2              =   580
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Name of Insurance Company, if applicable"
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
            Left            =   4440
            TabIndex        =   49
            Top             =   650
            Width           =   2985
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "$"
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
            TabIndex        =   48
            Top             =   220
            Width           =   90
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gross total payment"
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
            Left            =   4440
            TabIndex        =   47
            Top             =   220
            Width           =   1395
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "From"
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
            Left            =   7440
            TabIndex        =   46
            Top             =   220
            Width           =   345
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "To"
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
            Left            =   9600
            TabIndex        =   45
            Top             =   220
            Width           =   195
         End
      End
      Begin VB.Frame frQ8 
         Height          =   1095
         Left            =   0
         TabIndex        =   42
         Top             =   7080
         Width           =   11535
         Begin VB.TextBox txtRemarks 
            Appearance      =   0  'Flat
            DataField       =   "F9_COMMENTS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   4440
            MaxLength       =   300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Tag             =   "00-Any further information or remarks"
            Top             =   240
            Width           =   7005
         End
         Begin VB.Label lblQ8 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "8. Any further information or remarks."
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
            Height          =   435
            Left            =   120
            TabIndex        =   43
            Top             =   220
            Width           =   3690
         End
      End
      Begin VB.Label lblClaim 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Claim Number:"
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
         TabIndex        =   92
         Top             =   0
         Width           =   1020
      End
      Begin VB.Label lblClaimData 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         DataField       =   "F9_WCBNBR"
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
         Left            =   4080
         TabIndex        =   91
         Top             =   0
         Width           =   90
      End
      Begin VB.Label lblIncidentNo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         DataField       =   "F9_CASE"
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
         Left            =   1320
         TabIndex        =   90
         Top             =   0
         Width           =   90
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
         Left            =   0
         TabIndex        =   89
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "F9_EMPNBR"
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
      Left            =   11520
      TabIndex        =   37
      Top             =   1560
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "F9_COMPNO"
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
      Left            =   11520
      TabIndex        =   38
      Top             =   1320
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEHSF9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x%
Dim fGLBNew
Dim rsDATA As New ADODB.Recordset
Dim oRtnDate, oRtnTime, oRtnAP

Private Function chkHSForm9()
Dim SQLQ As String, Msg As String, dd&, tdat As Variant
Dim rsTemp As New ADODB.Recordset
Dim Part1, Part2

chkHSForm9 = False

On Error GoTo chkHSForm9_Err

'1
If Len(dlpRtrnWork.Text) > 0 Then
    If Not IsDate(dlpRtrnWork.Text) Then
        MsgBox "Date Commenced is not a valid date.", vbExclamation
        dlpRtrnWork.SetFocus
        Exit Function
    End If
End If

If Len(Trim(medRtrnWorkTime.Text)) > 0 And optRtrnWorkAP(0).Value = False And optRtrnWorkAP(1).Value = False Then
    MsgBox "One of the 'Time of Incident' time indicator selection for '1. Has the worker returned to work since the injury?' is required.", vbExclamation
    'optRtrnWorkAP.SetFocus
    'Exit Function
End If

If Len(medRtrnWorkTime.Text) = 5 Then
    Part1 = Left(medRtrnWorkTime, 2)
    Part2 = Right(medRtrnWorkTime, 2)
    If Not Left(Part1, 2) = "__" Or Not Right(Part2, 2) = "__" Then
        If Not IsNumeric(Part1) Or Not IsNumeric(Part2) Then
            MsgBox "Invalid 'Time' for '1. Has the worker returned to work since the injury?'.", vbExclamation
            medRtrnWorkTime.SetFocus
            Exit Function
        End If
        If CInt(Part1) > 12 Or CInt(Part2) > 59 Then
            MsgBox "Invalid 'Time' for '1. Has the worker returned to work since the injury?'.", vbExclamation
            medRtrnWorkTime.SetFocus
            Exit Function
        End If
    End If
ElseIf Len(medRtrnWorkTime.Text) <> 0 Then
    MsgBox "Invalid 'Time' for '1. Has the worker returned to work since the injury?'.", vbExclamation
    medRtrnWorkTime.SetFocus
    Exit Function
End If

'2
If Len(dlpAftLOffFDate.Text) > 0 Then
    If Not IsDate(dlpAftLOffFDate.Text) Then
        MsgBox "Worker worked after the first layoff 'From' date is not a valid date.", vbExclamation
        dlpAftLOffFDate.SetFocus
        Exit Function
    End If
End If

If Len(dlpAftLOffTDate.Text) > 0 Then
    If Not IsDate(dlpAftLOffTDate.Text) Then
        MsgBox "Worker worked after the first layoff 'To' date is not a valid date.", vbExclamation
        dlpAftLOffTDate.SetFocus
        Exit Function
    End If
End If

If IsDate(dlpAftLOffFDate.Text) And IsDate(dlpAftLOffTDate.Text) Then
    If CVDate(dlpAftLOffFDate.Text) > CVDate(dlpAftLOffTDate.Text) Then
        MsgBox "Worker worked after the first layoff 'From' Date cannot be greater than To Date.", vbExclamation
        dlpAftLOffFDate.SetFocus
        Exit Function
    End If
End If

If IsDate(dlpAftLOffFDate.Text) And Not IsDate(dlpAftLOffTDate.Text) Then
    MsgBox "To Date worker worked after the first layoff is required if From Date is entered.", vbExclamation
    dlpAftLOffTDate.SetFocus
    Exit Function
End If
If Not IsDate(dlpAftLOffFDate.Text) And IsDate(dlpAftLOffTDate.Text) Then
    MsgBox "From Date worker worked after the first layoff is required if To Date is entered.", vbExclamation
    dlpAftLOffFDate.SetFocus
    Exit Function
End If

If Len(medAftLOffFTime.Text) = 5 Then
    Part1 = Left(medAftLOffFTime, 2)
    Part2 = Right(medAftLOffFTime, 2)
    If Not Left(Part1, 2) = "__" Or Not Right(Part2, 2) = "__" Then
        If Not IsNumeric(Part1) Or Not IsNumeric(Part2) Then
            MsgBox "Invalid 'From' time for '2. If the worker worked after the first layoff, please enter dates.'.", vbExclamation
            medAftLOffFTime.SetFocus
            Exit Function
        End If
        If CInt(Part1) > 12 Or CInt(Part2) > 59 Then
            MsgBox "Invalid 'From' time for '2. If the worker worked after the first layoff, please enter dates.'.", vbExclamation
            medAftLOffFTime.SetFocus
            Exit Function
        End If
    End If
ElseIf Len(medAftLOffFTime.Text) <> 0 Then
    MsgBox "Invalid 'From' time for '2. If the worker worked after the first layoff, please enter dates.'.", vbExclamation
    medAftLOffFTime.SetFocus
    Exit Function
End If

If Len(medAftLOffTTime.Text) = 5 Then
    Part1 = Left(medAftLOffTTime, 2)
    Part2 = Right(medAftLOffTTime, 2)
    If Not Left(Part1, 2) = "__" Or Not Right(Part2, 2) = "__" Then
        If Not IsNumeric(Part1) Or Not IsNumeric(Part2) Then
            MsgBox "Invalid 'To' time for '2. If the worker worked after the first layoff, please enter dates.'.", vbExclamation
            medAftLOffTTime.SetFocus
            Exit Function
        End If
        If CInt(Part1) > 12 Or CInt(Part2) > 59 Then
            MsgBox "Invalid 'To' time for '2. If the worker worked after the first layoff, please enter dates.'.", vbExclamation
            medAftLOffTTime.SetFocus
            Exit Function
        End If
    End If
ElseIf Len(medAftLOffTTime.Text) <> 0 Then
    MsgBox "Invalid 'To' time for '2. If the worker worked after the first layoff, please enter dates.'.", vbExclamation
    medAftLOffTTime.SetFocus
    Exit Function
End If

'5
If Len(medServWorth.Text) > 0 Then
    If medServWorth.Text > 100 Then
        MsgBox "Service Worth in Percentage cannot be greater than 100%.", vbExclamation
        medServWorth.SetFocus
        Exit Function
    End If
End If

'7
If Len(dlpCoveredFrom.Text) > 0 Then
    If Not IsDate(dlpCoveredFrom.Text) Then
        MsgBox "Benefit/Payments Period From is not a valid date.", vbExclamation
        dlpCoveredFrom.SetFocus
        Exit Function
    End If
End If

If Len(dlpCoveredTo.Text) > 0 Then
    If Not IsDate(dlpCoveredTo.Text) Then
        MsgBox "Benefit/Payments Period To is not a valid date.", vbExclamation
        dlpCoveredTo.SetFocus
        Exit Function
    End If
End If

If IsDate(dlpCoveredFrom.Text) And IsDate(dlpCoveredTo.Text) Then
    If CVDate(dlpCoveredFrom.Text) > CVDate(dlpCoveredTo.Text) Then
        MsgBox "Benefit/Payments Period From Date cannot be greater than To Date.", vbExclamation
        dlpCoveredFrom.SetFocus
        Exit Function
    End If
End If

If IsDate(dlpCoveredFrom.Text) And Not IsDate(dlpCoveredTo.Text) Then
    MsgBox "To Date for Benefit/Payments Period is required if From Date is entered.", vbExclamation
    dlpCoveredTo.SetFocus
    Exit Function
End If
If Not IsDate(dlpCoveredFrom.Text) And IsDate(dlpCoveredTo.Text) Then
    MsgBox "From Date for Benefit/Payments Period is required if To Date is entered.", vbExclamation
    dlpCoveredFrom.SetFocus
    Exit Function
End If


chkHSForm9 = True

Exit Function

chkHSForm9_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSForm9", "HR_OHS_FORM9", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Sub cmdCancel_Click()
Dim bk

On Error GoTo Can_Err

fGLBNew = False

If Not (rsDATA.EOF And rsDATA.BOF) Then rsDATA.CancelUpdate

Call Display_Value

Data1.Refresh

'Call ST_UPD_MODE(True)  ' reset screen's attributes
'Call SET_UP_MODE

Exit Sub

Can_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_OHS_FORM9", "Cancel")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Sub cmdClose_Click()

Call NextForm
Unload Me

End Sub

'Sub cmdDelete_Click()
'Dim a As Integer, Msg As String, x%
'
'On Error GoTo Del_Err
'
'If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'    MsgBox "No Records Found"
'    Exit Sub
'End If
'
'Msg = "Are You Sure You Want To Delete "
'Msg = Msg & "This Record?"
'
'a% = MsgBox(Msg, 36, "Confirm Delete")
'If a% <> 6 Then Exit Sub
'
'If glbtermopen Then
'   gdbAdoIhr001X.BeginTrans
'   rsDATA.Delete
'   gdbAdoIhr001X.CommitTrans
'   Data1.Refresh
'Else
'   gdbAdoIhr001.BeginTrans
'   rsDATA.Delete
'   gdbAdoIhr001.CommitTrans
'   Data1.Refresh
'End If
'
'If Data1.Recordset.EOF And Data1.Recordset.BOF Then
'    Call Display_Value
'End If
'
'fglbNew = False
'
''Call ST_UPD_MODE(True)
'Call SET_UP_MODE
'
'Exit Sub
'
'Del_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HROHSCOS", "Delete")
'Screen.MousePointer = DEFAULT
'
'If gintRollBack% = False Then
'    Resume Next
'Else
'    Unload Me
'End If
'
'End Sub

Sub cmdModify_Click()
Dim x%

On Error GoTo Mod_Err

Call SET_UP_MODE

'Call ST_UPD_MODE(True)

'Store all values
oRtnDate = dlpRtrnWork.Text
oRtnTime = medRtrnWorkTime.Text
oRtnAP = txtRtrnWorkAP.Text

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_OHS_FORM9", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Sub cmdNew_Click()
'Dim SQLQ As String
'
'fglbNew = True
'
'Call SET_UP_MODE
'
'On Error GoTo AddN_Err
'
'Me.vbxTrueGrid.Enabled = False
'
'Call Set_Control("B", Me)
'
'If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
'
'wcb(1, 1) = ""
'
'lblCNum.Caption = "001"
'medShifts = 0
'
'Exit Sub
'
'AddN_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HROHSCOS", "Add")
'If gintRollBack% = False Then
'    Resume Next
'Else
'    Unload Me
'End If
'
'End Sub

Function cmdOK_Click()
Dim x%
Dim SQLQ As String
Dim xID

On Error GoTo Add_Err

cmdOK_Click = False

If Not chkHSForm9() Then Exit Function

lblCNum.Caption = "001"

rsDATA.Requery

If fGLBNew Then rsDATA.AddNew

'Update Incident screen with Return Date and Time
Call Update_Incident_Return_DateTime

Call UpdUStats(Me) ' update user's stats (who did it and when)

Call Set_Control("U", Me, rsDATA)

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    xID = Data1.Recordset!EC_ID
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    xID = Data1.Recordset!EC_ID
End If

Data1.Refresh
Data1.Recordset.Find "EC_ID=" & xID

If Not Data1.Recordset.EOF Then
    If glbtermopen Then
        SQLQ = "SELECT " & FldList1
        SQLQ = SQLQ & " FROM Term_OHS_FORM9"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " AND F9_CASE = " & Data1.Recordset!EC_CASE
        SQLQ = SQLQ & " AND F9_WCBNBR = '" & Data1.Recordset!EC_WCBNBR & "'"
    Else
        SQLQ = "SELECT " & FldList1
        SQLQ = SQLQ & " FROM HR_OHS_FORM9"
        SQLQ = SQLQ & " WHERE F9_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND F9_CASE = " & Data1.Recordset!EC_CASE
        SQLQ = SQLQ & " AND F9_WCBNBR = '" & Data1.Recordset!EC_WCBNBR & "'"
    End If
    SQLQ = SQLQ & " ORDER BY F9_CASE DESC"
    Data2.RecordSource = SQLQ
    Data2.Refresh
End If

cmdOK_Click = True

'Call Populate_Form9

Call SET_UP_MODE

fGLBNew = False

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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OHS_FORM9", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s WSIB Form 9"
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

RHeading = lblEEName & "'s WSIB Form 9"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError

If glbtermopen Then
    SQLQ = "SELECT " & FldList & " FROM Term_HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & "WHERE TERM_SEQ=" & glbTERM_Seq
    SQLQ = SQLQ & " AND EC_WCBNBR IS NOT NULL AND EC_WCBNBR <> ''"
Else
    SQLQ = "SELECT " & FldList & " FROM HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & "WHERE EC_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND EC_WCBNBR IS NOT NULL AND EC_WCBNBR <> ''"
End If
SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
Data1.RecordSource = SQLQ
Data1.Refresh

If Not Data1.Recordset.EOF Then
    'Add new record in Form 9 table if not already existing
    Call Create_Default_Form9
    
    'Retrieve Form 9 record for the selected Claimed Incident
    If glbtermopen Then
        SQLQ = "SELECT " & FldList1
        SQLQ = SQLQ & " FROM Term_OHS_FORM9"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " AND F9_CASE = " & Data1.Recordset!EC_CASE
        SQLQ = SQLQ & " AND F9_WCBNBR = '" & Data1.Recordset!EC_WCBNBR & "'"
    Else
        SQLQ = "SELECT " & FldList1
        SQLQ = SQLQ & " FROM HR_OHS_FORM9"
        SQLQ = SQLQ & " WHERE F9_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND F9_CASE = " & Data1.Recordset!EC_CASE
        SQLQ = SQLQ & " AND F9_WCBNBR = '" & Data1.Recordset!EC_WCBNBR & "'"
    End If
    SQLQ = SQLQ & " ORDER BY F9_CASE DESC"
    Data2.RecordSource = SQLQ
    Data2.Refresh
Else
    'Clear the values from the control if not Claimed Incident found for this employee
    Call Clear_Controls
End If

EERetrieve = True

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

Private Sub cmdPrintWF9_Test_Click()
'    Dim fieldValue, fValue
'    Dim objMyForm
'    Dim ErrorMessage
'    Dim nCount
'
'    Set objMyForm = CreateObject("CutePDF.Document")  'Create form object
'    objMyForm.Initialize ("FS21-010-94171023-00658222") 'Initialize object by serial number of the license
'
'    If objMyForm.openFile("U:\HR Systems VB6\Custom Features 7x\Health & Safety\Form 9\Form9revised.pdf") = False Then  'Open a PDF file's form
'        ErrorMessage = objMyForm.GetLastError()
'    End If
'
''    nCount = objMyForm.numberOfFields                      'Get total number of fields in the form
''    For X = 1 To nCount Step 1
''        fieldValue = objMyForm.getFieldName(X)                'Get field name first
''        fValue = objMyForm.getFieldValue(X)                   'Then get field value by name
''        Debug.Print fieldValue & " - " & fValue
''    Next
'
'    'For X = 0 To 331
'    '    fieldValue = objMyForm.getFieldName(X)
'    '    Debug.Print fieldValue
'    'Next
'    fieldValue = objMyForm.getFieldValue("CHECK1")
'    fieldValue = objMyForm.getFieldValue("CHECK2")
'    fieldValue = objMyForm.getFieldValue("CHECK3")
'    fieldValue = objMyForm.getFieldValue("RE")
End Sub

Private Sub cmdPrintWF9_Click()
    'This will generate the WSIB Form 9
    
    'Save the Form 9 data first
    If Not cmdOK_Click() Then Exit Sub

    'Call CutePDF_Test
    Call Generate_WSIB_Form9
End Sub

Private Sub Generate_WSIB_Form9()
    Dim rsEmp As New ADODB.Recordset
    Dim rsForm9 As New ADODB.Recordset
    Dim SQLQ As String
    Dim xPathToSaveIn As String
    Dim xFileName As String
    Dim xFileExtension As String
    
    Dim objMyForm
    Dim ErrorMessage
    Dim nCount, nI
    Dim nReturn
    
    Screen.MousePointer = HOURGLASS

    Set objMyForm = CreateObject("CutePDF.Document")  'Create form object
    objMyForm.Initialize ("FS21-010-94171023-00658222") 'Initialize object by serial number of the license
    
    'Open an encrypted PDF form file from an URL with password 'cutepdf'
    'If objMyForm.openFile("ftp://www.ftpsite.com/Encrypted_Form.pdf", "cutepdf") = False Then
    'Open the Form 9 template
    If objMyForm.openFile(glbIHRREPORTS & "Form9.pdf") = False Then
        ErrorMessage = objMyForm.GetLastError()
        Screen.MousePointer = DEFAULT
        MsgBox "Cannot find the template for Form 9. WSIB Form 9 cannot be generated.", vbCritical, "info:HR - Missing Form 9 Template"
        Exit Sub
    End If
    
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME, ED_ADDR1, ED_CITY, ED_PROV, ED_PCODE, "
    SQLQ = SQLQ & " ED_DOB, ED_SIN"
    SQLQ = SQLQ & " FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsEmp.EOF Then
        'Employee missing in the employee file - cannot generate the WSIB Form 9
        MsgBox "Employee data missing cannot generate the WSIB Form 9.", vbExclamation, "info:HR - Form 9 Generation"
        rsEmp.Close
        Set rsEmp = Nothing
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HR_OHS_FORM9"
    SQLQ = SQLQ & " WHERE F9_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND F9_CASE =" & Data1.Recordset!EC_CASE
    SQLQ = SQLQ & " AND F9_WCBNBR =" & Data1.Recordset!EC_WCBNBR
    rsForm9.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsForm9.EOF Then
        'Employee Health & Safety data missing - cannot generate the WSIB Form 7
        MsgBox "Employee Health & Safety data in the Form 9 is missing, cannot generate the WSIB Form 9.", vbExclamation, "info:HR - Form 9 Generation"
        rsForm9.Close
        Set rsForm9 = Nothing
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
    
    
    MDIMain.panHelp(1).Caption = "Please wait, generating Form 9...."
    MDIMain.panHelp(0).FloodType = 1
       
    MDIMain.panHelp(0).FloodPercent = 5

    'Start assigning values to Form 9
    
    'Claim #
    If Not IsNull(rsForm9("F9_CASE")) Then
        nReturn = objMyForm.setField("claim number", rsForm9("F9_CASE"))
    End If
    
    'Employee Information
    nReturn = objMyForm.setField("Last name", rsEmp("ED_SURNAME"))
    nReturn = objMyForm.setField("First Name", rsEmp("ED_FNAME"))
    nReturn = objMyForm.setField("ADDRESS", rsEmp("ED_ADDR1"))
    nReturn = objMyForm.setField("CITY/TOWN", rsEmp("ED_CITY"))
    nReturn = objMyForm.setField("PROVINCE", rsEmp("ED_PROV"))
    nReturn = objMyForm.setField("POSTAL  CODE", Mid(Replace(rsEmp("ED_PCODE"), " ", ""), 1, 3) & " " & Mid(Replace(rsEmp("ED_PCODE"), " ", ""), 4, 3))
    
    MDIMain.panHelp(0).FloodPercent = 10
    
    'Date of Injury
    nReturn = objMyForm.setField("DID", Day(rsForm9("F9_OCCDATE")))
    nReturn = objMyForm.setField("DIM", month(rsForm9("F9_OCCDATE")))
    nReturn = objMyForm.setField("DIY", Right(Year(rsForm9("F9_OCCDATE")), 2))
    
    'SIN #
    nReturn = objMyForm.setField("SIN1", Mid(rsEmp("ED_SIN"), 1, 3))
    nReturn = objMyForm.setField("SIN2", Mid(rsEmp("ED_SIN"), 4, 3))
    nReturn = objMyForm.setField("SIN3", Mid(rsEmp("ED_SIN"), 7, 3))
    
    'Date of Birth
    nReturn = objMyForm.setField("DBD", Day(rsEmp("ED_DOB")))
    nReturn = objMyForm.setField("DBM", month(rsEmp("ED_DOB")))
    nReturn = objMyForm.setField("DBY", Right(Year(rsEmp("ED_DOB")), 2))
    
    MDIMain.panHelp(0).FloodPercent = 20
    
    '1. Has the worker returned to work since the injury?
    If Not IsNull(rsForm9("F9_RETURNED_WORK")) Then
        nReturn = objMyForm.setField("Return to work", rsForm9("F9_RETURNED_WORK"))
    End If
    
    '1. If so, give date commenced
    If Not IsNull(Data1.Recordset!EC_RETURN) Then
        nReturn = objMyForm.setField("DCD", Day(Data1.Recordset!EC_RETURN))
        nReturn = objMyForm.setField("DCM", month(Data1.Recordset!EC_RETURN))
        nReturn = objMyForm.setField("DCY", Right(Year(Data1.Recordset!EC_RETURN), 2))
    End If
    
    'Time
    If Not IsNull(Data1.Recordset!EC_RETURN_TM) Then
        nReturn = objMyForm.setField("TIME1", Data1.Recordset!EC_RETURN_TM)
    End If
    
    'Time AM/PM
    If Not IsNull(Data1.Recordset!EC_RETURN_TM) Then
        If Not IsNull(Data1.Recordset!EC_RETURN_AMPM) Then
            If Data1.Recordset!EC_RETURN_AMPM = "A" Then
                nReturn = objMyForm.setField("CHECK1", "A.M.")
            ElseIf Data1.Recordset!EC_RETURN_AMPM = "P" Then
                nReturn = objMyForm.setField("CHECK1", "P.M.")
            End If
        End If
    End If
    
    
    '2. If the worker worked after the first layoff, please enter dates.
    'From Date
    If Not IsNull(rsForm9("F9_AFT_LOFF_FDATE")) Then
        nReturn = objMyForm.setField("FDLD", Day(rsForm9("F9_AFT_LOFF_FDATE")))
        nReturn = objMyForm.setField("FDLM", month(rsForm9("F9_AFT_LOFF_FDATE")))
        nReturn = objMyForm.setField("FDLY", Right(Year(rsForm9("F9_AFT_LOFF_FDATE")), 2))
    End If
    
    'From Time
    If Not IsNull(rsForm9("F9_AFT_LOFF_FTM")) Then
        nReturn = objMyForm.setField("TIME2", rsForm9("F9_AFT_LOFF_FTM"))
    End If
    
    'From Time AM/PM
    If Not IsNull(rsForm9("F9_AFT_LOFF_FTM")) Then
        If Not IsNull(rsForm9("F9_AFT_LOFF_FAMPM")) Then
            If rsForm9("F9_AFT_LOFF_FAMPM") = "A" Then
                nReturn = objMyForm.setField("CHECK2", "A.M.")
            ElseIf rsForm9("F9_AFT_LOFF_FAMPM") = "P" Then
                nReturn = objMyForm.setField("CHECK2", "P.M.")
            End If
        End If
    End If
    
    'To Date
    If Not IsNull(rsForm9("F9_AFT_LOFF_TDATE")) Then
        nReturn = objMyForm.setField("TDLD", Day(rsForm9("F9_AFT_LOFF_TDATE")))
        nReturn = objMyForm.setField("TDLM", month(rsForm9("F9_AFT_LOFF_TDATE")))
        nReturn = objMyForm.setField("TDLY", Right(Year(rsForm9("F9_AFT_LOFF_TDATE")), 2))
    End If
    
    'To Time
    If Not IsNull(rsForm9("F9_AFT_LOFF_TTM")) Then
        nReturn = objMyForm.setField("TIME3", rsForm9("F9_AFT_LOFF_TTM"))
    End If
    
    'To Time AM/PM
    If Not IsNull(rsForm9("F9_AFT_LOFF_TTM")) Then
        If Not IsNull(rsForm9("F9_AFT_LOFF_TAMPM")) Then
            If rsForm9("F9_AFT_LOFF_TAMPM") = "A" Then
                nReturn = objMyForm.setField("CHECK3", "A.M.")
            ElseIf rsForm9("F9_AFT_LOFF_TAMPM") = "P" Then
                nReturn = objMyForm.setField("CHECK3", "P.M.")
            End If
        End If
    End If
    
    MDIMain.panHelp(0).FloodPercent = 40
    
    '3. For Rotating Shift Workers Only, please complete the following:
    'Total number of shifts lost:
    If Not IsNull(rsForm9("F9_TOT_SHIFT_LOST")) Then
        nReturn = objMyForm.setField("SHIFTLOST", rsForm9("F9_TOT_SHIFT_LOST"))
    End If

    'Number of pay hours per shift:
    If Not IsNull(rsForm9("F9_PAY_HRS_SHIFT")) Then
        nReturn = objMyForm.setField("payhrs", rsForm9("F9_PAY_HRS_SHIFT"))
    End If

    
    '4. Did worker return as soon as able? (Give your opinion) If not, give date and time you consider
    'worker was able. On what do you base your opinion?
    If Not IsNull(rsForm9("F9_WORKER_RTN_OPINION")) Then
        nReturn = objMyForm.setField("worker return", rsForm9("F9_WORKER_RTN_OPINION"))
    End If
        
        
    '5. If unable to do former work, what kind of work is worker doing or able to do?
    If Not IsNull(rsForm9("F9_KIND_OF_WORK")) Then
        nReturn = objMyForm.setField("kind of work", rsForm9("F9_KIND_OF_WORK"))
    End If
        
    'If only able to do other than former work what do you consider services worth?
    If Not IsNull(rsForm9("F9_SERV_WORTH")) Then
        nReturn = objMyForm.setField("services worth", rsForm9("F9_SERV_WORTH"))
    End If
    If Not IsNull(rsForm9("F9_SERV_WORTH_PCT")) Then
        nReturn = objMyForm.setField("percent", rsForm9("F9_SERV_WORTH_PCT"))
    End If
        
    'When, if ever, will worker in your opinion be able to do former work?
    If Not IsNull(rsForm9("F9_WHEN_FORMER_WORK")) Then
        nReturn = objMyForm.setField("former work", rsForm9("F9_WHEN_FORMER_WORK"))
    End If
    
    
    MDIMain.panHelp(0).FloodPercent = 60
    
    '6. Provide the worker's average gross weekly earnings since returning to work.
    If Not IsNull(rsForm9("F9_AVG_GROSS_WEEKLY")) Then
        nReturn = objMyForm.setField("Ave. gross earn", rsForm9("F9_AVG_GROSS_WEEKLY"))
    End If
    
    'Are these earnings reduced in any way?
    If Not IsNull(rsForm9("F9_REDUCED_YN")) Then
        If rsForm9("F9_REDUCED_YN") <> 0 Then
            nReturn = objMyForm.setField("RE", "yes")
        Else
            nReturn = objMyForm.setField("RE", "no")
        End If
    End If
        
    
    '7. If the worker received any benefits or payments from your company or any other
    'insurance plan for the period of disablement please provide the following
    'Gross total Payment
    If Not IsNull(rsForm9("F9_GROSS_TOT")) Then
        nReturn = objMyForm.setField("gross total", rsForm9("F9_GROSS_TOT"))
    End If
    
    'From
    If Not IsNull(rsForm9("F9_DATE_COV_FDATE")) Then
        nReturn = objMyForm.setField("DCDF", Day(rsForm9("F9_DATE_COV_FDATE")))
        nReturn = objMyForm.setField("DCMF", month(rsForm9("F9_DATE_COV_FDATE")))
        nReturn = objMyForm.setField("DCYF", Right(Year(rsForm9("F9_DATE_COV_FDATE")), 2))
    End If
    
    'To
    If Not IsNull(rsForm9("F9_DATE_COV_TDATE")) Then
        nReturn = objMyForm.setField("DCDT", Day(rsForm9("F9_DATE_COV_TDATE")))
        nReturn = objMyForm.setField("DCMT", month(rsForm9("F9_DATE_COV_TDATE")))
        nReturn = objMyForm.setField("DCYT", Right(Year(rsForm9("F9_DATE_COV_TDATE")), 2))
    End If
    
    'Name of Insurance
    If Not IsNull(rsForm9("F9_INSURANCE")) Then
        nReturn = objMyForm.setField("Name of Insurance Co", rsForm9("F9_INSURANCE"))
    End If
            
        
    '8 Any further information or remarks.
    If Not IsNull(rsForm9("F9_COMMENTS")) Then
        nReturn = objMyForm.setField("FURTHER INFORMATION", rsForm9("F9_COMMENTS"))
    End If
    
    MDIMain.panHelp(0).FloodPercent = 70
    
    'Get the location to save the file in
    xPathToSaveIn = GetComPreferEmail("WSIBFORM7PATH")
    If Len(xPathToSaveIn) = 0 Then
        xPathToSaveIn = glbIHRREPORTS
    End If
    
    If Right(xPathToSaveIn, 1) <> "\" Then xPathToSaveIn = xPathToSaveIn & "\"
        
    'Set the attached document keys
    If Not IsNull(rsForm9("F9_DOCKEY")) Then
        glbDocKey = rsForm9("F9_DOCKEY")
    Else
        glbDocKey = 0
    End If
    glbJob = rsForm9("F9_CASE")
    
    'Save completed form file into a new PDF file
    objMyForm.saveFile (xPathToSaveIn & glbLEE_ID & "_WSIBForm9.pdf")
        
    MDIMain.panHelp(0).FloodPercent = 80
    
    'If above is successfull, i.e. nReturn = 1, then save the completed Form 9 pdf into the Incident
    'Document Attachment table in the infoHR_DOC database as part of other incident documents.
    If nReturn = 1 Then
        'Save the document in the Incident Attachment
        glbDocName = "INCIDENT"
        xFileName = xPathToSaveIn & glbLEE_ID & "_WSIBForm9.pdf"
        xFileExtension = GetFileExtension(xFileName)
    
        'Create Incident Attachment Record to get the DOCNO - glbDocTmp
        'Add in HRDOC_HEALTH_SAFETY/TERM_HRDOC_HEALTH_SAFETY
        glbDocTmp = Add_Incident_Attachment_Record(glbJob, rsForm9("F9_OCCDATE"))
        
        'Add in HRDOC_HEALTH_SAFETY2/TERM_HRDOC_HEALTH_SAFETY2
        'Add teh FRM9 code first if not existing
        Call CheckHRTABLCode("DOCT", "FRM9", "Form 9")
        'Release 8.0 - Grant permission to this Form 9 Document Type code for this user as well so the user can see the
        'Document of this Document Type
        Call Grant_DocumentTypeCode_Security(glbUserID, "FRM9", "Form 9")
        
        Call AppendIncident(glbLEE_ID, xFileName, xFileExtension, "FRM9", "FRM9 - " & Format(Now, "mm/dd/yyyy"))
    
        MDIMain.panHelp(0).FloodPercent = 95
    
        'Delete the WSIB Form 9 file now that as it has been saved into the Incident Attachment record
        If (Dir(xFileName)) <> "" Then
            'Call FileAttributeForm7(xFileName, "-r", xPathToSaveIn)
            'Call Pause(5)
            Kill xFileName
        End If
        
        'Close all the recordsets
        rsForm9.Close
        Set rsForm9 = Nothing
        
        rsEmp.Close
        Set rsEmp = Nothing
    
        MDIMain.panHelp(0).FloodPercent = 100
        
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = ""
        Screen.MousePointer = DEFAULT
        
        MsgBox "WSIB Form 9 generated successfully." & vbCrLf & vbCrLf & "Please go to the Incident Documents screen to view or print Form 9.", vbInformation, "WSIB - Form 9 Generation"
    Else
        'Close all the recordsets
        rsForm9.Close
        Set rsForm9 = Nothing
        
        rsEmp.Close
        Set rsEmp = Nothing
        
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = ""
        Screen.MousePointer = DEFAULT
        
        MsgBox "WSIB Form 9 cannot be generated.", vbCritical, "WSIB - Form 9 Generation"
    End If
    
End Sub

Private Sub cmdPrintWF9_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub dlpAftLOffFDate_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub dlpAftLOffTDate_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub dlpCoveredFrom_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub dlpCoveredTo_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub dlpRtrnWork_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
glbOnTop = "frmEHSF9"
End Sub

Private Sub Form_GotFocus()
glbOnTop = "frmEHSF9"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

'ReDim wcb(1, 3) 'laura nov 14, 1997

glbOnTop = "frmEHSF9"

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
    Data2.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
    Data2.ConnectionString = glbAdoIHRDB
End If

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

If Len(glbLEE_SName) < 1 Then Exit Sub

Screen.MousePointer = HOURGLASS

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "WSIB Cost - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = ShowEmpnbr(lblEEID)

'Call ST_UPD_MODE(True)

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

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= 12000 Then
        scrControl.Value = 0
        ScrFrame.Top = 2700
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        If Me.Height < 9100 Then
            scrControl.Max = 7000
        Else
            scrControl.Max = 2900
        End If
        scrControl.Left = Me.Width - scrControl.Width - 250
        If Me.Height - scrControl.Top - 780 > 0 Then
            scrControl.Height = Me.Height - scrControl.Top - 780
        End If
    End If
    
'    'Horizontal Scroll
'    scrHScroll.Width = Me.Width - 120
'    'ScrFrame.Height = Me.ScaleHeight - (scrHScroll.Height + 200)
'    If Me.Width >= 10750 Then
'        scrHScroll.Value = 0
'        scrHScroll.Visible = False
'    Else
'        scrHScroll.Visible = True
'        If Me.Width < 7500 Then
'            scrHScroll.Max = 200
'        Else
'            scrHScroll.Max = 30
'        End If
'        scrHScroll.Top = Me.Height - 800
'        scrHScroll.Width = Me.Width - 120
'    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

MDIMain.panHelp(0).Caption = "Select function from the menu."
Set frmEHSF9 = Nothing ' carmen may 00
Call NextForm
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

txtReturnedWork.Enabled = TF
dlpRtrnWork.Enabled = TF
medRtrnWorkTime.Enabled = TF
Frame10.Enabled = TF
dlpAftLOffFDate.Enabled = TF
medAftLOffFTime.Enabled = TF
Frame1.Enabled = TF
dlpAftLOffTDate.Enabled = TF
medAftLOffTTime.Enabled = TF
Frame3.Enabled = TF
medTotShiftLost.Enabled = TF
medNoPayHrsShift.Enabled = TF
txtWrkRtnOpinion.Enabled = TF
txtKindofWork.Enabled = TF
txtServWorth.Enabled = TF
medServWorth.Enabled = TF
txtWhenFormerWork.Enabled = TF
medAvgWeekEarn.Enabled = TF
Frame9.Enabled = TF
medGrossTot.Enabled = TF
dlpCoveredFrom.Enabled = TF
dlpCoveredTo.Enabled = TF
txtInsurance.Enabled = TF
txtRemarks.Enabled = TF

cmdPrintWF9.Enabled = TF

End Sub

Private Sub medAftLOffFTime_GotFocus()
Call SetPanHelp(ActiveControl)
medAftLOffFTime.Mask = "##:##"
End Sub

Private Sub medAftLOffFTime_LostFocus()
If medAftLOffFTime.Text = "__:__" Then
    medAftLOffFTime.Mask = ""
    medAftLOffFTime.Text = ""
Else
    medAftLOffFTime.Mask = "##:##"
End If
End Sub

Private Sub medAftLOffTTime_GotFocus()
Call SetPanHelp(ActiveControl)
medAftLOffTTime.Mask = "##:##"
End Sub

Private Sub medAftLOffTTime_LostFocus()
If medAftLOffTTime.Text = "__:__" Then
    medAftLOffTTime.Mask = ""
    medAftLOffTTime.Text = ""
Else
    medAftLOffTTime.Mask = "##:##"
End If
End Sub

Private Sub medAvgWeekEarn_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medGrossTot_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medNoPayHrsShift_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medRtrnWorkTime_GotFocus()
Call SetPanHelp(ActiveControl)
medRtrnWorkTime.Mask = "##:##"
End Sub

Private Sub medRtrnWorkTime_LostFocus()
If medRtrnWorkTime.Text = "__:__" Then
    medRtrnWorkTime.Mask = ""
    medRtrnWorkTime.Text = ""
Else
    medRtrnWorkTime.Mask = "##:##"
End If
End Sub

Private Sub medServWorth_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medTotShiftLost_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optAftLOffFAP_Click(Index As Integer)
    If optAftLOffFAP(0).Value = True Then
        txtAftLOffFAP.Text = "A"
    ElseIf optAftLOffFAP(1).Value = True Then
        txtAftLOffFAP.Text = "P"
    Else
        txtAftLOffFAP.Text = ""
    End If
End Sub

Private Sub optAftLOffFAP_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optAftLOffTAP_Click(Index As Integer)
    If optAftLOffTAP(0).Value = True Then
        txtAftLOffTAP.Text = "A"
    ElseIf optAftLOffTAP(1).Value = True Then
        txtAftLOffTAP.Text = "P"
    Else
        txtAftLOffTAP.Text = ""
    End If
End Sub

Private Sub optAftLOffTAP_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optEarningReducedYN_Click(Index As Integer)
    If optEarningReducedYN(0).Value = True Then
        txtEarningReducedYN.Text = "1"
    ElseIf optEarningReducedYN(1).Value Then
        txtEarningReducedYN.Text = "0"
    Else
        txtEarningReducedYN.Text = ""
    End If
End Sub

Private Sub optEarningReducedYN_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optRtrnWorkAP_Click(Index As Integer)
    If optRtrnWorkAP(0).Value = True Then
        txtRtrnWorkAP.Text = "A"
    ElseIf optRtrnWorkAP(1).Value = True Then
        txtRtrnWorkAP.Text = "P"
    Else
        txtRtrnWorkAP.Text = ""
    End If
End Sub

Private Sub optRtrnWorkAP_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub scrControl_Change()
ScrFrame.Top = 2700 - scrControl.Value
End Sub

Private Sub txtAftLOffFAP_Change()
    If txtAftLOffFAP.Text = "A" Then
        optAftLOffFAP(0).Value = True
    ElseIf txtAftLOffFAP.Text = "P" Then
        optAftLOffFAP(1).Value = True
    Else
        optAftLOffFAP(0).Value = False
        optAftLOffFAP(1).Value = False
    End If
End Sub

Private Sub txtAftLOffTAP_Change()
    If txtAftLOffTAP.Text = "A" Then
        optAftLOffTAP(0).Value = True
    ElseIf txtAftLOffTAP.Text = "P" Then
        optAftLOffTAP(1).Value = True
    Else
        optAftLOffTAP(0).Value = False
        optAftLOffTAP(1).Value = False
    End If
End Sub

Private Sub txtEarningReducedYN_Change()
    If txtEarningReducedYN.Text = "" Then
        optEarningReducedYN(0).Value = False
        optEarningReducedYN(1).Value = False
    ElseIf txtEarningReducedYN.Text <> "0" Then
        optEarningReducedYN(0).Value = True
    Else
        optEarningReducedYN(1).Value = True
    End If
End Sub

Private Sub txtInsurance_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtKindofWork_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtRemarks_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtReturnedWork_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtRtrnWorkAP_Change()
    If txtRtrnWorkAP.Text = "A" Then
        optRtrnWorkAP(0).Value = True
    ElseIf txtRtrnWorkAP.Text = "P" Then
        optRtrnWorkAP(1).Value = True
    Else
        optRtrnWorkAP(0).Value = False
        optRtrnWorkAP(1).Value = False
    End If
End Sub

Private Sub txtServWorth_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtWhenFormerWork_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtWrkRtnOpinion_GotFocus()
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
    
    If glbtermopen Then
        SQLQ = "SELECT " & FldList & " FROM Term_HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
        SQLQ = SQLQ & " AND EC_WCBNBR IS NOT NULL AND EC_WCBNBR <> ''"
    Else
        SQLQ = "SELECT " & FldList & " FROM HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & " WHERE EC_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND EC_WCBNBR IS NOT NULL AND EC_WCBNBR <> ''"
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

    Data1.RecordSource = SQLQ
    Data1.Refresh
    
    If Not Data1.Recordset.EOF Then
        'Create default Form 9 record for the claimed incident selected
        Call Create_Default_Form9
        
        'Retrieve the Form 9 record for the selected claimed incident
        If glbtermopen Then
            SQLQ = "SELECT " & FldList1
            SQLQ = SQLQ & " FROM Term_OHS_FORM9"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
            SQLQ = SQLQ & " AND F9_CASE = " & Data1.Recordset!EC_CASE
            SQLQ = SQLQ & " AND F9_WCBNBR = '" & Data1.Recordset!EC_WCBNBR & "'"
        Else
            SQLQ = "SELECT " & FldList1
            SQLQ = SQLQ & " FROM HR_OHS_FORM9"
            SQLQ = SQLQ & " WHERE F9_EMPNBR = " & glbLEE_ID
            SQLQ = SQLQ & " AND F9_CASE = " & Data1.Recordset!EC_CASE
            SQLQ = SQLQ & " AND F9_WCBNBR = '" & Data1.Recordset!EC_WCBNBR & "'"
        End If
        SQLQ = SQLQ & " ORDER BY F9_CASE DESC"
        Data2.RecordSource = SQLQ
        Data2.Refresh
    End If
        
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim SQLQ As String

On Error GoTo Tab1_Err

'Display the selected record on the form controls
Call Display_Value

Exit Sub

Tab1_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_OHS_FORM9", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
    Dim SQLQ
    
    If Not Data1.Recordset.EOF Then
        'For the selected claimed incident create the default Form 9 record if not already existing
        Call Create_Default_Form9
        
        If Data2.Recordset.EOF Or Data2.Recordset.BOF Then
            Call Set_Control("B", Me)
            If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
            If glbtermopen Then
                rsDATA.Open Data2.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
            Else
                rsDATA.Open Data2.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            End If
            Call SET_UP_MODE
            'Me.cmdModify_Click
            Exit Sub
        End If
    
        'Retrieve the Form 9 record for the selected Claimed Incident
        If glbtermopen Then
            SQLQ = "SELECT " & FldList1 & " FROM Term_OHS_FORM9 "
            SQLQ = SQLQ & " WHERE F9_CASE=" & Data1.Recordset!EC_CASE
            SQLQ = SQLQ & " AND F9_WCBNBR='" & Data1.Recordset!EC_WCBNBR & "'"
            SQLQ = SQLQ & " AND F9_EMPNBR =" & glbTERM_ID
            SQLQ = SQLQ & " AND TERM_SEQ=" & glbTERM_Seq
            If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
            rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            SQLQ = "SELECT " & FldList1 & " FROM HR_OHS_FORM9 "
            SQLQ = SQLQ & "WHERE F9_CASE = " & Data1.Recordset!EC_CASE
            SQLQ = SQLQ & " AND F9_WCBNBR='" & Data1.Recordset!EC_WCBNBR & "'"
            SQLQ = SQLQ & " AND F9_EMPNBR =" & glbLEE_ID
            If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
            rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        SQLQ = SQLQ & " ORDER BY F9_CASE DESC"
    Else
        'Refresh the Form 9 record source to make it empty as there are not claimed incidents so not Form 9
        'records.
        If glbtermopen Then
            Data2.RecordSource = "SELECT " & FldList1 & " FROM Term_OHS_FORM9 WHERE 1=2"
        Else
            Data2.RecordSource = "SELECT " & FldList1 & " FROM HR_OHS_FORM9 WHERE 1=2"
        End If
        
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbtermopen Then
            rsDATA.Open Data2.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDATA.Open Data2.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        
        'No Claimed Incident record found
        Call Set_Control("B", Me)
        
        'Clear the form control values
        'Call Clear_Controls
        
        Call SET_UP_MODE
        'Me.cmdModify_Click
        Exit Sub
        
'        If glbtermopen Then
'            SQLQ = "SELECT " & FldList1 & " FROM Term_OHS_FORM9 "
'            SQLQ = SQLQ & " WHERE "
'            SQLQ = SQLQ & " F9_EMPNBR =" & glbTERM_ID
'            SQLQ = SQLQ & " AND TERM_SEQ=" & glbTERM_Seq
'            If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
'            rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
'        Else
'            SQLQ = "SELECT " & FldList1 & " FROM HR_OHS_FORM9 "
'            SQLQ = SQLQ & "WHERE "
'            SQLQ = SQLQ & " F9_EMPNBR =" & glbLEE_ID
'            If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
'            rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'        End If
'        SQLQ = SQLQ & " ORDER BY F9_CASE DESC"
            
    End If
    
    Data2.RecordSource = SQLQ
    Data2.Refresh
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    
    Call Set_Control("R", Me, rsDATA)
    
    'Retrieve Return Date information
    If Not Data1.Recordset.EOF Then
        dlpRtrnWork.Text = IIf(Not IsNull(Data1.Recordset!EC_RETURN), Data1.Recordset!EC_RETURN, "")
        
        If Not IsNull(Data1.Recordset!EC_RETURN_TM) And Data1.Recordset!EC_RETURN_TM <> "" Then
            medRtrnWorkTime.Text = Data1.Recordset!EC_RETURN_TM
        Else
            medRtrnWorkTime.Mask = ""
            medRtrnWorkTime.Text = ""
        End If
        
        txtRtrnWorkAP.Text = IIf(Not IsNull(Data1.Recordset!EC_RETURN_AMPM), Data1.Recordset!EC_RETURN_AMPM, "")
    Else
        Call Clear_Controls
    End If
    
    'Populate Form 9
    'Call Populate_Form9
    
    Call SET_UP_MODE
    
    Me.cmdModify_Click

End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fGLBNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fGLBNew = True
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_HSWF9 And glbWSIBModule
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

If fGLBNew Then
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

Me.vbxTrueGrid.Enabled = True

Call ST_UPD_MODE(TF)

End Sub

Private Sub lblEEID_Change()

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    frmEHSF9.Caption = "WSIB Form 9 - " & Left$(glbLEE_SName, 5)
    frmEHSF9.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If

If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
'lblEEID = glbLEE_ID
lblEENum = ShowEmpnbr(lblEEID)

End Sub

Private Function FldList()
Dim SQLQ

SQLQ = ""
SQLQ = SQLQ & "EC_ID, EC_EMPNBR, EC_CASE, EC_OCCDATE, EC_CODE, EC_COMPNO,EC_WCBNBR,EC_WCBFDTE,EC_RETURN_TM,EC_RETURN_AMPM,EC_RETURN"

If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"

FldList = SQLQ

End Function

Private Function FldList1()
Dim SQLQ
SQLQ = ""
SQLQ = SQLQ & "* "

If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"
FldList1 = SQLQ
End Function

Private Sub Create_Default_Form9()
    Dim rsDATA1 As New ADODB.Recordset
    Dim SQLQ As String
    Dim xSal As Double

    'Form 9
    If Not Data1.Recordset.EOF Then
        If rsDATA1.State <> 0 Then: If rsDATA1.EOF Then rsDATA1.Close Else If rsDATA1.EditMode = adEditAdd Then rsDATA1.CancelUpdate: rsDATA1.Close Else rsDATA1.Close
        If glbtermopen Then
            SQLQ = "SELECT " & FldList1
            SQLQ = SQLQ & " FROM Term_OHS_FORM9"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
            SQLQ = SQLQ & " AND F9_CASE = " & Data1.Recordset!EC_CASE
            SQLQ = SQLQ & " AND F9_WCBNBR = '" & Data1.Recordset!EC_WCBNBR & "'"
            rsDATA1.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
            Data2.RecordSource = SQLQ
        Else
            SQLQ = "SELECT " & FldList1
            SQLQ = SQLQ & " FROM HR_OHS_FORM9"
            SQLQ = SQLQ & " WHERE F9_EMPNBR = " & glbLEE_ID
            SQLQ = SQLQ & " AND F9_CASE = " & Data1.Recordset!EC_CASE
            SQLQ = SQLQ & " AND F9_WCBNBR = '" & Data1.Recordset!EC_WCBNBR & "'"
            rsDATA1.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            Data2.RecordSource = SQLQ
        End If
        
        If rsDATA1.EOF Then
            rsDATA1.AddNew
            rsDATA1("F9_COMPNO") = "001"
            rsDATA1("F9_EMPNBR") = glbLEE_ID
            rsDATA1("F9_CASE") = Data1.Recordset!EC_CASE
            rsDATA1("F9_WCBNBR") = Data1.Recordset!EC_WCBNBR
            
            If glbtermopen Then
                rsDATA1("TERM_SEQ") = glbTERM_Seq
            End If
        End If
        
        rsDATA1("F9_OCCDATE") = Data1.Recordset!EC_OCCDATE
        rsDATA1("F9_WCBFDTE") = Data1.Recordset!EC_WCBFDTE
        
        'Calculate the Average Weekly gross earnings
        If IsNull(rsDATA1("F9_AVG_GROSS_WEEKLY")) Or rsDATA1("F9_AVG_GROSS_WEEKLY") = "" Then
            xSal = Calculate_Average_Weekly_Gross_Earnings(glbLEE_ID)
            rsDATA1("F9_AVG_GROSS_WEEKLY") = xSal
        End If
        
        rsDATA1("F9_LDATE") = Date
        rsDATA1("F9_LTIME") = Time$
        rsDATA1("F9_LUSER") = glbUserID
        
        rsDATA1.Update
        
        rsDATA1.Close
        Set rsDATA1 = Nothing
    End If

End Sub

Private Sub Update_Incident_Return_DateTime()
    Dim SQLQ As String
    Dim rsHS As New ADODB.Recordset
    
    If glbtermopen Then
        SQLQ = "SELECT " & FldList & " FROM Term_HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & " WHERE EC_CASE=" & Data1.Recordset!EC_CASE
        SQLQ = SQLQ & " AND EC_WCBNBR ='" & Data1.Recordset!EC_WCBNBR & "'"
        SQLQ = SQLQ & " AND EC_EMPNBR =" & glbTERM_ID
        SQLQ = SQLQ & " AND TERM_SEQ=" & glbTERM_Seq
        rsHS.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "SELECT " & FldList & " FROM HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & "WHERE EC_CASE = " & Data1.Recordset!EC_CASE
        SQLQ = SQLQ & " AND EC_WCBNBR ='" & Data1.Recordset!EC_WCBNBR & "'"
        SQLQ = SQLQ & " AND EC_EMPNBR =" & glbLEE_ID
        rsHS.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    
    If Not rsHS.EOF Then
        rsHS("EC_RETURN_TM") = medRtrnWorkTime.Text
        rsHS("EC_RETURN_AMPM") = txtRtrnWorkAP.Text
        If Trim(dlpRtrnWork.Text) <> "" Then
            rsHS("EC_RETURN") = dlpRtrnWork.Text
        End If
        rsHS.Update
    End If
    rsHS.Close
    Set rsHS = Nothing
    
End Sub

Function isChangedHS()

Dim tmDat As New ADODB.Recordset

isChangedHS = False

If Data1.Recordset.EOF Then Exit Function

If Data2.Recordset.EOF Or Data2.Recordset.BOF Then
    Exit Function
End If

If dlpRtrnWork.Text <> IIf(IsNull(Data1.Recordset!EC_RETURN), "", Data1.Recordset!EC_RETURN) Then GoTo isChange
If Left(medRtrnWorkTime, 2) <> "__" And Right(medRtrnWorkTime, 2) <> "__" Then
    If medRtrnWorkTime.Text <> IIf(IsNull(Data1.Recordset!EC_RETURN_TM), "", Data1.Recordset!EC_RETURN_TM) Then GoTo isChange
End If
If txtRtrnWorkAP.Text <> IIf(IsNull(Data1.Recordset!EC_RETURN_AMPM), "", Data1.Recordset!EC_RETURN_AMPM) Then GoTo isChange

Exit Function

isChange:
    isChangedHS = True
End Function

Private Function Clear_Controls()

    txtReturnedWork.Text = ""
    dlpRtrnWork.Text = ""
    medRtrnWorkTime.Mask = ""
    medRtrnWorkTime.Text = ""
    
    dlpAftLOffFDate.Text = ""
    medAftLOffFTime.Mask = ""
    medAftLOffFTime.Text = ""

    dlpAftLOffTDate.Text = ""
    medAftLOffTTime.Mask = ""
    medAftLOffTTime.Text = ""
    
    medTotShiftLost.Text = ""
    medNoPayHrsShift.Text = ""
    txtWrkRtnOpinion.Text = ""
    txtKindofWork.Text = ""
    txtServWorth.Text = ""
    medServWorth.Text = ""
    txtWhenFormerWork.Text = ""
    medAvgWeekEarn.Text = ""

    medGrossTot.Text = ""
    dlpCoveredFrom.Text = ""
    dlpCoveredTo.Text = ""
    txtInsurance.Text = ""
    txtRemarks.Text = ""
End Function

Private Function Calculate_Average_Weekly_Gross_Earnings(xEmpNbr)
    Dim rsSal As New ADODB.Recordset
    Dim SQLQ As String
    Dim xSal As Double
    
    xSal = 0
    
    If glbtermopen Then
        SQLQ = "SELECT SH_EMPNBR, SH_SALARY, SH_SALCD, SH_WHRS FROM Term_SALARY_HISTORY WHERE"
        SQLQ = SQLQ & " TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " AND SH_EMPNBR = " & xEmpNbr
        SQLQ = SQLQ & " AND SH_CURRENT <> 0"
        SQLQ = SQLQ & " ORDER BY SH_EDATE DESC"
        rsSal.Open SQLQ, gdbAdoIhr001X, adOpenDynamic, adLockOptimistic
    Else
        SQLQ = "SELECT SH_EMPNBR, SH_SALARY, SH_SALCD, SH_WHRS FROM HR_SALARY_HISTORY WHERE"
        SQLQ = SQLQ & " SH_EMPNBR = " & xEmpNbr
        SQLQ = SQLQ & " AND SH_CURRENT <> 0"
        SQLQ = SQLQ & " ORDER BY SH_EDATE DESC"
        rsSal.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    End If
    If Not rsSal.EOF Then
        If rsSal("SH_SALCD") = "H" Then
            If IsNull(rsSal("SH_WHRS")) Or rsSal("SH_WHRS") = "" Then
                xSal = 0
            Else
                xSal = Round2DEC(rsSal("SH_SALARY") * rsSal("SH_WHRS"))
            End If
        ElseIf rsSal("SH_SALCD") = "A" Then
            xSal = Round2DEC(rsSal("SH_SALARY") / 52)
        ElseIf rsSal("SH_SALCD") = "M" Then
            xSal = Round2DEC((rsSal("SH_SALARY") * 12) / 52)
        ElseIf rsSal("SH_SALCD") = "D" Then
            xSal = Round2DEC(rsSal("SH_SALARY") * 5)
        End If
    End If
    rsSal.Close
    Set rsSal = Nothing
    
     Calculate_Average_Weekly_Gross_Earnings = xSal
    
End Function

Private Function Round2DEC(tmpNUM, Optional HourlyRate As String)    'laura nov 10, 1997
Dim strNUM As String, x%

If glbFrench Then
    tmpNUM = Replace(Replace(tmpNUM, ",", "."), " ", "")
End If

If glbCompDecHR <> 2 And glbCompDecHR <> 3 And glbCompDecHR <> 4 Then
    glbCompDecHR = 2  'THIS SHOULD NOT HAPPEN BUT IS A VALID DEFAULT
End If
If glbCompSerial = "S/N - 2375W" Then   'City of Timmins
    If GetEmpData(glbLEE_ID, "ED_REGION") <> "S" Then
        Round2DEC = Round(tmpNUM, 2)
    Else
        Round2DEC = Round(tmpNUM, glbCompDecHR)
    End If
Else
    Round2DEC = Round(Val(tmpNUM), glbCompDecHR)
End If
End Function

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
    
    xDocDesc = Left("Form 9 -" & Now, 30)
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

