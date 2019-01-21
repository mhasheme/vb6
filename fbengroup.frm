VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmBENGR 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Benefit Group Master"
   ClientHeight    =   10365
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11715
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10365
   ScaleWidth      =   11715
   WindowState     =   2  'Maximized
   Begin VB.Frame scrFrame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   97
      Top             =   0
      Width           =   9975
      Begin VB.ComboBox cmbGroups 
         Height          =   315
         Left            =   3720
         TabIndex        =   98
         Text            =   "Combo1"
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label Label5 
         Caption         =   "Select Benefit Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   99
         Top             =   150
         Width           =   1935
      End
   End
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   0
      Max             =   50
      SmallChange     =   4
      TabIndex        =   42
      Top             =   9105
      Width           =   10335
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fbengroup.frx":0000
      Height          =   1995
      Left            =   0
      OleObjectBlob   =   "fbengroup.frx":0014
      TabIndex        =   0
      Top             =   480
      Width           =   9675
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8880
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.VScrollBar scrControl 
      Height          =   5445
      LargeChange     =   350
      Left            =   9840
      Max             =   100
      SmallChange     =   350
      TabIndex        =   41
      Top             =   2600
      Width           =   300
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   40
      Top             =   9600
      Width           =   11715
      _Version        =   65536
      _ExtentX        =   20664
      _ExtentY        =   1349
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
      Begin VB.CommandButton cmdRecaSalDep 
         Caption         =   "Recalculate Salary Dependent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.CommandButton cmdCopyBenefitGroup 
         Caption         =   "&Copy Benefit Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   38
         Top             =   0
         Width           =   2475
      End
      Begin VB.CommandButton cmdRecalcPP 
         Caption         =   "Pay Period"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdRecal 
         Caption         =   "&Recalculate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   35
         Top             =   0
         Width           =   1275
      End
      Begin VB.CommandButton cmdRecalAll 
         Caption         =   "&Recalculate All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1740
         TabIndex        =   36
         Top             =   0
         Width           =   1695
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   10920
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
   Begin VB.Frame scrFrame 
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   0
      TabIndex        =   43
      Top             =   2640
      Width           =   9735
      Begin Threed.SSPanel panDetals 
         Height          =   6480
         Left            =   120
         TabIndex        =   44
         Top             =   0
         Width           =   9615
         _Version        =   65536
         _ExtentX        =   16960
         _ExtentY        =   11430
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   0
         Begin VB.TextBox txtPolicy 
            Appearance      =   0  'Flat
            DataField       =   "BM_POLICY"
            Height          =   315
            Left            =   1800
            MaxLength       =   25
            TabIndex        =   32
            Tag             =   "00-Policy Number"
            Top             =   4380
            Width           =   4215
         End
         Begin VB.TextBox txtTAXBEN 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "BM_TAXBEN"
            Height          =   285
            Left            =   7740
            MaxLength       =   1
            TabIndex        =   26
            Tag             =   "00-Taxable Benefit    Y=Yes     N=No"
            Top             =   3360
            Width           =   855
         End
         Begin VB.TextBox txtPer 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "BM_PER"
            Enabled         =   0   'False
            Height          =   285
            Left            =   7710
            MaxLength       =   5
            TabIndex        =   20
            Tag             =   "10-Enter number of units"
            Top             =   2700
            Width           =   870
         End
         Begin VB.ComboBox comPreAftTax 
            Height          =   315
            ItemData        =   "fbengroup.frx":90D4
            Left            =   7380
            List            =   "fbengroup.frx":90DE
            TabIndex        =   23
            Tag             =   "Pre Tax/After Tax"
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox txtPreAftTax 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            DataField       =   "BM_PTAX"
            Height          =   315
            Left            =   8760
            TabIndex        =   62
            Top             =   3120
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Frame Frame2 
            Height          =   465
            Left            =   6630
            TabIndex        =   59
            Top             =   1680
            Width           =   1875
            Begin VB.OptionButton optRound 
               Caption         =   "Next"
               Height          =   225
               Index           =   1
               Left            =   1080
               TabIndex        =   61
               Top             =   150
               Width           =   735
            End
            Begin VB.OptionButton optRound 
               Caption         =   "Nearest"
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   60
               Top             =   150
               Value           =   -1  'True
               Width           =   885
            End
         End
         Begin VB.ComboBox comSalDepn 
            Height          =   315
            ItemData        =   "fbengroup.frx":90F6
            Left            =   4500
            List            =   "fbengroup.frx":9100
            TabIndex        =   11
            Text            =   "No"
            Top             =   1050
            Width           =   735
         End
         Begin VB.ComboBox comRndFactor 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "fbengroup.frx":910D
            Left            =   4530
            List            =   "fbengroup.frx":9144
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Tag             =   "Rounding Factor"
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox txtCovType 
            Appearance      =   0  'Flat
            DataField       =   "BM_COVER"
            Height          =   285
            Left            =   1695
            MaxLength       =   1
            TabIndex        =   10
            Tag             =   "00-Type of Coverage (Single or Family)"
            Top             =   1080
            Width           =   330
         End
         Begin VB.TextBox memComments 
            Appearance      =   0  'Flat
            DataField       =   "BM_COMMENTS"
            Height          =   1305
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Tag             =   "00-Comments - free form"
            Top             =   4920
            Width           =   8565
         End
         Begin VB.TextBox txtRoundFactor 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            DataField       =   "BM_ROUND"
            Height          =   225
            Left            =   5640
            TabIndex        =   58
            Top             =   1770
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox txtSalDepn 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            DataField       =   "BM_SALARYDEPENDANT"
            Height          =   225
            Left            =   5340
            TabIndex        =   12
            Top             =   1080
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox Updstats 
            BorderStyle     =   0  'None
            DataField       =   "BM_LDATE"
            Height          =   285
            Index           =   0
            Left            =   8400
            TabIndex        =   57
            Text            =   "Text1"
            Top             =   6120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox Updstats 
            BorderStyle     =   0  'None
            DataField       =   "BM_LTIME"
            Height          =   285
            Index           =   1
            Left            =   7080
            TabIndex        =   56
            Text            =   "Text1"
            Top             =   6120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox Updstats 
            BorderStyle     =   0  'None
            DataField       =   "BM_LUSER"
            Height          =   285
            Index           =   2
            Left            =   5760
            TabIndex        =   55
            Text            =   "Text1"
            Top             =   6120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtCovAmount 
            BorderStyle     =   0  'None
            DataField       =   "BM_AMT"
            Height          =   285
            Left            =   3000
            TabIndex        =   54
            Text            =   "Text1"
            Top             =   2760
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtMCost 
            BorderStyle     =   0  'None
            DataField       =   "BM_MTHCCOST"
            Height          =   285
            Left            =   3000
            TabIndex        =   53
            Text            =   "Text1"
            Top             =   3360
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtAnComp 
            BorderStyle     =   0  'None
            DataField       =   "BM_CCOST"
            Height          =   285
            Left            =   3000
            TabIndex        =   52
            Text            =   "Text1"
            Top             =   3720
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtMeCost 
            BorderStyle     =   0  'None
            DataField       =   "BM_MTHECOST"
            Height          =   285
            Left            =   4200
            TabIndex        =   51
            Text            =   "Text1"
            Top             =   3360
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtEeCost 
            BorderStyle     =   0  'None
            DataField       =   "BM_ECOST"
            Height          =   285
            Left            =   4200
            TabIndex        =   50
            Text            =   "Text1"
            Top             =   3720
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtTCost 
            BorderStyle     =   0  'None
            DataField       =   "BM_TCOST"
            Height          =   285
            Left            =   6480
            TabIndex        =   64
            TabStop         =   0   'False
            Text            =   "Text1"
            Top             =   3840
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtPerOrDoll 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "BM_PERORDOLL"
            Height          =   285
            Left            =   8640
            TabIndex        =   9
            Top             =   690
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtDWM 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "BM_DWM"
            Height          =   285
            Left            =   8880
            TabIndex        =   45
            Top             =   390
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ComboBox cmbPerOrDoll 
            Height          =   315
            ItemData        =   "fbengroup.frx":9198
            Left            =   7200
            List            =   "fbengroup.frx":91A5
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Tag             =   "40-Select Dallor or Percentage"
            Top             =   720
            Width           =   1215
         End
         Begin VB.ComboBox cmbDWM 
            Height          =   315
            ItemData        =   "fbengroup.frx":91C0
            Left            =   7740
            List            =   "fbengroup.frx":91CD
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Tag             =   "40-Select Day, Week or Month"
            Top             =   390
            Width           =   1095
         End
         Begin VB.TextBox txtWaitPeriod 
            Appearance      =   0  'Flat
            DataField       =   "BM_WaitPeriod"
            Height          =   285
            Left            =   7200
            MaxLength       =   3
            TabIndex        =   4
            Tag             =   "10-Waiting Period (in months)"
            Top             =   420
            Width           =   480
         End
         Begin VB.CheckBox chkCost 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   225
            Left            =   7200
            TabIndex        =   2
            Top             =   120
            Width           =   195
         End
         Begin Threed.SSFrame frmAP 
            Height          =   465
            Left            =   0
            TabIndex        =   46
            Top             =   2160
            Width           =   8535
            _Version        =   65536
            _ExtentX        =   15055
            _ExtentY        =   820
            _StockProps     =   14
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
            Font3D          =   1
            Begin Threed.SSOption optActual 
               Height          =   225
               Index           =   0
               Left            =   1890
               TabIndex        =   47
               Tag             =   "Choose actual or premium"
               Top             =   180
               Width           =   1365
               _Version        =   65536
               _ExtentX        =   2408
               _ExtentY        =   397
               _StockProps     =   78
               Caption         =   "Actual"
               ForeColor       =   16711680
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
            Begin Threed.SSOption optActual 
               Height          =   225
               Index           =   1
               Left            =   4680
               TabIndex        =   48
               TabStop         =   0   'False
               Tag             =   "Choose actual or premium"
               Top             =   150
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
               _ExtentY        =   397
               _StockProps     =   78
               Caption         =   "Premium"
               ForeColor       =   16711680
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
            Begin VB.Label lblAP 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               DataField       =   "BM_PREMIUM"
               DataSource      =   "Data2"
               ForeColor       =   &H80000008&
               Height          =   225
               Left            =   720
               TabIndex        =   49
               Top             =   240
               Visible         =   0   'False
               Width           =   435
            End
         End
         Begin INFOHR_Controls.DateLookup dlpEDate 
            DataField       =   "BM_EDATE"
            Height          =   285
            Left            =   1380
            TabIndex        =   6
            Tag             =   "41-Effective Date of coverage"
            Top             =   720
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin MSMask.MaskEdBox medPayPeriodAmount 
            DataField       =   "BM_PPAMT"
            Height          =   285
            Left            =   4500
            TabIndex        =   7
            Tag             =   "20-Amount charged for every pay period"
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   "#,##0.0000;(#,##0.0000)"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox medMaxAmnt 
            DataField       =   "BM_MAXDOL"
            Height          =   285
            Left            =   7200
            TabIndex        =   13
            Tag             =   "20-Enter Maximum Amount"
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   "$#,##0.0000;($#,##0.0000)"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox medCovAmount 
            Height          =   285
            Left            =   1800
            TabIndex        =   18
            Tag             =   "20-Amount of Coverage"
            Top             =   2700
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   "$##,##0.00;($##,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPPComp 
            DataField       =   "BM_PCC"
            Height          =   285
            Left            =   1800
            TabIndex        =   21
            Tag             =   "11-Percentage paid by company"
            Top             =   3030
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   "##0.00%"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medMCCOST 
            Height          =   285
            Left            =   1800
            TabIndex        =   24
            Tag             =   "21-Monthly company cost"
            Top             =   3360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
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
            Format          =   "$#,##0.0000;($#,##0.0000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medCompCost 
            Height          =   285
            Left            =   1800
            TabIndex        =   27
            Tag             =   "11-Cost of Benefit to Company"
            Top             =   3690
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
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
            Format          =   "$#,##0.0000;($#,##0.0000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medUnitCost 
            DataField       =   "BM_UNITCOST"
            Height          =   285
            Left            =   4620
            TabIndex        =   19
            Tag             =   "20-Enter Unit Cost"
            Top             =   2700
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Format          =   "$#,##0.000000;($#,##0.000000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPPE 
            DataField       =   "BM_PCE"
            Height          =   285
            Left            =   4620
            TabIndex        =   22
            Tag             =   "11-Percentage paid by employee"
            Top             =   3030
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Format          =   "##0.00%"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medMECOST 
            Height          =   285
            Left            =   4620
            TabIndex        =   25
            Tag             =   "21-Monthly employee cost"
            Top             =   3360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
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
            Format          =   "$#,##0.0000;($#,##0.0000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medEECost 
            Height          =   285
            Left            =   4620
            TabIndex        =   28
            Tag             =   "11-Cost of benefit to Employee"
            Top             =   3690
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
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
            Format          =   "$#,##0.0000;($#,##0.0000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medMinCover 
            DataField       =   "BM_MINIMUM"
            Height          =   285
            Left            =   1695
            TabIndex        =   14
            Tag             =   "20-Minimum of Coverage"
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
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
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medMaxCover 
            DataField       =   "BM_MAXIMUM"
            Height          =   285
            Left            =   4530
            TabIndex        =   15
            Tag             =   "20-Maximum of Coverage"
            Top             =   1425
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
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
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medSalFactor 
            DataField       =   "BM_FACTOR"
            Height          =   285
            Left            =   1695
            TabIndex        =   16
            Tag             =   "20-Salary Factor"
            Top             =   1800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
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
            Format          =   "#,##0.000000000;(#,##0.000000000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medTCost 
            Height          =   285
            Left            =   6885
            TabIndex        =   29
            Tag             =   "21-Total Cost of the Coverage"
            Top             =   3690
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Format          =   "$##,###.0000;($##,###.0000)"
            PromptChar      =   "_"
         End
         Begin INFOHR_Controls.CodeLookup clpBGroup 
            DataField       =   "BM_BENEFIT_GROUP"
            Height          =   285
            Left            =   1380
            TabIndex        =   1
            Tag             =   "01-Benefit - Group Code"
            Top             =   30
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "BGMF"
            MaxLength       =   10
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "BM_BCODE"
            Height          =   285
            Index           =   4
            Left            =   1380
            TabIndex        =   3
            Tag             =   "01-Benefit - Code"
            Top             =   360
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "BNCD"
            MaxLength       =   10
         End
         Begin MSMask.MaskEdBox medCYTD 
            DataField       =   "BM_CYTD"
            Height          =   315
            Left            =   1800
            TabIndex        =   30
            Tag             =   "11-Current YTD Company"
            Top             =   4020
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
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
            Format          =   "$#,##0.0000;($#,##0.0000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medEYTD 
            DataField       =   "BM_EYTD"
            Height          =   315
            Left            =   4620
            TabIndex        =   31
            Tag             =   "11-Current YTD Employee"
            Top             =   4020
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
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
            Format          =   "$#,##0.0000;($#,##0.0000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medRateLevel 
            DataField       =   "BM_RATELEVEL"
            Height          =   285
            Left            =   8115
            TabIndex        =   33
            Tag             =   "10-Rate Level"
            Top             =   4395
            Visible         =   0   'False
            Width           =   480
            _ExtentX        =   847
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
            Format          =   "0"
            PromptChar      =   " "
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rate Level"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   46
            Left            =   7080
            TabIndex        =   101
            Top             =   4440
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label lblTitle 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Policy Number"
            Height          =   255
            Index           =   30
            Left            =   60
            TabIndex        =   100
            Top             =   4380
            Width           =   1395
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pre Tax/After Tax"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   33
            Left            =   6000
            TabIndex        =   96
            Top             =   3060
            Width           =   1455
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Salary Dependent"
            Height          =   315
            Index           =   25
            Left            =   3000
            TabIndex        =   95
            Top             =   1065
            Width           =   1335
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum Coverage"
            Height          =   315
            Index           =   26
            Left            =   60
            TabIndex        =   94
            Top             =   1470
            Width           =   1515
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum Coverage"
            Height          =   315
            Index           =   28
            Left            =   3000
            TabIndex        =   93
            Top             =   1425
            Width           =   1515
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Salary Factor"
            Height          =   315
            Index           =   27
            Left            =   60
            TabIndex        =   92
            Top             =   1845
            Width           =   1515
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Rounding Factor"
            Height          =   315
            Index           =   29
            Left            =   3000
            TabIndex        =   91
            Top             =   1785
            Width           =   1515
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   14
            Left            =   6000
            TabIndex        =   90
            Top             =   3690
            Width           =   615
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Taxable Benefit"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6000
            TabIndex        =   89
            Top             =   3360
            Width           =   1455
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Per"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   8
            Left            =   6000
            TabIndex        =   88
            Top             =   2700
            Width           =   300
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   13
            Left            =   3360
            TabIndex        =   87
            Top             =   3720
            Width           =   825
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3360
            TabIndex        =   86
            Top             =   3390
            Width           =   1095
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "% Paid Employee"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   10
            Left            =   3360
            TabIndex        =   85
            Top             =   3060
            Width           =   1425
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Cost"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   3360
            TabIndex        =   84
            Top             =   2730
            Width           =   795
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Company"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   12
            Left            =   720
            TabIndex        =   83
            Top             =   3690
            Width           =   780
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Company"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   720
            TabIndex        =   82
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Company"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   81
            Top             =   3030
            Width           =   975
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Annual:"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   11
            Left            =   60
            TabIndex        =   80
            Top             =   3690
            Width           =   735
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Monthly:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   60
            TabIndex        =   79
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "% Paid "
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   9
            Left            =   60
            TabIndex        =   78
            Top             =   3030
            Width           =   735
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Coverage Amount"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   6
            Left            =   60
            TabIndex        =   77
            Top             =   2700
            Width           =   1515
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum Amount"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   5880
            TabIndex        =   76
            Top             =   1110
            Width           =   1470
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pay Period Amount"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   3000
            TabIndex        =   75
            Top             =   750
            Width           =   1620
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Coverage"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   74
            Top             =   1140
            Width           =   825
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Effective Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   73
            Top             =   780
            Width           =   1245
         End
         Begin VB.Label lblBenefit 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Benefit"
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
            Left            =   60
            TabIndex        =   72
            Top             =   420
            Width           =   615
         End
         Begin VB.Label lblBen 
            BackStyle       =   0  'Transparent
            Caption         =   "Benefit Group"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   71
            Top             =   90
            Width           =   1215
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Comments"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   32
            Left            =   60
            TabIndex        =   70
            Top             =   4680
            Width           =   735
         End
         Begin VB.Label lblRound 
            BackColor       =   &H00E0E0E0&
            DataField       =   "BM_NEXTNEAREST"
            Height          =   345
            Left            =   8160
            TabIndex        =   69
            Top             =   1710
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Dollar/Percentage"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   37
            Left            =   5880
            TabIndex        =   68
            Top             =   780
            Width           =   1305
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Waiting Period"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   31
            Left            =   5880
            TabIndex        =   67
            Top             =   450
            Width           =   1215
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   39
            Left            =   3360
            TabIndex        =   66
            Top             =   4050
            Width           =   825
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Current YTD:  Company"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   38
            Left            =   60
            TabIndex        =   65
            Top             =   4020
            Width           =   1680
         End
         Begin VB.Label lblCost 
            Caption         =   "Use Cost Table"
            Height          =   315
            Left            =   5880
            TabIndex        =   63
            Top             =   120
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frmBENGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SN_CASEY = "S/N - 2214W"

Dim Actn
Dim ChangingFields As Boolean
Dim fglbESQLQ, fglbWSQLQ
Dim fglbSDate As Variant
Dim XUpdCount
Dim fglbNew
Dim RSDATA As New ADODB.Recordset
Dim OBCode, OCOVER, OTCOST, OPremium, OEDate, OPPE, OPCC
Dim OPPAMT, OMAXDOL, OBNAME, OBRELATE, ODOB, OPER
Dim OUNITCOST, OBAMT
Dim OMTHCOMP, OMTHEMP, OTAXBEN
Dim setFlag As Boolean
Dim salFlag As Boolean
Dim fglbBGroup
Dim xWellingtonFlg As Boolean
Dim flgSendEmail As Boolean
Dim flgBenUpdated As Boolean
Dim MailBody

Private Sub cmbDWM_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbPerOrDoll_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdCopyBenefitGroup_Click()
    
    glbWFC_IPPopFormName = ""
    
    frmBENGRCopy.clpBGroupOld = clpBGroup.Text
    frmBENGRCopy.txtPolicy.Text = txtPolicy.Text
    frmBENGRCopy.memComments.Text = memComments.Text
    frmBENGRCopy.Show 1
    '------------------refreshing the form
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    Data1.Refresh
    
    'Ticket #22526
    Call INIData
    Call optActual_Click(0, 1)
    'End
    
    Call SET_UP_MODE
    
End Sub

Private Sub cmdRecal_Click()
Dim a As Integer, Msg As String, x%
Dim Response%

Msg = "Are You Sure You Want To Do This?"
a% = MsgBox(Msg, 36, "Confirm Recalculate")
If a% <> 6 Then Exit Sub


If Data1.Recordset.EOF Then
    MsgBox "Nothing to recalculate"
    Exit Sub
End If
Call UPDBGroup(Trim(clpBGroup.Text), Trim(clpCode(4).Text), Trim(txtCovType.Text), GroupMasterRecal)
If glbCompSerial = "S/N - 2380W" Then Call CalcPP(Trim(clpCode(4).Text), Trim(clpBGroup.Text))
MsgBox "     Recalculate is Finished.     "
End Sub

Private Sub UPDBGroup(BGroup, BCode, Cover, BSource As BenefitUpdateSource)
Dim rsEmp As New ADODB.Recordset
Dim rsGroup As New ADODB.Recordset
Dim SQLQ
Dim xEmpnbr
Dim BCodeCover
Dim xTotalRecs, XUpdCount As Integer

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = "Update Benefit"
        
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute "UPDATE HRBENFT SET BF_LUSER ='" & glbUserID & "' WHERE BF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_BENEFIT_GROUP='" & BGroup & "')"
gdbAdoIhr001.CommitTrans

DoEvents
SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_BENEFIT_GROUP='" & BGroup & "'"
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
xTotalRecs = rsEmp.RecordCount

Do Until rsEmp.EOF
    MDIMain.panHelp(0).FloodPercent = (XUpdCount / xTotalRecs) * 100
    XUpdCount = XUpdCount + 1
    
    xEmpnbr = rsEmp("ED_EMPNBR")
    BCodeCover = BCode & "_" & Cover

    'Release 8.1 - Email Notification - only if Benefit got updated
    glbBenAdded = "False"
    glbBenChanged = "False"
    glbBenDeleted = "False"
    glbBenEffDate = ""
    
    'Update Employee's Benefit
    Call updateBenefit(xEmpnbr, BGroup, "A", BSource, BCodeCover)
    
    'Release 8.1 - Employee Updated - Send Email Notification
    If flgSendEmail Then
        If BSource = GroupMasterAdd And glbBenAdded = "True" Then
            Call EmailNotification("ADD", xEmpnbr, BCode, glbBenEffDate)
        ElseIf BSource = GroupMasterEdit And glbBenChanged <> "False" Then
            Call EmailNotification("UPDATE", xEmpnbr, BCode, glbBenEffDate)
        ElseIf BSource = GroupMasterDelete And glbBenDeleted = "True" Then
            Call EmailNotification("DELETE", xEmpnbr, BCode, glbBenEffDate)
        End If
    End If
    
    DoEvents
    rsEmp.MoveNext
Loop
rsEmp.Close

'City of Timmins
If glbCompSerial = "S/N - 2375W" Then
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute "UPDATE TERM_HRBENFT SET BF_LUSER ='" & glbUserID & "' WHERE BF_EMPNBR IN (SELECT ED_EMPNBR FROM TERM_HREMP WHERE ED_BENEFIT_GROUP='" & BGroup & "')"
    gdbAdoIhr001.CommitTrans
    
    DoEvents
    XUpdCount = 0
    SQLQ = "SELECT ED_EMPNBR FROM TERM_HREMP WHERE ED_BENEFIT_GROUP='" & BGroup & "'"
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    xTotalRecs = rsEmp.RecordCount
    
    Do Until rsEmp.EOF
        MDIMain.panHelp(0).FloodPercent = (XUpdCount / xTotalRecs) * 100
        XUpdCount = XUpdCount + 1
        
        xEmpnbr = rsEmp("ED_EMPNBR")
        BCodeCover = BCode & "_" & Cover
    
        Call updateBenefit_TERM(xEmpnbr, BGroup, "T", BSource, BCodeCover)
        DoEvents
        rsEmp.MoveNext
    Loop
    rsEmp.Close
End If


If glbWFC Then 'Ticket #15818
    Call WFCCNDBeneAuditFlag
End If
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""

End Sub

Private Sub cmdRecal_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdRecalAll_Click()
Dim rsEmp As New ADODB.Recordset
Dim rsGroup As New ADODB.Recordset
Dim SQLQ, xEmpnbr, NewBGroup
Dim a As Integer, Msg As String, x%
Dim Response%

Msg = "Are You Sure You Want To Do This?"
a% = MsgBox(Msg, 36, "Confirm Recalculate All")
If a% <> 6 Then Exit Sub

cmdRecalcPP.Enabled = False
SQLQ = "SELECT DISTINCT BM_BENEFIT_GROUP FROM HR_BENEFITS_GROUP "
If fglbBGroup <> "ALLGROUP" Then
    SQLQ = SQLQ & " WHERE BM_BENEFIT_GROUP='" & fglbBGroup & "'"
End If
Screen.MousePointer = HOURGLASS
rsGroup.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
Do Until rsGroup.EOF
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_BENEFIT_GROUP='" & rsGroup("BM_BENEFIT_GROUP") & "'"
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Do Until rsEmp.EOF
        xEmpnbr = rsEmp("ED_EMPNBR")
        NewBGroup = rsGroup("BM_BENEFIT_GROUP")
        Call updateBenefit(xEmpnbr, NewBGroup, "A", GroupMasterRecal)
        rsEmp.MoveNext
    Loop
    rsEmp.Close
    rsGroup.MoveNext
Loop
If glbCompSerial = "S/N - 2380W" Then Call CalcPP
Screen.MousePointer = DEFAULT
cmdRecalcPP.Enabled = True
MsgBox "     Recalculate is Finished.     "
End Sub

Private Sub cmdRecalAll_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdRecalcPP_Click()
    cmdRecalcPP.Enabled = False

    Call CalcPP

    cmdRecalcPP.Enabled = True
End Sub

Private Sub cmdRecaSalDep_Click()
Dim rsEmp As New ADODB.Recordset
Dim rsGroup As New ADODB.Recordset
Dim SQLQ, xEmpnbr, NewBGroup
Dim a As Integer, Msg As String, x%
Dim Response%
Dim xCunt As Long
Dim xCuntTot As Long


Msg = "Are You Sure You Want To Do This?"
a% = MsgBox(Msg, 36, "Confirm Recalculate Salary Dependent")
If a% <> 6 Then Exit Sub

cmdRecaSalDep.Enabled = False
SQLQ = "SELECT DISTINCT BM_BENEFIT_GROUP FROM HR_BENEFITS_GROUP "
SQLQ = SQLQ & " WHERE (1=1) "
If fglbBGroup <> "ALLGROUP" Then
    SQLQ = SQLQ & "AND BM_BENEFIT_GROUP='" & fglbBGroup & "'"
End If
SQLQ = SQLQ & "AND BM_SALARYDEPENDANT = 'Y' "
Screen.MousePointer = HOURGLASS
rsGroup.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

MDIMain.panHelp(0).FloodType = 1
Do Until rsGroup.EOF
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_BENEFIT_GROUP='" & rsGroup("BM_BENEFIT_GROUP") & "'"
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsEmp.EOF Then
        xCunt = 0
        xCuntTot = rsEmp.RecordCount
    End If
    Do Until rsEmp.EOF
        MDIMain.panHelp(0).FloodPercent = (xCunt / xCuntTot) * 100: xCunt = xCunt + 1
        DoEvents
        
        xEmpnbr = rsEmp("ED_EMPNBR")
        NewBGroup = rsGroup("BM_BENEFIT_GROUP")
        Call updateBenefit(xEmpnbr, NewBGroup, "A", GroupMasterRecal)
        rsEmp.MoveNext
    Loop
    rsEmp.Close
    rsGroup.MoveNext
Loop

Screen.MousePointer = DEFAULT
cmdRecaSalDep.Enabled = True
MsgBox "     Recalculate Salary Dependent is Finished.     "
MDIMain.panHelp(0).FloodType = 0

End Sub

Private Sub comPreAftTax_Change()
    If comPreAftTax = "Pre Tax" Then
        txtPreAftTax = "P"
    ElseIf comPreAftTax = "After Tax" Then
        txtPreAftTax = "A"
    Else
        txtPreAftTax = ""
    End If
End Sub

Private Sub comPreAftTax_Click()
    Call comPreAftTax_Change
End Sub

Private Sub comRndFactor_LostFocus()
Call Set_SalCover
End Sub

Private Sub comSalDepn_Click()
Dim x
If comSalDepn = "Yes" Then
    lblAP.Caption = "P"
    Set_SalCover
End If
txtSalDepn = Left(comSalDepn, 1)
comSalDepn_Change
End Sub

Private Sub Form_Activate()
    Call SET_UP_MODE
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Keepfocus As Boolean
    If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
    Keepfocus = Not isUpdated(Me)
    Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub lblRound_Change()
If lblRound = "R" Then
    optRound(0) = True
Else
    optRound(1) = True
End If
End Sub

Private Sub medCYTD_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medEYTD_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPPE_LostFocus()
If Len(medPPE) > 0 Then
    If IsNumeric(medPPE) Then
        If medTCost <> "From System" Then
            medEECost = Val(medTCost) * Val(medPPE) / 100
        End If
        medPPE = Val(medPPE) / 100
        If medTCost <> "From System" Then
            setFlag = False
        End If
        Call setTotal
        
        medEECost.Visible = True
    End If
End If
End Sub

Private Sub medRateLevel_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medTCost_LostFocus()
If medTCost <> "From System" Then
    setFlag = False
End If
Call setTotal
End Sub

Private Sub optRound_Click(Index As Integer)
lblRound = IIf(optRound(0), "R", "N")
Call Set_SalCover
End Sub

Private Sub optRound_LostFocus(Index As Integer)
lblRound = IIf(optRound(0), "R", "N")
Call Set_SalCover
End Sub

Private Sub optRound_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Call Set_SalCover
End Sub

Private Sub medMCCOST_Change()
    If glbCompSerial = SN_CASEY And Not ChangingFields Then CalcBeneCasey
End Sub

Private Sub medMECOST_Change()
    If glbCompSerial = SN_CASEY And Not ChangingFields Then CalcBeneCasey
    If glbCompSerial = "S/N - 2380W" And comSalDepn.Text = "No" And IsNumeric(medMECOST) Then
        Select Case clpBGroup
        Case "GHON", "GHQC", "CAMPBELL", "CAMPBC", "GHQC113", "GHON113", "CAMPBC113"   'Ticket #18963, Ticket #24537 - more codes
            medPayPeriodAmount = medEECost / 52
        Case Else
            medPayPeriodAmount = medMECOST / 2
        End Select
    End If
End Sub

Private Function chkMUBENEFITS()
Dim x%, SQLQ As String, Msg As String, a%
Dim rsTBen As New ADODB.Recordset, xFlagBen As Boolean
chkMUBENEFITS = False

On Error GoTo chkMUBENEFITS_Err

Call medTCost_Change

If medEECost.Text = "" Then medEECost.Text = 0
If medCompCost.Text = "" Then medCompCost.Text = 0
If medMCCOST.Text = "" Then medMCCOST.Text = 0
If medMECOST.Text = "" Then medMECOST.Text = 0
If medTCost.Text = "" Then medTCost.Text = 0

If Len(clpCode(4).Text) > 0 And clpCode(4).Caption = "Unassigned" Then
    MsgBox "If code entered it must be known"
    clpCode(4).SetFocus
    Exit Function
End If
If Len(clpCode(4).Text) < 1 Then
    MsgBox "Benefit is a required field"
    clpCode(4).SetFocus
    Exit Function
End If


If Len(clpBGroup.Text) > 0 And clpBGroup.Caption = "Unassigned" Then
    MsgBox "If code entered it must be known"
    clpBGroup.SetFocus
    Exit Function
End If
If Len(clpBGroup.Text) < 1 Then
    MsgBox "Benefit Group is a required field"
    clpBGroup.SetFocus
    Exit Function
End If


If Len(dlpEDate.Text) > 0 Then
    If Not IsDate(dlpEDate.Text) Then
        MsgBox "Effective Date is not a valid date."
        dlpEDate.SetFocus
        Exit Function
    End If
Else
'    If Actn = "A" Then
'        MsgBox "Effective Date is required."
'        dlpEDate.SetFocus
'        Exit Function
'    End If
End If

If glbCompSerial = "S/N - 2385W" Then 'Ticket #14402
    If Len(txtCovType.Text) = 0 Then
        MsgBox "Coverage is required."
        txtCovType.SetFocus
        Exit Function
    End If
End If

If Actn = "D" Then GoTo BpChk

'--------------------------- Add/Modify
'If Actn = "M" And Len(txtCovType) = 0 Then
'    MsgBox "Coverage is required."
'    txtCovType.SetFocus
'    Exit Function
'End If
If Len(medPayPeriodAmount) > 0 Then
   If Not IsNumeric(medPayPeriodAmount) Then
       MsgBox "Pay Period Amount is invalid"
       medPayPeriodAmount.SetFocus
       Exit Function
   End If
Else
   medPayPeriodAmount = 0
End If

If Len(medMaxAmnt) > 0 Then
    If Not IsNumeric(medMaxAmnt) Then
        MsgBox "Maximum Amount is invalid"
        medMaxAmnt.SetFocus
        Exit Function
    End If
Else
    medMaxAmnt = 0
End If

If Len(txtWaitPeriod.Text) > 0 Then
    If Len(txtDWM.Text) = 0 Then
        MsgBox "Waiting Period must be Day(s) or Week(s) or Month(s)"
        medMaxAmnt.SetFocus
        Exit Function
    End If
End If

'--------added by Jaddy 11/2/99 begin
If comSalDepn = "Yes" Then
    If Len(medMinCover) > 0 Then
        If Not IsNumeric(medMinCover) Then
            MsgBox "Minimum Coverage must be numeric ", 16
            If medMinCover.Enabled Then medMinCover.SetFocus
            Exit Function
        End If
    Else
        medMinCover = 0
    End If
    If Len(medMaxCover) > 0 Then
        If Not IsNumeric(medMaxCover) Then
            MsgBox "Maximum Coverage must be numeric ", 16
            If medMaxCover.Enabled Then medMaxCover.SetFocus
            Exit Function
        Else
            If Val(medMaxCover) > 0 And Val(medMinCover) > 0 Then
                If Val(medMaxCover) < Val(medMinCover) Then
                    MsgBox "Maximum Coverage Must Be Greater Then Minimum Coverage", 16
                    If medMaxCover.Enabled Then medMaxCover.SetFocus
                    Exit Function
                End If
            End If
        End If
    Else
        medMaxCover = 0
    End If
    If Len(Trim(medSalFactor)) > 0 Then
        If Not IsNumeric(medSalFactor) Then
            MsgBox "Salary Factor be numeric ", 16
            If medSalFactor.Enabled Then medSalFactor.SetFocus
            Exit Function
'        Else
'            If Val(medSalFactor) = 0 Then
'                MsgBox "Salary Factor Must Be Greater Then 0", 16
'                If medSalFactor.Enabled Then medSalFactor.SetFocus
'                Exit Function
'            End If
        End If
    Else
        medSalFactor = 0
'          MsgBox "Salary Factor Must Be Greater Then 0", 16
'          If medSalFactor.Enabled Then medSalFactor.SetFocus
'          Exit Function
    End If

Else
'--------added by Jaddy 11/2/99 end
    If Len(medCovAmount) > 0 Then
        If Not IsNumeric(medCovAmount) Then
            MsgBox "Coverage Amount is invalid", 48
            medCovAmount.SetFocus
            Exit Function
        End If
        If optActual(1) And medCovAmount = 0 Then
            MsgBox "Coverage Amount is required", 48
            medCovAmount.SetFocus
            Exit Function
        End If
    Else
        If optActual(1) Then
            MsgBox "Coverage Amount is required", 48
            medCovAmount.SetFocus
            Exit Function
        Else
            medCovAmount = 0
        End If
        
    End If
End If 'jaddy 11/2/99
If Len(medUnitCost) > 0 Then
    If Not IsNumeric(medUnitCost) Then
        MsgBox "Per Unit is invalid.", 48
        medUnitCost.SetFocus
        Exit Function
    End If
    If optActual(1) And medUnitCost = 0 Then
        MsgBox "Per Unit is required."
        medUnitCost.SetFocus
        Exit Function
    End If
Else
    If optActual(1) Then
        MsgBox "Per Unit is required."
        medUnitCost.SetFocus
        Exit Function
    End If
    medUnitCost = 0
End If
If Len(txtPer) > 0 Then
    If Not IsNumeric(txtPer) Then
        MsgBox "Per Unit Cost is invalid.", 48
        txtPer.SetFocus
        Exit Function
    End If
    If optActual(1) And txtPer = 0 Then
        MsgBox "Per Unit Cost is required."
        txtPer.SetFocus
        Exit Function
    End If
Else
    If optActual(1) Then
        MsgBox "Per Unit Cost is required."
        txtPer.SetFocus
        Exit Function
    End If
    txtPer = 0
End If

If Len(medPPComp) <= 0 Then
    MsgBox "Company Percentage Paid is required"
    medPPComp.SetFocus
    Exit Function
End If
If Not IsNumeric(medPPComp) Then
    MsgBox "Company Percentage Paid is invalid", 48
    medPPComp.SetFocus
    Exit Function
End If
If medPPComp > 1 Or medPPComp < 0 Then
    MsgBox "Company Percentage Paid is invalid", 48
    medPPComp.SetFocus
    Exit Function
End If
If Len(medMCCOST) > 0 And medMCCOST <> "From System" Then    'laura 02/27/98
  If Not IsNumeric(medMCCOST) Then
      MsgBox "Monthly Company Cost paid is invalid", 48
      medMCCOST.SetFocus
      Exit Function
  End If
Else
   If comSalDepn <> "Yes" Then medMCCOST = 0      'jaddy 11/3/99
End If
If Len(medMECOST) > 0 And medMECOST <> "From System" Then       'laura 02/27/98
  If Not IsNumeric(medMECOST) Then
      MsgBox "Monthly Employee Cost paid is invalid", 48
      medMECOST.SetFocus
      Exit Function
  End If
Else
  If comSalDepn <> "Yes" Then medMECOST = 0     'jaddy 11/3/99
End If
'------Jaddy 11/3/99 changed begin
If optActual(0) Then
    If Len(medTCost) > 0 Then
        If Not IsNumeric(medTCost) Then
            MsgBox "Total Cost is invalid", 48
            medTCost.SetFocus
            Exit Function
        Else
'            If medTCost = 0 Then
'                MsgBox "Total Cost is required"
'                medTCost.SetFocus
'                Exit Function
'            End If
        End If
    Else
        medTCost = 0
'        MsgBox "Total Cost is required"
'        medTCost.SetFocus
'        Exit Function
    End If
End If

'Ticket #22464 - Goodmans
If glbCompSerial = "S/N - 2290W" And Len(medRateLevel.Text) > 0 Then
    If Not IsNumeric(medRateLevel.Text) Then
        MsgBox "Invalid " & lblTitle(46).Caption
        medRateLevel.SetFocus
        Exit Function
    End If
End If

''Jerry doesn't want to check duplicate here, commented by Frank
''Frank 11/03/2003 check duplicated benefit code - Begin
'SQLQ = "SELECT * FROM HR_BENEFITS_GROUP "
'SQLQ = SQLQ & " WHERE BM_BENEFIT_GROUP = '" & clpBGroup & "' "
'SQLQ = SQLQ & "AND BM_BCODE= '" & clpCode(4) & "' "
'If fglbNew <> True Then
'SQLQ = SQLQ & " AND BM_BENE_ID <> " & rsDATA!BM_BENE_ID
'End If
'SQLQ = SQLQ & " ORDER BY BM_BENEFIT_GROUP, BM_BCODE, BM_EDATE"
'rsTBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
'xFlagBen = False
'If Not rsTBen.EOF Then
'    xFlagBen = True
'End If
'rsTBen.Close
'If xFlagBen Then
'    Msg = "Duplicate Benefit Code entered. Continue? Yes/No "
'    a% = MsgBox(Msg, 36, "Confirm")
'    If a% <> 6 Then Exit Function
'End If
''Frank 11/03/2003 check duplicated benefit code - End

'------Jaddy 11/3/99 changed end
BpChk:
chkMUBENEFITS = True

Exit Function

chkMUBENEFITS_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEbenefit", "HRBENFT", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub comSalDepn_Change()
'txtSalDepn = Left(comSalDepn, 1)
If comSalDepn = "Yes" Then
    comRndFactor.Enabled = True
    medMinCover.Enabled = True
    medMaxCover.Enabled = True
    medSalFactor.Enabled = True
    medCovAmount.Enabled = False
    
Else
    comRndFactor.Enabled = False
    medMinCover.Enabled = False
    medMaxCover.Enabled = False
    medSalFactor.Enabled = False
    medCovAmount.Enabled = True
    If comSalDepn = "No" Then
        comRndFactor.ListIndex = 0
    End If
    medMinCover.Text = ""
    medMaxCover.Text = ""
    medSalFactor.Text = ""
'    medCovAmount = 0
End If

Call Set_SalCover
'Call comSalDepn_Click
End Sub

Private Sub comSalDepn_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Set_SalCover()
If salFlag = False Then
    If comSalDepn = "Yes" Then
        medCovAmount = "From System"
    Else
        medCovAmount = txtCovAmount
    End If
    salFlag = True
Else
    salFlag = False
End If
End Sub

Private Sub comRndFactor_Change()
    If Val(txtRoundFactor) <> Val(comRndFactor.ItemData(comRndFactor.ListIndex)) Then
        txtRoundFactor = Val(comRndFactor.ItemData(comRndFactor.ListIndex))
    End If
    Call Set_SalCover
End Sub

Private Sub comRndFactor_Click()
    If Val(txtRoundFactor) <> Val(comRndFactor.ItemData(comRndFactor.ListIndex)) Then
        txtRoundFactor = Val(comRndFactor.ItemData(comRndFactor.ListIndex))
    End If
    Call Set_SalCover
End Sub

Private Sub comRndFactor_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medSalFactor_Change()
Call Set_SalCover
End Sub

Private Sub medSalFactor_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMinCover_Change()
Call Set_SalCover
End Sub

Private Sub medMinCover_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMaxCover_Change()
Call Set_SalCover
End Sub

Private Sub medMaxCover_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Load()
setFlag = False
salFlag = False

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = "FRMBENGR"

lblAP.Caption = "A"

Screen.MousePointer = HOURGLASS

Dim SQLQ

Data1.ConnectionString = glbAdoIHRDB

SQLQ = "SELECT  * FROM HR_BENEFITS_GROUP"

Data1.RecordSource = SQLQ
Data1.Refresh

'If Data1.Recordset.EOF Then
fglbNew = False
'Else
'fglbNew = True
'End If


Call setRptCaption(Me)

If glbLinamar Then
    clpCode(4).MaxLength = 8
'    clpCode(5).MaxLength = 8
End If

If glbCompSerial = SN_CASEY Then
    medPPComp.Enabled = False
    medTCost.Enabled = False
    medMECOST.Enabled = True
    medMCCOST.Enabled = True
End If

If glbCompSerial = "S/N - 2380W" Then
    cmdRecalcPP.Visible = True
End If

If glbCompSerial = "S/N - 2262W" Then   'To avoid the Runtime error 28 - Out of stack space
    xWellingtonFlg = True
End If

If glbCompSerial = "S/N - 2385W" Then 'Ticket #14402
    lblTitle(4).FontBold = True
End If

If glbVadim Then
    lblTitle(46).Visible = True
    medRateLevel.Visible = True
End If

If glbCompSerial = "S/N - 2290W" Then 'Ticket #22464 - Goodmans
    lblTitle(46).Visible = True
    medRateLevel.Visible = True
    lblTitle(46).Caption = "Sequence #"
    medRateLevel.Tag = "10-Sequence #"
    medRateLevel.MaxLength = 2
End If

If glbWFC Then 'Ticket #28772 Franks 06/20/2016
    WFCScreenSetup
End If

Call INI_Controls(Me)

'txtRoundFactor = Val(comRndFactor.ItemData(comRndFactor.ListIndex))

medPPComp = 0

Call Display_Value
Call INIData
Call optActual_Click(0, 1)

clpBGroup.TextBoxWidth = 1500
'If glbMulti Then textMulti.Visible = True
Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
Set frmBENGR = Nothing  'carmen apr 2000

End Sub

Private Sub lblAP_Change()
    If lblAP.Caption = "A" And optActual(0).Value <> True Then
    optActual(0).Value = True
    ElseIf lblAP.Caption = "P" And optActual(1).Value <> True Then
    optActual(1).Value = True
End If

End Sub

Private Sub medCompCost_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCovAmount_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCovAmount_Change()
If optActual(1).Value = True Then
    If comSalDepn = "Yes" Then
        medTCost = "From System"
    Else
        If Not IsNumeric(medCovAmount) Then medCovAmount = 0
        If Not IsNumeric(txtPer) Then txtPer = 0
        If Not IsNumeric(medUnitCost) Then medUnitCost = 0
        If txtPer > 0 And medUnitCost > 0 Then
            medTCost = medCovAmount / txtPer * medUnitCost
        Else
            medTCost = 0
        End If
    End If
End If

End Sub

Private Sub medEECost_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMaxAmnt_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMCCOST_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMECOST_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPayPeriodAmount_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPPComp_GotFocus()
Call SetPanHelp(ActiveControl)
medPPComp = Val(medPPComp) * 100
End Sub

Private Sub medPPComp_Change()
    '    If glbCompSerial <> SN_CASEY Then
    '        medPPE = 1 - Val(medPPComp) / 100
    '        Call medTCost_Change
    '    End If
End Sub

Private Sub medPPComp_LostFocus()
    medPPComp = Val(medPPComp) / 100
    If medTCost <> "From System" Then
        setFlag = False
    End If
    Call setTotal
End Sub

Private Sub medPPE_GotFocus()
Call SetPanHelp(ActiveControl)
medPPE = Val(medPPE) * 100
End Sub

Private Sub medTCost_Change()
If medTCost = "From System" Then
    medEECost = "From System"
    medCompCost = "From System"
    medMECOST = "From System"
    medMCCOST = "From System"
Else
    If medTCost <> "From System" Then
        setFlag = False
    End If
    
    If glbCompSerial = "S/N - 2262W" Then   'To avoid the Runtime error 28 - Out of stack space
        If xWellingtonFlg = True Then
            Call setTotal
        End If
    Else
        Call setTotal
    End If
End If
End Sub

Private Sub medTCost_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medUnitCost_Change()
Call medCovAmount_Change
End Sub

Private Sub medUnitCost_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optActual_Click(Index As Integer, Value As Integer)
If optActual(1).Value = True Then
    txtPer.Enabled = True
    medUnitCost.Enabled = True
    'medCovAmount.Enabled = True  'jaddy 11/2/99
    medTCost = "From System"
    medTCost.Enabled = False
Else
    'txtPer = 0
    txtPer.Enabled = False
    'medUnitCost = 0
    medUnitCost.Enabled = False
    'medCovAmount.Enabled = False  'jaddy 11/2/99
    medTCost.Enabled = True
    medTCost = ""
    If IsNumeric(txtTCost.Text) Then medTCost = txtTCost.Text
    medEECost = 0
    medCompCost = 0
    medMECOST = 0
    medMCCOST = 0
End If

If optActual(0).Value = True And lblAP.Caption <> "A" Then
    lblAP.Caption = "A"
ElseIf optActual(1).Value = True And lblAP.Caption <> "P" Then
    lblAP.Caption = "P"
End If

End Sub

Private Sub optActual_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optActual_LostFocus(Index As Integer)
If optActual(0).Value = True And lblAP.Caption <> "A" Then
    lblAP.Caption = "A"
ElseIf optActual(1).Value = True And lblAP.Caption <> "P" Then
    lblAP.Caption = "P"
End If
End Sub

Private Sub scrControl_Change()
'panDetals.Top = 500 + vbxTrueGrid.Height - scrControl.Value * 2.5
'scrFrame.Top = 1920 - scrControl.Value
scrFrame.Top = 2640 - scrControl.Value
End Sub

Private Sub scrHScroll_Change()
scrFrame.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
End Sub

Private Sub txtAnComp_Change()
    If medCompCost <> txtAnComp.Text Then
        medCompCost = txtAnComp.Text
    End If
End Sub

Private Sub txtCovAmount_Change()
    If medCovAmount.Text <> txtCovAmount.Text Then
        medCovAmount.Text = txtCovAmount.Text
    End If
End Sub

Private Sub txtCovType_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtCovType_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtDWM_Change()
cmbDWM.ListIndex = -1
Select Case txtDWM
Case "D"
    cmbDWM.ListIndex = 0
Case "W"
    cmbDWM.ListIndex = 1
Case "M"
    cmbDWM.ListIndex = 2
End Select
End Sub

Private Sub txtPerorDoll_Change()
cmbPerOrDoll.ListIndex = -1
Select Case txtPerOrDoll
Case "D"
    cmbPerOrDoll.ListIndex = 0
Case "P"
    cmbPerOrDoll.ListIndex = 1
End Select
End Sub

Private Sub txtEeCost_Change()
    If medEECost <> txtEeCost Then
        medEECost = txtEeCost
    End If
End Sub

Private Sub txtMCost_Change()
    If medMCCOST <> txtMCost Then
        medMCCOST = txtMCost
    End If
End Sub

Private Sub txtMeCost_Change()
    If medMECOST <> txtMeCost Then
        medMECOST = txtMeCost
    End If
End Sub

Private Sub txtPer_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtPer_Change()
Call medCovAmount_Change
End Sub

Private Sub txtPreAftTax_Change()
If txtPreAftTax = "P" Then
    comPreAftTax.ListIndex = 0
ElseIf txtPreAftTax = "A" Then
    comPreAftTax.ListIndex = 1
Else
    comPreAftTax = ""
End If
End Sub

Private Sub txtRoundFactor_Change()
    Dim c As Long
    
    'If txtRoundFactor <> Val(comRndFactor.ItemData(comRndFactor.ListIndex)) Then
        For c = 0 To comRndFactor.ListCount - 1
            If comRndFactor.ItemData(c) = Val(txtRoundFactor) Then
                comRndFactor.ListIndex = c
                Exit For
            End If
        Next c
    'End If
    
End Sub

Private Sub txtSalDepn_Change()
'comSalDepn = IIf(txtSalDepn = "Y", "Yes", "No")
If txtSalDepn = "Y" Then
    comSalDepn = "Yes"
ElseIf txtSalDepn = "N" Then
    comSalDepn = "No"
Else
    comSalDepn = ""
End If
End Sub

Private Sub txtTAXBEN_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub CalcBeneCasey()
    Dim Total As Double, ECost As Double, CCost As Double
    
    If ChangingFields Then Exit Sub
    ChangingFields = True   ' Flag to tell other code we're changing fields, prevents loop
    ECost = Val(medMECOST.Text)
    CCost = Val(medMCCOST.Text)
    Total = ECost + CCost
    If Total = 0 Then
        medPPE.Text = "0"
        medPPComp.Text = "0"
        medEECost.Text = "0"
        medCompCost.Text = "0"
    Else
        medPPE.Text = ECost / Total
        medPPComp.Text = CCost / Total
        medEECost.Text = ECost * 12
        medCompCost.Text = CCost * 12
        medTCost.Text = Total * 12
    End If
    ChangingFields = False
End Sub

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

panDetals.Enabled = TF
cmdRecal.Enabled = TF
cmdRecalAll.Enabled = TF

If glbCElgin Then cmbDWM.Enabled = False

End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_BenefitGroupSetup 'gSec_Upd_Basic
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

Private Sub Form_Resize()

'If Me.Height >= vbxTrueGrid.Height + panDetals.Height + panControls.Height + 230 Then
'    scrControl.Value = 0
'    panDetals.Top = vbxTrueGrid.Height + 500 '240
'    scrControl.Visible = False
'    Exit Sub
'End If
'If Me.Height < vbxTrueGrid.Height + scrControl.Top + panControls.Height Then Exit Sub
'scrControl.Visible = True
'scrControl.Max = vbxTrueGrid.Height + panDetals.Height + panControls.Height - Me.Height + 250
'scrControl.Left = Me.Width - scrControl.Width - 120
'scrControl.Height = Me.Height - scrControl.Top - panControls.Height - 400
'panDetals.Width = Me.Width


panDetals.Height = 6480 '6135
If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= 10290 Then
        scrControl.Value = 0
        scrFrame.Top = 2640 '1920
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        If Me.Height < 7000 Then
            scrControl.Max = 4100
        Else
            scrControl.Max = 1900
        End If
        scrControl.Left = Me.Width - scrControl.Width - 120
        If Me.Height - scrControl.Top - 780 > 0 Then
            scrControl.Height = Me.Height - scrControl.Top - 780
        End If
    End If
    
    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 120
    'scrFrame.Height = Me.ScaleHeight - (scrHScroll.Height - 200)  '
    If Me.Width >= 10900 Then
        scrHScroll.Value = 0
        scrHScroll.Visible = False
    Else
        scrHScroll.Visible = True
        If Me.Width < 9000 Then
            scrHScroll.Max = 80
        Else
            scrHScroll.Max = 30
        End If
        scrHScroll.Top = Me.Height - 800
        scrHScroll.Width = Me.Width - 120
    End If
    scrFrame.Refresh
End If

End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, x%
Dim Response%

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

Dim xGroup, xCode, xCover
xGroup = Trim(clpBGroup.Text)
xCode = Trim(clpCode(4).Text)
xCover = Trim(txtCovType.Text)

gdbAdoIhr001.BeginTrans
RSDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

Call Display_Value

'Commenting this out for now as currently we are not deleting the Benefits from Employee's Benefit screen when a Benefit for a Group
'is delete. That part of the code has been commented by Frank since 2012 Ticket #22243 under updateBenefit() function under UpdData.bas.
'Release 8.1 - Send Email Notification but prompt the user first
'flgSendEmail = False
'If gsEMAIL_ONBENEFIT Then
'    Msg = "Do you want to send Email Notification automatically for each employees affected by this delete?"
'    Response% = MsgBox(Msg, vbYesNo + vbQuestion, "Benefit Delete Email Notification")
'    If Response% = vbNo Then
'        flgSendEmail = False
'    Else
'        flgSendEmail = True
'    End If
'End If

Call UPDBGroup(xGroup, xCode, xCover, GroupMasterDelete)

fglbNew = False

Exit Sub
Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRBENFTGROUP", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Sub cmdView_Click()
Dim RHeading As String
    
'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = "Benefit Group's "
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub

Sub cmdPrint_Click()
Dim RHeading As String
RHeading = "Benefit Group's "
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub

Sub cmdNew_Click()
Dim SQLQ As String, Msg$, x%
Dim Title$, DgDef As Variant, Response%

On Error GoTo AddN_Err

fglbNew = True

Call SET_UP_MODE

If Not gSec_BenefitGroupSetup Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Actn = "A"

Call Set_Control("B", Me)

OBCode = ""
OCOVER = ""
OTCOST = ""
OPremium = ""
OPPE = ""
OPCC = ""
OPPAMT = ""
OMAXDOL = ""
OEDate = ""
OPER = ""
OBAMT = ""
OUNITCOST = ""
'   If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
'lblCNum.Caption = "001"
medPPComp = 1
medPPE = 0
medTCost = 0
medUnitCost = 0
txtPer = 0
comSalDepn = "No"
medCovAmount = 0
optActual(0) = True
If glbCElgin Then txtDWM = "D"

If glbCompSerial = "S/N - 2214W" And optActual(0) Then
    medMCCOST.Enabled = True
    medMECOST.Enabled = True
End If
clpBGroup.SetFocus


Screen.MousePointer = DEFAULT

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "ATTEND", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Sub cmdOK_Click()
Dim rsBen As New ADODB.Recordset
Dim x
Dim xID
Dim Msg, Response%

On Error GoTo Add_Err

txtCovAmount.Text = Val(medCovAmount)
txtMCost.Text = Val(medMCCOST)
txtMeCost = Val(medMECOST)
txtEeCost = Val(medEECost)
txtAnComp.Text = Val(medCompCost)
txtTCost.Text = Val(medTCost)
txtDWM.Text = Left(cmbDWM, 1)
txtPerOrDoll = Left(cmbPerOrDoll, 1)
lblRound = IIf(optRound(0), "R", "N")

RSDATA.Requery

If Not chkMUBENEFITS() Then Exit Sub

'Release 8.1 - Send Email Notification but prompt the user first
flgSendEmail = False
If gsEMAIL_ONBENEFIT Then
    Msg = "Do you want to send Email Notification automatically for each employees affected by this update?"
    Response% = MsgBox(Msg, vbYesNo + vbQuestion, "Benefit Update Email Notification")
    If Response% = vbNo Then
        flgSendEmail = False
    Else
        flgSendEmail = True
    End If
End If

'rsDATA.Requery
If fglbNew Then RSDATA.AddNew

Call UpdUStats(Me)
        
gdbAdoIhr001.BeginTrans
Call Set_Control("U", Me, RSDATA)
RSDATA.Update
gdbAdoIhr001.CommitTrans

RSDATA.Resync

xID = RSDATA!BM_BENE_ID

Data1.Refresh

Data1.Recordset.Find "BM_BENE_ID=" & xID

If fglbNew Then
    Call UPDBGroup(Trim(clpBGroup.Text), Trim(clpCode(4).Text), Trim(txtCovType.Text), GroupMasterAdd)
Else
    Call UPDBGroup(Trim(clpBGroup.Text), Trim(clpCode(4).Text), Trim(txtCovType.Text), GroupMasterEdit)
End If

If glbCompSerial = "S/N - 2380W" Then Call CalcPP(Trim(clpCode(4).Text), Trim(clpBGroup.Text))


fglbNew = False

Call SET_UP_MODE

Exit Sub

Add_Err:
If Err = 3022 Then
    Data1.Recordset.CancelUpdate    ' no dups
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRBENFT", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub



Sub Display_Value()
Dim SQLQ
setFlag = True
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    RSDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
Else
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    SQLQ = "SELECT * FROM HR_BENEFITS_GROUP "
    SQLQ = SQLQ & " WHERE BM_BENE_ID = " & Data1.Recordset!BM_BENE_ID
    SQLQ = SQLQ & " ORDER BY BM_BENEFIT_GROUP, BM_BCODE, BM_EDATE"
    RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

    If RSDATA.EOF Or RSDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, RSDATA)
    chkCost.Value = IIf(GetCost, 1, 0)
End If
setFlag = False
Call SET_UP_MODE
End Sub

Function GetCost()
Dim rsCost As New ADODB.Recordset
Dim SQLQ
On Error Resume Next

If clpCode(4).Text = "OMER" Then
'Ticket #20872 Franks 09/27/2011 - make this for all customers
''If glbGP And clpCode(4).Text = "OMER" Then
    ''Ticket #17643
    ''If Benefit Code = OMER and there is a OMERS Formula record for the current year, the Benefit Group record will be setup like this
    ''Assume the "Use Cost Table" is checked.
    'If glbCompSerial = "S/N - 2172W" Then
    '    GetCost = OMER_UseCostTable
    'End If
    GetCost = OMER_UseCostTable
Else
    'SQLQ = "SELECT CU_BGROUP,CU_BCODE FROM HR_BENEFIT_COST "
    SQLQ = "SELECT CU_BENEFIT_GROUP,CU_BCODE FROM HR_BENEFIT_COST "
    SQLQ = SQLQ & " WHERE CU_BCODE='" & clpCode(4).Text & "'"
    SQLQ = SQLQ & " AND (CU_BENEFIT_GROUP='" & clpBGroup.Text & "' OR CU_BENEFIT_GROUP='' OR CU_BENEFIT_GROUP IS NULL)"
    rsCost.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If rsCost.EOF Then
        GetCost = False
    Else
        GetCost = True
    End If
End If

End Function

Sub cmdCancel_Click()
Dim x, bk

On Error GoTo Can_Err

Call Display_Value

fglbNew = False

Call SET_UP_MODE

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_BENFTS_GROUP", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub txtTAXBEN_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtTCost_Change()
    If medTCost.Text <> txtTCost.Text Then
        medTCost.Text = txtTCost.Text
    End If
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
        
        SQLQ = "SELECT  * FROM HR_BENEFITS_GROUP"
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
Call Set_SalCover
Call setTotal
End Sub

Private Sub INIData()
Dim rsTD As New ADODB.Recordset
Dim SQLQ
SQLQ = "SELECT TB_NAME, TB_KEY, TB_DESC FROM HRTABL WHERE TB_NAME = 'BGMF'"
rsTD.Open SQLQ, gdbAdoIhr001, adOpenStatic
cmbGroups.Clear
cmbGroups.AddItem "Show All Groups"
Do Until rsTD.EOF
    cmbGroups.AddItem rsTD("TB_KEY") & " - " & rsTD("TB_DESC")
    rsTD.MoveNext
Loop
cmbGroups.ListIndex = 0
End Sub

Private Sub cmbGroups_Click()
Dim SQLQ

If cmbGroups = "Show All Groups" Then
    fglbBGroup = "ALLGROUP"
    Data1.RecordSource = "SELECT * FROM HR_BENEFITS_GROUP ORDER BY BM_BENEFIT_GROUP, BM_BCODE, BM_EDATE"
Else
    fglbBGroup = Left(cmbGroups.Text, InStr(cmbGroups.Text, "-") - 1)
    SQLQ = "SELECT * FROM HR_BENEFITS_GROUP "
    SQLQ = SQLQ & " WHERE BM_BENEFIT_GROUP =" & "'" & fglbBGroup & "'"
    SQLQ = SQLQ & " ORDER BY BM_BENEFIT_GROUP, BM_BCODE, BM_EDATE"
    Data1.RecordSource = SQLQ
End If
Data1.Refresh

End Sub

Sub setTotal()

If setFlag = False Then
    If glbCompSerial = SN_CASEY Then Exit Sub

    If glbCompSerial = "S/N - 2262W" Then 'Wellington - Ticket #10718
        If clpCode(4).Text = "5ADB" Or clpCode(4).Text = "5GRB" Or clpCode(4).Text = "5LTB" Or _
            clpCode(4).Text = "6ADB" Or clpCode(4).Text = "6GRB" Or clpCode(4).Text = "6LTB" Or _
            clpCode(4).Text = "8GRB" Or clpCode(4).Text = "4ADW" Or clpCode(4).Text = "4GRW" Or _
            clpCode(4).Text = "4LTW" Or clpCode(4).Text = "1GRB" Then
            
            xWellingtonFlg = False
            medTCost = Round(Val(medTCost), 2)
        Else
            xWellingtonFlg = True
        End If
    End If
    
    medEECost = Val(medTCost) * Val(medPPE)
    medCompCost = Val(medTCost) * Val(medPPComp)
    
    'changed by Bryan Ticket#10444
    If glbCompSerial = "S/N - 2262W" Then 'Wellington
        If Right(Round(CStr(Val(medEECost) / 12), 3), 1) = 5 Then
            If clpCode(4).Text = "5ADB" Or clpCode(4).Text = "5GRB" Or clpCode(4).Text = "5LTB" Or _
                clpCode(4).Text = "6ADB" Or clpCode(4).Text = "6GRB" Or clpCode(4).Text = "6LTB" Or _
                clpCode(4).Text = "8GRB" Or clpCode(4).Text = "4ADW" Or clpCode(4).Text = "4GRW" Or _
                clpCode(4).Text = "4LTW" Or clpCode(4).Text = "1GRB" Then
                
                medMECOST = Val(medEECost) / 12
            Else
                medMECOST = Round((Val(medEECost) / 12) - 0.005, 2)
            End If
        Else
            medMECOST = Val(medEECost) / 12
        End If
        If Right(Round(CStr(Val(medCompCost) / 12), 3), 1) = 5 Then
            If clpCode(4).Text = "5ADB" Or clpCode(4).Text = "5GRB" Or clpCode(4).Text = "5LTB" Or _
                clpCode(4).Text = "6ADB" Or clpCode(4).Text = "6GRB" Or clpCode(4).Text = "6LTB" Or _
                clpCode(4).Text = "8GRB" Or clpCode(4).Text = "4ADW" Or clpCode(4).Text = "4GRW" Or _
                clpCode(4).Text = "4LTW" Or clpCode(4).Text = "1GRB" Then
                
                medMCCOST = Val(medCompCost) / 12
            Else
                medMCCOST = Round((Val(medCompCost) / 12) - 0.005, 2)
            End If
        Else
            medMCCOST = Val(medCompCost) / 12
        End If
    ElseIf (glbCompSerial <> "S/N - 2335W") Or (glbCompSerial = "S/N - 2335W" And InStr(1, glbSeleDiv, "HSV") > 0) Then
        medMECOST = Round(Val(medEECost) / 12, 2)
        medMCCOST = Round(Val(medCompCost) / 12, 2)
    Else
        medMECOST = 0
        medMCCOST = 0
    End If
    setFlag = True
End If

End Sub

Private Function AUDITBENF(xEmpnbr, xCode, xAmount) As Boolean
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String, ACTX As String
Dim strFields As String
On Error GoTo AUDIT_ERR
AUDITBENF = False
ACTX = "M"
rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEmpnbr, gdbAdoIhr001, adOpenKeyset

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
rsTB.Close

'rsTB.Open "SELECT * FROM HRBENFT WHERE BF_EMPNBR=" & xEmpNbr, gdbAdoIhr001, adOpenKeyset, adCmdText
'If rsTB.EOF Then GoTo MODNOUPD

'strfields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_COVER, AU_EDATE, AU_MAXDOL, AU_PPAMT, "
strFields = strFields & "AU_MTHCCOST, AU_MTHECOST, AU_BCODE, AU_BNAME, AU_BRELATE, AU_BDOB, AU_TAXBEN, AU_COVER, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, "
strFields = strFields & "AU_OLDPPMT, AU_MAXDOL, AU_EDATE, AU_PER, AU_BAMT, AU_UNITCOST, AU_BCODE, AU_BNAME, "
strFields = strFields & "AU_BRELATE, AU_BDOB, AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False

MODUPD:

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv
rsTA("AU_PPAMT") = xAmount 'rsTB("BF_PPAMT") AU_BCODE
rsTA("AU_BCODE") = xCode

Dim rsEmp As New ADODB.Recordset
Dim SQLQ
SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmp.EOF Then
    If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
End If
rsEmp.Close

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = xEmpnbr 'glbLEE_ID
rsTA("AU_LDATE") = Date
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
rsTA.Update

MODNOUPD:
AUDITBENF = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

Private Sub CalcPP(Optional xCode As String, Optional xGroup As String)
    Dim rs As New ADODB.Recordset
    Dim rsIn As New ADODB.Recordset
    Dim SQLQ As String, WSQLQ As String
    Dim x As Boolean
    Dim I As Long, xTot As Long, oPayP As Double
    
    WSQLQ = ""
    If IsEmpty(xCode) = False Then
        If xCode <> "" Then
            WSQLQ = "WHERE BF_BCODE='" & xCode & "' and BF_GROUP='" & xGroup & "'"
        End If
    End If
    
    SQLQ = "SELECT BF_EMPNBR, BF_PPAMT, BF_MTHECOST, BF_ECOST, BF_GROUP, BF_BCODE, BF_LUSER, BF_LDATE, BF_LTIME FROM HRBENFT "
    SQLQ = SQLQ & WSQLQ
    
    rs.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic, adCmdText
    I = 0
    If Not rs.EOF Then
        xTot = rs.RecordCount
    End If
    MDIMain.panHelp(0).FloodType = 1
    Do While Not rs.EOF
    'If rs.EOF = False And rs.BOF = False Then
        MDIMain.panHelp(0).FloodPercent = (I / xTot) * 100: I = I + 1
        oPayP = rs("BF_PPAMT")
        Select Case rs("BF_GROUP")
        Case "GHON", "GHQC", "CAMPBELL", "CAMPBC", "GHQC113", "GHON113", "CAMPBC113"   'Ticket #18963, Ticket #24537 - more codes
            rs("BF_PPAMT") = rs("BF_ECOST") / 52
        Case Else
            rs("BF_PPAMT") = rs("BF_MTHECOST") / 2
        End Select
        rs("BF_LUSER") = glbUserID
        rs("BF_LTIME") = Time$
        rs("BF_LDATE") = Format(Now, "SHORT DATE")
        rs.Update
        If oPayP <> rs("BF_PPAMT") Then
            x = AUDITBENF(rs("BF_EMPNBR"), rs("BF_BCODE"), rs("BF_PPAMT"))
        End If
        rs.MoveNext
    'End If
    Loop
    rs.Close

    MDIMain.panHelp(0).FloodType = 0
End Sub

Private Sub EmailNotification(xUpdType, xEmpnbr, xBCode, xEDate)
                            
    'Benefits Added
    If xUpdType = "ADD" Then
        MailBody = "The New Benefit:" & vbCrLf & vbCrLf
        MailBody = MailBody & "Employee #: " & xEmpnbr & vbCrLf
        MailBody = MailBody & "Name: " & GetEmpData(xEmpnbr, "ED_SURNAME") & ", " & GetEmpData(xEmpnbr, "ED_FNAME") & vbCrLf
        MailBody = MailBody & "New Benefit: " & GetTABLDesc("BNCD", xBCode) & vbCrLf
        MailBody = MailBody & "Effective Date: " & Format(CVDate(xEDate), "SHORT DATE") & vbCrLf
    End If
                            
    'Benefits Updated
    If xUpdType = "UPDATE" Then
        MailBody = "The Updated Benefit:" & vbCrLf & vbCrLf
        MailBody = MailBody & "Employee #: " & xEmpnbr & vbCrLf
        MailBody = MailBody & "Name: " & GetEmpData(xEmpnbr, "ED_SURNAME") & ", " & GetEmpData(xEmpnbr, "ED_FNAME") & vbCrLf
        MailBody = MailBody & "Updated Benefit: " & GetTABLDesc("BNCD", xBCode) & vbCrLf
        MailBody = MailBody & "Effective Date: " & Format(CVDate(xEDate), "SHORT DATE") & vbCrLf
    End If
    
    'Benefits Deleted
    If xUpdType = "DELETE" Then
        MailBody = "The Deleted Benefit:" & vbCrLf & vbCrLf
        MailBody = MailBody & "Employee #: " & xEmpnbr & vbCrLf
        MailBody = MailBody & "Name: " & GetEmpData(xEmpnbr, "ED_SURNAME") & ", " & GetEmpData(xEmpnbr, "ED_FNAME") & vbCrLf
        MailBody = MailBody & "Deleted Benefit: " & GetTABLDesc("BNCD", xBCode) & vbCrLf
        MailBody = MailBody & "Effective Date: " & Format(CVDate(xEDate), "SHORT DATE") & vbCrLf
    End If
    
    Call imgEmail_Click(xUpdType, xEmpnbr)
    
    Screen.MousePointer = DEFAULT

End Sub

Public Sub imgEmail_Click(xType, xEmpnbr)
Dim xEmail
Dim xToEmail As String

On Error GoTo Email_Err

        If Not UserEmailExist Then
            Exit Sub
        End If
        xEmail = GetCurEmpEmail
        
        If Len(xEmail) > 0 Then
            'Ticket #18090 - begin
            If glbCompSerial = "S/N - 2382W" Then  'Samuel
                xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT", glbLEE_ID)
                If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                    xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT")
                End If
            Else
                'Ticket #20317 - More Emails for everyone
                xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT", glbLEE_ID)
                If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                    xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT")
                End If
            End If
            'Ticket #18090 - end
            
            frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONBENEFIT")
            
            If glbCompSerial = "S/N - 2382W" Then  'Samuel
            Else
                frmSendEmail.txtCC.Text = xEmail
            End If
                        
            'Email Subject line based on the Type of Email
            If xType = "DELETE" Then
                frmSendEmail.txtSubject.Text = "info:HR Benefit Delete Notice - " & GetEmpData(xEmpnbr, "ED_SURNAME") & ", " & GetEmpData(xEmpnbr, "ED_FNAME")
            ElseIf xType = "UPDATE" Then
                frmSendEmail.txtSubject.Text = "info:HR Benefit Update Notice - " & GetEmpData(xEmpnbr, "ED_SURNAME") & ", " & GetEmpData(xEmpnbr, "ED_FNAME")
            Else
                frmSendEmail.txtSubject.Text = "info:HR New Benefit Notice - " & GetEmpData(xEmpnbr, "ED_SURNAME") & ", " & GetEmpData(xEmpnbr, "ED_FNAME")
            End If
            frmSendEmail.txtBody.Text = MailBody
            
            'Not showing the Email Send window as this is a mass update
            'frmSendEmail.Show 1
            frmSendEmail.cmdSend_Click
        Else
            'If Len(glbLEE_SName) = 0 Then
            '    MsgBox "There is no email on Status/Dates screen for employee. "
            'Else
            '    MsgBox "There is no email on Status/Dates screen for employee " & GetEmpData(xEmpNbr, "ED_SURNAME") & ", " & GetEmpData(xEmpNbr, "ED_FNAME")
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

Private Sub WFCScreenSetup()
    cmdRecalAll.Visible = False
    cmdRecaSalDep.Left = 1740
    cmdRecaSalDep.Visible = True
End Sub
