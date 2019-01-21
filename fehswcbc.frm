VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEHSWCBC 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   " WSIB Cost "
   ClientHeight    =   10650
   ClientLeft      =   150
   ClientTop       =   180
   ClientWidth     =   11580
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
   ScaleHeight     =   10650
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "CC_RateGrp"
      Height          =   285
      Index           =   1
      Left            =   2250
      TabIndex        =   19
      Tag             =   "00-Rate Group Code"
      Top             =   7260
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECGP"
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fehswcbc.frx":0000
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "fehswcbc.frx":0014
      TabIndex        =   0
      Top             =   600
      Width           =   9015
   End
   Begin INFOHR_Controls.DateLookup dlpFromTo 
      DataField       =   "CC_TDATE"
      Height          =   285
      Index           =   1
      Left            =   6815
      TabIndex        =   4
      Tag             =   "42-Cost To Date"
      Top             =   3562
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpFromTo 
      DataField       =   "CC_FDATE"
      Height          =   285
      Index           =   0
      Left            =   6815
      TabIndex        =   3
      Tag             =   "42-Cost From Date"
      Top             =   3210
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpSDate 
      DataField       =   "CC_STMTDT"
      Height          =   285
      Left            =   2230
      TabIndex        =   2
      Tag             =   "41-WSIB Cost Statement Date"
      Top             =   3210
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   10320
      Top             =   9720
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
      Height          =   660
      Left            =   0
      TabIndex        =   50
      Top             =   9990
      Width           =   11580
      _Version        =   65536
      _ExtentX        =   20426
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
   Begin VB.ComboBox cmbWCB 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "Choose WSIB this cost is related to"
      Top             =   2700
      Width           =   8895
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CC_LDATE"
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
      Left            =   2640
      MaxLength       =   25
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   8700
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CC_LTIME"
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
      Left            =   4440
      MaxLength       =   25
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8700
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CC_LUSER"
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
      Left            =   6240
      MaxLength       =   25
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8700
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   11580
      _Version        =   65536
      _ExtentX        =   20426
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
         Left            =   7440
         TabIndex        =   56
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   135
         Width           =   720
      End
   End
   Begin MSMask.MaskEdBox medTemp 
      DataField       =   "CC_TEMPC"
      Height          =   285
      Left            =   2550
      TabIndex        =   5
      Tag             =   "20-Cost for Temporary Compensation"
      Top             =   3562
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medPension 
      DataField       =   "CC_PENSION"
      Height          =   285
      Left            =   2550
      TabIndex        =   6
      Tag             =   "20-Cost related to Pension"
      Top             =   3914
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medRehab 
      DataField       =   "CC_REHAB"
      Height          =   285
      Left            =   2550
      TabIndex        =   8
      Tag             =   "20-Cost related to Rehabilitation"
      Top             =   4266
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medNonE 
      DataField       =   "CC_NONECO"
      Height          =   285
      Left            =   2550
      TabIndex        =   10
      Tag             =   "20-Non Economic Loss Award"
      Top             =   4618
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medFutureE 
      DataField       =   "CC_FUTECOA"
      Height          =   285
      Left            =   2550
      TabIndex        =   12
      Tag             =   "20-Cost related to Loss of Earning Pension Award"
      Top             =   4970
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medRetPen 
      DataField       =   "CC_RETIREP"
      Height          =   285
      Left            =   2550
      TabIndex        =   13
      Tag             =   "20-Cost related to Retirement Pension"
      Top             =   5322
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medReEmp 
      DataField       =   "CC_REEMPL"
      Height          =   285
      Left            =   2550
      TabIndex        =   15
      Tag             =   "20-Cost related to Re-Employment"
      Top             =   5680
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medPartial2 
      DataField       =   "CC_PPD1"
      Height          =   285
      Left            =   2565
      TabIndex        =   17
      Tag             =   "20-Cost related to PPD Supplement # 135/2"
      Top             =   6525
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medPartial4 
      DataField       =   "CC_PPD2"
      Height          =   285
      Left            =   2565
      TabIndex        =   18
      Tag             =   "20-Cost related to PPD Supplement # 135/4"
      Top             =   6892
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medHealth 
      DataField       =   "CC_HEALTH"
      Height          =   285
      Left            =   7125
      TabIndex        =   7
      Tag             =   "20-Cost of Health Care"
      Top             =   3914
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medSurvivor 
      DataField       =   "CC_SURVBF"
      Height          =   285
      Left            =   7125
      TabIndex        =   9
      Tag             =   "20-Survivor Benefit Costs"
      Top             =   4266
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medOther 
      DataField       =   "CC_OTHER"
      Height          =   285
      Left            =   7125
      TabIndex        =   11
      Tag             =   "20-Other (user) specified costs"
      Top             =   4618
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medFutureES 
      DataField       =   "CC_FUTECOS"
      Height          =   285
      Left            =   7125
      TabIndex        =   14
      Tag             =   "20-Future Economic Loss Supplemental cost"
      Top             =   5322
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medSIEF 
      DataField       =   "CC_SIEF_PC"
      Height          =   285
      Left            =   7125
      TabIndex        =   16
      Tag             =   "20-SIEF Percentage"
      Top             =   5680
      Width           =   1485
      _ExtentX        =   2619
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
      Left            =   5880
      TabIndex        =   55
      Top             =   8520
      Width           =   1095
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
      Left            =   6960
      TabIndex        =   54
      Top             =   8520
      Width           =   1935
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
      Left            =   2520
      TabIndex        =   53
      Top             =   8520
      Width           =   975
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
      Left            =   3480
      TabIndex        =   52
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SIEF Percentage"
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
      Left            =   5295
      TabIndex        =   51
      Top             =   5725
      Width           =   1395
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Permanent Partial Disability: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   49
      Top             =   6240
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      Height          =   195
      Left            =   5985
      TabIndex        =   48
      Top             =   3607
      Width           =   705
   End
   Begin VB.Label lblFutureES 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Future  Economic  Loss  Supplement"
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
      Left            =   4095
      TabIndex        =   47
      Top             =   5367
      Width           =   2595
   End
   Begin VB.Label lblWCBNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "lblWCBNo"
      DataField       =   "CC_WCBNBR"
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
      Height          =   225
      Left            =   5880
      TabIndex        =   46
      Top             =   4995
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label lblCase 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      DataField       =   "CC_CASE"
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
      Left            =   5160
      TabIndex        =   45
      Tag             =   "11-Unique ID of Incident"
      Top             =   5015
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      DataSource      =   "Data2"
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
      Height          =   225
      Left            =   4200
      TabIndex        =   44
      Top             =   5000
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label lblOther 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Other"
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
      Left            =   6300
      TabIndex        =   43
      Top             =   4663
      Width           =   390
   End
   Begin VB.Label lblSurv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Survivor's  Benefits"
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
      Left            =   5340
      TabIndex        =   42
      Top             =   4311
      Width           =   1350
   End
   Begin VB.Label lblHCare 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Health  Care"
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
      Left            =   5805
      TabIndex        =   41
      Top             =   3959
      Width           =   885
   End
   Begin VB.Label lblFDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5805
      TabIndex        =   40
      Top             =   3255
      Width           =   885
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   210
      TabIndex        =   39
      Top             =   7305
      Width           =   825
   End
   Begin VB.Label lblPerm4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Supplement  135 / 4"
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
      Left            =   210
      TabIndex        =   38
      Top             =   6937
      Width           =   1455
   End
   Begin VB.Label lblPerm2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Supplement  135 / 2"
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
      Left            =   210
      TabIndex        =   37
      Top             =   6570
      Width           =   1455
   End
   Begin VB.Label lblReE 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Employment"
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
      Left            =   210
      TabIndex        =   36
      Top             =   5725
      Width           =   1110
   End
   Begin VB.Label lblRetPen 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Retirement  Pension"
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
      Left            =   210
      TabIndex        =   35
      Top             =   5367
      Width           =   1425
   End
   Begin VB.Label lblFutureE 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Loss of Earning Pension Award"
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
      Left            =   210
      TabIndex        =   34
      Top             =   5015
      Width           =   2205
   End
   Begin VB.Label lblNonE 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Non-Economic Loss  Award"
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
      Left            =   210
      TabIndex        =   33
      Top             =   4663
      Width           =   1965
   End
   Begin VB.Label lblRehab 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rehabilitation  Costs"
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
      Left            =   210
      TabIndex        =   32
      Top             =   4311
      Width           =   1440
   End
   Begin VB.Label lblPen 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pension"
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
      Left            =   210
      TabIndex        =   31
      Top             =   3959
      Width           =   570
   End
   Begin VB.Label lblTemp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Temporary  Compensation"
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
      Left            =   210
      TabIndex        =   30
      Top             =   3607
      Width           =   1845
   End
   Begin VB.Label lblStatement 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Statement  Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   210
      TabIndex        =   29
      Top             =   3255
      Width           =   1395
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "CC_EMPNBR"
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
      Left            =   1710
      TabIndex        =   27
      Top             =   8820
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "CC_COMPNO"
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
      Left            =   315
      TabIndex        =   28
      Top             =   8820
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEHSWCBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X%
Dim fglbNew
Dim wcb() As Variant
Dim fglbComboWCB% ' is there data in combo box? were wcbs found?
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control

Private Function chkHSWCBCs()

Dim SQLQ As String, Msg As String, dd&, tdat As Variant

chkHSWCBCs = False

On Error GoTo chkHSWCBCs_Err

If Len(dlpSDate.Text) >= 1 Then
    If Not IsDate(dlpSDate.Text) Then
        MsgBox "Statement Date is not a valid date."
        dlpSDate.SetFocus
        Exit Function
    End If
Else
    MsgBox "Statement Date is required."
    dlpSDate.SetFocus
    Exit Function
End If

If Len(dlpFromTo(0).Text) >= 1 Then
    If Not IsDate(dlpFromTo(0).Text) Then
        MsgBox "From date is not a valid date."
        dlpFromTo(0).SetFocus
        Exit Function
    End If
Else
    MsgBox "From date is required."
    dlpFromTo(0).SetFocus
    Exit Function
End If

If Len(dlpFromTo(1).Text) >= 1 Then
    If Not IsDate(dlpFromTo(1).Text) Then
        MsgBox "To date is not a valid date."
        dlpFromTo(1).SetFocus
        Exit Function
    End If
Else
    MsgBox "To date is required."
    dlpFromTo(1).SetFocus
    Exit Function
End If

dd& = DateDiff("d", CVDate(dlpFromTo(0).Text), CVDate(dlpFromTo(1).Text))

If dd& < 0 Then
    MsgBox "From date must be earlier than To date."
    dlpFromTo(0).SetFocus
    Exit Function
End If

tdat = wcb(cmbWCB.ListIndex + 1, 3)

dd& = DateDiff("d", CVDate(dlpSDate.Text), CVDate(tdat))

If dd& > 0 Then
    MsgBox "Statement date must be later than File date."
    dlpSDate.SetFocus
    Exit Function
End If

If Len(clpCode(1).Text) >= 1 And clpCode(1).Caption = "Unassigned" Then
    MsgBox "Rate Group code must be valid"
     clpCode(1).SetFocus
    Exit Function
End If


'If Len(medTemp) = 0 Then medTemp = 0#
'If Len(medHealth) = 0 Then medHealth = 0
'If Len(medPension) = 0 Then medPension = 0
'If Len(medSurvivor) = 0 Then medSurvivor = 0
'If Len(medRehab) = 0 Then medRehab = 0
'If Len(medNonE) = 0 Then medNonE = 0
'If Len(medFutureES) = 0 Then medFutureES = 0
'If Len(medFutureE) = 0 Then medFutureE = 0
'If Len(medRetPen) = 0 Then medRetPen = 0
'If Len(medReEmp) = 0 Then medReEmp = 0
'If Len(medPartial2) = 0 Then medPartial2 = 0
'If Len(medPartial4) = 0 Then medPartial4 = 0
'If Len(medOther) = 0 Then medOther = 0
Dim Ctrol As Control

Set Ctrol = medTemp: If Not chkNumeric(Ctrol) Then Exit Function
Set Ctrol = medHealth: If Not chkNumeric(Ctrol) Then Exit Function
Set Ctrol = medPension: If Not chkNumeric(Ctrol) Then Exit Function
Set Ctrol = medSurvivor: If Not chkNumeric(Ctrol) Then Exit Function
Set Ctrol = medRehab: If Not chkNumeric(Ctrol) Then Exit Function
Set Ctrol = medNonE: If Not chkNumeric(Ctrol) Then Exit Function
Set Ctrol = medFutureES: If Not chkNumeric(Ctrol) Then Exit Function
Set Ctrol = medFutureE: If Not chkNumeric(Ctrol) Then Exit Function
Set Ctrol = medRetPen: If Not chkNumeric(Ctrol) Then Exit Function
Set Ctrol = medReEmp: If Not chkNumeric(Ctrol) Then Exit Function
Set Ctrol = medPartial2: If Not chkNumeric(Ctrol) Then Exit Function
Set Ctrol = medPartial4: If Not chkNumeric(Ctrol) Then Exit Function
Set Ctrol = medOther: If Not chkNumeric(Ctrol) Then Exit Function

'added by Bryan 24/oct/05 Ticket#2259
If cmbWCB.ListIndex > -1 Then
    Dim rs As New ADODB.Recordset
    
    SQLQ = "Select EC_WCBCDTE "
    If glbtermopen Then
        SQLQ = SQLQ & " FROM Term_HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " and EC_WCBNBR = '" & wcb(cmbWCB.ListIndex + 1, 1) & "'"
    Else
        SQLQ = SQLQ & " FROM HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & " where EC_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " and EC_WCBNBR = '" & wcb(cmbWCB.ListIndex + 1, 1) & "'"
    End If
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not IsNull(rs("EC_WCBCDTE")) And IsDate(rs("EC_WCBCDTE")) Then
        If DateDiff("d", dlpSDate.Text, rs("EC_WCBCDTE")) < 0 Then
            MsgBox "Statements cannot be entered after the claim is closed"
            Exit Function
        End If
    End If
    rs.Close
    Set rs = Nothing
End If
chkHSWCBCs = True

Exit Function

chkHSWCBCs_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSWCBCs", "HROHSCOS", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function



Private Sub cmbWCB_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbWCB_LostFocus()
'Dim X%
'X% = cmbWCB.ListIndex
'If X% >= 0 Then
'    lblWCBNo.Caption = wcb(X% + 1, 1)
'    lblCase.Caption = wcb(X% + 1, 2)
'End If

End Sub


'Private Sub cmdCAction_Click()
'frmEHSCorrective.Show
'Unload Me
'End Sub

Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err

'Data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh
fglbNew = False
''' Sam add July 2002 * Remove Binding Control

Call Display_Value
Data1.Refresh

'Call ST_UPD_MODE(True)  ' reset screen's attributes
'Call SET_UP_MODE


Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HROHSCOS", "Cancel")
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

End Sub

'Private Sub cmdClose_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdContact_Click()
'frmEHSContact.Show
'Unload Me
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, X%

On Error GoTo Del_Err
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "No Records Found"
    Exit Sub
End If

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

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
fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HROHSCOS", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdDelete_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdIncident_Click()
'frmEHSINCIDENT.Show
'Unload Me
'End Sub

'Private Sub cmdInjLoc_Click()
'frmEHSINJURY.Show
'Unload Me
'End Sub

Sub cmdModify_Click()
Dim X%

'If Not gSec_Upd_Health_Safety Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If

On Error GoTo Mod_Err

Call SET_UP_MODE
'Call ST_UPD_MODE(True)
dlpSDate.SetFocus

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HROHSCOS", "Modify")
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

Sub cmdNew_Click()
Dim SQLQ As String

'If Not gSec_Upd_Health_Safety Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If
fglbNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
On Error GoTo AddN_Err
'If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'    Me.vbxTrueGrid.Enabled = False
'    Data1.RecordSource = "HROHSCOS"
'    Data1.Refresh
'    fglbEmptyNew = True
'End If

'Data1.Recordset.AddNew
''' Sam add July 2002 * Remove Binding Control
Call Set_Control("B", Me)


If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID

wcb(1, 1) = ""
fglbComboWCB% = popComboWCB()

lblCNum.Caption = "001"
medTemp = 0
medHealth = 0
medPension = 0
medSurvivor = 0
medRehab = 0
medNonE = 0
medFutureES = 0
medFutureE = 0
medRetPen = 0
medReEmp = 0
medPartial2 = 0
medPartial4 = 0
medOther = 0
cmbWCB.ListIndex = 0
cmbWCB.SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HROHSCOS", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub CmdNew_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim X%

On Error GoTo Add_Err

If Not chkHSWCBCs() Then Exit Sub
rsDATA.Requery
If fglbNew Then rsDATA.AddNew
Call UpdUStats(Me) ' update user's stats (who did it and when)

rsDATA!CC_FUTECOA = medFutureE

X% = cmbWCB.ListIndex
If X% >= 0 Then
    lblWCBNo.Caption = wcb(X% + 1, 1)
    lblCase.Caption = wcb(X% + 1, 2)
End If


If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    Call Set_Control("U", Me, rsDATA)
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    Call Set_Control("U", Me, rsDATA)
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    Data1.Refresh
End If

'Call ST_UPD_MODE(True)
Call SET_UP_MODE
If NextFormIF(" WSIB Cost") Then
    Call cmdNew_Click
End If
fglbNew = False
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HROHSCOS", "Update")
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

RHeading = lblEEName & "'s WSIB Cost"
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

RHeading = lblEEName & "'s WSIB Cost"
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

'Private Sub cmdWCBMed_Click()
'frmEHSWCB.Show
'Unload Me
'End Sub


Function EERetrieve()
Dim SQLQ As String
EERetrieve = False
On Error GoTo EERError
If glbtermopen Then
    SQLQ = "SELECT * from Term_HROHSCOS "
    SQLQ = SQLQ & "WHERE TERM_SEQ = " & glbTERM_Seq & " ORDER BY CC_WCBNBR,CC_STMTDT"
Else
    SQLQ = "SELECT * from HROHSCOS "
    SQLQ = SQLQ & "WHERE CC_EMPNBR = " & glbLEE_ID & " ORDER BY CC_WCBNBR,CC_STMTDT"
End If
Data1.RecordSource = SQLQ
Data1.Refresh

fglbComboWCB% = popComboWCB()

EERetrieve = True
Exit Function
EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "OCH Retrieve", "HROHSCOS", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
Exit Function
End Function
Private Sub Form_Activate()
glbOnTop = "FRMEHSWCBC"
End Sub

Private Sub Form_GotFocus()
glbOnTop = "FRMEHSWCBC"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

ReDim wcb(1, 3) 'laura nov 14, 1997
glbOnTop = "FRMEHSWCBC"
If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
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
fglbComboWCB% = popComboWCB()

Screen.MousePointer = HOURGLASS
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "WSIB Cost - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = ShowEmpnbr(lblEEID)

'Release 8.1
If glbCompSerial = "S/N - 2417W" Then
    lblOther.Caption = "Physician/Admin Fees"
    vbxTrueGrid.Columns(16).Caption = "Physician/Admin Fees $"
Else
    lblOther.Caption = "Other"
    vbxTrueGrid.Columns(16).Caption = "Other $"
End If

Call ST_UPD_MODE(True)

If Not gSec_Upd_HSCost Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
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
Set frmEHSWCBC = Nothing ' carmen may 00
Call NextForm
End Sub


Private Sub medFutureE_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Private Sub medFutureES_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub medHealth_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Private Sub medNonE_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medNonE_ValidationError(InvalidText As String, StartPosition As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medOther_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub medPartial2_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Private Sub medPartial4_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub medPension_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub medReEmp_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Private Sub medRehab_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub medRetPen_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub medSurvivor_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub medTemp_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Private Function popComboWCB() ' returns number of records found
Dim snapWCBs As New ADODB.Recordset
Dim X%, SQLQ As String, cmbItems$, Msg$

On Error GoTo popComboWCB_Err

cmbWCB.Clear

fglbComboWCB% = 0         ' if not found - no depts

SQLQ = "Select EC_EMPNBR,EC_WCBNBR,EC_CASE,EC_OCCDATE,EC_TYPE,EC_CLASS,EC_CAUSECD,EC_WCBFDTE,EC_WCBFDTE,EC_WCBCDTE "
If glbtermopen Then
    SQLQ = SQLQ & " FROM Term_HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " and EC_WCBNBR <> ' ' "
    snapWCBs.Open SQLQ, gdbAdoIhr001X, adOpenStatic
Else
    SQLQ = SQLQ & " FROM HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " where EC_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " and EC_WCBNBR <> ' ' "
    snapWCBs.Open SQLQ, gdbAdoIhr001, adOpenStatic
End If
    



If snapWCBs.EOF And snapWCBs.BOF Then
    cmbWCB.AddItem "No Worker's Safety Insurance Board Records Found "
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdPrint.Enabled = False
    snapWCBs.Close
    Exit Function
End If
'cmdModify.Enabled = True
'cmdNew.Enabled = True
'cmdPrint.Enabled = True

snapWCBs.MoveLast
X% = snapWCBs.RecordCount
ReDim wcb(X%, 3) As Variant
X% = 0

'EOF?
snapWCBs.MoveFirst
While Not snapWCBs.EOF
    X% = X% + 1
    wcb(X%, 1) = CStr(snapWCBs("EC_WCBNBR"))
    cmbItems$ = "WSIB # " & wcb(X%, 1) & " - "
    wcb(X%, 2) = CStr(snapWCBs("EC_CASE"))
    cmbItems$ = cmbItems$ & "Case # " & wcb(X%, 2) & " - "

    If Not IsNull(snapWCBs("EC_OCCDATE")) Then
        cmbItems$ = cmbItems$ & "occurred " & CStr(snapWCBs("EC_OCCDATE")) & " - "
    End If
    If Not IsNull(snapWCBs("EC_TYPE")) Then
        cmbItems$ = cmbItems$ & "Type " & snapWCBs("EC_TYPE") & " - "
    End If
    If Not IsNull(snapWCBs("EC_CLASS")) Then
        cmbItems$ = cmbItems$ & "Class " & snapWCBs("EC_CLASS") & " - "
    End If
    If Not IsNull(snapWCBs("EC_CAUSECD")) Then
        cmbItems$ = cmbItems$ & "Cause " & snapWCBs("EC_CAUSECD") & " - "
    End If
    If Not IsNull(snapWCBs("EC_WCBFDTE")) Then
        wcb(X%, 3) = CStr(snapWCBs("EC_WCBFDTE"))
        cmbItems$ = cmbItems$ & "From " & CStr(snapWCBs("EC_WCBFDTE")) & " - "
    End If
    If Not IsNull(snapWCBs("EC_WCBCDTE")) Then
        cmbItems$ = cmbItems$ & "C" & " - "
    End If
    cmbWCB.AddItem cmbItems$
    snapWCBs.MoveNext
Wend

fglbComboWCB% = X%  ' can't always trust recordcount and list need it
popComboWCB = fglbComboWCB%
snapWCBs.Close

Exit Function

popComboWCB_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Snap create", "HROHSCOS", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function



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

'cmdWCBMed.Enabled = FT
'cmdIncident.Enabled = FT
'cmdTCause.Enabled = FT
'cmdCAction.Enabled = FT
'cmdInjLoc.Enabled = FT
'cmdContact.Enabled = FT

cmbWCB.Enabled = TF
medFutureE.Enabled = TF
medFutureES.Enabled = TF
medHealth.Enabled = TF
medNonE.Enabled = TF
medOther.Enabled = TF
medPartial2.Enabled = TF
medPartial4.Enabled = TF
medPension.Enabled = TF
medReEmp.Enabled = TF
medRehab.Enabled = TF
medRetPen.Enabled = TF
medSurvivor.Enabled = TF
medTemp.Enabled = TF
 clpCode(1).Enabled = TF
dlpFromTo(0).Enabled = TF
dlpFromTo(1).Enabled = TF
dlpSDate.Enabled = TF
'vbxTrueGrid.Enabled = FT
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'   cmdModify.Enabled = False
End If

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
            SQLQ = "SELECT * from Term_HROHSCOS "
            SQLQ = SQLQ & "WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "SELECT * from HROHSCOS "
            SQLQ = SQLQ & "WHERE CC_EMPNBR = " & glbLEE_ID
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
 '   If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
 '   End If
End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$
Dim SQLQ As String, X%, WCBN$, WCBN2$

On Error GoTo Tab1_Err
Call Display_Value

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    'MsgBox "No Records Found"
Else
    If Not IsNull(Data1.Recordset("CC_WCBNBR")) Then
        WCBN$ = Data1.Recordset("CC_WCBNBR")
        cmbWCB.ListIndex = -1
        For X% = 1 To fglbComboWCB%
            WCBN2$ = wcb(X%, 1)
            If WCBN2$ = WCBN$ Then
                cmbWCB.ListIndex = X% - 1
                Exit For
            End If
        Next X%
        
    End If
End If
Exit Sub

Tab1_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HROHSCOS", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
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
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        Call SET_UP_MODE
        'Me.cmdModify_Click
        Exit Sub
    End If

    
If glbtermopen Then
    SQLQ = "SELECT * from Term_HROHSCOS "
    SQLQ = SQLQ & "WHERE CC_WCBC_ID = " & Data1.Recordset!CC_WCBC_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    
Else
    SQLQ = "SELECT * from HROHSCOS "
    SQLQ = SQLQ & "WHERE CC_WCBC_ID = " & Data1.Recordset!CC_WCBC_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
Call SET_UP_MODE
'Me.cmdModify_Click
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
UpdateRight = gSec_Upd_HSCost
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
    frmEHSWCBC.Caption = " WSIB Cost Statements - " & Left$(glbLEE_SName, 5)
    frmEHSWCBC.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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
Function chkNumeric(Ctrol As Control)
chkNumeric = False
If Len(Ctrol) = 0 Then
    Ctrol = 0
Else
    If Not IsNumeric(Ctrol.Text) Then
        MsgBox Mid(Ctrol.Tag, 4) & " Must be numeric"
        Ctrol.SetFocus
        Exit Function
    End If
End If
chkNumeric = True
End Function
