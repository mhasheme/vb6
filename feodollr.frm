VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEODOLLAR 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Dollar Entitlements"
   ClientHeight    =   8490
   ClientLeft      =   105
   ClientTop       =   1380
   ClientWidth     =   11475
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
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdActAmtDetails 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3550
      TabIndex        =   6
      Tag             =   "Actual Amount Details"
      Top             =   4156
      Width           =   375
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   270
      Left            =   8010
      TabIndex        =   37
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   375
      Left            =   1800
      Top             =   7320
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      DataField       =   "DE_COMMENTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Tag             =   "00-Comments"
      Top             =   6150
      Width           =   6615
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feodollr.frx":0000
      Height          =   2025
      Left            =   120
      OleObjectBlob   =   "feodollr.frx":0014
      TabIndex        =   0
      Tag             =   "Listing of Dollar Entitlements"
      Top             =   480
      Width           =   9615
   End
   Begin INFOHR_Controls.DateLookup dlpPDate 
      DataField       =   "DE_PAIDDATE"
      Height          =   285
      Left            =   1965
      TabIndex        =   11
      Tag             =   "40-Date Paid"
      Top             =   5804
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpFDate 
      DataField       =   "DE_FDATE"
      Height          =   285
      Left            =   1965
      TabIndex        =   2
      Tag             =   "41-Starting date"
      Top             =   3463
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "DE_TYPE"
      Height          =   285
      Index           =   1
      Left            =   1965
      TabIndex        =   1
      Tag             =   "01-Entitlement - Code"
      Top             =   3120
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOL"
   End
   Begin VB.TextBox txtRefer 
      Appearance      =   0  'Flat
      DataField       =   "DE_REFNBR"
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
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   9
      Tag             =   "00-Reference Number"
      Top             =   5118
      Width           =   1230
   End
   Begin VB.TextBox txtPaidTo 
      Appearance      =   0  'Flat
      DataField       =   "DE_PAIDTO"
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
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   10
      Tag             =   "00-Paid To"
      Top             =   5461
      Width           =   2985
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   7800
      Top             =   7320
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   2
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
      Caption         =   "Ado1"
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
      TabIndex        =   27
      Top             =   7830
      Width           =   11475
      _Version        =   65536
      _ExtentX        =   20241
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
      Begin VB.CommandButton cmdRecalcActualALL 
         Caption         =   "Recalculate ALL"
         Height          =   375
         Left            =   5160
         TabIndex        =   40
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdRecalcActual 
         Caption         =   "Recalculate Actual Amount"
         Height          =   495
         Left            =   1680
         TabIndex        =   29
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdFlag 
         Caption         =   "COE &Update"
         Height          =   495
         Left            =   240
         TabIndex        =   28
         Top             =   120
         Width           =   1215
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   7005
         Top             =   135
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
      DataField       =   "DE_LDATE"
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
      Left            =   10320
      MaxLength       =   25
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DE_LTIME"
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
      Left            =   10680
      MaxLength       =   25
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DE_LUSER"
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
      Left            =   9960
      MaxLength       =   25
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11475
      _Version        =   65536
      _ExtentX        =   20241
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
         TabIndex        =   39
         Top             =   135
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
         Left            =   120
         TabIndex        =   19
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
         Left            =   1440
         TabIndex        =   18
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         TabIndex        =   17
         Top             =   135
         Width           =   1740
      End
   End
   Begin Threed.SSCheck chkCOEFlag 
      DataField       =   "DE_COST_OF_EMPLOYMENT"
      Height          =   225
      Left            =   330
      TabIndex        =   8
      Tag             =   "Include entitlement on COE report ?"
      Top             =   4805
      Width           =   2145
      _Version        =   65536
      _ExtentX        =   3784
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "Cost of Employment           "
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
   Begin MSMask.MaskEdBox medEntitleAmnt 
      DataField       =   "DE_ENTITLE"
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Tag             =   "20-Amount of entitlement during the period"
      Top             =   3806
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
   Begin MSMask.MaskEdBox medActualAmnt 
      DataField       =   "DE_ACTUAL"
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Tag             =   "20-Actual amount during the period"
      Top             =   4149
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.DateLookup dlpTDate 
      DataField       =   "DE_TDATE"
      Height          =   285
      Left            =   5970
      TabIndex        =   3
      Tag             =   "41-Ending date"
      Top             =   3463
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin VB.Label lblExceeded 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Exceeded Entitlement"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3600
      TabIndex        =   41
      Top             =   4537
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label lblImport 
      Caption         =   "Document"
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
      Left            =   6360
      TabIndex        =   38
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgSec 
      Height          =   240
      Left            =   7590
      Picture         =   "feodollr.frx":4F18
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgNoSec 
      Height          =   240
      Left            =   7590
      Picture         =   "feodollr.frx":5062
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblDOH1 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   36
      Top             =   2640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblDOH 
      AutoSize        =   -1  'True
      Caption         =   "Original Hire Date"
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
      Left            =   360
      TabIndex        =   35
      Top             =   2640
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lblTitle 
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
      Height          =   255
      Index           =   10
      Left            =   330
      TabIndex        =   34
      Top             =   6150
      Width           =   1305
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   5040
      TabIndex        =   33
      Top             =   3508
      Width           =   705
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reference Number"
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
      Left            =   330
      TabIndex        =   32
      Top             =   5163
      Width           =   1680
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Paid To"
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
      Left            =   330
      TabIndex        =   31
      Top             =   5506
      Width           =   885
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Paid"
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
      Left            =   330
      TabIndex        =   30
      Top             =   5849
      Width           =   705
   End
   Begin VB.Label lblVar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Top             =   4492
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Variance"
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
      Left            =   330
      TabIndex        =   26
      Top             =   4537
      Width           =   765
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Amount"
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
      Left            =   330
      TabIndex        =   25
      Top             =   4194
      Width           =   1245
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Entitlement Amount"
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
      Left            =   330
      TabIndex        =   24
      Top             =   3851
      Width           =   1650
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   330
      TabIndex        =   23
      Top             =   3508
      Width           =   885
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Entitlement"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   22
      Top             =   3165
      Width           =   960
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "DE_EMPNBR"
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
      TabIndex        =   20
      Top             =   6480
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "DE_COMPNO"
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
      Left            =   30
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEODOLLAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim Actn, OBDOLLAR, OADOLLAR, OTYPE, OCOEFLAG
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim fglbNew As Integer


Private Function AUDITODOL(ACTX)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String

On Error GoTo AUDIT_ERR

AUDITODOL = False


'rsTB.Open "HREMP", gdbAdoIhr001, adOpenKeyset, , adCmdTableDirect
'rsTB.Find "ED_EMPNBR = " & glbLEE_ID
rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    'xPT = rsTB("ED_PT")
    'xDiv = rsTB("ED_DIV")
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
'strFields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_DOLENT, AU_SAFETY, AU_UNIFORM, AU_EQUIP, AU_CLEAN, "
strFields = strFields & "AU_BDOLLAR, AU_ADOLLAR, AU_COEFLAG, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID, "
strFields = strFields & "AU_DOFDATE, AU_DOTDATE, AU_REFNBR, AU_PAIDTO, AU_PAIDDATE"
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False

If ACTX = "D" Or ACTX = "A" Then GoTo MODUPD
If OBDOLLAR <> medEntitleAmnt Or OTYPE <> clpCode(1).Text Then GoTo MODUPD
If OADOLLAR <> medActualAmnt Or OCOEFLAG <> chkCOEFlag Then GoTo MODUPD
GoTo MODNOUPD

MODUPD:
rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM"
rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP"
rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL"
rsTA("AU_DOLENT_TABL") = "EDOL"   '"EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv

If ACTX = "D" Then
    rsTA("AU_DOLENT") = clpCode(1).Text
    If clpCode(1).Text = "SAFE" Then rsTA("AU_SAFETY") = "-"
    If clpCode(1).Text = "UNIF" Then rsTA("AU_UNIFORM") = "-"
    If clpCode(1).Text = "EQUI" Then rsTA("AU_EQUIP") = "-"
    If clpCode(1).Text = "CLEA" Then rsTA("AU_CLEAN") = "-"
    rsTA("AU_DOFDATE") = dlpFDate.Text
    rsTA("AU_DOTDATE") = dlpTDate.Text
Else
    rsTA("AU_DOLENT") = clpCode(1).Text
    If ACTX = "A" Then
        If clpCode(1).Text = "SAFE" Then rsTA("AU_SAFETY") = "Y"
        If clpCode(1).Text = "UNIF" Then rsTA("AU_UNIFORM") = "Y"
        If clpCode(1).Text = "EQUI" Then rsTA("AU_EQUIP") = "Y"
        If clpCode(1).Text = "CLEA" Then rsTA("AU_CLEAN") = "Y"
    End If
    rsTA("AU_BDOLLAR") = CDbl(medEntitleAmnt)       'Ticket #22781
    rsTA("AU_ADOLLAR") = CDbl(medActualAmnt)        'Ticket #22781
    If chkCOEFlag = True Then
        rsTA("AU_COEFLAG") = "Y"
    Else
        rsTA("AU_COEFLAG") = "N"
    End If
    rsTA("AU_DOFDATE") = dlpFDate.Text
    rsTA("AU_DOTDATE") = dlpTDate.Text
    rsTA("AU_REFNBR") = txtRefer.Text
    rsTA("AU_PAIDTO") = txtPaidTo.Text
    If IsDate(dlpPDate.Text) Then 'Ticket #20274 Franks 05/05/2011
        rsTA("AU_PAIDDATE") = dlpPDate.Text
    End If
End If

Dim rsEmp As New ADODB.Recordset
Dim SQLQ
SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmp.EOF Then
    If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
End If
rsEmp.Close
    
rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = glbLEE_ID
rsTA("AU_LDATE") = Date
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
rsTA.Update

MODNOUPD:
AUDITODOL = True
Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '23July99 js

End Function

Private Sub Calc_Var()
Dim AVar
AVar = 0
If Len(medEntitleAmnt) <= 0 Then medEntitleAmnt = 0
If Not IsNumeric(medEntitleAmnt) Then medEntitleAmnt = 0
If Len(medActualAmnt) <= 0 Then medActualAmnt = 0
If Not IsNumeric(medActualAmnt) Then medActualAmnt = 0
AVar = medEntitleAmnt - medActualAmnt
lblVar = Format(AVar, "Currency")
If Val(medActualAmnt) > Val(medEntitleAmnt) Then
    lblExceeded.Visible = True
Else
    lblExceeded.Visible = False
End If
End Sub

Private Sub chkCOEFlag_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Function chkEODollar()
Dim SQLQ As String, Msg As String, dd#
Dim rsEmp As New ADODB.Recordset

chkEODollar = False

On Error GoTo chkEODollar_Err

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Entitlement code is a required field"
    clpCode(1).SetFocus
    Exit Function
End If

If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Entitlement code must be valid"
    clpCode(1).SetFocus
    Exit Function
End If

If Len(dlpFDate.Text) >= 1 Then
    If Not IsDate(dlpFDate.Text) Then
        MsgBox "From Date is not a valid date."
        dlpFDate.SetFocus
        Exit Function
    End If
Else
    MsgBox "From Date is required."
    dlpFDate.SetFocus
    Exit Function
End If

If Len(dlpTDate.Text) >= 1 Then
    If Not IsDate(dlpTDate.Text) Then
        MsgBox "To Date is not a valid date."
        dlpTDate.SetFocus
        Exit Function
    End If
Else
    MsgBox "To Date is required."
    dlpTDate.SetFocus
    Exit Function
End If

dd# = DateDiff("d", CVDate(dlpFDate.Text), CVDate(dlpTDate.Text))
If dd# < 0 Then
    MsgBox "From date must be earlier than To Date"
    dlpFDate.SetFocus
    Exit Function
End If

If Len(Trim(medActualAmnt)) <= 0 Then medActualAmnt = 0
If Len(Trim(medEntitleAmnt)) <= 0 Then medEntitleAmnt = 0

'Hemu 05/09/2003 Begin - Date Paid and Original Hire Date
If Len(dlpPDate.Text) > 0 Then
    If Not IsDate(dlpPDate.Text) Then
        MsgBox "Date Paid is not a valid date"
        dlpPDate.SetFocus
        Exit Function
    End If
    
    rsEmp.Open "SELECT ED_DOH FROM HREMP WHERE ED_EMPNBR = " & lblEENum, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If DaysBetween(rsEmp("ED_DOH"), dlpPDate.Text) < 0 Then
            MsgBox "Date Paid can not be prior to Original Hire date"
            dlpPDate.SetFocus
            rsEmp.Close
            Exit Function
        End If
    End If
    rsEmp.Close
End If
'Hemu 05/09/2003 End

chkEODollar = True

Exit Function

chkEODollar_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkODollar", "HRDOLENT", "edit/Add")
Resume Next

End Function

Sub cmdCancel_Click()
Dim X
On Error GoTo Can_Err
fglbNew = False
rsDATA.CancelUpdate
Call Display_Value


'Call ST_UPD_MODE(True)  ' reset screen's attributes
'Call SET_UP_MODE

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRDOLENT", "Cancel")
Resume Next

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEODOLLAR" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, X

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

If Not AUDITODOL("D") Then MsgBox "ERROR : AUDIT FILE"

If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001X.CommitTrans
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.Execute "Delete from Term_HRDOC_DOLENT where DE_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " and DE_DOCKEY=" & glbDocKey & " " '
    End If
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001.CommitTrans
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.Execute "delete from HRDOC_HRDOLENT where DE_TYPE='" & UCase(glbDocName) & "' AND DE_EMPNBR = " & glbLEE_ID & " and DE_DOCKEY=" & glbDocKey & " "
    End If
    Data1.Refresh
End If
'If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
'End If
fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRDOLENT", "Delete")

Resume Next
Unload Me

End Sub

Private Sub cmdActAmtDetails_Click()
    'Ticket #28789 - Get the Actual Amount Details
    glbDolType = clpCode(1).Text
    glbDolFDate = dlpFDate.Text
    glbDolTDate = dlpTDate.Text
    
    frmEODOLLARDTL.Show 1
    DoEvents
    
    'Ticket #28789 - Recompute the Total Actual Amounts
    'If Not glbtermopen Then
        Call ReCompute_ActualAmount("UpdMst")
        Call Calc_Var
    'End If
End Sub

Private Sub cmdActAmtDetails_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdFlag_Click()
Dim SQLQ As String, Msg$, Response%, Title$, DgDef As Variant

On Error GoTo MAll_Err

Msg$ = "How would you like to mark all COE flags?"
Msg = Msg$
Title$ = "Mark all completed?"   ' zzz
DgDef = MB_YESNOCANCEL + MB_ICONQUESTION + MB_DEFBUTTON3  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
Dim rsHR As New ADODB.Recordset

If glbtermopen Then
    rsHR.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    rsHR.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

If Response = IDYES Then    ' Evaluate response
    Screen.MousePointer = HOURGLASS
    With rsHR
        Do Until .EOF
            .ActiveConnection.BeginTrans
            !DE_COST_OF_EMPLOYMENT = True
            .Update
            .ActiveConnection.CommitTrans
            .MoveNext
            DoEvents
        Loop
    End With
    Data1.Refresh
    Call Display_Value
    Screen.MousePointer = DEFAULT
End If

If Response = IDNO Then    ' Evaluate response
    Screen.MousePointer = HOURGLASS
    With rsHR
        Do Until .EOF
            .ActiveConnection.BeginTrans
            !DE_COST_OF_EMPLOYMENT = False
            .Update
            .ActiveConnection.CommitTrans
            .MoveNext
            DoEvents
        Loop
    End With
    Data1.Refresh
    Call Display_Value
    Screen.MousePointer = DEFAULT

End If

Exit Sub

MAll_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdMarkAll", "HRDOLENT", "Mark All")
Call RollBack '23July99 js

End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Actn = "M"
OBDOLLAR = medEntitleAmnt
OTYPE = clpCode(1).Text
OADOLLAR = medActualAmnt
OCOEFLAG = chkCOEFlag

'Call ST_UPD_MODE(True)
Call SET_UP_MODE
'clpCode(1).Enabled = True
'clpCode(1).SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HRDOLENT", "Modify")
Call RollBack '23July99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
    Dim SQLQ As String
    
    fglbNew = True
    
    'Call ST_UPD_MODE(True)
    Call SET_UP_MODE
    
    On Error GoTo AddN_Err
    
    If gsAttachment_DB And Not glbtermopen Then
        lblImport.Visible = True
        imgSec.Visible = False
        imgNoSec.Visible = True
        cmdImport.Visible = True
    End If
    
    Actn = "A"
    OBDOLLAR = ""
    OTYPE = ""
    OADOLLAR = ""
    OCOEFLAG = False
    
    Call Set_Control("B", Me)
    
    rsDATA.AddNew
        
    dlpFDate.Text = 1
    If glbCompSerial = "S/N - 2288W" Or _
        glbCompSerial = "S/N - 2418W" Then ' Musashi Auto tkt# 10866 Ticket #17786 charton hobbs
        dlpFDate.Text = ""
        dlpTDate.Text = ""
    Else
        dlpFDate.Text = CVDate(GetMonth("January") & " 1, " & Year(Now))
        dlpTDate.Text = CVDate(GetMonth("December") & " 31, " & Year(Now))
    
    End If
    
    lblVar = "$0.00"
    
    'Ticket #28789 - Jerry said to set the COE flag to ON for all new records by default.
    chkCOEFlag.Value = True
    cmdActAmtDetails.Enabled = False
    cmdRecalcActual.Enabled = False

    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
    
    lblCNum.Caption = "001"
    clpCode(1).SetFocus
    
    MDIMain.MainToolBar.ButtonS(8).Enabled = True
    MDIMain.MainToolBar.ButtonS(9).Enabled = True
    
    Exit Sub

AddN_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRDOLENT", "Add")
    Resume Next
    
End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim xID As Long

On Error GoTo Add_Err

fglbNew = False

If Not chkEODollar() Then Exit Sub

If Data1.Recordset.RecordCount > 0 Then
  If Not ChkDup(Actn) Then Exit Sub
End If

If Not glbtermopen Then
    If Not AUDITODOL(Actn) Then MsgBox "ERROR : AUDIT FILE"
End If

'Ticket #28789 - Actual Amount Details recalculate option
'If Not glbtermopen Then
    Call ReCompute_ActualAmount
    Call Calc_Var
'End If

'Ticket #22781
medEntitleAmnt.Text = CDbl(medEntitleAmnt)
medActualAmnt.Text = CDbl(medActualAmnt)

Call UpdUStats(Me)
Call Set_Control("U", Me, rsDATA)

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    xID = rsDATA("DE_ENTITLE_ID")
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    xID = rsDATA("DE_ENTITLE_ID")
    Data1.Refresh
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

'Call ST_UPD_MODE(True)
Call SET_UP_MODE

If NextFormIF("Entitlement") Then
    Call cmdNew_Click
End If

Exit Sub

Add_Err:
If Err = 3022 Then
    Data1.Refresh
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRDOLENT", "Update")
Resume Next
Unload Me

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Dollar Entitlements"
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

RHeading = lblEEName & "'s Dollar Entitlements"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Function ChkDup(AddChg)
Dim SQLQ, Logx, Msg$, SavReviewDate
Dim rsTB As New ADODB.Recordset
Dim Answer, OID
ChkDup = False

Logx = False
SavReviewDate = dlpTDate.Text
If AddChg = "A" Then
    SQLQ = "SELECT * FROM HRDOLENT WHERE DE_EMPNBR = " & lblEEID
    SQLQ = SQLQ & " AND DE_TYPE = '" & clpCode(1).Text & "'"
    SQLQ = SQLQ & " AND DE_TDATE = " & Date_SQL(dlpTDate.Text)
  
    rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    If Not rsTB.EOF Then
       Logx = True
    End If
Else
    OID = Data1.Recordset("DE_ENTITLE_ID")
    SQLQ = "SELECT * FROM HRDOLENT WHERE DE_EMPNBR = " & lblEEID
    SQLQ = SQLQ & " AND DE_TYPE = '" & clpCode(1).Text & "'"
    SQLQ = SQLQ & " AND DE_TDATE = " & Date_SQL(dlpTDate.Text)
    
    SQLQ = SQLQ & " AND DE_ENTITLE_ID <> " & OID
    rsTB.Open SQLQ, gdbAdoIhr001
    If rsTB.RecordCount > 0 Then Logx = True
End If
rsTB.Close
If Logx = True Then
    Msg$ = "Duplicate exist. OK to proceed"
    MsgBox Msg$
    clpCode(1).SetFocus
End If
'change required by Linamar
ChkDup = True

End Function

Function EERetrieve()
Dim SQLQ, SQLQ1 As String
Dim rsDOH As New ADODB.Recordset
EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS

If glbtermopen Then         'Lucy July 5, 2000
    SQLQ = "Select Term_DOLENT.*, DE_ENTITLE-DE_ACTUAL AS Variance  from Term_DOLENT"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY DE_FDATE DESC"
Else
    SQLQ = "Select HRDOLENT.*, DE_ENTITLE-DE_ACTUAL AS Variance  from HRDOLENT"
    SQLQ = SQLQ & " where DE_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY DE_FDATE DESC"
  
End If

Data1.RecordSource = SQLQ
Data1.Refresh

  If glbCompSerial = "S/N - 2288W" Then ' Musashi Auto tkt# 10866
    If glbtermopen Then
        SQLQ1 = "Select ED_DOH from Term_HREMP"
        SQLQ1 = SQLQ1 & " WHERE TERM_SEQ = " & glbTERM_Seq
    
    Else
        SQLQ1 = "Select ED_DOH from HREMP"
        SQLQ1 = SQLQ1 & " where ED_EMPNBR = " & glbLEE_ID
     
    End If
    Data2.RecordSource = SQLQ1
    Data2.Refresh

        If glbtermopen Then
            rsDOH.Open SQLQ1, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDOH.Open SQLQ1, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        If rsDOH.EOF = False And rsDOH.BOF = False Then
            lblDOH1.Caption = rsDOH("ED_DOH")
        End If
        rsDOH.Close
End If

EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HRDOLENT", "SELECT")
Resume Next

Exit Function

End Function


Private Sub cmdImport_Click()
    glbDocNewRecord = fglbNew
    glbDocName = "DollarEnt"
    If fglbNew Then
        glbDocKey = 0
    Else
        glbDocKey = rsDATA("DE_ENTITLE_ID")
    End If
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmEODOLLAR")
End Sub

Private Sub cmdRecalcActual_Click()
    Call ReCompute_ActualAmount("UpdMst")
    Call Calc_Var
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
'Me.cmdModify_Click
    glbOnTop = "FRMEODOLLAR"
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEODOLLAR"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "FRMEODOLLAR"

If glbtermopen Then         'Lucy July 5, 2000
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

If glbCompSerial = "S/N - 2288W" Then ' Musashi Auto tkt# 10866
    If glbtermopen Then
        Data2.ConnectionString = glbAdoIHRAUDIT
    Else
        Data2.ConnectionString = glbAdoIHRDB
    End If
    lblDOH.Visible = True
    lblDOH1.Visible = True
End If

'Ticket #17786 Charton-Hobbs Inc. - #2418
'If glbCompSerial = "S/N - 2418W" Then
 '   lblTitle(3).FontBold = False

'End If

Screen.MousePointer = HOURGLASS

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
    Me.Caption = "Dollar Entitlements - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = ShowEmpnbr(lblEEID)

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

'Ticket #28789 - Actual Amount Details recalculate option
'If glbtermopen Then
'    cmdRecalcActual.Visible = False
'    cmdRecalcActualALL.Visible = False
'End If

Call Display_Value

Call ST_UPD_MODE(False)

If Not gSec_Upd_Other_Entitlements Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
    cmdFlag.Enabled = False
End If

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

Private Sub Form_Unload(Cancel As Integer)

    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmEODOLLAR = Nothing
    Call NextForm
End Sub

Private Sub medActualAmnt_Change()
    Call Calc_Var
End Sub

Private Sub medActualAmnt_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub medActualAmnt_LostFocus()
    Call Calc_Var
End Sub

Private Sub medEntitleAmnt_Change()
    Call Calc_Var
End Sub

Private Sub medEntitleAmnt_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medEntitleAmnt_LostFocus()
    Call Calc_Var
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

''cmdFlag.Enabled = FT 'TF


'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT

medActualAmnt.Enabled = TF
'Ticket #28789 - Actual Amounts Details
cmdActAmtDetails.Enabled = TF

medEntitleAmnt.Enabled = TF
clpCode(1).Enabled = TF
dlpFDate.Enabled = TF
dlpTDate.Enabled = TF
chkCOEFlag.Enabled = TF
txtRefer.Enabled = TF
txtPaidTo.Enabled = TF
dlpPDate.Enabled = TF
'txtComments.Enabled = TF
txtComments.Locked = FT

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If
'vbxTrueGrid.Enabled = FT

glbDocName = "DollarEnt"
If gsAttachment_DB Then
    glbDocKey = 0
    If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
        If rsDATA.RecordCount > 0 Then
            If Not IsNull(rsDATA("DE_DOCKEY")) Then
                glbDocKey = rsDATA("DE_DOCKEY")
            Else
                glbDocKey = 0
            End If
        Else
            If Not IsNull(Data1.Recordset("DE_DOCKEY")) Then
                glbDocKey = Data1.Recordset("DE_DOCKEY")
            Else
                glbDocKey = 0
            End If
        End If
    End If
    
    Call DispimgIcon(Me, "frmEODOLLAR")
    If gSec_Upd_Other_Entitlements And Not glbtermopen Then
        If Data1.Recordset.BOF And Data1.Recordset.EOF Then
            cmdImport.Visible = False
        Else
            cmdImport.Visible = True
        End If
    End If
End If

End Sub

Private Sub Text1_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtComments_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub txtFDate_Change()
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtFDate_DblClick()
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub txtFDate_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub txtFDate_KeyPress(KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub

Private Sub txtPaidTo_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub
'Private Sub txtPDate_Change()
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtPDate_DblClick()
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub txtPDate_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub txtPDate_KeyPress(KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub

Private Sub txtRefer_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub
'Private Sub txtTDate_Change()
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtTDate_DblClick()
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub txtTDate_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub txtTDate_KeyPress(KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub

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
        
        If glbtermopen Then         'Lucy July 5, 2000
            SQLQ = "Select Term_DOLENT.*, DE_ENTITLE-DE_ACTUAL AS Variance  from Term_DOLENT"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select HRDOLENT.*, DE_ENTITLE-DE_ACTUAL AS Variance  from HRDOLENT"
            SQLQ = SQLQ & " where DE_EMPNBR = " & glbLEE_ID
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
Dim tdcode$
Dim SQLQ As String

On Error GoTo Tab1_Err
'If Not Fnd_Match_Data1() Then Exit Sub
Call Display_Value


Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HRDOLENT", "Add")
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
Dim SQLQ, SQLQ1 As String

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    If glbtermopen Then
        SQLQ = "Select Term_DOLENT.* from Term_DOLENT"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "Select HRDOLENT.* from HRDOLENT"
        SQLQ = SQLQ & " where DE_EMPNBR = " & glbLEE_ID
        If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
            rsDATA.CursorLocation = adUseServer
        End If
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    
    Call SET_UP_MODE
    Me.cmdModify_Click
     
    Exit Sub
End If
    
If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
If glbtermopen Then
    SQLQ = "Select Term_DOLENT.* from Term_DOLENT"
    SQLQ = SQLQ & " WHERE DE_ENTITLE_ID = " & Data1.Recordset!DE_ENTITLE_ID
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "Select HRDOLENT.*  from HRDOLENT"
    SQLQ = SQLQ & " where DE_ENTITLE_ID = " & Data1.Recordset!DE_ENTITLE_ID
    If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
        rsDATA.CursorLocation = adUseServer
    End If
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If
 
If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
Call Set_Control("R", Me, rsDATA)
Call SET_UP_MODE

'Ticket #28789 - Check if details available then the Actual Amount field should be locked
If Is_ActualAmountDetailsAvailable Then
    medActualAmnt.Enabled = False
    
    'Ticket #28789 - Compute Actual Amount total
    'If Not glbtermopen Then
        Call ReCompute_ActualAmount
        Call Calc_Var
    'End If
Else
    medActualAmnt.Enabled = True
End If

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
UpdateRight = gSec_Upd_Other_Entitlements
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
    cmdFlag.Enabled = False
    cmdRecalcActual.Enabled = False     'Ticket #28789 - Actual Amounts Details total
ElseIf rsDATA.EOF Then
    UpdateState = NoRecord
    TF = False
    cmdFlag.Enabled = False
    cmdRecalcActual.Enabled = False     'Ticket #28789 - Actual Amounts Details total
Else
    UpdateState = OPENING
    cmdFlag.Enabled = True
    cmdRecalcActual.Enabled = True      'Ticket #28789 - Actual Amounts Details total
    TF = True
End If

Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
Call ST_UPD_MODE(TF)
End Sub

Private Sub lblEEID_Change()

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    frmEODOLLAR.Caption = "Dollar Entitlements - " & Left$(glbLEE_SName, 5)
    frmEODOLLAR.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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
    SQLQ = getSQL("frmEODOLLAR")
    Call FillMemoFile(SQLQ, "DollarEnt")
End Sub

Private Sub ReCompute_ActualAmount(Optional UpdType)
    Dim rsDollarActDtl As New ADODB.Recordset
    Dim SQLQ As String
    Dim xBookMrk
    
    If Is_ActualAmountDetailsAvailable Then
        If glbtermopen Then
            SQLQ = "SELECT SUM(DA_ACTUAL) AS TOT_ACTUAL FROM Term_DOLENT_ACTDTL WHERE TERM_SEQ = " & glbTERM_Seq
            SQLQ = SQLQ & " AND DA_TYPE = '" & clpCode(1).Text & "'"
            SQLQ = SQLQ & " AND DA_FDATE = " & Date_SQL(dlpFDate.Text)
            SQLQ = SQLQ & " AND DA_TDATE = " & Date_SQL(dlpTDate.Text)
            rsDollarActDtl.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            SQLQ = "SELECT SUM(DA_ACTUAL) AS TOT_ACTUAL FROM HRDOLENT_ACTDTL WHERE DA_EMPNBR = " & glbLEE_ID
            SQLQ = SQLQ & " AND DA_TYPE = '" & clpCode(1).Text & "'"
            SQLQ = SQLQ & " AND DA_FDATE = " & Date_SQL(dlpFDate.Text)
            SQLQ = SQLQ & " AND DA_TDATE = " & Date_SQL(dlpTDate.Text)
            rsDollarActDtl.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        If Not rsDollarActDtl.EOF Then
            medActualAmnt.Text = IIf(IsNull(rsDollarActDtl("TOT_ACTUAL")), 0, rsDollarActDtl("TOT_ACTUAL"))
            
            If Val(medActualAmnt) > Val(medEntitleAmnt) Then
                lblExceeded.Visible = True
            Else
                lblExceeded.Visible = False
            End If
            
            'Update Master table?
            If Not IsMissing(UpdType) Then
                If UpdType = "UpdMst" Then
                    If glbtermopen Then
                        SQLQ = "UPDATE Term_DOLENT SET DE_ACTUAL = " & IIf(IsNull(rsDollarActDtl("TOT_ACTUAL")), 0, rsDollarActDtl("TOT_ACTUAL")) & " WHERE TERM_SEQ = " & glbTERM_Seq
                    Else
                        SQLQ = "UPDATE HRDOLENT SET DE_ACTUAL = " & IIf(IsNull(rsDollarActDtl("TOT_ACTUAL")), 0, rsDollarActDtl("TOT_ACTUAL")) & " WHERE DE_EMPNBR = " & glbLEE_ID
                    End If
                    SQLQ = SQLQ & " AND DE_TYPE = '" & clpCode(1).Text & "'"
                    SQLQ = SQLQ & " AND DE_FDATE = " & Date_SQL(dlpFDate.Text)
                    SQLQ = SQLQ & " AND DE_TDATE = " & Date_SQL(dlpTDate.Text)
                    gdbAdoIhr001.Execute SQLQ
                    xBookMrk = vbxTrueGrid.Bookmark
                    Data1.Refresh
                    Data1.Recordset.Bookmark = xBookMrk
                End If
            End If
        End If
        rsDollarActDtl.Close
        Set rsDollarActDtl = Nothing
    End If
End Sub

Private Function Is_ActualAmountDetailsAvailable() As Boolean
    Dim rsDollarActDtl As New ADODB.Recordset
    Dim SQLQ As String
    If glbtermopen Then
        SQLQ = "SELECT * FROM Term_DOLENT_ACTDTL WHERE TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " AND DA_TYPE = '" & clpCode(1).Text & "'"
        SQLQ = SQLQ & " AND DA_FDATE = " & Date_SQL(dlpFDate.Text)
        SQLQ = SQLQ & " AND DA_TDATE = " & Date_SQL(dlpTDate.Text)
        rsDollarActDtl.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "SELECT * FROM HRDOLENT_ACTDTL WHERE DA_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND DA_TYPE = '" & clpCode(1).Text & "'"
        SQLQ = SQLQ & " AND DA_FDATE = " & Date_SQL(dlpFDate.Text)
        SQLQ = SQLQ & " AND DA_TDATE = " & Date_SQL(dlpTDate.Text)
        rsDollarActDtl.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    If Not rsDollarActDtl.EOF Then
        Is_ActualAmountDetailsAvailable = True
        cmdRecalcActual.Enabled = True
    Else
        Is_ActualAmountDetailsAvailable = False
        cmdRecalcActual.Enabled = False
    End If
    rsDollarActDtl.Close
    Set rsDollarActDtl = Nothing

End Function
