VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmOTHERERN 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Earnings"
   ClientHeight    =   7950
   ClientLeft      =   105
   ClientTop       =   975
   ClientWidth     =   10815
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
   ScaleHeight     =   7950
   ScaleWidth      =   10815
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPostion 
      Caption         =   "P&ositions"
      Height          =   255
      Left            =   210
      TabIndex        =   10
      Tag             =   "Postions"
      Top             =   4710
      Width           =   975
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feothern.frx":0000
      Height          =   2235
      Left            =   120
      OleObjectBlob   =   "feothern.frx":0014
      TabIndex        =   0
      Tag             =   "Listing of Other Earnings"
      Top             =   570
      Width           =   9045
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "TDATE"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   6
      Tag             =   "41-Ending date of period"
      Top             =   4038
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "FDATE"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   4
      Tag             =   "41-Starting date of period"
      Top             =   3672
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EARN_TYPE"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Tag             =   "01-Earnings code"
      Top             =   2940
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EARN"
   End
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
      DataField       =   "COMMENTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Tag             =   "00-Enter Comments"
      Top             =   5400
      Width           =   8805
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   26
      Top             =   7530
      Width           =   10815
      _Version        =   65536
      _ExtentX        =   19076
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
      Begin MSAdodcLib.Adodc DATA1 
         Height          =   405
         Left            =   9600
         Top             =   0
         Visible         =   0   'False
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   714
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
      Begin VB.CommandButton cmdFlag 
         Appearance      =   0  'Flat
         Caption         =   "Flag &Updt"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   0
         Width           =   1215
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8250
         Top             =   150
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
      DataField       =   "LDATE"
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
      Left            =   4605
      MaxLength       =   25
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6930
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LTIME"
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
      Left            =   6180
      MaxLength       =   25
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6930
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LUSER"
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
      Left            =   7755
      MaxLength       =   25
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6930
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10815
      _Version        =   65536
      _ExtentX        =   19076
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
         Left            =   7320
         TabIndex        =   29
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label lbltitle 
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
         Top             =   120
         Width           =   1740
      End
   End
   Begin Threed.SSCheck chkCOEFlag 
      DataField       =   "COST_OF_EMPLOYMENT"
      Height          =   195
      Left            =   210
      TabIndex        =   8
      Tag             =   "Include earnings on COE report"
      Top             =   4404
      Width           =   2355
      _Version        =   65536
      _ExtentX        =   4154
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "Cost of Employment                "
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
   Begin MSMask.MaskEdBox medAmount 
      DataField       =   "ACT_DOLLAR"
      Height          =   285
      Left            =   2355
      TabIndex        =   2
      Tag             =   "20- Actual amount of earnings during this period"
      Top             =   3306
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   10
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
   Begin MSMask.MaskEdBox medPPE 
      Height          =   285
      Left            =   6720
      TabIndex        =   3
      Tag             =   "10-Percentage Earning"
      Top             =   3306
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
      MaxLength       =   10
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
   Begin MSMask.MaskEdBox medHours 
      Height          =   285
      Left            =   6720
      TabIndex        =   5
      Tag             =   "Earning Hours"
      Top             =   3672
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   10
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
   Begin INFOHR_Controls.CodeLookup clpJob 
      DataField       =   "OE_JOB"
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Tag             =   "01-Job Code"
      Top             =   4700
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   5
   End
   Begin MSMask.MaskEdBox medCorpEqui 
      Height          =   285
      Left            =   6705
      TabIndex        =   7
      Tag             =   "20- Actual amount of earnings during this period"
      Top             =   4038
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   10
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
   Begin MSMask.MaskEdBox MskTROIC 
      Height          =   285
      Left            =   6705
      TabIndex        =   9
      Tag             =   "01-High Dollars"
      Top             =   4400
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
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
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ROIC"
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
      Left            =   5760
      TabIndex        =   33
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corporate Equivalent"
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
      Left            =   5040
      TabIndex        =   32
      Top             =   4080
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Earning Hours"
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
      Left            =   5160
      TabIndex        =   31
      Top             =   3717
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Earning %"
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
      Left            =   5160
      TabIndex        =   30
      Top             =   3351
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   255
      Index           =   5
      Left            =   210
      TabIndex        =   28
      Top             =   5160
      Width           =   990
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   210
      TabIndex        =   25
      Top             =   4083
      Width           =   705
   End
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   210
      TabIndex        =   24
      Top             =   3717
      Width           =   885
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Left            =   210
      TabIndex        =   23
      Top             =   3345
      Width           =   1845
   End
   Begin VB.Label lbltitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Earnings"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   210
      TabIndex        =   22
      Top             =   2985
      Width           =   1455
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EMPNBR"
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
      Left            =   3870
      TabIndex        =   20
      Top             =   7035
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "COMPNO"
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
      Left            =   3030
      TabIndex        =   21
      Top             =   7035
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmOTHERERN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim OADOLLAR, OEARN, OCOEFLAG, Actn
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim fglbNew As Integer
Dim FRS As ADODB.Recordset
Dim OEarnPer, OEarnHrs
Dim fglbJobList As String

Private Function AUDITOEAR(ACTX)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String

On Error GoTo AUDIT_ERR

AUDITOEAR = False

'rsTB.Open "HREMP", gdbAdoIhr001, adOpenKeyset, , adCmdTableDirect
'rsTB.Find "ED_EMPNBR = " & glbLEE_ID
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

strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, "
strFields = strFields & "AU_TREAS_TABL, AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_EARN, AU_SERVICE, "
strFields = strFields & "AU_ADOLLAR, AU_COEFLAG, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID, "
strFields = strFields & "AU_DOFDATE, AU_DOTDATE, AU_EARNPCE, AU_EARNHOURS, AU_JOB "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False

If ACTX = "D" Or ACTX = "A" Then GoTo MODUPD

If OEARN <> clpCode(1).Text Then GoTo MODUPD

If OADOLLAR <> medAmount Or OCOEFLAG <> chkCOEFlag Then GoTo MODUPD
If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #19936 Franks 04/15/2011
    If OEarnPer <> medPPE.Text Then GoTo MODUPD
    If OEarnHrs <> medHours.Text Then GoTo MODUPD
End If

GoTo MODNOUPD

MODUPD:

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM"
rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP"
rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL"
rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv

If ACTX = "D" Then
    rsTA("AU_EARN") = clpCode(1).Text
    If clpCode(1).Text = "SERV" Then rsTA("AU_SERVICE") = "Y"
    rsTA("AU_DOFDATE") = dlpDate(0).Text
    rsTA("AU_DOTDATE") = dlpDate(1).Text
Else
    rsTA("AU_EARN") = clpCode(1).Text
    If ACTX = "A" Then If clpCode(1).Text = "SERV" Then rsTA("AU_SERVICE") = "Y"
       
    If OADOLLAR <> medAmount Then
        If medAmount <> "" Then                      '16Aug99 js
            rsTA("AU_ADOLLAR") = CDbl(medAmount)    'Ticket #22781
        Else
            rsTA("AU_ADOLLAR") = 0
        End If
    End If
    If OCOEFLAG <> chkCOEFlag Then
        If chkCOEFlag = True Then
            rsTA("AU_COEFLAG") = "Y"
        Else
            rsTA("AU_COEFLAG") = "N"
        End If
    End If
    rsTA("AU_DOFDATE") = dlpDate(0).Text
    rsTA("AU_DOTDATE") = dlpDate(1).Text
    If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
        If OEarnPer <> medPPE.Text Then
            If medPPE.Text <> "" Then
                rsTA("AU_EARNPCE") = medPPE.Text
            Else                                         '
                rsTA("AU_EARNPCE") = 0
            End If
        End If
        If OEarnHrs <> medHours.Text Then
            If medHours.Text <> "" Then
                rsTA("AU_EARNHOURS") = medHours.Text
            Else
                rsTA("AU_EARNHOURS") = 0
            End If
        End If
    End If
End If

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = glbLEE_ID
rsTA("AU_LDATE") = Date
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
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

'Ticket #24410 - City of Sarnia - Position Code added in Earnings table so updating the Audit table as well
rsTA("AU_JOB") = clpJob.Text

rsTA.Update

MODNOUPD:
AUDITOEAR = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '26July99 js

End Function

Private Sub chkCOEFlag_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Function chkEOTHERE()
Dim SQLQ As String, Msg As String
Dim dd&

chkEOTHERE = False

On Error GoTo chkEOTHERE_Err

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Earnings code is a required field"
    clpCode(1).SetFocus
    Exit Function
Else
    If clpCode(1).Caption = "Unassigned" Then
        MsgBox "Earnings code must be valid"
        clpCode(1).SetFocus
        Exit Function
    End If
End If
If Len(Trim(medAmount)) > 0 Then      'laura jan 12, 1998
    If Not IsNumeric(medAmount) Then
        MsgBox "Amount is invalid"
        medAmount.SetFocus
        Exit Function
    End If
Else
    medAmount = 0
End If

If Len(dlpDate(0).Text) < 1 Then
    MsgBox "From Date is required field"
    dlpDate(0).SetFocus
    Exit Function
End If

If Len(dlpDate(0).Text) >= 1 Then
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "From Date is not a valid date"
        dlpDate(0).SetFocus
        Exit Function
    End If
End If

If Len(dlpDate(1).Text) >= 1 Then
    If Not IsDate(dlpDate(1).Text) Then
        MsgBox "To Date is not a valid date"
        dlpDate(1).SetFocus
        Exit Function
    End If
Else
    MsgBox "To Date is required field"
    dlpDate(1).SetFocus
    Exit Function
End If

dd& = DateDiff("d", CVDate(dlpDate(0).Text), CVDate(dlpDate(1).Text))

If dd& < 0 Then
    MsgBox "To Date cannot precede From Date"
    dlpDate(1).SetFocus
    Exit Function
End If

'Add by Frank Dec 18,2001
If glbCompSerial = "S/N - 2214W" Then 'Casey House
    If Not chkCOEFlag Then
        MsgBox "Cost of Employment must be On"
        chkCOEFlag.SetFocus
        Exit Function
    End If
End If

chkEOTHERE = True

Exit Function

chkEOTHERE_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkOThern", "HREARN", "edit/Add")
Resume Next

End Function

Sub cmdCancel_Click()
Dim x
On Error GoTo Can_Err
fglbNew = False
rsDATA.CancelUpdate
Call Display_Value
'Call ST_UPD_MODE(True)  ' reset screen's attributes
Exit Sub
Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREARN", "Cancel")
Resume Next

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMOTHERERN" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, x

If DATA1.Recordset.BOF And DATA1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

If Not AUDITOEAR("D") Then MsgBox "ERROR - AUDIT FILE"

If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001X.CommitTrans
    DATA1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001.CommitTrans
    DATA1.Refresh
End If
If DATA1.Recordset.EOF And DATA1.Recordset.BOF Then
    Call Display_Value
End If

fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HREARN", "Delete")
Call RollBack '26July99 js

End Sub

Private Sub clpCode_LostFocus(Index As Integer)
If glbWFC And Index = 1 Then 'Type of Earnings
    If fglbNew Then
        If clpCode(1).Text = "BON4" Then
            dlpDate(0).Text = getWFCFiscalYearStartDate(Date)
            dlpDate(1).Text = getWFCFiscalYearToDate(Date)
        End If
    End If
End If
End Sub

Private Sub clpJob_Change()
If clpJob.Text <> "" Then
    clpJob.Caption = GetJobData(clpJob.Text, "JB_DESCR", "")
End If
End Sub

Private Sub clpJob_LostFocus()
If clpJob.Text <> "" Then
    clpJob.Caption = GetJobData(clpJob.Text, "JB_DESCR", "")
End If
End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdFlag_Click()
Dim SQLQ As String
Dim Msg$, Response%, Title$, DgDef As Variant

On Error GoTo MAll_Err

Msg$ = "How would you like to mark all COE flags?"
Msg = Msg$
Title$ = "Mark all completed?"   ' zzz
DgDef = MB_YESNOCANCEL + MB_ICONQUESTION + MB_DEFBUTTON3  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.

Dim rsHR As New ADODB.Recordset
If glbtermopen Then
    rsHR.Open DATA1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    rsHR.Open DATA1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If
If Response = IDYES Then    ' Evaluate response
    Screen.MousePointer = HOURGLASS
    With rsHR
        Do Until .EOF
            .ActiveConnection.BeginTrans
            !COST_OF_EMPLOYMENT = True
            .Update
            .ActiveConnection.CommitTrans
            .MoveNext
            DoEvents
        Loop
    End With
    DATA1.Refresh
    Call Display_Value
    Screen.MousePointer = DEFAULT
End If

If Response = IDNO Then    ' Evaluate response
    Screen.MousePointer = HOURGLASS
    With rsHR
        Do Until .EOF
            .ActiveConnection.BeginTrans
            !COST_OF_EMPLOYMENT = False
            .Update
            .ActiveConnection.CommitTrans
            .MoveNext
            DoEvents
        Loop
    End With
    DATA1.Refresh
    Call Display_Value
    Screen.MousePointer = DEFAULT
End If

Exit Sub

MAll_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdMarkAll", "HREARN", "Mark All")
Call RollBack '26July99 js

End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Actn = "M"
OADOLLAR = medAmount
OEARN = clpCode(1).Text
OCOEFLAG = chkCOEFlag
OEarnPer = medPPE.Text 'Ticket #19936
OEarnHrs = medHours.Text 'Ticket #19936

'Comment by Frank 11/06/03, it can cause the Error 5
''Call ST_UPD_MODE(True)
'Call SET_UP_MODE
''clpCode(1).Enabled = True
'If clpCode(1).Enabled Then clpCode(1).SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HREARN", "Modify")
Call RollBack '26July99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
Dim SQLQ As String
fglbNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
cmdFlag.Enabled = False
On Error GoTo AddN_Err

'DATA1.Recordset.AddNew
''' Sam add July 2002 * Remove Binding Control
Call Set_Control("B", Me)
rsDATA.AddNew

Actn = "A"
OADOLLAR = ""
OEARN = ""
OCOEFLAG = False
If glbWFC Then 'Ticket #29011 Franks 08/03/2016
    chkCOEFlag.Value = True
End If
OEarnPer = "" 'Ticket #19936
OEarnHrs = "" 'Ticket #19936
medAmount = 0

If glbWFC Then 'Ticket #29011 Franks 08/03/2016
    dlpDate(0).Text = CVDate(GetMonth("January") & " 1, " & Year(Now))
    dlpDate(1).Text = CVDate(GetMonth("December") & " 31, " & Year(Now))
Else
    dlpDate(0).Text = CVDate(GetMonth("January") & " 1, " & Year(Now))
    dlpDate(1).Text = CVDate(GetMonth("December") & " 31, " & Year(Now))
End If
'dlpDate(0).Text = glbCompEdFrom
'dlpDate(1).Text = glbCompEdTo
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"

'Ticket #24410 - City of Sarnia - Position Code
clpJob.Text = GetJHData(glbLEE_ID, "JH_JOB", "")

clpCode(1).SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HREARN", "Add")
Resume Next

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim xID As Long

On Error GoTo Add_Err

If Not chkEOTHERE Then Exit Sub

If fglbNew Then
    If Not AUDITOEAR("A") Then MsgBox "ERROR - AUDIT FILE"
Else
    If Not AUDITOEAR("M") Then MsgBox "ERROR - AUDIT FILE"
End If

'Ticket #22781
medAmount.Text = CDbl(medAmount)

Call UpdUStats(Me) ' update user's stats (who did it and when)
Call Set_Control("U", Me, rsDATA)

If glbCompSerial = "S/N - 2214W" Then 'Casey House
    rsDATA("ACT_DOLLAR") = Round53(Val(medAmount))
End If

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
DATA1.Refresh
fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(True)
If NextFormIF("Earning") Then
    Call cmdNew_Click
End If
Exit Sub

Add_Err:
If Err = 3022 Then
    DATA1.Recordset.CancelUpdate
    DATA1.Refresh
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREARN", "Update")
Resume Next
Unload Me

End Sub

Private Function Round53(xNumb)
Dim xInteger, xDecimal, xDecTmp
    xInteger = Int(xNumb * 1000)
    xDecimal = xNumb * 1000 - xInteger
    xDecTmp = 0
    If xDecimal >= 0 And xDecimal < 0.5 Then
        xDecTmp = 0
    End If
    If xDecimal >= 0.5 Then
        xDecTmp = 1
    End If
    Round53 = (xInteger + xDecTmp) / 1000
End Function
'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Other Earnings"
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

RHeading = lblEEName & "'s Other Earnings"
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

Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError
Screen.MousePointer = HOURGLASS
If glbtermopen Then
    SQLQ = "Select * from Term_EARN"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY FDATE DESC"
Else
    SQLQ = "Select * from HREARN"
    SQLQ = SQLQ & " where EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY FDATE DESC"
End If

DATA1.RecordSource = SQLQ
DATA1.Refresh
Set FRS = DATA1.Recordset.Clone

EERetrieve = True

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Earnings", "HREARN", "SELECT")
Resume Next

Exit Function

End Function

Private Sub cmdPostion_Click()
Dim OJOB As String, OJobD As String

OJOB = clpJob.Text
OJobD = clpJob.Caption

Load frmJOBS
frmJOBS.Show 1

If Len(glbPos) < 1 Then
    clpJob.Text = OJOB
    clpJob.Caption = OJobD
Else
    clpJob.Text = glbPos
    clpJob.Caption = glbPosDesc
End If
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
Me.cmdModify_Click
    glbOnTop = "FRMOTHERERN"
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMOTHERERN"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "FRMOTHERERN"


If glbWFC Then 'Ticket #29011 Franks 08/03/2016
    Call WFCScreenSetup
End If

If glbtermopen Then
    DATA1.ConnectionString = glbAdoIHRAUDIT
Else
    DATA1.ConnectionString = glbAdoIHRDB
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

vbxTrueGrid.FetchRowStyle = True
vbxTrueGrid.MarqueeStyle = 3

If Len(glbLEE_SName) < 1 Then Exit Sub
Screen.MousePointer = HOURGLASS
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "Other Earnings - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = ShowEmpnbr(lblEEID)

Call ST_UPD_MODE(True)
Call Display_Value

If Not gSec_Upd_Earnings Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
    cmdFlag.Enabled = False
End If


If glbCompSerial = "S/N - 2214W" Then 'Casey House
    medAmount.Format = "$##,##0.000"
End If

If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
    lbltitle(6).Visible = True
    medPPE.Visible = True
    medPPE.DataField = "EARNPCE"
    'Ticket #19936 Franks 04/15/2011
    lbltitle(7).Visible = True
    medHours.Visible = True
    medHours.DataField = "EARNHOURS"
Else
    vbxTrueGrid.Columns(5).Visible = False
End If

'Ticket #24410 - City of Sarnia - Position Code added
Call CR_JobHis_Snap

Call INI_Controls(Me)

'Ticket #24410 - City of Sarnia - Position Code added
clpJob.seleEMPCode = fglbJobList

Call TabOrderSetup

Screen.MousePointer = DEFAULT
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmOTHERERN = Nothing
    Call NextForm
End Sub

Private Sub medAmount_GotFocus()
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
cmdFlag.Enabled = TF

chkCOEFlag.Enabled = TF
medAmount.Enabled = TF
clpCode(1).Enabled = TF
dlpDate(0).Enabled = TF
dlpDate(1).Enabled = TF
'vbxTrueGrid.Enabled = FT
'memComments.Enabled = TF
memComments.Locked = FT
clpJob.Enabled = TF
MskTROIC.Enabled = TF

If DATA1.Recordset.EOF Or DATA1.Recordset.BOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If
End Sub

Private Sub medAmount_LostFocus()
    If Len(Trim(medAmount)) = 0 Then medAmount = 0
End Sub

Private Sub medHours_Change()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPPE_GotFocus()
Call SetPanHelp(ActiveControl)
medPPE = Val(medPPE) * 100
End Sub

Private Sub medPPE_LostFocus()
If Len(medPPE) > 0 Then
    If IsNumeric(medPPE) Then
        medPPE = Val(medPPE) / 100
    End If
End If
End Sub

Private Sub memComments_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo Eh
    'added by Bryan 18/Jan/06 Ticket#10222
    FRS.Requery
    FRS.Bookmark = Bookmark
    'Change the colour of a row
'    If FRS("BD_FREEZE") = True Then
'        RowStyle.ForeColor = vbRed
'    End If
'
Eh:
    Exit Sub
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
            SQLQ = "Select * from Term_EARN"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select * from HREARN"
            SQLQ = SQLQ & " where EMPNBR = " & glbLEE_ID
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        DATA1.RecordSource = SQLQ
        DATA1.Refresh
        Set FRS = DATA1.Recordset.Clone
        vbxTrueGrid.FetchRowStyle = True

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$
Dim SQLQ As String

On Error GoTo Tab1_Err
Call Display_Value

'If Not Fnd_Match_Data2() Then Exit Sub

' ' set description for code

Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HREARN", "Add")
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
    
    'Ticket #24410 - City of Sarnia - Position Code added
    Call CR_JobHis_Snap
    clpJob.seleEMPCode = fglbJobList
    
    
    If DATA1.Recordset.EOF Or DATA1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbtermopen Then
            rsDATA.Open DATA1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDATA.Open DATA1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        Call SET_UP_MODE
        Me.cmdModify_Click
        Exit Sub
    End If
    
    If glbtermopen Then
        SQLQ = "Select * from Term_EARN"
        SQLQ = SQLQ & " WHERE EARN_ID = " & DATA1.Recordset!EARN_ID
        SQLQ = SQLQ & " ORDER BY EARN_TYPE"
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        
    Else
        SQLQ = "Select * from HREARN"
        SQLQ = SQLQ & " where EARN_ID = " & DATA1.Recordset!EARN_ID
        SQLQ = SQLQ & " ORDER BY EARN_TYPE"
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If


    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call SET_UP_MODE
    Me.cmdModify_Click
    Call Set_Control("R", Me, rsDATA)

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
UpdateRight = gSec_Upd_Earnings
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
    frmOTHERERN.Caption = "Other Earnings - " & Left$(glbLEE_SName, 5)
    frmOTHERERN.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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

Private Sub TabOrderSetup()
    If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
        clpCode(1).TabIndex = 1
        dlpDate(0).TabIndex = 2
        dlpDate(1).TabIndex = 3
    End If
End Sub

Private Sub CR_JobHis_Snap()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String
Dim dynaJobHIS As New ADODB.Recordset

On Error GoTo JobHis_Err

fglbJobList = " "
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
    fglbJobList = ""
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
dynaJobHIS.Close
Set dynaJobHIS = Nothing

Screen.MousePointer = DEFAULT

Exit Sub

JobHis_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job List", "HR_JOB_History", "SELECT")
Screen.MousePointer = DEFAULT
Resume Next

End Sub

Private Sub WFCScreenSetup()
    clpJob.TextBoxWidth = 1315 '
    lbltitle(2).Caption = "Amount (in local currency)"
    lbltitle(8).Visible = True
    medCorpEqui.Visible = True
    medCorpEqui.DataField = "CORP_EQUI"
    
    'Ticket #29015 Franks 01/11/2017
    lbltitle(9).Visible = True 'ROIC
    MskTROIC.Visible = True
    MskTROIC.DataField = "IP_ROIC"
    
End Sub
