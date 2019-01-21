VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEODOLLARDTL 
   Appearance      =   0  'Flat
   Caption         =   "Dollary Entitlement - Actual Amounts Details"
   ClientHeight    =   7260
   ClientLeft      =   105
   ClientTop       =   1380
   ClientWidth     =   9885
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7260
   ScaleWidth      =   9885
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Entitlement"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   240
      TabIndex        =   24
      Top             =   2520
      Width           =   9495
      Begin INFOHR_Controls.DateLookup dlpFDate 
         DataField       =   "DA_FDATE"
         Height          =   285
         Left            =   1755
         TabIndex        =   25
         Tag             =   "41-Starting date"
         Top             =   840
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "DA_TYPE"
         Height          =   285
         Index           =   1
         Left            =   1755
         TabIndex        =   26
         Tag             =   "01-Entitlement - Code"
         Top             =   480
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOL"
      End
      Begin INFOHR_Controls.DateLookup dlpTDate 
         DataField       =   "DA_TDATE"
         Height          =   285
         Left            =   5760
         TabIndex        =   27
         Tag             =   "41-Ending date"
         Top             =   840
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
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
         Left            =   120
         TabIndex        =   32
         Top             =   525
         Width           =   960
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
         Left            =   120
         TabIndex        =   31
         Top             =   885
         Width           =   885
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   4830
         TabIndex        =   30
         Top             =   885
         Width           =   705
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
         Left            =   150
         TabIndex        =   29
         Top             =   120
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblDOH1 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2070
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   270
      Left            =   8010
      TabIndex        =   21
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      DataField       =   "DA_COMMENTS"
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
      TabIndex        =   5
      Tag             =   "00-Comments"
      Top             =   5310
      Width           =   6615
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feodollrDtl.frx":0000
      Height          =   2025
      Left            =   120
      OleObjectBlob   =   "feodollrDtl.frx":0014
      TabIndex        =   0
      Tag             =   "Listing of Dollar Entitlements"
      Top             =   480
      Width           =   9615
   End
   Begin INFOHR_Controls.DateLookup dlpPDate 
      DataField       =   "DA_PAIDDATE"
      Height          =   285
      Left            =   1965
      TabIndex        =   4
      Tag             =   "40-Date Paid"
      Top             =   4965
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin VB.TextBox txtRefer 
      Appearance      =   0  'Flat
      DataField       =   "DA_REFNBR"
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
      TabIndex        =   2
      Tag             =   "00-Reference Number"
      Top             =   4275
      Width           =   1230
   End
   Begin VB.TextBox txtPaidTo 
      Appearance      =   0  'Flat
      DataField       =   "DA_PAIDTO"
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
      TabIndex        =   3
      Tag             =   "00-Paid To"
      Top             =   4620
      Width           =   2985
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   16
      Top             =   6600
      Width           =   9885
      _Version        =   65536
      _ExtentX        =   17436
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
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   6000
         TabIndex        =   39
         Tag             =   "Print Division Listing"
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Tag             =   "Close and exit this screen"
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "Delete"
         Height          =   375
         Left            =   4800
         TabIndex        =   37
         Tag             =   "Delete the listed document"
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "New"
         Height          =   375
         Left            =   3870
         TabIndex        =   36
         Tag             =   "Attach a new document"
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2925
         TabIndex        =   35
         Tag             =   "Cancel the changes made"
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1995
         TabIndex        =   34
         Tag             =   "Save the changes made"
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1080
         TabIndex        =   33
         Tag             =   "Edit the Information"
         Top             =   120
         Width           =   915
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
      DataField       =   "DA_LDATE"
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
      Left            =   8160
      MaxLength       =   25
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DA_LTIME"
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
      Left            =   8520
      MaxLength       =   25
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DA_LUSER"
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
      Left            =   7800
      MaxLength       =   25
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9885
      _Version        =   65536
      _ExtentX        =   17436
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
         TabIndex        =   23
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   135
         Width           =   1740
      End
   End
   Begin MSMask.MaskEdBox medActualAmnt 
      DataField       =   "DA_ACTUAL"
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Tag             =   "20-Actual amount during the period"
      Top             =   3915
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   8040
      Top             =   6240
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   375
      Left            =   5160
      Top             =   6240
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
      TabIndex        =   22
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
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
      TabIndex        =   20
      Top             =   5310
      Width           =   1305
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
      TabIndex        =   19
      Top             =   4320
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
      TabIndex        =   18
      Top             =   4665
      Width           =   885
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Paid"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   330
      TabIndex        =   17
      Top             =   5010
      Width           =   705
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Amount"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   330
      TabIndex        =   15
      Top             =   3960
      Width           =   1245
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "DA_EMPNBR"
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
      Left            =   720
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "DA_COMPNO"
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
      Left            =   750
      TabIndex        =   14
      Top             =   5760
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEODOLLARDTL"
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
'If OBDOLLAR <> medEntitleAmnt Or OTYPE <> clpCode(1).Text Then GoTo MODUPD
If OADOLLAR <> medActualAmnt Then GoTo MODUPD 'Or OCOEFLAG <> chkCOEFlag Then GoTo MODUPD
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
    'rsTA("AU_BDOLLAR") = CDbl(medEntitleAmnt)       'Ticket #22781
    rsTA("AU_ADOLLAR") = CDbl(medActualAmnt)        'Ticket #22781
    'If chkCOEFlag = True Then
    '    rsTA("AU_COEFLAG") = "Y"
    'Else
    '    rsTA("AU_COEFLAG") = "N"
    'End If
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
'If Len(medEntitleAmnt) <= 0 Then medEntitleAmnt = 0
'If Not IsNumeric(medEntitleAmnt) Then medEntitleAmnt = 0
If Len(medActualAmnt) <= 0 Then medActualAmnt = 0
If Not IsNumeric(medActualAmnt) Then medActualAmnt = 0
'AVar = medEntitleAmnt - medActualAmnt
'lblVar = Format(AVar, "Currency")
End Sub

Private Function chkEODollar()
Dim SQLQ As String, Msg As String, dd#
Dim rsEmp As New ADODB.Recordset

chkEODollar = False

On Error GoTo chkEODollar_Err

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Entitlement code is a required field"
    'clpCode(1).SetFocus
    Exit Function
End If

If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Entitlement code must be valid"
    'clpCode(1).SetFocus
    Exit Function
End If

If Len(dlpFDate.Text) >= 1 Then
    If Not IsDate(dlpFDate.Text) Then
        MsgBox "From Date is not a valid date."
        'dlpFDate.SetFocus
        Exit Function
    End If
Else
    MsgBox "From Date is required."
    'dlpFDate.SetFocus
    Exit Function
End If

If Len(dlpTDate.Text) >= 1 Then
    If Not IsDate(dlpTDate.Text) Then
        MsgBox "To Date is not a valid date."
        'dlpTDate.SetFocus
        Exit Function
    End If
Else
    MsgBox "To Date is required."
    'dlpTDate.SetFocus
    Exit Function
End If

dd# = DateDiff("d", CVDate(dlpFDate.Text), CVDate(dlpTDate.Text))
If dd# < 0 Then
    MsgBox "From date must be earlier than To Date"
    'dlpFDate.SetFocus
    Exit Function
End If

'Ticket #28789 - Actual Amount is mandatory
If Len(Trim(medActualAmnt)) <= 0 Then medActualAmnt = 0
If medActualAmnt = 0 Or Not IsNumeric(medActualAmnt) Then
    MsgBox "Actual Amount is a required field"
    medActualAmnt.SetFocus
    Exit Function
End If

'If Len(Trim(medEntitleAmnt)) <= 0 Then medEntitleAmnt = 0

'Ticket #28789 - Date Paid is mandatory
If Len(Trim(dlpPDate.Text)) = 0 Then
    MsgBox "Date Paid is a required field"
    dlpPDate.SetFocus
    Exit Function
End If

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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkODollar", "HRDOLENT_ACTDTL", "edit/Add")
Resume Next

End Function

Sub cmdCancel_Click()
Dim X
On Error GoTo Can_Err

fglbNew = False

rsDATA.CancelUpdate
Call Display_Value

Call ST_UPD_MODE(False)  ' reset screen's attributes
'Call SET_UP_MODE

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRDOLENT_ACTDTL", "Cancel")
Resume Next

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Unload Me
'If glbOnTop = "FRMEODOLLARDTL" Then glbOnTop = ""
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
    'If gsAttachment_DB Then
    '    gdbAdoIhr001_DOC.Execute "Delete from Term_HRDOC_DOLENT_ACTDTL where DA_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " and DA_DOCKEY=" & glbDocKey & " " '
    'End If
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001.CommitTrans
    'If gsAttachment_DB Then
    '    gdbAdoIhr001_DOC.Execute "delete from HRDOC_HRDOLENT_ACTDTL where DA_TYPE='" & UCase(glbDocName) & "' AND DA_EMPNBR = " & glbLEE_ID & " and DA_DOCKEY=" & glbDocKey & " "
    'End If
    Data1.Refresh
End If
'If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
'End If
fglbNew = False
'Call ST_UPD_MODE(True)
'Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRDOLENT_ACTDTL", "Delete")

Resume Next
Unload Me

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Actn = "M"
'OBDOLLAR = medEntitleAmnt
'OTYPE = clpCode(1).Text
OADOLLAR = medActualAmnt
'OCOEFLAG = chkCOEFlag

Call ST_UPD_MODE(True)
'Call SET_UP_MODE
'clpCode(1).Enabled = True
'clpCode(1).SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HRDOLENT_ACTDTL", "Modify")
Call RollBack '23July99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
    Dim SQLQ As String
    
    fglbNew = True
    
    Call ST_UPD_MODE(True)
    'Call SET_UP_MODE
    
    On Error GoTo AddN_Err
    
    'If gsAttachment_DB And Not glbtermopen Then
    '    lblImport.Visible = True
    '    imgSec.Visible = False
    '    imgNoSec.Visible = True
    '    cmdImport.Visible = True
    'End If
    
    Actn = "A"
    OBDOLLAR = ""
    OTYPE = ""
    OADOLLAR = ""
    OCOEFLAG = False
    
    Call Set_Control("B", Me)
    
    rsDATA.AddNew
        
    'dlpFDate.Text = 1
    'If glbCompSerial = "S/N - 2288W" Or _
    '    glbCompSerial = "S/N - 2418W" Then ' Musashi Auto tkt# 10866 Ticket #17786 charton hobbs
    '    dlpFDate.Text = ""
    '    dlpTDate.Text = ""
    'Else
    '    dlpFDate.Text = CVDate(GetMonth("January") & " 1, " & Year(Now))
    '    dlpTDate.Text = CVDate(GetMonth("December") & " 31, " & Year(Now))
    'End If
    
    'lblVar = "$0.00"
    
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
    
    lblCNum.Caption = "001"
    'clpCode(1).SetFocus
    
    'Ticket #28789 - Set the values from Master table
    clpCode(1).Text = glbDolType
    dlpFDate.Text = glbDolFDate
    dlpTDate.Text = glbDolTDate
    
    'MDIMain.MainToolBar.ButtonS(8).Enabled = True
    'MDIMain.MainToolBar.ButtonS(9).Enabled = True
    
    Exit Sub

AddN_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRDOLENT_ACTDTL", "Add")
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

'If Not glbtermopen Then
'    If Not AUDITODOL(Actn) Then MsgBox "ERROR : AUDIT FILE"
'End If

'Ticket #22781
'medEntitleAmnt.Text = CDbl(medEntitleAmnt)
medActualAmnt.Text = CDbl(medActualAmnt)

Call UpdUStats(Me)
Call Set_Control("U", Me, rsDATA)

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    xID = rsDATA("DA_ENTITLE_ID")
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    xID = rsDATA("DA_ENTITLE_ID")
    Data1.Refresh
End If

'If gsAttachment_DB Then
'    If glbDocNewRecord Then 'New Record only
'        If Len(glbDocImpFile) > 0 Then
'            glbDocKey = xID
'            Call AttachmentAdd(glbLEE_ID, glbDocImpFile, glbDocType, glbDocDesc)
'        End If
'    End If
'    glbDocImpFile = ""
'End If


Call SET_UP_MODE
Call ST_UPD_MODE(False)

'If NextFormIF("Entitlement") Then
'    Call cmdNew_Click
'End If

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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRDOLENT_ACTDTL", "Update")
Resume Next
Unload Me

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Dollar Entitlement - Actual Amounts Details"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = lblEEName & "'s Dollar Entitlement - Actual Amounts Details"
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
    SQLQ = "SELECT * FROM HRDOLENT_ACTDTL WHERE DA_EMPNBR = " & lblEEID
    SQLQ = SQLQ & " AND DA_TYPE = '" & clpCode(1).Text & "'"
    SQLQ = SQLQ & " AND DA_FDATE = " & Date_SQL(dlpFDate.Text)
    SQLQ = SQLQ & " AND DA_TDATE = " & Date_SQL(dlpTDate.Text)
    SQLQ = SQLQ & " AND DA_PAIDDATE = " & Date_SQL(dlpPDate.Text)
    rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    If Not rsTB.EOF Then
       Logx = True
    End If
Else
    OID = Data1.Recordset("DA_ENTITLE_ID")
    SQLQ = "SELECT * FROM HRDOLENT_ACTDTL WHERE DA_EMPNBR = " & lblEEID
    SQLQ = SQLQ & " AND DA_TYPE = '" & clpCode(1).Text & "'"
    SQLQ = SQLQ & " AND DA_FDATE = " & Date_SQL(dlpFDate.Text)
    SQLQ = SQLQ & " AND DA_TDATE = " & Date_SQL(dlpTDate.Text)
    SQLQ = SQLQ & " AND DA_PAIDDATE = " & Date_SQL(dlpPDate.Text)
    SQLQ = SQLQ & " AND DA_ENTITLE_ID <> " & OID
    rsTB.Open SQLQ, gdbAdoIhr001
    If rsTB.RecordCount > 0 Then Logx = True
End If
rsTB.Close
If Logx = True Then
    Msg$ = "Duplicate exist. OK to proceed"
    MsgBox Msg$
    'clpCode(1).SetFocus
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


If glbtermopen Then
    SQLQ = "Select Term_DOLENT_ACTDTL.* from Term_DOLENT_ACTDTL"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " AND DA_TYPE = '" & clpCode(1).Text & "'"
    SQLQ = SQLQ & " AND DA_FDATE = " & Date_SQL(dlpFDate.Text)
    SQLQ = SQLQ & " AND DA_TDATE = " & Date_SQL(dlpTDate.Text)
    SQLQ = SQLQ & " ORDER BY DA_PAIDDATE DESC"
Else
    SQLQ = "Select HRDOLENT_ACTDTL.* from HRDOLENT_ACTDTL"
    SQLQ = SQLQ & " WHERE DA_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND DA_TYPE = '" & clpCode(1).Text & "'"
    SQLQ = SQLQ & " AND DA_FDATE = " & Date_SQL(dlpFDate.Text)
    SQLQ = SQLQ & " AND DA_TDATE = " & Date_SQL(dlpTDate.Text)
    SQLQ = SQLQ & " ORDER BY DA_PAIDDATE DESC"
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HRDOLENT_ACTDTL", "SELECT")
Resume Next

Exit Function

End Function

Private Sub cmdImport_Click()
    glbDocNewRecord = fglbNew
    glbDocName = "DollarEntActDtl"
    If fglbNew Then
        glbDocKey = 0
    Else
        glbDocKey = rsDATA("DA_ENTITLE_ID")
    End If
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmEODOLLARDTL")
End Sub

Private Sub Form_Activate()
'Call SET_UP_MODE
'Call ST_UPD_MODE(False)
'Me.cmdModify_Click
    glbOnTop = "FRMEODOLLARDTL"
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEODOLLARDTL"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "FRMEODOLLARDTL"

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

Screen.MousePointer = HOURGLASS

Screen.MousePointer = DEFAULT

'Ticket #28789 - Set the values from Master table
clpCode(1).Text = glbDolType
dlpFDate.Text = glbDolFDate
dlpTDate.Text = glbDolTDate

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
    'Me.Show
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

Call Display_Value

'If Not gSec_Upd_Other_Entitlements Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
'    cmdFlag.Enabled = False
'End If

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
    Set frmEODOLLARDTL = Nothing
    'Call NextForm
End Sub

Private Sub medActualAmnt_Change()
    'Call Calc_Var
End Sub

Private Sub medActualAmnt_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medActualAmnt_LostFocus()
    'Call Calc_Var
End Sub

Private Sub medEntitleAmnt_Change()
    'Call Calc_Var
End Sub

Private Sub medEntitleAmnt_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medEntitleAmnt_LostFocus()
    'Call Calc_Var
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

cmdOK.Enabled = TF
cmdCancel.Enabled = TF

''cmdFlag.Enabled = FT 'TF

cmdClose.Enabled = FT
cmdModify.Enabled = FT
cmdNew.Enabled = FT
cmdDelete.Enabled = FT
cmdPrint.Enabled = FT

medActualAmnt.Enabled = TF
'medEntitleAmnt.Enabled = TF
'clpCode(1).Enabled = TF
'dlpFDate.Enabled = TF
'dlpTDate.Enabled = TF
'chkCOEFlag.Enabled = TF
txtRefer.Enabled = TF
txtPaidTo.Enabled = TF
dlpPDate.Enabled = TF
'txtComments.Enabled = TF
txtComments.Locked = FT

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End If
'vbxTrueGrid.Enabled = FT

'glbDocName = "DollarEntActDtl"
'If gsAttachment_DB Then
'    glbDocKey = 0
'    If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
'        If rsDATA.RecordCount > 0 Then
'            If Not IsNull(rsDATA("DA_DOCKEY")) Then
'                glbDocKey = rsDATA("DA_DOCKEY")
'            Else
'                glbDocKey = 0
'            End If
'        Else
'            If Not IsNull(Data1.Recordset("DA_DOCKEY")) Then
'                glbDocKey = Data1.Recordset("DA_DOCKEY")
'            Else
'                glbDocKey = 0
'            End If
'        End If
'    End If
    
'    Call DispimgIcon(Me, "frmEODOLLARDTL")
'    If gSec_Upd_Other_Entitlements And Not glbtermopen Then
'        If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'            cmdImport.Visible = False
'        Else
'            cmdImport.Visible = True
'        End If
'    End If
'End If

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
        SQLQ = "Select Term_DOLENT_ACTDTL.* FROM Term_DOLENT_ACTDTL"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " AND DA_TYPE = '" & clpCode(1).Text & "'"
        SQLQ = SQLQ & " AND DA_FDATE = " & Date_SQL(dlpFDate.Text)
        SQLQ = SQLQ & " AND DA_TDATE = " & Date_SQL(dlpTDate.Text)
    Else
        SQLQ = "Select HRDOLENT_ACTDTL.* FROM HRDOLENT_ACTDTL"
        SQLQ = SQLQ & " WHERE DA_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND DA_TYPE = '" & clpCode(1).Text & "'"
        SQLQ = SQLQ & " AND DA_FDATE = " & Date_SQL(dlpFDate.Text)
        SQLQ = SQLQ & " AND DA_TDATE = " & Date_SQL(dlpTDate.Text)
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HRDOLENT_ACTDTL", "Display Row")
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
        
    'Ticket #28789 - Set the values from Master table
    clpCode(1).Text = glbDolType
    dlpFDate.Text = glbDolFDate
    dlpTDate.Text = glbDolTDate
    
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    If glbtermopen Then
        SQLQ = "Select Term_DOLENT_ACTDTL.* FROM Term_DOLENT_ACTDTL"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " AND DA_TYPE = '" & clpCode(1).Text & "'"
        SQLQ = SQLQ & " AND DA_FDATE = " & Date_SQL(dlpFDate.Text)
        SQLQ = SQLQ & " AND DA_TDATE = " & Date_SQL(dlpTDate.Text)
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "Select HRDOLENT_ACTDTL.* FROM HRDOLENT_ACTDTL"
        SQLQ = SQLQ & " WHERE DA_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND DA_TYPE = '" & clpCode(1).Text & "'"
        SQLQ = SQLQ & " AND DA_FDATE = " & Date_SQL(dlpFDate.Text)
        SQLQ = SQLQ & " AND DA_TDATE = " & Date_SQL(dlpTDate.Text)
        If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
            rsDATA.CursorLocation = adUseServer
        End If
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    
    Call SET_UP_MODE
    'Me.cmdModify_Click
     
    Exit Sub
End If
    
If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
If glbtermopen Then
    SQLQ = "Select Term_DOLENT_ACTDTL.* FROM Term_DOLENT_ACTDTL"
    SQLQ = SQLQ & " WHERE DA_ENTITLE_ID = " & Data1.Recordset!DA_ENTITLE_ID
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "Select HRDOLENT_ACTDTL.* FROM HRDOLENT_ACTDTL"
    SQLQ = SQLQ & " WHERE DA_ENTITLE_ID = " & Data1.Recordset!DA_ENTITLE_ID
    If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
        rsDATA.CursorLocation = adUseServer
    End If
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If
 
If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
Call Set_Control("R", Me, rsDATA)
Call SET_UP_MODE
Call ST_UPD_MODE(False)
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
    frmEODOLLARDTL.Caption = "Dollar Entitlements - " & Left$(glbLEE_SName, 5)
    frmEODOLLARDTL.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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
    SQLQ = getSQL("frmEODOLLARDTL")
    Call FillMemoFile(SQLQ, "DollarEntActDtl")
End Sub


