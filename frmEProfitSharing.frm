VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEProfitSharing 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Profit Sharing"
   ClientHeight    =   8700
   ClientLeft      =   105
   ClientTop       =   1380
   ClientWidth     =   11760
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
   ScaleHeight     =   8700
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
      DataField       =   "PS_NOTE"
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
      Left            =   2715
      MaxLength       =   500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Tag             =   "00-Comments"
      Top             =   6600
      Width           =   6555
   End
   Begin VB.ComboBox comPSType 
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
      ItemData        =   "frmEProfitSharing.frx":0000
      Left            =   7980
      List            =   "frmEProfitSharing.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Tag             =   "Type"
      Top             =   5145
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtPSType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   225
      Left            =   9600
      TabIndex        =   34
      Top             =   5415
      Visible         =   0   'False
      Width           =   345
   End
   Begin INFOHR_Controls.DateLookup dlpPDate 
      DataField       =   "PS_DATEPAID"
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Tag             =   "40-Date Paid"
      Top             =   5880
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpFDate 
      DataField       =   "PS_FDATE"
      Height          =   285
      Left            =   2400
      TabIndex        =   9
      Tag             =   "41-Starting date"
      Top             =   5505
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
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
      TabIndex        =   25
      Top             =   8040
      Width           =   11760
      _Version        =   65536
      _ExtentX        =   20743
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
      DataField       =   "PS_LDATE"
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
      Left            =   7800
      MaxLength       =   25
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PS_LTIME"
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
      Left            =   8160
      MaxLength       =   25
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PS_LUSER"
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
      Left            =   7440
      MaxLength       =   25
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11760
      _Version        =   65536
      _ExtentX        =   20743
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
         TabIndex        =   28
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   135
         Width           =   1740
      End
   End
   Begin MSMask.MaskEdBox medPSAmnt 
      DataField       =   "PS_AMOUNT"
      Height          =   285
      Left            =   2715
      TabIndex        =   12
      Tag             =   "20-Amount of entitlement during the period"
      Top             =   6240
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
   Begin INFOHR_Controls.DateLookup dlpTDate 
      DataField       =   "PS_TDATE"
      Height          =   285
      Left            =   6480
      TabIndex        =   10
      Tag             =   "41-Ending date"
      Top             =   5505
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PS_SECTION"
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   3
      Tag             =   "00-Section"
      Top             =   3720
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
      MaxLength       =   8
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PS_ADMINBY"
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   0
      Tag             =   "00-Administered By"
      Top             =   2640
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PS_EMP"
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   6
      Tag             =   "00-Enter Status Code"
      Top             =   4800
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      DataField       =   "PS_DEPTNO"
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Tag             =   "00-Department"
      Top             =   3360
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      DataField       =   "PS_DIV"
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Tag             =   "00-Division"
      Top             =   3000
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PS_PCODE"
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   7
      Tag             =   "00-Enter Profit Sharing Type"
      Top             =   5160
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "PSTY"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PS_REGION"
      Height          =   285
      Index           =   4
      Left            =   2400
      TabIndex        =   5
      Tag             =   "00-Region"
      Top             =   4440
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PS_LOC"
      Height          =   285
      Index           =   5
      Left            =   2400
      TabIndex        =   4
      Tag             =   "00-Location - Code"
      Top             =   4080
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmEProfitSharing.frx":0017
      Height          =   2025
      Left            =   0
      OleObjectBlob   =   "frmEProfitSharing.frx":002B
      TabIndex        =   39
      Tag             =   "Listing of Dollar Entitlements"
      Top             =   480
      Width           =   11655
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   13
      Left            =   240
      TabIndex        =   38
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   12
      Left            =   240
      TabIndex        =   37
      Top             =   4080
      Width           =   750
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
      Index           =   11
      Left            =   240
      TabIndex        =   36
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   315
      Index           =   10
      Left            =   240
      TabIndex        =   35
      Top             =   5145
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   33
      Top             =   3360
      Width           =   990
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   32
      Top             =   3000
      Width           =   690
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Section"
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   31
      Top             =   3720
      Width           =   660
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   30
      Top             =   2655
      Width           =   1365
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   29
      Top             =   4800
      Width           =   1620
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   5520
      TabIndex        =   27
      Top             =   5505
      Width           =   705
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Paid"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   26
      Top             =   5880
      Width           =   1185
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   24
      Top             =   6240
      Width           =   1650
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   23
      Top             =   5505
      Width           =   1605
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "PS_EMPNBR"
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
      Left            =   9030
      TabIndex        =   21
      Top             =   6360
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "PS_COMPNO"
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
      Left            =   7350
      TabIndex        =   22
      Top             =   6360
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEProfitSharing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim Actn, OBDOLLAR, OADOLLAR, OTYPE, OCOEFLAG
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim fglbNew As Integer


'Private Function AUDITODOL(ACTX)
'Dim rsTA As New ADODB.Recordset
'Dim rsTB As New ADODB.Recordset
'Dim xADD As Boolean, xPT As String, xDiv As String
'Dim strFields As String
'
'On Error GoTo AUDIT_ERR
'
'AUDITODOL = False
'
'
''rsTB.Open "HREMP", gdbAdoIhr001, adOpenKeyset, , adCmdTableDirect
''rsTB.Find "ED_EMPNBR = " & glbLEE_ID
'rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
'
'If Not rsTB.EOF Then
'    'xPT = rsTB("ED_PT")
'    'xDiv = rsTB("ED_DIV")
'    If IsNull(rsTB("ED_PT")) Then
'        xPT = ""
'    Else
'        xPT = rsTB("ED_PT")
'    End If
'    If IsNull(rsTB("ED_DIV")) Then
'        xDiv = ""
'    Else
'        xDiv = rsTB("ED_DIV")
'    End If
'Else
'    xPT = ""
'    xDiv = ""
'End If
''strFields added by Bryan 02/Dec/05 Ticket#9899
'strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCOPS_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCOPS_TABL, AU_TREAS_TABL, "
'strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_DOLENT, AU_SAFETY, AU_UNIFORM, AU_EQUIP, AU_CLEAN, "
'strFields = strFields & "AU_BDOLLAR, AU_ADOLLAR, AU_COEFLAG, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID, "
'strFields = strFields & "AU_DOFDATE, AU_DOTDATE, AU_REFNBR, AU_PAIDTO, AU_PAIDDATE"
'rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
'
'xADD = False
'
'If ACTX = "D" Or ACTX = "A" Then GoTo MODUPD
'If OBDOLLAR <> medPSAmnt Or OTYPE <> clpCode(1).Text Then GoTo MODUPD
'GoTo MODNOUPD
'
'MODUPD:
'rsTA.AddNew
'rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM"
'rsTA("AU_SUPCOPS_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP"
'rsTA("AU_BCOPS_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL"
'rsTA("AU_DOLENT_TABL") = "EDOL"   '"EARN"
'rsTA("AU_NEWEMP") = "N"
'rsTA("AU_PTUPL") = xPT
'rsTA("AU_DIVUPL") = xDiv
'
'If ACTX = "D" Then
'
'    rsTA("AU_DOFDATE") = dlpFDate.Text
'    rsTA("AU_DOTDATE") = dlpTDate.Text
'Else
'
'    rsTA("AU_BDOLLAR") = medPSAmnt
'
'    rsTA("AU_DOFDATE") = dlpFDate.Text
'    rsTA("AU_DOTDATE") = dlpTDate.Text
'    If IsDate(dlpPDate.Text) Then 'Ticket #20274 Franks 05/05/2011
'        rsTA("AU_PAIDDATE") = dlpPDate.Text
'    End If
'End If
'
'Dim rsEMP As New ADODB.Recordset
'Dim SQLQ
'SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
'rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
'If Not rsEMP.EOF Then
'    If Not IsNull(rsEMP("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEMP("ED_PAYROLL_ID")
'End If
'rsEMP.Close
'
'rsTA("AU_COMPNO") = "001"
'rsTA("AU_EMPNBR") = glbLEE_ID
'rsTA("AU_LDATE") = Date
'rsTA("AU_LUSER") = glbUserID
'rsTA("AU_LTIME") = Time$
'rsTA("AU_UPLOAD") = "N"
'rsTA("AU_TYPE") = ACTX
'rsTA.Update
'
'MODNOUPD:
'AUDITODOL = True
'Exit Function
'
'AUDIT_ERR:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
'Call RollBack '23July99 js
'
'End Function

Private Sub Calc_Var()
Dim AVar
'AVar = 0
'If Len(medPSAmnt) <= 0 Then medPSAmnt = 0
'If Not IsNumeric(medPSAmnt) Then medPSAmnt = 0
'If Len(medActualAmnt) <= 0 Then medActualAmnt = 0
'If Not IsNumeric(medActualAmnt) Then medActualAmnt = 0

End Sub

Private Function chkProfitSharing()
Dim SQLQ As String, Msg As String, dd#
Dim rsEMP As New ADODB.Recordset

chkProfitSharing = False

On Error GoTo chkProfitSharing_Err

If Len(clpCode(1).Text) < 1 Then
    MsgBox lblTitle(8).Caption & " code is a required field"
    clpCode(1).SetFocus
    Exit Function
End If
If clpCode(1).Caption = "Unassigned" Then
    MsgBox lblTitle(8).Caption & " code must be valid"
    clpCode(1).SetFocus
    Exit Function
End If
If Len(clpDiv.Text) < 1 Then
    MsgBox lblTitle(1).Caption & " code is a required field"
    clpDiv.SetFocus
    Exit Function
End If
If clpDiv.Caption = "Unassigned" Then
    MsgBox lblTitle(1).Caption & " code must be valid"
    clpDiv.SetFocus
    Exit Function
End If
If Len(clpDept.Text) < 1 Then
    MsgBox lblTitle(5).Caption & " code is a required field"
    clpDept.SetFocus
    Exit Function
End If
If clpDept.Caption = "Unassigned" Then
    MsgBox lblTitle(5).Caption & " code must be valid"
    clpDept.SetFocus
    Exit Function
End If
If Len(clpCode(2).Text) < 1 Then
    MsgBox lblTitle(9).Caption & " code is a required field"
    clpCode(2).SetFocus
    Exit Function
End If
If clpCode(2).Caption = "Unassigned" Then
    MsgBox lblTitle(9).Caption & " code must be valid"
    clpCode(2).SetFocus
    Exit Function
End If
'Ticket #22675 Franks 11/27/2012 - begin
If Len(clpCode(5).Text) < 1 Then 'Location
    MsgBox lblTitle(12).Caption & " code is a required field"
    clpCode(5).SetFocus
    Exit Function
End If
If clpCode(5).Caption = "Unassigned" Then
    MsgBox lblTitle(12).Caption & " code must be valid"
    clpCode(5).SetFocus
    Exit Function
End If
If Len(clpCode(4).Text) < 1 Then 'Location
    MsgBox lblTitle(13).Caption & " code is a required field"
    clpCode(4).SetFocus
    Exit Function
End If
If clpCode(4).Caption = "Unassigned" Then
    MsgBox lblTitle(13).Caption & " code must be valid"
    clpCode(4).SetFocus
    Exit Function
End If
'Ticket #22675 Franks 11/27/2012 - end
If Len(clpCode(3).Text) < 1 Then
    MsgBox lblTitle(6).Caption & " code is a required field"
    clpCode(3).SetFocus
    Exit Function
End If
If clpCode(3).Caption = "Unassigned" Then
    MsgBox lblTitle(6).Caption & " code must be valid"
    clpCode(3).SetFocus
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

'Ticket #21651 Franks 03/16/2012
'If Len(comPSType.Text) = 0 Then
'    MsgBox "Type is required."
'    comPSType.SetFocus
'    Exit Function
'End If
If Len(clpCode(0).Text) = 0 Then
    MsgBox "Type is required."
    clpCode(0).SetFocus
    Exit Function
End If

If Len(dlpPDate.Text) > 0 Then
    If Not IsDate(dlpPDate.Text) Then
        MsgBox "Date Paid is not a valid date"
        dlpPDate.SetFocus
        Exit Function
    End If
    
    rsEMP.Open "SELECT ED_DOH FROM HREMP WHERE ED_EMPNBR = " & lblEENum, gdbAdoIhr001, adOpenStatic
    If Not rsEMP.EOF Then
        If DaysBetween(rsEMP("ED_DOH"), dlpPDate.Text) < 0 Then
            MsgBox "Date Paid can not be prior to Original Hire date"
            dlpPDate.SetFocus
            rsEMP.Close
            Exit Function
        End If
    End If
    rsEMP.Close
Else
    MsgBox "Date Paid is required."
    dlpPDate.SetFocus
    Exit Function
End If
If Len(medPSAmnt.Text) = 0 Then
    MsgBox "Amount is required."
    medPSAmnt.SetFocus
    Exit Function
Else
    If Not IsNumeric(medPSAmnt.Text) Then
        MsgBox "Amount must be numeric."
        medPSAmnt.SetFocus
        Exit Function
    End If
End If

chkProfitSharing = True

Exit Function

chkProfitSharing_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkProfitSharing", "HR_PROFIT_SHARING", "edit/Add")
Resume Next

End Function

Sub cmdCancel_Click()
Dim x
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_PROFIT_SHARING", "Cancel")
Resume Next

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "frmEProfitSharing" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, x

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

'If Not AUDITODOL("D") Then MsgBox "ERROR : AUDIT FILE"

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

Call Display_Value

fglbNew = False
Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_PROFIT_SHARING", "Delete")

Resume Next
Unload Me

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub


Sub cmdModify_Click()

On Error GoTo Mod_Err

Actn = "M"
OBDOLLAR = medPSAmnt
'OTYPE = clpCode(1).Text

Call SET_UP_MODE
'clpCode(1).Enabled = True
'clpCode(1).SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_PROFIT_SHARING", "Modify")
Call RollBack '23July99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
    Dim SQLQ As String
    
    fglbNew = True

    Call SET_UP_MODE
    
    On Error GoTo AddN_Err
    
    
    Actn = "A"
    OBDOLLAR = ""
    'OTYPE = ""
    'OADOLLAR = ""
    'OCOEFLAG = False
    
    Call Set_Control("B", Me)
    
    rsDATA.AddNew

    'dlpFDate.Text = CVDate(GetMonth("January") & " 1, " & Year(Now))
    'dlpTDate.Text = CVDate(GetMonth("December") & " 31, " & Year(Now))
    dlpFDate.Text = ""
    dlpTDate.Text = ""

    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
    
    Call getDefaultsFromHREMP
    
    lblCNum.Caption = "001"
    clpCode(1).SetFocus
    
    MDIMain.MainToolBar.ButtonS(8).Enabled = True
    MDIMain.MainToolBar.ButtonS(9).Enabled = True
    
    Exit Sub

AddN_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_PROFIT_SHARING", "Add")
    Resume Next
    
End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim xID As Long

On Error GoTo Add_Err

fglbNew = False

If Not chkProfitSharing() Then Exit Sub

If Data1.Recordset.RecordCount > 0 Then
  If ChkDup(Actn) Then Exit Sub
End If

'If Not glbtermopen Then
'    If Not AUDITODOL(Actn) Then MsgBox "ERROR : AUDIT FILE"
'End If

Call UpdUStats(Me)
Call Set_Control("U", Me, rsDATA)

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    xID = rsDATA("PS_ID")
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    xID = rsDATA("PS_ID")
    Data1.Refresh
End If

Call SET_UP_MODE

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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_PROFIT_SHARING", "Update")
Resume Next
Unload Me

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Profit Sharing"
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

RHeading = lblEEName & "'s Profit Sharing"
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
    SQLQ = "SELECT * FROM HR_PROFIT_SHARING WHERE PS_EMPNBR = " & lblEEID
    SQLQ = SQLQ & " AND PS_FDATE = " & Date_SQL(dlpFDate.Text)
    SQLQ = SQLQ & " AND PS_TDATE = " & Date_SQL(dlpTDate.Text)
    SQLQ = SQLQ & " AND PS_PCODE = '" & clpCode(0).Text & "' "
    rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    If Not rsTB.EOF Then
       Logx = True
    End If
Else
    OID = Data1.Recordset("PS_ID")
    SQLQ = "SELECT * FROM HR_PROFIT_SHARING WHERE PS_EMPNBR = " & lblEEID
    SQLQ = SQLQ & " AND PS_FDATE = " & Date_SQL(dlpFDate.Text)
    SQLQ = SQLQ & " AND PS_TDATE = " & Date_SQL(dlpTDate.Text)
    SQLQ = SQLQ & " AND PS_PCODE = '" & clpCode(0).Text & "' "
    SQLQ = SQLQ & " AND PS_ID <> " & OID
    rsTB.Open SQLQ, gdbAdoIhr001
    If rsTB.RecordCount > 0 Then Logx = True
End If
rsTB.Close
If Logx = True Then
    Msg$ = "Duplicate exist. OK to proceed"
    MsgBox Msg$
    clpCode(1).SetFocus
End If


ChkDup = Logx ' True

End Function
Function EERetrieve()
Dim SQLQ, SQLQ1 As String
Dim rsDOH As New ADODB.Recordset
EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS


If glbtermopen Then         'Lucy July 5, 2000
    SQLQ = "Select * FROM Term_PROFIT_SHARING"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY PS_FDATE DESC"
Else
    SQLQ = "Select * from HR_PROFIT_SHARING"
    SQLQ = SQLQ & " WHERE PS_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY PS_FDATE DESC"
  
End If

Data1.RecordSource = SQLQ
Data1.Refresh


EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HR_PROFIT_SHARING", "SELECT")
Resume Next

Exit Function

End Function





Private Sub clpCode_LostFocus(Index As Integer)
    If Index = 0 Then
        If fglbNew Then
            Call ToDateCalculate(dlpFDate.Text)
        End If
    End If
End Sub

Private Sub comPSType_Click()
txtPSType.Text = Left(comPSType.Text, 1)
End Sub

Private Sub comPSType_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comPSType_LostFocus()
'Call ToDateCalculate(dlpFDate.Text)
End Sub

Private Sub dlpFDate_LostFocus()
    If fglbNew Then
        Call ToDateCalculate(dlpFDate.Text)
    End If
End Sub

Private Sub Form_Activate()
    Call SET_UP_MODE
    glbOnTop = "frmEProfitSharing"
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "frmEProfitSharing"
End Sub

Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer

glbOnTop = "frmEProfitSharing"

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

Screen.MousePointer = HOURGLASS

Screen.MousePointer = DEFAULT

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
    
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

'Ticket #22453 Franks 08/24/2012 - begin
If glbNoNONE Then
    If glbUNION = "NONE" Then
        MsgBox "You Do Not Have Authority For This Transaction"
        glbOnTop = Empty
        Unload Me
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
End If
If glbNoEXEC Then       'Hemu -EXE
    If glbUNION = "EXEC" Then
        MsgBox "You Do Not Have Authority For This Transaction"
        glbOnTop = Empty
        Unload Me
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
End If
'Ticket #22453 Franks 08/24/2012 - end
    
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
    Me.Caption = "Profit Sharing - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = ShowEmpnbr(lblEEID)

comPSType.Clear
comPSType.AddItem "Annual"
comPSType.AddItem "Quarterly"
comPSType.ListIndex = -1

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Call LabelMasterFun

Call Display_Value

Call ST_UPD_MODE(False)

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
    Set frmEProfitSharing = Nothing
    Call NextForm
End Sub

Private Sub medPSAmnt_Change()
    Call Calc_Var
End Sub

Private Sub medPSAmnt_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPSAmnt_LostFocus()
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

clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
clpDiv.Enabled = TF
clpDept.Enabled = TF
medPSAmnt.Enabled = TF

dlpFDate.Enabled = TF
dlpTDate.Enabled = TF
dlpPDate.Enabled = TF
'comPSType.Enabled = TF
clpCode(0).Enabled = TF
memComments.Enabled = TF

End Sub

Private Sub Text1_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub




Private Sub txtPSType_Change()
If txtPSType.Text = "A" Or txtPSType.Text = "Q" Then
    If txtPSType.Text = "A" Then
        comPSType.ListIndex = 0
    Else
        comPSType.ListIndex = 1
    End If
Else
    comPSType.ListIndex = -1
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
        
        If glbtermopen Then         'Lucy July 5, 2000
            SQLQ = "Select Term_PROFIT_SHARING.* from Term_PROFIT_SHARING"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select HR_PROFIT_SHARING.* from HR_PROFIT_SHARING"
            SQLQ = SQLQ & " where PS_EMPNBR = " & glbLEE_ID
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_PROFIT_SHARING", "Add")
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
        SQLQ = "Select * from Term_PROFIT_SHARING "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "Select HR_PROFIT_SHARING.* from HR_PROFIT_SHARING"
        SQLQ = SQLQ & " where PS_EMPNBR = " & glbLEE_ID
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
    SQLQ = "Select * from Term_PROFIT_SHARING "
    SQLQ = SQLQ & " WHERE PS_ID = " & Data1.Recordset!PS_ID
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "Select HR_PROFIT_SHARING.*  from HR_PROFIT_SHARING"
    SQLQ = SQLQ & " where PS_ID = " & Data1.Recordset!PS_ID
    If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
        rsDATA.CursorLocation = adUseServer
    End If
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
UpdateRight = gSec_Upd_Profit_Sharing
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
    frmEProfitSharing.Caption = "Profit Sharing - " & Left$(glbLEE_SName, 5)
    frmEProfitSharing.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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

Private Sub LabelMasterFun()
lblTitle(8).Caption = lStr("Administered By")
lblTitle(1).Caption = lStr("Division")
lblTitle(5).Caption = lStr("Department")
lblTitle(9).Caption = lStr("Section")
lblTitle(12).Caption = lStr("Location")
lblTitle(13).Caption = lStr("Region")
'lblTitle(6).Caption = lStr("Employment Status")
vbxTrueGrid.Columns(0).Caption = lStr("Administered By")
vbxTrueGrid.Columns(1).Caption = lStr("Division")
vbxTrueGrid.Columns(2).Caption = lStr("Department")
vbxTrueGrid.Columns(3).Caption = lStr("Section")
vbxTrueGrid.Columns(4).Caption = lStr("Location")
vbxTrueGrid.Columns(5).Caption = lStr("Region")
'vbxTrueGrid.Columns(4).Caption = lStr("Employment Status")
End Sub

Private Sub getDefaultsFromHREMP()
Dim rsTmpE As New ADODB.Recordset
Dim SQLQ As String

    If glbtermopen Then
        SQLQ = "SELECT * FROM Term_HREMP WHERE TERM_SEQ = " & glbTERM_Seq & " "
    Else
        SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
    End If
    rsTmpE.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTmpE.EOF Then
        If Not IsNull(rsTmpE("ED_ADMINBY")) Then clpCode(1).Text = rsTmpE("ED_ADMINBY")
        If Not IsNull(rsTmpE("ED_DIV")) Then clpDiv.Text = rsTmpE("ED_DIV")
        If Not IsNull(rsTmpE("ED_DEPTNO")) Then clpDept.Text = rsTmpE("ED_DEPTNO")
        If Not IsNull(rsTmpE("ED_SECTION")) Then clpCode(2).Text = rsTmpE("ED_SECTION")
        If Not IsNull(rsTmpE("ED_EMP")) Then clpCode(3).Text = rsTmpE("ED_EMP")
    End If
    rsTmpE.Close

End Sub

Private Sub ToDateCalculate(xFromDate)
Dim toDate
Dim xtmpdate
    toDate = dlpTDate.Text
    If IsDate(xFromDate) Then
        ''If comPSType.Text = "Quarterly" Or comPSType.Text = "Annual" Then
        ''    If comPSType.Text = "Quarterly" Then
        ''        xtmpdate = DateAdd("M", 3, xFromDate)
        ''    End If
        ''    If comPSType.Text = "Annual" Then
        ''        xtmpdate = DateAdd("M", 12, xFromDate)
        ''    End If
        ''End If
        'Ticket #21651 Franks 03/16/2012
        If clpCode(0).Text = "QU" Then
            xtmpdate = DateAdd("M", 3, xFromDate)
        Else
            xtmpdate = DateAdd("M", 12, xFromDate)
        End If
        toDate = DateAdd("d", -1, xtmpdate)
    End If
    dlpTDate.Text = toDate
End Sub
