VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEUserDef 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "User Defined Table"
   ClientHeight    =   9630
   ClientLeft      =   -150
   ClientTop       =   765
   ClientWidth     =   10050
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
   ScaleHeight     =   9630
   ScaleWidth      =   10050
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      DataField       =   "UD_TEXT2"
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
      Left            =   1995
      MaxLength       =   25
      TabIndex        =   16
      Tag             =   "00-Text"
      Top             =   7200
      Width           =   3225
   End
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      DataField       =   "UD_TEXT1"
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
      Left            =   1995
      MaxLength       =   25
      TabIndex        =   15
      Tag             =   "00-Text"
      Top             =   6840
      Width           =   3225
   End
   Begin VB.TextBox txtCode1TabName 
      Appearance      =   0  'Flat
      DataField       =   "UD_CODE1_TABL"
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
      Left            =   7080
      MaxLength       =   25
      TabIndex        =   45
      TabStop         =   0   'False
      Text            =   "COD1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtFlag 
      Appearance      =   0  'Flat
      DataField       =   "UD_FLAG5"
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
      Left            =   7680
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   6375
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkFlag 
      Alignment       =   1  'Right Justify
      Caption         =   "UFlag 5 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   5520
      TabIndex        =   14
      Top             =   6345
      Width           =   1875
   End
   Begin VB.TextBox txtFlag 
      Appearance      =   0  'Flat
      DataField       =   "UD_FLAG4"
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
      Left            =   7680
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   6015
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkFlag 
      Alignment       =   1  'Right Justify
      Caption         =   "UFlag 4 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   5520
      TabIndex        =   13
      Tag             =   " "
      Top             =   6000
      Width           =   1875
   End
   Begin VB.TextBox txtFlag 
      Appearance      =   0  'Flat
      DataField       =   "UD_FLAG3"
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
      Left            =   7680
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5655
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkFlag 
      Alignment       =   1  'Right Justify
      Caption         =   "UFlag 3 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   5520
      TabIndex        =   12
      Top             =   5625
      Width           =   1875
   End
   Begin VB.TextBox txtFlag 
      Appearance      =   0  'Flat
      DataField       =   "UD_FLAG2"
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
      Left            =   7680
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   5295
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkFlag 
      Alignment       =   1  'Right Justify
      Caption         =   "UFlag 2 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   5520
      TabIndex        =   11
      Top             =   5265
      Width           =   1875
   End
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      DataField       =   "UD_COMMENTS"
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
      Left            =   1995
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Tag             =   "00-Comments"
      Top             =   7590
      Width           =   6615
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feuserdef.frx":0000
      Height          =   2085
      Left            =   90
      OleObjectBlob   =   "feuserdef.frx":0014
      TabIndex        =   18
      Tag             =   "Listing of Associations"
      Top             =   630
      Width           =   9435
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "UD_DATE1"
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Tag             =   "41-Date"
      Top             =   4920
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   556
      TextBoxWidth    =   1215
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7080
      Top             =   9120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      TabIndex        =   32
      Top             =   8970
      Width           =   10050
      _Version        =   65536
      _ExtentX        =   17727
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
         Left            =   9090
         Top             =   90
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
         GridSource      =   "vbxTrueGrid"
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.CheckBox chkFlag 
      Alignment       =   1  'Right Justify
      Caption         =   "UFlag 1 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   5520
      TabIndex        =   10
      Top             =   4905
      Width           =   1875
   End
   Begin VB.TextBox txtFlag 
      Appearance      =   0  'Flat
      DataField       =   "UD_FLAG1"
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
      Left            =   7680
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4935
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "UD_LDATE"
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
      Left            =   3120
      MaxLength       =   25
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   8760
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "UD_LTIME"
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
      Left            =   4920
      MaxLength       =   25
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8760
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "UD_LUSER"
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
      Left            =   6600
      MaxLength       =   25
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8760
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   10050
      _Version        =   65536
      _ExtentX        =   17727
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
         Left            =   6840
         TabIndex        =   48
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
         TabIndex        =   26
         Top             =   155
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
         TabIndex        =   25
         Top             =   132
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
         Left            =   3100
         TabIndex        =   24
         Top             =   132
         Width           =   1740
      End
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "UD_CODE3"
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Tag             =   "01-Code"
      Top             =   3720
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "COD3"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "UD_CODE4"
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   3
      Tag             =   "01-Code"
      Top             =   4080
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "COD4"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "UD_CODE5"
      Height          =   285
      Index           =   4
      Left            =   1680
      TabIndex        =   4
      Tag             =   "01-Code"
      Top             =   4440
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "COD5"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "UD_CODE1"
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Tag             =   "01-Code"
      Top             =   3000
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "COD1"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "UD_CODE2"
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Tag             =   "01-Code"
      Top             =   3360
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "COD2"
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "UD_DATE2"
      Height          =   315
      Index           =   1
      Left            =   1680
      TabIndex        =   6
      Tag             =   "40-Date"
      Top             =   5280
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   556
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "UD_DATE3"
      Height          =   315
      Index           =   2
      Left            =   1680
      TabIndex        =   7
      Tag             =   "40-Date"
      Top             =   5640
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   556
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "UD_DATE4"
      Height          =   315
      Index           =   3
      Left            =   1680
      TabIndex        =   8
      Tag             =   "40-Date"
      Top             =   6000
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   556
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "UD_DATE5"
      Height          =   315
      Index           =   4
      Left            =   1680
      TabIndex        =   9
      Tag             =   "40-Date"
      Top             =   6360
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   556
      TextBoxWidth    =   1215
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "UText 2"
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
      Index           =   12
      Left            =   240
      TabIndex        =   47
      Top             =   7215
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "UText 1"
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
      Index           =   31
      Left            =   240
      TabIndex        =   46
      Top             =   6855
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date 5"
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
      TabIndex        =   40
      Top             =   6420
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date 4"
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
      Left            =   240
      TabIndex        =   39
      Top             =   6060
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date 3"
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
      Left            =   240
      TabIndex        =   38
      Top             =   5700
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Code 5"
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
      TabIndex        =   37
      Top             =   4485
      Width           =   1350
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Code 4"
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
      Left            =   240
      TabIndex        =   36
      Top             =   4125
      Width           =   1350
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Code 3"
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
      Left            =   240
      TabIndex        =   35
      Top             =   3765
      Width           =   1350
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Code 2"
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
      Left            =   240
      TabIndex        =   34
      Top             =   3405
      Width           =   1350
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "UComments"
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
      Left            =   240
      TabIndex        =   33
      Top             =   7560
      Width           =   1305
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date 2"
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
      Left            =   240
      TabIndex        =   31
      Top             =   5340
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date 1"
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
      Left            =   240
      TabIndex        =   30
      Top             =   4980
      Width           =   1140
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Code 1"
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
      Left            =   240
      TabIndex        =   29
      Top             =   3045
      Width           =   1350
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "UD_EMPNBR"
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
      Left            =   2160
      TabIndex        =   27
      Top             =   8880
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "UD_COMPNO"
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
      Left            =   480
      TabIndex        =   28
      Top             =   8880
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEUserDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fGLBNew As Boolean
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim Ctrl As Control 'Sam add July 2002 * Remove ADO

Private Function chkEUserDefine()
Dim oCode As String, OCodeD As String
Dim x As Integer

chkEUserDefine = False

On Error GoTo chkEASSOC_Err

'If Len(clpCode(1).Text) < 1 Then
'    MsgBox "Association code is a required field"
'    clpCode(1).SetFocus
'    Exit Function
'End If

For x = 0 To 4
    If clpCode(x).Caption = "Unassigned" Then
        Select Case x
            Case 0:
                MsgBox lblTitle(1).Caption & " code must be valid"
            Case 0:
                MsgBox lblTitle(2).Caption & " code must be valid"
            Case 0:
                MsgBox lblTitle(5).Caption & " code must be valid"
            Case 0:
                MsgBox lblTitle(6).Caption & " code must be valid"
            Case 0:
                MsgBox lblTitle(7).Caption & " code must be valid"
        End Select
        
        clpCode(x).SetFocus
        Exit Function
    End If
Next x

For x = 0 To 4
    If chkFlag(x).Value = 1 Then
        txtFlag(x).Text = "1"
    Else
        txtFlag(x).Text = "0"
    End If
Next x

For x = 0 To 4
    If Len(dlpDate(x).Text) > 0 Then
        If Not IsDate(dlpDate(x).Text) Then
            Select Case x
                Case 0:
                    MsgBox lblTitle(3).Caption & " is not a valid date."
                Case 1:
                    MsgBox lblTitle(4).Caption & " is not a valid date."
                Case 2:
                    MsgBox lblTitle(8).Caption & " is not a valid date."
                Case 3:
                    MsgBox lblTitle(9).Caption & " is not a valid date."
                Case 4:
                    MsgBox lblTitle(11).Caption & " is not a valid date."
            End Select
            dlpDate(x).SetFocus
            Exit Function
        End If
    End If
Next x

chkEUserDefine = True

Exit Function

chkEASSOC_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEUserDefine", "HR_USERDEFINE_TABLE", "edit/Add")
Call RollBack '23July99 js

End Function

Sub cmdCancel_Click()
Dim x
On Error GoTo Can_Err

rsDATA.CancelUpdate
Call Display_Value


fGLBNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(True)  ' reset screen's attributes

fGLBNew = False
Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Cancel", "HR_USERDEFINE_TABLE", "Cancel")
Call RollBack '23July99 js

End Sub

Sub cmdClose_Click()
'Call NextForm
Unload Me
If glbOnTop = "FRMEUSERDEF" Then glbOnTop = ""

End Sub

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

fGLBNew = False

Call SET_UP_MODE
'Call ST_UPD_MODE(True)

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDelete", "HR_USERDEFINE_TABLE", "Delete")
Call RollBack '23July99 js

End Sub

Sub cmdNew_Click()
Dim SQLQ As String

On Error GoTo AddN_Err

fGLBNew = True
Call SET_UP_MODE

Call Set_Control("B", Me)
rsDATA.AddNew

If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err


Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_USERDEFINE_TABLE", "Add")
Call RollBack '23July99 js

End Sub

Sub cmdOK_Click()
Dim x
On Error GoTo Add_Err

If Not chkEUserDefine() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)

If glbCompSerial = "S/N - 2188W" Then  'City of Chatham Kent
    txtCode1TabName.Text = "ESCD"
Else
    txtCode1TabName.Text = "COD1"
End If

Call Set_Control("U", Me, rsDATA)
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


'Call ST_UPD_MODE(True)

fGLBNew = False
Call SET_UP_MODE
Me.vbxTrueGrid.SetFocus
'If NextFormIF("Association") Then
'    Call cmdNew_Click
'End If
Exit Sub

Add_Err:
If Err = 3022 Then
    'Data1.UpdateControls  ' no dups
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_USERDEFINE_TABLE", "Update")
Call RollBack '23July99 js

End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s " & lStr("User Defined Table")
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

RHeading = lblEEName & "'s " & lStr("User Defined Table")
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub

Function EERetrieve()
Dim SQLQ As String

Screen.MousePointer = HOURGLASS

EERetrieve = False

On Error GoTo EERError

If glbtermopen Then         'Lucy July 5, 2000
    SQLQ = "Select * from Term_USERDEFINE_TABLE"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY UD_CODE1"
Else
    SQLQ = "Select * from HR_USERDEFINE_TABLE"
    SQLQ = SQLQ & " where UD_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY UD_CODE1"
End If

Data1.RecordSource = SQLQ
Data1.Refresh
EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "User Define", "HR_USERDEFINE_TABLE", "SELECT")
Call RollBack '23July99 js

Exit Function
End Function

Private Sub chkFlag_Click(Index As Integer)
    If chkFlag(Index).Value = 1 Then
        txtFlag(Index).Text = "1"
    Else
        txtFlag(Index).Text = "0"
    End If
End Sub

Private Sub chkFlag_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
    glbOnTop = "FRMEUSERDEF"
    Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEUSERDEF"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "FRMEUSERDEF"
If glbtermopen Then         'Lucy July 5, 2000
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

Call setCaption(lblTitle(1))
Call setCaption(lblTitle(2))
Call setCaption(lblTitle(5))
Call setCaption(lblTitle(6))
Call setCaption(lblTitle(7))

Call setCaption(lblTitle(3))
Call setCaption(lblTitle(4))
Call setCaption(lblTitle(8))
Call setCaption(lblTitle(9))
Call setCaption(lblTitle(11))

Call setCaption(lblTitle(10))

Call setCaption(chkFlag(0))
chkFlag(1).Caption = lStr(Trim(chkFlag(1).Caption))
chkFlag(2).Caption = lStr(Trim(chkFlag(2).Caption))
chkFlag(3).Caption = lStr(Trim(chkFlag(3).Caption))
chkFlag(4).Caption = lStr(Trim(chkFlag(4).Caption))

Call setCaption(lblTitle(31))
Call setCaption(lblTitle(12))

If glbCompSerial = "S/N - 2188W" Then  'City of Chatham Kent
    clpCode(0).TablName = "ESCD"    'Code 1 Table Name is Course Code for Chatham Kent
    txtCode1TabName.Text = "ESCD"
Else
    txtCode1TabName.Text = "COD1"
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
Me.vbxTrueGrid.SetFocus

lblEENum.Caption = ShowEmpnbr(lblEEID)

Call Display_Value
Call ST_UPD_MODE(True)             '
Call INI_Controls(Me)

vbxTrueGrid.Columns(0).Caption = lStr(vbxTrueGrid.Columns(0).Caption)
vbxTrueGrid.Columns(1).Caption = lStr(vbxTrueGrid.Columns(1).Caption)
vbxTrueGrid.Columns(2).Caption = lStr(vbxTrueGrid.Columns(2).Caption)
vbxTrueGrid.Columns(3).Caption = lStr(vbxTrueGrid.Columns(3).Caption)
vbxTrueGrid.Columns(4).Caption = lStr(vbxTrueGrid.Columns(4).Caption)
vbxTrueGrid.Columns(5).Caption = lStr(vbxTrueGrid.Columns(5).Caption)
vbxTrueGrid.Columns(6).Caption = lStr(vbxTrueGrid.Columns(6).Caption)
vbxTrueGrid.Columns(7).Caption = lStr(vbxTrueGrid.Columns(7).Caption)
vbxTrueGrid.Columns(8).Caption = lStr(vbxTrueGrid.Columns(8).Caption)
vbxTrueGrid.Columns(9).Caption = lStr(vbxTrueGrid.Columns(9).Caption)
vbxTrueGrid.Columns(10).Caption = lStr(vbxTrueGrid.Columns(10).Caption)
vbxTrueGrid.Columns(11).Caption = lStr(vbxTrueGrid.Columns(11).Caption)
vbxTrueGrid.Columns(12).Caption = lStr(vbxTrueGrid.Columns(12).Caption)
vbxTrueGrid.Columns(13).Caption = lStr(vbxTrueGrid.Columns(13).Caption)
vbxTrueGrid.Columns(14).Caption = lStr(vbxTrueGrid.Columns(14).Caption)
vbxTrueGrid.Columns(15).Caption = lStr(vbxTrueGrid.Columns(15).Caption)
vbxTrueGrid.Columns(16).Caption = lStr(vbxTrueGrid.Columns(16).Caption)


MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

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
'    Call NextForm
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

chkFlag(0).Enabled = TF
chkFlag(1).Enabled = TF
chkFlag(2).Enabled = TF
chkFlag(3).Enabled = TF
chkFlag(4).Enabled = TF

clpCode(0).Enabled = TF
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
clpCode(4).Enabled = TF

dlpDate(0).Enabled = TF
dlpDate(1).Enabled = TF
dlpDate(2).Enabled = TF
dlpDate(3).Enabled = TF
dlpDate(4).Enabled = TF

txtText(0).Enabled = TF
txtText(1).Enabled = TF

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then

End If

End Sub

Private Sub txtFlag_Change(Index As Integer)
    If txtFlag(Index) = "-1" Or txtFlag(Index) = "1" Then
        chkFlag(Index) = 1
    Else
        chkFlag(Index) = 0
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
    
    If glbtermopen Then         'Lucy July 5, 2000
        SQLQ = "Select * from Term_USERDEFINE_TABLE"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = "Select * from HR_USERDEFINE_TABLE"
        SQLQ = SQLQ & " where UD_EMPNBR = " & glbLEE_ID
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
       
    Data1.RecordSource = SQLQ
    Data1.Refresh

End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$
Dim SQLQ As String

On Error GoTo Tab1_Err

Call Display_Value

Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_USERDEFINE_TABLE", "Add")
Call RollBack '23July99 js

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
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    If glbtermopen Then
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    Call SET_UP_MODE
    Exit Sub
End If
      
If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    
    If glbtermopen Then
    SQLQ = "Select * from Term_USERDEFINE_TABLE"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & Data1.Recordset!TERM_SEQ
    SQLQ = SQLQ & " ORDER BY UD_CODE1"
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "Select * from HR_USERDEFINE_TABLE"
    SQLQ = SQLQ & " where UD_ID = " & Data1.Recordset!UD_ID
    SQLQ = SQLQ & " ORDER BY UD_CODE1"
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
Call Set_Control("R", Me, rsDATA)
Call SET_UP_MODE

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
UpdateRight = gSec_Upd_UserDefineTbl
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
Call ST_UPD_MODE(TF)

End Sub

Private Sub lblEEID_Change()

    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
        frmEUserDef.Caption = lStr("User Defined Table") & " - " & Left$(glbLEE_SName, 5)
        frmEUserDef.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID

    lblEENum = ShowEmpnbr(lblEEID)
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = glbLEE_ProdLine
    Else
        lblEEProdLine = ""
    End If
End Sub

