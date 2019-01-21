VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEHSCorrective 
   AutoRedraw      =   -1  'True
   Caption         =   "Corrective Actions Data"
   ClientHeight    =   8985
   ClientLeft      =   -135
   ClientTop       =   600
   ClientWidth     =   11400
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtAssigned 
      Appearance      =   0  'Flat
      DataField       =   "CR_ASSIGNED"
      Height          =   285
      Left            =   3135
      TabIndex        =   31
      Tag             =   "00-Employee Number of individual's supervisor"
      Top             =   4020
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAssignedToName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4290
      TabIndex        =   56
      Tag             =   "00-Employee Name of individual's supervisor"
      Top             =   4070
      Visible         =   0   'False
      Width           =   3890
   End
   Begin VB.TextBox txtAssignedSName 
      Appearance      =   0  'Flat
      DataField       =   "CR_ASSIGNED_SURNAME"
      Height          =   285
      Index           =   0
      Left            =   8520
      TabIndex        =   49
      Tag             =   "00-Employee Surname Name of individual's supervisor"
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAssignedFName 
      Appearance      =   0  'Flat
      DataField       =   "CR_ASSIGNED_FNAME"
      Height          =   285
      Index           =   0
      Left            =   8280
      TabIndex        =   48
      Tag             =   "00-Employee First Name of individual's supervisor"
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtProDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2900
      MaxLength       =   50
      TabIndex        =   4
      Tag             =   "00-Describe Specific Corrective Action"
      Top             =   7440
      Width           =   5610
   End
   Begin VB.Frame frExtra 
      Height          =   1095
      Left            =   5160
      TabIndex        =   40
      Top             =   4560
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtInvMemName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   4290
         TabIndex        =   59
         Tag             =   "00-Employee Name of individual's supervisor"
         Top             =   800
         Visible         =   0   'False
         Width           =   3890
      End
      Begin VB.TextBox txtInvMemName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   4290
         TabIndex        =   58
         Tag             =   "00-Employee Name of individual's supervisor"
         Top             =   470
         Visible         =   0   'False
         Width           =   3890
      End
      Begin VB.TextBox txtInvMemName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   4290
         TabIndex        =   57
         Tag             =   "00-Employee Name of individual's supervisor"
         Top             =   140
         Visible         =   0   'False
         Width           =   3890
      End
      Begin VB.TextBox txtInvMemFName 
         Appearance      =   0  'Flat
         DataField       =   "CR_VERIFIED1_FNAME"
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   55
         Tag             =   "00-Employee First Name of individual's supervisor"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMemFName 
         Appearance      =   0  'Flat
         DataField       =   "CR_VERIFIED2_FNAME"
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   54
         Tag             =   "00-Employee First Name of individual's supervisor"
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMemFName 
         Appearance      =   0  'Flat
         DataField       =   "CR_VERIFIED3_FNAME"
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   53
         Tag             =   "00-Employee First Name of individual's supervisor"
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMemSName 
         Appearance      =   0  'Flat
         DataField       =   "CR_VERIFIED1_SURNAME"
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   52
         Tag             =   "00-Employee Surname Name of individual's supervisor"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMemSName 
         Appearance      =   0  'Flat
         DataField       =   "CR_VERIFIED2_SURNAME"
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   51
         Tag             =   "00-Employee Surname Name of individual's supervisor"
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMemSName 
         Appearance      =   0  'Flat
         DataField       =   "CR_VERIFIED3_SURNAME"
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   50
         Tag             =   "00-Employee Surname Name of individual's supervisor"
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMem 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   7680
         TabIndex        =   43
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMem 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   7680
         TabIndex        =   42
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtInvMem 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   7680
         TabIndex        =   41
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin INFOHR_Controls.EmployeeLookup elpInvMem 
         Height          =   285
         Index           =   0
         Left            =   2580
         TabIndex        =   10
         Tag             =   "00-Employee Number"
         Top             =   90
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.EmployeeLookup elpInvMem 
         Height          =   285
         Index           =   1
         Left            =   2580
         TabIndex        =   11
         Tag             =   "00-Employee Number"
         Top             =   420
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.EmployeeLookup elpInvMem 
         Height          =   285
         Index           =   2
         Left            =   2580
         TabIndex        =   12
         Tag             =   "00-Employee Number"
         Top             =   750
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Verified By 1"
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   46
         Top             =   90
         Width           =   915
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Verified By 2"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   45
         Top             =   420
         Width           =   915
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Verified By 3"
         Height          =   195
         Index           =   10
         Left            =   360
         TabIndex        =   44
         Top             =   750
         Width           =   915
      End
   End
   Begin VB.Frame frComment 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      TabIndex        =   33
      Top             =   5280
      Width           =   9495
      Begin VB.TextBox txtComments 
         Appearance      =   0  'Flat
         DataField       =   "CR_Comments"
         Height          =   1005
         Left            =   2880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Tag             =   "00-Comments"
         Top             =   60
         Width           =   5895
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   39
         Top             =   0
         Width           =   1785
      End
      Begin VB.Label lblUpdateDate 
         Caption         =   "Updated Date"
         Height          =   255
         Left            =   6240
         TabIndex        =   38
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblUpdDateDesc 
         Height          =   255
         Left            =   7440
         TabIndex        =   37
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblUpdateBy 
         Caption         =   "Updated By"
         Height          =   255
         Left            =   2880
         TabIndex        =   36
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblUserDesc 
         Height          =   255
         Left            =   3840
         TabIndex        =   35
         Top             =   1200
         Width           =   2295
      End
   End
   Begin INFOHR_Controls.DateLookup dlpTarget 
      DataField       =   "CR_TARGETDATE"
      Height          =   315
      Left            =   2580
      TabIndex        =   6
      Tag             =   "Target Date"
      Top             =   4350
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.EmployeeLookup elpAssigned 
      Height          =   285
      Left            =   2580
      TabIndex        =   5
      Tag             =   "Assigned To"
      Top             =   4020
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   503
      RefreshDescriptionWhen=   2
   End
   Begin VB.CheckBox chkCompleted 
      Caption         =   "Check1"
      DataField       =   "CR_COMPLETED"
      Height          =   195
      Left            =   2910
      TabIndex        =   7
      Tag             =   "Completed"
      Top             =   4710
      Width           =   285
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "CR_TDATE"
      Height          =   285
      Index           =   0
      Left            =   2580
      TabIndex        =   1
      Tag             =   "41-Date  occurred"
      Top             =   2970
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "Fehscorr.frx":0000
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "Fehscorr.frx":0014
      TabIndex        =   0
      Top             =   600
      Width           =   8655
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "CR_Code"
      Height          =   285
      Left            =   2580
      TabIndex        =   3
      Tag             =   "01-Corrective Actions Code"
      Top             =   3690
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECCR"
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   11040
      Top             =   8280
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
      Caption         =   "Ado3"
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   11040
      Top             =   8040
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
      Caption         =   "Ado1"
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
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      DataField       =   "CR_Case"
      Height          =   285
      Left            =   4440
      MaxLength       =   8
      TabIndex        =   13
      Tag             =   "11- incident Number"
      Top             =   3330
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox comShift 
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2900
      TabIndex        =   2
      Tag             =   "01-Incident Number"
      Top             =   3315
      Width           =   1575
   End
   Begin VB.TextBox Updstats 
      DataField       =   "CR_LDate"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   6030
      MaxLength       =   25
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   8730
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      DataField       =   "CR_LTime"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   7830
      MaxLength       =   25
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8730
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      DataField       =   "CR_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   9510
      MaxLength       =   25
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   8730
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11400
      _Version        =   65536
      _ExtentX        =   20108
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
         TabIndex        =   32
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label lblEENumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   160
         Width           =   1005
      End
      Begin VB.Label lblEENum 
         AutoSize        =   -1  'True
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
         TabIndex        =   19
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         AutoSize        =   -1  'True
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
         TabIndex        =   18
         Top             =   135
         Width           =   720
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   9960
      Top             =   7920
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
   Begin INFOHR_Controls.DateLookup dlpCompleted 
      DataField       =   "CR_COMPLETE_DATE"
      Height          =   315
      Left            =   7260
      TabIndex        =   8
      Tag             =   "Completed Date"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      DataField       =   "CR_DURATION"
      Height          =   285
      Left            =   2580
      TabIndex        =   9
      Tag             =   "01-Duration Code"
      Top             =   4980
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECDU"
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Describe Specific Action"
      Height          =   195
      Index           =   7
      Left            =   360
      TabIndex        =   47
      Top             =   7485
      Width           =   2595
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Duration"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   30
      Top             =   4995
      Width           =   1965
   End
   Begin VB.Label lblCompleted 
      BackStyle       =   0  'Transparent
      Caption         =   "Completed Date"
      Height          =   315
      Left            =   5340
      TabIndex        =   29
      Top             =   4710
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Target Date"
      Height          =   255
      Left            =   330
      TabIndex        =   28
      Top             =   4380
      Width           =   2055
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Assigned To"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   27
      Top             =   4035
      Width           =   2085
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Completed"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   26
      Top             =   4710
      Width           =   2535
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Corrective Action Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   25
      Top             =   3720
      Width           =   2355
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Incident Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   24
      Top             =   3390
      Width           =   1545
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Corrective Action Date "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   23
      Top             =   3030
      Width           =   2355
   End
   Begin VB.Label lblEEID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "CR_Empnbr"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5070
      TabIndex        =   21
      Top             =   8850
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "CR_CompNo"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3930
      TabIndex        =   22
      Top             =   8850
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmEHSCorrective"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim oldTarget
Dim oldAssigned

Function chkHSCorrective()

Dim SQLQ As String, Msg As String, dd#

chkHSCorrective = False

On Error GoTo chkHSCorrective_Err
If Len(dlpDate(0).Text) >= 1 Then
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "Corrective Date Not Valid."
        dlpDate(0).SetFocus
        Exit Function
    End If
Else
    MsgBox "Corrective Date is required."
    dlpDate(0).SetFocus
    Exit Function
End If

Dim tTime As Variant
Dim Part1$, Part2$

'~~

If Len(txtShift) < 1 Then
    MsgBox "Incident Number is a required field"
    comShift.SetFocus
    Exit Function
End If
If Not IfIncidentNo(Val(txtShift)) Then
    MsgBox "Incident Number Not Valid"
    comShift.SetFocus
    Exit Function
End If

If Len(clpCode.Text) < 1 Then
    MsgBox "Corrective Actions Code is a required field"
    clpCode.SetFocus
    Exit Function
End If

If clpCode.Caption = "Unassigned" Then
    MsgBox "Corrective Actions code must be valid"
    clpCode.SetFocus
    Exit Function
End If

If Len(elpAssigned.Text) > 0 And elpAssigned.Caption = "Unassigned" Then  'Ticket #23915 Franks 06/132013
    elpAssigned.SetFocus
    MsgBox "Invalid Assigned To"
    Exit Function
End If

chkHSCorrective = True

Exit Function

chkHSCorrective_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HR_OHS_CORRECTIVE", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function


Sub cmdCancel_Click()
Dim X
On Error GoTo Can_Err

fglbNew = False
Call Display_Value

'Call ST_UPD_MODE(True)  ' reset screen's attributes
'Call SET_UP_MODE
'Me.vbxTrueGrid.SetFocus
Call txtAssigned_Change

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_OHS_CORRECTIVE", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If



End Sub

'Sub cmdCancel_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEHSCORRECTIVE" Then glbOnTop = ""

End Sub

'Sub cmdClose_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdContact_Click()
'frmEHSContact.Show
'Unload Me
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, INo&, X

If Not gSec_Upd_HSCorrectiveAct Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If


On Error GoTo Del_Err


Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Record?  "

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

Call updFollow("D")

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

Me.vbxTrueGrid.SetFocus
fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_OHS_CORRECTIVE", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub



Sub cmdNew_Click()
Dim SQLQ As String

If Not gSec_Upd_HSCorrectiveAct Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

fglbNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
On Error GoTo AddN_Err
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    Me.vbxTrueGrid.Enabled = False
End If
Me.vbxTrueGrid.Enabled = False
'data1.Recordset.AddNew
''' Sam add July 2002 * Remove Binding Control
Call Set_Control("B", Me)

oldTarget = ""
oldAssigned = ""
If elpAssigned.Text = "" Then
    elpAssigned.Caption = ""
    txtAssignedToName.Text = ""
End If

If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"


dlpDate(0).SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_OHS_CORRECTIVE", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Sub CmdNew_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim X
On Error GoTo Add_Err

If Not chkHSCorrective() Then Exit Sub
Call updFollow("U")

Call UpdUStats(Me) ' update user's stats (who did it and when)


rsDATA.Requery
If fglbNew Then
    rsDATA.AddNew
    rsDATA("CR_CODE_TABL") = "ECCR"
End If

If IsNumeric(txtAssigned.Text) And elpAssigned.Caption <> "Unassigned" Then
    txtAssignedFName(0).Text = GetEmpData(txtAssigned.Text, "ED_FNAME")
    txtAssignedSName(0).Text = GetEmpData(txtAssigned.Text, "ED_SURNAME")
End If

If IsNumeric(txtInvMem(0).Text) And elpInvMem(0).Caption <> "Unassigned" Then
    txtInvMemFName(0).Text = GetEmpData(txtInvMem(0).Text, "ED_FNAME")
    txtInvMemSName(0).Text = GetEmpData(txtInvMem(0).Text, "ED_SURNAME")
End If
If IsNumeric(txtInvMem(1).Text) And elpInvMem(1).Caption <> "Unassigned" Then
    txtInvMemFName(1).Text = GetEmpData(txtInvMem(1).Text, "ED_FNAME")
    txtInvMemSName(1).Text = GetEmpData(txtInvMem(1).Text, "ED_SURNAME")
End If
If IsNumeric(txtInvMem(2).Text) And elpInvMem(2).Caption <> "Unassigned" Then
    txtInvMemFName(2).Text = GetEmpData(txtInvMem(2).Text, "ED_FNAME")
    txtInvMemSName(2).Text = GetEmpData(txtInvMem(2).Text, "ED_SURNAME")
End If

Call Set_Control("U", Me, rsDATA)
If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
Else
    gdbAdoIhr001.BeginTrans
    Call Set_Control("U", Me, rsDATA)
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
End If
Data1.Refresh

fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE

Me.vbxTrueGrid.Enabled = True

'Me.vbxTrueGrid.SetFocus
If NextFormIF("Corrective Action") Then
    Call cmdNew_Click
End If
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OHS_CORRECTIVE", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

'Sub cmdOK_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Corrective Actions"
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

RHeading = lblEEName & "'s Corrective Actions"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub

Private Sub chkCompleted_Click()
    If chkCompleted Then
        lblCompleted.Visible = True
        dlpCompleted.Visible = True
    Else
        lblCompleted.Visible = False
        dlpCompleted.Visible = False
    End If
End Sub



'Sub cmdPrint_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdTCause_Click()
'frmEHSCause.Show
'Unload Me
'End Sub

'Private Sub cmdTCause_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Sub cmdWCBMed_Click()
'frmEHSWCB.Show
'Unload Me
'End Sub

'Private Sub cmdWCBMed_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdWSIB_Click()
'frmEHSWCBC.Show
'Unload Me
'End Sub

Sub comShift_Change()
'txtShift = comShift  'JDY
End Sub

Sub comShift_Click()
'txtShift = comShift      'JDY
End Sub

Sub comShift_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub




Private Sub comShift_LostFocus()
txtShift = comShift  'JDY
End Sub

Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HR_OHS_CORRECTIVE", "SELECT")


End Sub




Function EERetrieve()
Dim SQLQ As String
EERetrieve = False

Screen.MousePointer = HOURGLASS
On Error GoTo EERError


If glbtermopen Then         'Lucy July 5, 2000
    SQLQ = "Select * from Term_OHS_CORRECTIVE"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY CR_Case DESC"
Else
    SQLQ = "Select * from HR_OHS_CORRECTIVE"
    SQLQ = SQLQ & " where CR_Empnbr = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY CR_Case DESC"
End If

Data1.RecordSource = SQLQ
Data1.Refresh


If glbtermopen Then     'Lucy July 5, 2000
    SQLQ = "SELECT EC_EMPNBR, EC_CASE, EC_OCCDATE from Term_HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
Else
    SQLQ = "SELECT EC_EMPNBR, EC_CASE, EC_OCCDATE from HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " WHERE EC_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
End If

Data3.RecordSource = SQLQ
Data3.Refresh

EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function


EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "OCH Retrieve", "HR_OHS_CORRECTIVE", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
Exit Function
End Function

Private Sub elpAssigned_Change()
    'txtAssigned.Text = getEmpnbr(elpAssigned.Text)
    'If elpAssigned.Caption = "Unassigned" Then
    '    If Not Data1.Recordset.EOF Then
    '        txtAssignedToName.Text = Data1.Recordset!CR_ASSIGNED_SURNAME & ", " & Data1.Recordset!CR_ASSIGNED_FNAME
    '        txtAssignedToName.Visible = True
    '    Else
    '        txtAssignedToName.Visible = False
    '    End If
    'Else
    '    txtAssignedToName.Visible = False
    'End If
    txtAssigned.Text = getEmpnbr(elpAssigned.Text)
    If elpAssigned.Caption = "Unassigned" And Len(elpAssigned.Text) > 0 Then 'Ticket #23915 Franks 06/13/2013
        If Not Data1.Recordset.EOF Then
            'txtAssignedToName.Text = Data1.Recordset!CR_ASSIGNED_SURNAME & ", " & Data1.Recordset!CR_ASSIGNED_FNAME
            txtAssignedToName.Text = "Unassigned" 'Ticket #23915 Franks 06/13/2013
            txtAssignedToName.Visible = True
        Else
            txtAssignedToName.Visible = False
        End If
    Else
        txtAssignedToName.Visible = False
    End If
End Sub

Private Sub elpInvMem_Change(Index As Integer)
    txtInvMem(Index).Text = getEmpnbr(elpInvMem(Index).Text)
    If elpInvMem(Index).Caption = "Unassigned" Then
        If Not Data1.Recordset.EOF Then
            Select Case Index
                Case 0: txtInvMemName(Index).Text = Data1.Recordset!CR_VERIFIED1_SURNAME & ", " & Data1.Recordset!CR_VERIFIED1_FNAME
                Case 1: txtInvMemName(Index).Text = Data1.Recordset!CR_VERIFIED2_SURNAME & ", " & Data1.Recordset!CR_VERIFIED2_FNAME
                Case 2: txtInvMemName(Index).Text = Data1.Recordset!CR_VERIFIED3_SURNAME & ", " & Data1.Recordset!CR_VERIFIED3_FNAME
            End Select
            
            If Trim(txtInvMemName(Index)) = "," Then
                txtInvMemName(Index).Visible = False
            Else
                txtInvMemName(Index).Visible = True
            End If
        Else
            txtInvMemName(Index).Visible = False
        End If
    Else
        txtInvMemName(Index).Visible = False
    End If
End Sub

Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMEHSCORRECTIVE"

End Sub

Sub Form_GotFocus()
glbOnTop = "FRMEHSCORRECTIVE"
End Sub

Sub Form_Load()
Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found
Dim X%
Dim SQLQ1
glbOnTop = "FRMEHSCORRECTIVE"

If glbtermopen Then         'Lucy July 5, 2000
    Data1.ConnectionString = glbAdoIHRAUDIT
    Data3.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
    Data3.ConnectionString = glbAdoIHRDB
End If


Screen.MousePointer = HOURGLASS

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
    lblTitle(1).Caption = "Verified (in place and effective)"
    txtAssigned.DataField = "CR_ASSIGNED"
    txtProDesc.DataField = "CR_DESCCORR"
    txtInvMem(0).DataField = "CR_VERIFIEDBY1"
    txtInvMem(1).DataField = "CR_VERIFIEDBY2"
    txtInvMem(2).DataField = "CR_VERIFIEDBY3"
    'Hemu
    lblTitle(7).Top = 4045
    txtProDesc.Top = 4020
    lblTitle(3).Top = 4390
    elpAssigned.Top = 4360
    txtAssigned.Top = 4360
    txtAssignedToName.Top = 4410
    Label1.Top = 4740
    dlpTarget.Top = 4700
    lblTitle(1).Top = 5110
    chkCompleted.Top = 5110
    lblTitle(6).Top = 5490
    clpCode2.Top = 5440
    'frComment.Top = 5390
    'Hemu
    
    frComment.Top = 6880
    frExtra.Left = 0
    frExtra.Top = 5750 '5300 '7080
    frExtra.Width = 8775
    frExtra.BorderStyle = 0
    frExtra.Visible = True
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

comShift.Clear
Do Until Data3.Recordset.EOF                  'JDY
  comShift.AddItem Data3.Recordset("EC_CASE") 'JDY
  Data3.Recordset.MoveNext                    'JDY
Loop

If glbLinHS Then
    If Len(glbDivDesc) > 0 Then   ' dont do on add new until in
        Me.Caption = "Corrective Actions Data - " & glbDivDesc
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv
    lblEENumber.Caption = lStr("Division")
Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
        Me.Caption = "Corrective Actions Data - " & Left$(glbLEE_SName, 8)
        Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    lblEENum.Caption = ShowEmpnbr(lblEEID)
End If

Call ST_UPD_MODE(False)

Call Display_Value

If Not gSec_Upd_HSCorrectiveAct Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
End If

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False


End Sub

Sub Form_LostFocus()
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

Sub Form_Unload(Cancel As Integer)

MDIMain.panHelp(0).Caption = "Select function from the menu."
Set frmEHSCorrective = Nothing 'carmen 18 may 00
Call NextForm
End Sub

Function IfIncidentNo(InciNo As Double)
  IfIncidentNo = False
  If Data3.Recordset.BOF And Data3.Recordset.EOF Then
     Exit Function
  End If
  Data3.Recordset.MoveFirst
  Data3.Recordset.Find "EC_Case=" & InciNo
  If Data3.Recordset.EOF Then Exit Function
  IfIncidentNo = True


End Function

Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

glbOHSEdit% = TF

dlpDate(0).Enabled = TF
txtComments.Enabled = TF
comShift.Enabled = TF
clpCode.Enabled = TF
txtProDesc.Enabled = TF
elpAssigned.Enabled = TF
dlpTarget.Enabled = TF
chkCompleted.Enabled = TF
clpCode2.Enabled = TF
txtComments.Enabled = TF
elpInvMem(0).Enabled = TF
elpInvMem(1).Enabled = TF
elpInvMem(2).Enabled = TF

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
'cmdContact.Enabled = FT
'cmdInjLoc.Enabled = FT
'cmdWSIB.Enabled = FT
'vbxTrueGrid.Enabled = FT
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
Else
'    cmdModify.Enabled = True
End If
End Sub

Private Sub txtAssigned_Change()
    elpAssigned = ShowEmpnbr(txtAssigned.Text)
    If IsNumeric(txtAssigned.Text) And elpAssigned.Caption <> "Unassigned" Then
        txtAssignedFName(0).Text = GetEmpData(txtAssigned.Text, "ED_FNAME")
        txtAssignedSName(0).Text = GetEmpData(txtAssigned.Text, "ED_SURNAME")
    Else
        'Ticket #23915 Franks 06/13/2013 - begin
        If Len(txtAssigned.Text) = 0 Then
            elpAssigned.Text = ""
            elpAssigned.Caption = ""
        End If
        txtAssignedToName.Text = elpAssigned.Caption
        'Ticket #23915 Franks 06/13/2013 - end
    End If
End Sub

'Sub clpCode_DblClick()  '(Index As Integer)
'Dim oCode As String, OCodeD As String
'oCode = clpCode
'OCodeD = clpCode(1)
'Call Get_Code(CodeCodes(1, 1), CodeCodes(1, 2))
'If glbCodeRef Then Call ReCreatSnap(1)
'If Len(glbCode) < 1 Then
'    clpCode.Text = oCode
'    clpCode(1).Caption = OCodeD
'Else
'    clpCode.Text = glbCode
'    clpCode(1).Caption = glbCodeDesc
'    clpCode(1).Visible = True
'End If
'End Sub
'Sub clpCode_GotFocus()  '(Index As Integer)
'Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub clpCode_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
'End Sub
'Sub clpCode_LostFocus()  'Index As Integer)
'      '(Index) ' set description for code
'End Sub

Private Sub txtComments_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtInvMem_Change(Index As Integer)
    elpInvMem(Index) = ShowEmpnbr(txtInvMem(Index).Text)
    If IsNumeric(txtInvMem(Index).Text) And elpInvMem(Index).Caption <> "Unassigned" Then
        txtInvMemFName(Index).Text = GetEmpData(txtInvMem(Index).Text, "ED_FNAME")
        txtInvMemSName(Index).Text = GetEmpData(txtInvMem(Index).Text, "ED_SURNAME")
    End If
End Sub

Private Sub txtProDesc_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

'Private Sub txtDate_Change(Index As Integer)
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtDate_DblClick(Index As Integer)
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Sub txtDate_GotFocus(Index As Integer)
'Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub txtDate_KeyPress(Index As Integer, KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub
Sub txtShift_Change()
 
  If Not (Val(txtShift) = 0) Then
    comShift = txtShift
  Else
    comShift = ""
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

Sub vbxTrueGrid_GotFocus()
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
            SQLQ = "Select * from Term_OHS_CORRECTIVE"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select * from HR_OHS_CORRECTIVE"
            SQLQ = SQLQ & " where CR_Empnbr = " & glbLEE_ID
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
'    End If
End If

End Sub


''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
    Dim SQLQ
'    elpAssigned.Caption = ""
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
    
If glbtermopen Then
    SQLQ = "Select * from Term_OHS_CORRECTIVE"
    SQLQ = SQLQ & " WHERE CR_ID = " & Data1.Recordset!CR_ID
    SQLQ = SQLQ & " ORDER BY CR_Case DESC"
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "Select * from HR_OHS_CORRECTIVE"
    SQLQ = SQLQ & " where CR_ID = " & Data1.Recordset!CR_ID
    SQLQ = SQLQ & " ORDER BY CR_Case DESC"
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

End If

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
Call SET_UP_MODE

oldTarget = dlpTarget.Text
oldAssigned = elpAssigned.Text

'Ticket #23915 Franks 06/13/2013 - begin
If Len(txtAssigned.Text) > 0 Then
    'George Jan 10,2005 for an employee who was assigned to the corrective.
    If Not IsNull(rsDATA.Fields("CR_TERM_EMPNAME")) Then  'Ticket #23915 Franks 06/13/2013
        If IsNull(oldAssigned) Or Len(oldAssigned) = 0 Then
            elpAssigned.Caption = rsDATA.Fields("CR_TERM_EMPNAME")
        End If
    Else
        If IsNull(oldAssigned) Or Len(oldAssigned) = 0 Then
            elpAssigned.Caption = "Unassigned"
            txtAssignedToName.Text = elpAssigned.Caption
        End If
    End If
Else
    elpAssigned.Caption = ""
    txtAssignedToName.Text = ""
End If
'Ticket #23915 Franks 06/13/2013 - end


End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value

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
UpdateRight = gSec_Upd_HSCorrectiveAct
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
If glbLinHS Then
    If Len(glbDivDesc) > 0 Then   ' dont do on add new until in
        Me.Caption = "Corrective Actions Data - " & glbDivDesc
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv
    
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = ""
    End If
Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
        frmEHSCorrective.Caption = "Corrective Actions Data - " & Left$(glbLEE_SName, 5)
        frmEHSCorrective.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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


Private Function updFollow(xType)   'Laura on 11/2/97
Dim newline As String
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
Dim rsFollow As New ADODB.Recordset
Dim Edit1 As Boolean
Dim xComments
newline = Chr$(13) & Chr$(10)
updFollow = False

On Error GoTo CrFollow_Err
If IsNumeric(getEmpnbr(oldAssigned)) And IsDate(oldTarget) Then
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & getEmpnbr(oldAssigned)
    SQLQ = SQLQ & " AND EF_FREAS = 'HSCR'"
    SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(oldTarget)
    rsFollow.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Edit1 = True
Else
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = 0" 'only for adding new records
    rsFollow.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Edit1 = False
End If
If xType = "D" Or chkCompleted <> 0 Then
    If Edit1 Then
        Do Until rsFollow.EOF
            xComments = rsFollow("EF_COMMENTS")
            If IsNull(xComments) Then xComments = ""
            If InStr(xComments, glbLEE_ID & "") <> 0 Then
                rsFollow.Delete
                'MsgBox "A Follow Up Record was deleted!"
            End If
            rsFollow.MoveNext
        Loop
    End If
Else
    If fglbNew Or rsFollow.EOF Or Edit1 = False Then
        If IsDate(dlpTarget.Text) And IsNumeric(getEmpnbr(elpAssigned.Text)) Then
            rsFollow.AddNew
            rsFollow("EF_COMPNO") = "001"
            rsFollow("EF_EMPNBR") = getEmpnbr(elpAssigned.Text)
            rsFollow("EF_FDATE") = CVDate(dlpTarget.Text)
            rsFollow("EF_FREAS_TABL") = "FURE"
            'Ticket #24257 - Do not update Admin By for them only
            If glbCompSerial <> "S/N - 2262W" Then
                rsFollow("EF_ADMINBY_TABL") = "EDAB"
                rsFollow("EF_ADMINBY") = GetEmpData(rsFollow("EF_EMPNBR"), "ED_ADMINBY", Null)
            End If
            rsFollow("EF_FREAS") = "HSCR"
            rsFollow("EF_COMMENTS") = "For employee #" & glbLEE_ID & "(" & lblEEName & ")"
            rsFollow("EF_LDATE") = Date
            rsFollow("EF_LTIME") = Time$
            rsFollow("EF_LUSER") = glbUserID
            rsFollow.Update
            'MsgBox "A Follow Up Record was created!"
        End If
    Else
        If oldTarget <> dlpTarget.Text Or getEmpnbr(oldAssigned) <> getEmpnbr(elpAssigned.Text) Then
        
            Do Until rsFollow.EOF
                xComments = rsFollow("EF_COMMENTS")
                If IsNull(xComments) Then xComments = ""
                If InStr(xComments, glbLEE_ID & "") <> 0 Then
                    If IsDate(dlpTarget.Text) And IsNumeric(getEmpnbr(elpAssigned.Text)) Then
    
                        rsFollow("EF_COMPNO") = "001"
                        rsFollow("EF_EMPNBR") = getEmpnbr(elpAssigned.Text)
                        rsFollow("EF_FDATE") = CVDate(dlpTarget.Text)
                        
                        rsFollow("EF_FREAS_TABL") = "FURE"
                        
                        'Ticket #24257 - Do not update Admin By for them only
                        If glbCompSerial <> "S/N - 2262W" Then
                            rsFollow("EF_ADMINBY_TABL") = "EDAB"
                            rsFollow("EF_ADMINBY") = GetEmpData(rsFollow("EF_EMPNBR"), "ED_ADMINBY", Null)
                        End If
                        
                        rsFollow("EF_FREAS") = "HSCR"
                        rsFollow("EF_COMMENTS") = "For employee #" & glbLEE_ID & "(" & lblEEName & ")"
                        rsFollow("EF_LDATE") = Date
                        rsFollow("EF_LTIME") = Time$
                        rsFollow("EF_LUSER") = glbUserID
                        rsFollow.Update
                        'MsgBox "A Follow Up Record was updated!"
                    Else
                        rsFollow.Delete
                        'MsgBox "A Follow Up Record was deleted!"

                    End If
                End If
                rsFollow.MoveNext
            Loop
        End If
    End If
End If

updFollow = True
  
Exit Function

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Function
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Function
