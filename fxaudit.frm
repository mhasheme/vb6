VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmAUDIT 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audit Master File Update"
   ClientHeight    =   8670
   ClientLeft      =   4380
   ClientTop       =   3915
   ClientWidth     =   11115
   DrawMode        =   1  'Blackness
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
   Icon            =   "fxaudit.frx":0000
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8670
   ScaleWidth      =   11115
   Tag             =   "Audit Master File Update"
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkEmploymentDates 
      Height          =   225
      Left            =   5400
      TabIndex        =   19
      Top             =   4800
      Width           =   225
   End
   Begin VB.CheckBox chkExcelRpt 
      Height          =   225
      Left            =   5400
      TabIndex        =   49
      Top             =   6615
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CheckBox chkEmergContact 
      Height          =   225
      Left            =   2520
      TabIndex        =   18
      Top             =   6615
      Width           =   225
   End
   Begin VB.CheckBox chkStatus 
      Height          =   225
      Left            =   2520
      TabIndex        =   17
      Top             =   6255
      Width           =   225
   End
   Begin VB.CheckBox chkTerm 
      Height          =   225
      Left            =   5400
      TabIndex        =   21
      Top             =   5535
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CheckBox chkBank 
      Height          =   225
      Left            =   5400
      TabIndex        =   20
      Top             =   5175
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CheckBox chkNameAdd 
      Height          =   225
      Left            =   2520
      TabIndex        =   16
      Top             =   5895
      Width           =   225
   End
   Begin VB.CheckBox chkBenefits 
      Height          =   225
      Left            =   2520
      TabIndex        =   15
      Top             =   5535
      Width           =   225
   End
   Begin VB.ComboBox comGroup 
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Tag             =   "Final sorting of records - no totals"
      Top             =   7680
      Width           =   2325
   End
   Begin VB.CheckBox chkNewHires 
      Height          =   225
      Left            =   2520
      TabIndex        =   14
      Top             =   5175
      Width           =   225
   End
   Begin VB.CheckBox chkSalary 
      Height          =   225
      Left            =   2520
      TabIndex        =   13
      Top             =   4815
      Width           =   225
   End
   Begin INFOHR_Controls.CodeLookup clpDIV 
      Height          =   285
      Left            =   2190
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.DateLookup dlpTo 
      Height          =   285
      Left            =   2190
      TabIndex        =   9
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3090
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpFrom 
      Height          =   285
      Left            =   2190
      TabIndex        =   8
      Tag             =   "40-Date from and including this date forward"
      Top             =   2760
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   7680
      Top             =   6120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.ComboBox cmbUpload 
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
      Left            =   2505
      TabIndex        =   10
      Tag             =   "Choose Upload flag."
      Text            =   "Combo1"
      Top             =   3480
      Width           =   975
   End
   Begin Threed.SSCheck chkPage 
      Height          =   225
      Left            =   2520
      TabIndex        =   22
      Tag             =   "Page break after Employee changes"
      Top             =   7200
      Width           =   225
      _Version        =   65536
      _ExtentX        =   397
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "Page Break"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   27.01
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Font3D          =   3
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7080
      Top             =   6120
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
   End
   Begin VB.Frame frmAT 
      Height          =   435
      Left            =   210
      TabIndex        =   31
      Top             =   360
      Width           =   5115
      Begin VB.OptionButton optAT 
         Caption         =   "Terminated Employee"
         Height          =   255
         Index           =   1
         Left            =   2490
         TabIndex        =   2
         Top             =   150
         Width           =   2175
      End
      Begin VB.OptionButton optAT 
         Caption         =   "Active Employee"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   150
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   2190
      TabIndex        =   7
      Tag             =   "10-Enter Employee Number"
      Top             =   2430
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpPP 
      DataField       =   "SH_PAYP"
      Height          =   285
      Left            =   2205
      TabIndex        =   11
      Tag             =   "00-Enter pay period code"
      Top             =   3900
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDPP"
   End
   Begin INFOHR_Controls.EmployeeLookup elpUser 
      Height          =   315
      Left            =   2190
      TabIndex        =   12
      Top             =   4320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      ShowDescription =   0   'False
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv1 
      Height          =   285
      Left            =   2190
      TabIndex        =   3
      Tag             =   "00-Specific Division Desired"
      Top             =   1000
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   1
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   2190
      TabIndex        =   40
      Tag             =   "00-Enter Administered By Code"
      Top             =   8115
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   2190
      TabIndex        =   4
      Top             =   1360
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDSE"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   2190
      TabIndex        =   6
      Tag             =   "00-Enter Region Code"
      Top             =   2070
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDRG"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   2190
      TabIndex        =   5
      Tag             =   "00-Enter Location Code"
      Top             =   1725
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDLC"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Dates Only"
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
      Left            =   3120
      TabIndex        =   51
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label lblExcelRpt 
      BackStyle       =   0  'Transparent
      Caption         =   "Excel Friendly Version"
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
      Left            =   3120
      TabIndex        =   50
      Top             =   6600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Emergency Contact Only"
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
      Left            =   300
      TabIndex        =   48
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
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
      TabIndex        =   47
      Top             =   1725
      Width           =   1695
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      TabIndex        =   46
      Top             =   2055
      Width           =   1710
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status Only"
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
      Left            =   300
      TabIndex        =   45
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label lblTerm 
      BackStyle       =   0  'Transparent
      Caption         =   "Leaves and Termination Only"
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
      Left            =   3120
      TabIndex        =   44
      Top             =   5520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblBank 
      BackStyle       =   0  'Transparent
      Caption         =   "Banking Information Only"
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
      Left            =   3120
      TabIndex        =   43
      Top             =   5160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblSection 
      AutoSize        =   -1  'True
      Caption         =   "Section"
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
      Left            =   300
      TabIndex        =   42
      Top             =   1360
      Width           =   1620
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
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
      TabIndex        =   41
      Top             =   8160
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name/Address Only"
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
      Left            =   300
      TabIndex        =   39
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Benefits Only"
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
      Left            =   300
      TabIndex        =   38
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
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
      TabIndex        =   37
      Top             =   1005
      Width           =   1755
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User"
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
      Left            =   300
      TabIndex        =   36
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Final Sort"
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
      Left            =   300
      TabIndex        =   35
      Top             =   7710
      Width           =   660
   End
   Begin VB.Label lblBewHires 
      BackStyle       =   0  'Transparent
      Caption         =   "New Hires Only"
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
      Left            =   300
      TabIndex        =   34
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label lblSalary 
      BackStyle       =   0  'Transparent
      Caption         =   "Salary Changes Only"
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
      Left            =   300
      TabIndex        =   33
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label lblPP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Period"
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
      TabIndex        =   32
      Top             =   3900
      Width           =   930
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Facility"
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
      Index           =   0
      Left            =   240
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblTo 
      Appearance      =   0  'Flat
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
      Left            =   1680
      TabIndex        =   29
      Top             =   3120
      Width           =   240
   End
   Begin VB.Label lblFrom 
      Appearance      =   0  'Flat
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
      Left            =   1650
      TabIndex        =   28
      Top             =   2805
      Width           =   420
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Page Break on Employee"
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
      Left            =   300
      TabIndex        =   27
      Top             =   7200
      Width           =   1800
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Upload Flag"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   300
      TabIndex        =   26
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Label lblFromTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Range"
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
      TabIndex        =   25
      Top             =   2805
      Width           =   870
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number  "
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
      TabIndex        =   24
      Top             =   2460
      Width           =   1380
   End
End
Attribute VB_Name = "frmAUDIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DeletedRecs As Long
Dim xTrainMatrixPath

Const SW_SHOW = 5

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Function chkAudit()
Dim dd As Long

chkAudit = False
On Error GoTo chkEOTHERE_Err

If glbLinamar Then
    If Len(clpDIV) > 0 Then
        If clpDIV.Caption = "Unassigned" Then
            MsgBox "If Facility Entered - they must exist"
            clpDIV.SetFocus
            Exit Function
        End If
    End If
Else
    If Not clpDiv1.ListChecker Then
        Exit Function
    End If
End If

If Len(dlpFrom.Text) > 0 Then
    If Not IsDate(dlpFrom.Text) Then
        MsgBox "Invalid From date"
        dlpFrom.SetFocus
        Exit Function
    End If
End If
If Len(dlpTo.Text) > 0 Then
    If Not IsDate(dlpTo.Text) Then
        MsgBox "Invalid To date"
        dlpTo.SetFocus
        Exit Function
    End If
End If
If Len(dlpFrom.Text) > 0 And Len(dlpTo.Text) > 0 Then
    dd = DateDiff("d", CVDate(dlpFrom.Text), CVDate(dlpTo.Text))
    If dd < 0 Then
        MsgBox "From date must be earlier than To Date"
        dlpFrom.SetFocus
        Exit Function
    End If
End If
If Not elpEEID.ListChecker Then
    Exit Function
End If

chkAudit = True
Exit Function

chkEOTHERE_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkAudit", "HRAUDIT", "Update")
Resume Next

End Function

Private Sub chkPage_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbUpload_GotFocus()
    Call SetPanHelp(ActiveControl)
    MDIMain.panHelp(2).Caption = "Req."
End Sub

Public Sub cmdClose_Click()
    Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub Cri_PP()
    Dim PPCri As String
    
    If Len(clpPP.Text) > 0 Then
      PPCri = "{HR_SALARY_HISTORY.SH_PAYP} in ['" & clpPP.Text & "'] "
      If glbOracle Then
        PPCri = PPCri & "AND {HR_SALARY_HISTORY.SH_CURRENT}<>0 "
      Else
        PPCri = PPCri & "AND {HR_SALARY_HISTORY.SH_CURRENT} "
      End If
      If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
      glbstrSelCri = glbstrSelCri & PPCri
    End If
End Sub

Private Sub Cri_AdminBy()
    Dim AdminByCri As String
    
    If Len(clpCode(1).Text) > 0 Then
      AdminByCri = "{HREMP.ED_ADMINBY} = ['" & clpCode(1).Text & "'] "
      If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
      glbstrSelCri = glbstrSelCri & AdminByCri
    End If
End Sub

Public Sub cmdDelete_Click()
Dim x As Integer
Dim DgDef, Title As String, Msg As String, Response As Integer

If glbLinamar Then
    If Len(clpDIV) = 0 Then
        MsgBox "Facility is a required field"
        clpDIV.SetFocus
        Exit Sub
    End If
End If

Title = "Mass Audit File Delete"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg = "Are You Sure You Want To Delete ALL records for this criteria?"
Response = MsgBox(Msg, DgDef, Title)    ' Get user response.

If Response = IDNO Then    ' Evaluate response
    Exit Sub
End If

Screen.MousePointer = HOURGLASS

x = modDelRecs()

Screen.MousePointer = DEFAULT

If DeletedRecs = 0 Then
    MsgBox "No records found for given selection criteria."
Else
    MsgBox DeletedRecs & " records deleted successfully"
End If

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "Other Earnings", "Delete")
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdPrint_Click()
On Error GoTo PrntErr
Dim x As Integer

Screen.MousePointer = HOURGLASS
If chkAudit() Then
    If Not PrtForm("Audit Master Update Criteria", Me) Then
        Exit Sub
    End If
    ' cmdView.Enabled = False
    ' cmdPrint.Enabled = False
    ' cmdDelete.Enabled = False
     x = Cri_SetAll()
     
    'Ticket #28815 - Jerry said to open for all
    'If glbWFC Then 'Ticket #27605 Franks 10/16/2015
        If chkExcelRpt.Value Then
            Screen.MousePointer = DEFAULT
            Call set_PrintState(True)
            Exit Sub
        End If
    'End If
    
     Me.vbxCrystal.Destination = 1
     MDIMain.Timer1.Enabled = False
     Me.vbxCrystal.Action = 1
     vbxCrystal.Reset
     MDIMain.Timer1.Enabled = True
    '  cmdView.Enabled = True
    '  cmdPrint.Enabled = True
    '  If gSec_Upd_Audit Then cmdDelete.Enabled = True
End If
Screen.MousePointer = DEFAULT

Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdView_Click()
Dim x As Integer

On Error GoTo ViewErr

Screen.MousePointer = HOURGLASS

If chkAudit() Then
    '  cmdView.Enabled = False
    '  cmdPrint.Enabled = False
    '  cmdDelete.Enabled = False
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    x = Cri_SetAll()
    
    'Ticket #28815 - Jerry said to open for all
    'If glbWFC Then 'Ticket #27605 Franks 10/16/2015
        If chkExcelRpt.Value Then
            Screen.MousePointer = DEFAULT
            Call set_PrintState(True)
            Exit Sub
        End If
    'End If
    
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    '  cmdView.Enabled = True
    '  cmdPrint.Enabled = True
    '  If gSec_Upd_Audit Then cmdDelete.Enabled = True
End If

Screen.MousePointer = DEFAULT

Exit Sub

ViewErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Cri_EE()
Dim EECri As String

If Len(elpEEID.Text) > 0 Then
    EECri = "{HRAUDIT.AU_EMPNBR} in [" & getEmpnbr(elpEEID.Text) & "] "
    
    If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
    
    glbstrSelCri = glbstrSelCri & EECri
End If

End Sub

Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY As Integer, dtMM As Integer, dtDD As Integer


If Len(dlpFrom.Text) = 0 And Len(dlpTo.Text) = 0 Then Exit Sub
TempCri = "({HRAUDIT.AU_LDATE} "
If Len(dlpFrom.Text) > 0 And Len(dlpTo.Text) > 0 Then
  dtYYY = Year(dlpFrom.Text)
  dtMM = month(dlpFrom.Text)
  dtDD = Day(dlpFrom.Text)
  TempCri = TempCri & " in Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ") "
  dtYYY = Year(dlpTo.Text)
  dtMM = month(dlpTo.Text)
  dtDD = Day(dlpTo.Text)
  TempCri = TempCri & " to Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ")) "
Else
  If Len(dlpFrom.Text) > 0 Then
    TempCri = TempCri & " >= "
    dtYYY = Year(dlpFrom.Text)
    dtMM = month(dlpFrom.Text)
    dtDD = Day(dlpFrom.Text)
    TempCri = TempCri & " Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ")) "
  End If
  If Len(dlpTo.Text) > 0 Then
    TempCri = TempCri & " <= "
    dtYYY = Year(dlpTo.Text)
    dtMM = month(dlpTo.Text)
    dtDD = Day(dlpTo.Text)
    TempCri = TempCri & " Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ")) "
  End If
End If
If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
glbstrSelCri = glbstrSelCri & TempCri

End Sub

Private Function Cri_SetAll()
On Error GoTo modSetCriteria_Err
Dim x As Integer

Cri_SetAll = False

Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

Call glbCri_DeptUN("")

' call cri models set both glbiONeWhere and strSelCri
If glbLinamar Then
    Call Cri_Div
Else
    Call Cri_Div1
End If

Call Cri_Section 'Ticket #19437 11/12/2010 Frank
Call Cri_Loc
Call Cri_Region
Call Cri_EE
Call Cri_PP
Call Cri_AdminBy 'Ticket #18352 04/27/2010 Frank
Call Cri_FTDates
Call Cri_Upload
Call Cri_Checks
Call Cri_Sorts
Call Cri_User

If glbWFC Then 'Ticket #27605 Franks 10/16/2015
    If chkExcelRpt.Value Then
        Call WFCExRPTWRK
        Call WFCExcelRpt
        Exit Function
    End If
Else
    'Ticket #28815 - Jerry said to open for all
    If chkExcelRpt.Value Then
        Call WFCExRPTWRK
        Call AllExcelRpt
        Exit Function
    End If
End If

Call setRptLabel(Me, 2)

If optAT(0) <> 0 Then
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzAudit.rpt"
Else
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzAudit2.rpt"
End If

Me.vbxCrystal.SelectionFormula = glbstrSelCri

If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    If optAT(0) <> 0 Then
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
    Else
        Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
    End If
    Me.vbxCrystal.DataFiles(1) = glbIHRDB
    Me.vbxCrystal.DataFiles(2) = glbIHRAUDIT
    Me.vbxCrystal.DataFiles(3) = glbIHRDB
End If

If chkPage Then
  Me.vbxCrystal.SectionFormat(0) = "GH1;T;F;X;X;X;X;X"
  Me.vbxCrystal.SectionFormat(1) = "GF1;X;X;T;X;X;X;X"
Else
  Me.vbxCrystal.SectionFormat(0) = "GH1;T;X;X;X;X;X;X"
  Me.vbxCrystal.SectionFormat(1) = "GF1;X;F;X;X;X;X;X"
End If

If glbPayWeb Then
  Me.vbxCrystal.Formulas(10) = "xWCB = IF LENGTH ({HRAUDIT.AU_WCB}) > 0 THEN  {HRAUDIT.AU_TYPE} + '                          E.I. Reduced Rate'"
End If

If glbSQL Then 'Ticket #18267, make this function for Samuel and all SQL customers
'If glbWFC Then 'Ticket #12867
    Me.vbxCrystal.Formulas(10) = "WFCNoEXECuser = " & glbNoEXEC & " "
    Me.vbxCrystal.Formulas(11) = "WFCNoNONEuser = " & glbNoNONE & " "
End If

'Ticket #22682 - Release 8.0: Testing
'If glbSQL Or glbOracle Then
'    vbxCrystal.SubreportToChange = "rzAudit2"
'    vbxCrystal.Connect = RptODBC_SQL
'    vbxCrystal.SubreportToChange = ""
'End If

' window title if appropriate
Me.vbxCrystal.WindowTitle = "Audit Master File Report"

Cri_SetAll = True

'For x = 0 To 1000
'    If Me.vbxCrystal.Formulas(x) <> "" Then Debug.Print Me.vbxCrystal.Formulas(x)
'Next

Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Audit Master", "HRAUDIT Report", "Select")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

Private Sub Cri_Upload()
Dim EECri As String

If cmbUpload.ListIndex > 0 Then
    If cmbUpload.ListIndex = 1 Then
        EECri = "{HRAUDIT.AU_UPLOAD} = 'Y' "
    End If
    
    If cmbUpload.ListIndex = 2 Then
        EECri = "{HRAUDIT.AU_UPLOAD} = 'N' "
    End If
    
    If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
    
    glbstrSelCri = glbstrSelCri & EECri
End If
End Sub

Private Sub Form_Activate()
    glbOnTop = "FRMAUDIT"
    Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Dim SQLQ As String

glbOnTop = "FRMAUDIT"
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

If glbLinamar Then
    lblDiv.Visible = False
    clpDiv1.Visible = False
Else
    Call setCaption(lblDiv)
End If

'Ticket #19437
lblSection.Caption = lStr("Section")
'Ticket #22423 Franks 08/30/2012
lblLocation.Caption = lStr("Location")
lblRegion.Caption = lStr("Region")

Screen.MousePointer = HOURGLASS
If glbLinamar Then
    clpPP.Visible = False
    lblPP.Visible = False
End If
Data1.ConnectionString = glbAdoIHRAUDIT

'Hemu - Talked to Jerry about this read - not really required. It's taking time to load
'the page - 09/12/2008
'SQLQ = "SELECT AU_EMPNBR FROM HRAUDIT WHERE AU_EMPNBR IN(SELECT ED_EMPNBR FROM HREMP "
'SQLQ = SQLQ & in_SQL(glbIHRDB)
'SQLQ = SQLQ & " WHERE " & glbSeleDeptUn & ")"
'Data1.RecordSource = SQLQ
'Data1.Refresh

'If Data1.Recordset.EOF And Data1.Recordset.EOF Then
'  MsgBox "ACTIVE AUDIT FILE IS EMPTY"
'  Screen.MousePointer = DEFAULT
'End If

cmbUpload.AddItem "All"
cmbUpload.AddItem "Yes"
cmbUpload.AddItem "No"
cmbUpload.ListIndex = 0

comGroup.Clear
comGroup.AddItem "Date Changed"
comGroup.AddItem "Employee Number"
comGroup.AddItem "Employee Name"
'Ticket #22682 - Release 8.0: Add User to the final sort
comGroup.AddItem "User"
comGroup.ListIndex = 0

If Not gSec_Upd_Audit Then     'May99 js
'    cmdDelete.Enabled = False   '
End If                          '
If glbLinamar Then
    lblTitle(0).Visible = True
    clpDIV.Visible = True
    frmAT.Visible = True
End If
elpUser.LookupType = 2

If glbCompSerial = "S/N - 2382W" Then 'Ticket #18352 Samuel - add Admin By
    lblAdmin.Caption = lStr("Administered By")
    lblAdmin.Top = 650
    lblAdmin.Visible = True
    clpCode(1).Top = 650
    clpCode(1).Visible = True
    frmAT.Top = 120
    
    'Ticket #21181 Franks 11/09/2011 - begin
    lblBank.Visible = True
    lblTerm.Visible = True
    chkBank.Visible = True
    chkTerm.Visible = True
    'Ticket #21181 Franks 11/09/2011 - end
End If

If glbWFC Then 'Ticket #27605 Franks 10/16/2015
    lblExcelRpt.Visible = True
    chkExcelRpt.Visible = True
Else
    'Ticket #28815 - Jerry said to open for all
    lblExcelRpt.Visible = True
    chkExcelRpt.Visible = True
End If

If gsTRAININGMATRIX Then
    xTrainMatrixPath = GetComPreferEmail("TRAININGMATRIX")
End If
If Len(xTrainMatrixPath) = 0 Then
    xTrainMatrixPath = glbIHRREPORTS
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

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmAUDIT = Nothing  'carmen may 2000
End Sub

Private Function modDelRecs()
'''On Error GoTo cmdDel_Err
Dim SQLQ As String, SQLW As String, SQL1 As String, SQLQ1 As String
Dim TmpDeletedRecs As Long, DeletedRecs1 As Long, TmpDeletedRecs1 As Long, TmpDeletedRecs2 As Long, DeletedEmp0Recs As Long, DeletedEmp0Recs2 As Long
Dim SQLQ2, SQLQ_0 As String

modDelRecs = False

glbstrSelCri = ""
Screen.MousePointer = HOURGLASS

SQLQ = "Delete FROM HRAUDIT WHERE 1=1 "

' do selection for pay period if they entered one
If Len(clpPP.Text) > 0 Then
    SQLQ = SQLQ & "AND AU_EMPNBR IN (SELECT SH_EMPNBR FROM HR_SALARY_HISTORY "
    If Not glbSQL Then
        SQLQ = SQLQ & in_SQL(glbIHRDB)
    End If
    SQLQ = SQLQ & "WHERE SH_CURRENT<>0 AND SH_PAYP='" & clpPP.Text & "') "
End If

' pay period selection end
If glbLinamar Then
    ' do selection for only emps we have security for
    SQLQ = SQLQ & "AND AU_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLW = "WHERE " & glbSeleDeptUn & ")"
Else
    SQLW = ""
End If

SQLQ1 = SQLQ

If Len(elpEEID.Text) > 0 Then SQLW = SQLW & " AND AU_EMPNBR in (" & getEmpnbr(elpEEID.Text) & ") "
If Len(dlpFrom.Text) > 0 Then SQLW = SQLW & " AND AU_LDATE >= " & Date_SQL(dlpFrom.Text)
If Len(dlpTo.Text) > 0 Then SQLW = SQLW & " AND AU_LDATE <= " & Date_SQL(dlpTo.Text)
If glbLinamar Then
    If Len(clpDIV) > 0 Then SQLW = SQLW & " AND RIGHT(AU_EMPNBR,3)=" & clpDIV
Else
    If Len(clpDiv1.Text) > 0 Then SQLW = SQLW & " AND AU_DIVUPL IN ('" & getCodes(clpDiv1.Text) & "') "
End If
If Len(elpUser.Text) > 0 Then SQLW = SQLW & "AND Lower(AU_LUSER) = '" & LCase(elpUser.Text) & "' "
If cmbUpload.ListIndex > 0 Then
  If cmbUpload.ListIndex = 1 Then SQLW = SQLW + " AND AU_UPLOAD = 'Y' "
  If cmbUpload.ListIndex = 2 Then SQLW = SQLW + " AND AU_UPLOAD = 'N' "
End If

glbstrSelCri = ""
If glbSQL Or glbOracle Then
    Call glbCri_DeptUN("")
    glbstrSelCri = Trim(Replace(Replace(glbstrSelCri, "{", ""), "}", ""))
    If LCase(Left(Trim(glbstrSelCri), 3)) = "and" Then
        glbstrSelCri = Mid(glbstrSelCri, 4, Len(glbstrSelCri) - 3)
    End If
    glbstrSelCri = " AND (AU_EMPNBR in (SELECT ED_EMPNBR FROM HREMP WHERE " & glbstrSelCri & ") OR AU_EMPNBR in (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & Replace(glbstrSelCri, "HREMP.", "Term_HREMP.") & ")  )"
    
    SQLW = SQLW & glbstrSelCri
End If

If glbLinamar Then
    SQLW = SQLW & " AND AU_TYPE<>'R'"
End If

' Ticket #19591 Franks 12/30/2010 - begin
If chkSalary.Value = vbChecked Then
    SQLQ = SQLQ & " AND NOT (AU_SALARY IS NULL) "
End If
If chkNewHires.Value = vbChecked Then
    SQLQ = SQLQ & " AND (AU_NEWEMP = 'Y') "
End If
If chkBenefits.Value = vbChecked Then
    SQLQ = SQLQ & " AND NOT (AU_BCODE IS NULL) "
End If
If chkNameAdd.Value = vbChecked Then
    SQLQ = SQLQ & " AND (NOT (AU_SURNAME IS NULL) OR NOT (AU_ADDR1 IS NULL)) "
End If
' Ticket #19591 Franks 12/30/2010 - end

'Ticket #22334 - Begin
If chkStatus.Value = vbChecked Then
    SQLQ = SQLQ & " AND NOT (AU_EMP IS NULL) "
End If
'Ticket #22334 - End

'Ticket #21181 Franks 11/09/2011 - begin
If glbCompSerial = "S/N - 2382W" Then
    If chkBank.Value = vbChecked Then
        SQLQ = SQLQ & " AND (NOT (AU_BANK IS NULL) OR NOT (AU_BANK2 IS NULL) OR NOT (AU_BANK3 IS NULL)) "
    End If
End If
'Ticket #21181 Franks 11/09/2011 - end
SQLQ = SQLQ & SQLW
gdbAdoIhr001X.Execute SQLQ, DeletedRecs

'--------------------------------------------------------------------------------------------
'Delete Audit records with AU_DIVUPL = blank or null
If Not glbLinamar Or Len(clpDiv1.Text) > 0 Then
    SQL1 = ""
    If Len(elpEEID.Text) > 0 Then SQL1 = SQL1 & " AND AU_EMPNBR in (" & getEmpnbr(elpEEID.Text) & ") "
    If Len(dlpFrom.Text) > 0 Then SQL1 = SQL1 & " AND AU_LDATE >= " & Date_SQL(dlpFrom.Text)
    If Len(dlpTo.Text) > 0 Then SQL1 = SQL1 & " AND AU_LDATE <= " & Date_SQL(dlpTo.Text)
    
    'If Len(clpDiv1.Text) > 0 Then SQL1 = SQL1 & " AND AU_DIVUPL IN ('" & getCodes(clpDiv1.Text) & "') "
    If Len(clpDiv1.Text) > 0 Then SQL1 = SQL1 & " AND ((AU_DIVUPL IS NULL OR AU_DIVUPL = '') AND AU_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_DIV IN ('" & getCodes(clpDiv1.Text) & "')))"
    
    If Len(elpUser.Text) > 0 Then SQL1 = SQL1 & " AND Lower(AU_LUSER) = '" & LCase(elpUser.Text) & "' "
    If cmbUpload.ListIndex > 0 Then
      If cmbUpload.ListIndex = 1 Then SQL1 = SQL1 + " AND AU_UPLOAD = 'Y' "
      If cmbUpload.ListIndex = 2 Then SQL1 = SQL1 + " AND AU_UPLOAD = 'N' "
    End If
    SQL1 = SQL1 & glbstrSelCri
    SQLQ1 = SQLQ1 & SQL1
    ' Ticket #19591 Franks 12/30/2010 - begin
    If chkSalary.Value = vbChecked Then
        SQLQ1 = SQLQ1 & " AND NOT (AU_SALARY IS NULL) "
    End If
    If chkNewHires.Value = vbChecked Then
        SQLQ1 = SQLQ1 & " AND (AU_NEWEMP = 'Y') "
    End If
    If chkBenefits.Value = vbChecked Then
        SQLQ1 = SQLQ1 & " AND NOT (AU_BCODE IS NULL) "
    End If
    If chkNameAdd.Value = vbChecked Then
        SQLQ1 = SQLQ1 & " AND (NOT (AU_SURNAME IS NULL) OR NOT (AU_ADDR1 IS NULL)) "
    End If
    ' Ticket #19591 Franks 12/30/2010 - end
    
    'Ticket #22334 - Begin
    If chkStatus.Value = vbChecked Then
        SQLQ1 = SQLQ1 & " AND NOT (AU_EMP IS NULL) "
    End If
    'Ticket #22334 - End
    
    gdbAdoIhr001X.Execute SQLQ1, DeletedRecs1
End If
'--------------------------------------------------------------------------------------------

' dkostka - 08/20/2001 - Added code to remove records for terminated emps too
SQLQ = "DELETE FROM HRAUDIT WHERE 1=1 "
If glbLinamar Then
    SQLQ = SQLQ & "AND AU_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP "
End If
SQLQ = SQLQ & SQLW

' do selection for pay period if they entered one
If Len(clpPP.Text) > 0 Then
    SQLQ = SQLQ & "AND AU_EMPNBR IN (SELECT SH_EMPNBR FROM Term_SALARY_HISTORY "
    SQLQ = SQLQ & "WHERE SH_CURRENT<>0 AND SH_PAYP='" & clpPP.Text & "') "
End If
' pay period selection end

SQLQ = SQLQ & glbstrSelCri
SQLQ2 = SQLQ
' Ticket #19591 Franks 12/30/2010 - begin
If chkSalary.Value = vbChecked Then
    SQLQ2 = SQLQ2 & " AND NOT (AU_SALARY IS NULL) "
End If
If chkNewHires.Value = vbChecked Then
    SQLQ2 = SQLQ2 & " AND (AU_NEWEMP = 'Y') "
End If
If chkBenefits.Value = vbChecked Then
    SQLQ2 = SQLQ2 & " AND NOT (AU_BCODE IS NULL) "
End If
If chkNameAdd.Value = vbChecked Then
    SQLQ2 = SQLQ2 & " AND (NOT (AU_SURNAME IS NULL) OR NOT (AU_ADDR1 IS NULL)) "
End If
' Ticket #19591 Franks 12/30/2010 - end

'Ticket #22334 - Begin
If chkStatus.Value = vbChecked Then
    SQLQ2 = SQLQ2 & " AND NOT (AU_EMP IS NULL) "
End If
'Ticket #22334 - End


'gdbAdoIhr001X.Execute SQLQ, TmpDeletedRecs
gdbAdoIhr001X.Execute SQLQ2, TmpDeletedRecs
'DeletedRecs = DeletedRecs + TmpDeletedRecs

'--------------------------------------------------------------------------------------------
'Delete Audit records with AU_DIVUPL = blank or null - Terminated employees
If Not glbLinamar Or Len(clpDiv1.Text) > 0 Then
    SQLQ2 = "DELETE FROM HRAUDIT WHERE 1=1 "
    SQLQ2 = SQLQ2 & SQL1
    
    ' do selection for pay period if they entered one
    If Len(clpPP.Text) > 0 Then
        SQLQ2 = SQLQ2 & "AND AU_EMPNBR IN (SELECT SH_EMPNBR FROM Term_SALARY_HISTORY "
        SQLQ2 = SQLQ2 & "WHERE SH_CURRENT<>0 AND SH_PAYP='" & clpPP.Text & "') "
    End If
    ' Ticket #19591 Franks 12/30/2010 - begin
    If chkSalary.Value = vbChecked Then
        SQLQ2 = SQLQ2 & " AND NOT (AU_SALARY IS NULL) "
    End If
    If chkNewHires.Value = vbChecked Then
        SQLQ2 = SQLQ2 & " AND (AU_NEWEMP = 'Y') "
    End If
    If chkBenefits.Value = vbChecked Then
        SQLQ2 = SQLQ2 & " AND NOT (AU_BCODE IS NULL) "
    End If
    If chkNameAdd.Value = vbChecked Then
        SQLQ2 = SQLQ2 & " AND (NOT (AU_SURNAME IS NULL) OR NOT (AU_ADDR1 IS NULL)) "
    End If
    ' Ticket #19591 Franks 12/30/2010 - end
    
    'Ticket #22334 - Begin
    If chkStatus.Value = vbChecked Then
        SQLQ2 = SQLQ2 & " AND NOT (AU_EMP IS NULL) "
    End If
    'Ticket #22334 - End
    
    SQLQ2 = SQLQ2 & glbstrSelCri
    
    ' pay period selection end
    gdbAdoIhr001X.Execute SQLQ2, TmpDeletedRecs1
    'DeletedRecs = DeletedRecs + TmpDeletedRecs1 + DeletedRecs1
End If
'--------------------------------------------------------------------------------------------

'Ticket #15576 v7.8 make HRAUDIT2 for customers
'If glbCompSerial = "S/N - 2380W" Then   'VitalAire Ticket #12142
'HRAUDIT2
    TmpDeletedRecs = 0
    SQLQ = Replace(SQLQ, "HRAUDIT", "HRAUDIT2")
    
    'Ticket #22682 - Release 8.0: Deleting from HRAUDIT2 the Emergency Contact if checked
    If chkEmergContact.Value = vbChecked Then
        SQLQ = SQLQ & " AND NOT (AU_ECONT IS NULL)"
    End If
    
    gdbAdoIhr001X.Execute SQLQ, TmpDeletedRecs2
    'DeletedRecs = DeletedRecs + TmpDeletedRecs + TmpDeletedRecs2
'End If

'Ticket #16768
SQLQ_0 = "DELETE FROM HRAUDIT WHERE AU_EMPNBR = 0"
gdbAdoIhr001X.Execute SQLQ_0, DeletedEmp0Recs

'Ticket #22682 - Release 8.0: Deleting Employee # = 0 from HRAUDIT2 as well
SQLQ_0 = "DELETE FROM HRAUDIT2 WHERE AU_EMPNBR = 0"
gdbAdoIhr001X.Execute SQLQ_0, DeletedEmp0Recs2


DeletedRecs = DeletedRecs + DeletedRecs1 + DeletedEmp0Recs2 + TmpDeletedRecs + TmpDeletedRecs1 + TmpDeletedRecs2
  

modDelRecs = True

Exit Function

cmdDel_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modDelRecs", "HRAUDIT", "Delete")

Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub Cri_Div()

Dim DivCri As String

If Len(clpDIV.Text) > 0 Then
    DivCri = "(RIGHT(TOTEXT({HRAUDIT.AU_EMPNBR},0),3) = '" & clpDIV.Text & "')"
End If

If Len(DivCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = DivCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & DivCri
    End If
    glbiOneWhere = True
End If

End Sub



Public Property Get ChangeAction() As UpdateStateEnum
ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = Reports
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Audit
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property
Public Property Get Updateble() As Boolean
Updateble = False
End Property
Public Property Get Deleteble() As Boolean
    Deleteble = True
End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Call set_Buttons
MDIMain.MainToolBar.ButtonS(10).Visible = True
MDIMain.MainToolBar.ButtonS(10).Enabled = True

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Cri_Checks()
'Added by Bryan 6/Jul/05 for ticket#8857
    Dim TempCri As String
        
    If Not glbLinamar Then
        If Not clpDiv1.ListChecker Then
            Exit Sub
        End If
    End If
    
    If chkNewHires.Value = vbChecked Then
        TempCri = "({HRAUDIT.AU_NEWEMP} = 'Y') "
    End If
    
    If chkSalary.Value = vbChecked Then
        If Len(TempCri) >= 1 Then
            TempCri = TempCri & "AND (isnull({HRAUDIT.AU_SALARY})=false ) "
        Else
            TempCri = "(isnull({HRAUDIT.AU_SALARY})=false ) "
        End If
    End If
    
    If chkBenefits.Value = vbChecked Then
        If Len(TempCri) >= 1 Then
            TempCri = TempCri & "AND (isnull({HRAUDIT.AU_BCODE})=false ) "
        Else
            TempCri = "(isnull({HRAUDIT.AU_BCODE})=false ) "
        End If
    End If
    
    If chkNameAdd.Value = vbChecked Then
        If Len(TempCri) >= 1 Then
            TempCri = TempCri & "AND ((isnull({HRAUDIT.AU_FNAME})=false) OR (isnull({HRAUDIT.AU_SURNAME})=false) OR (isnull({HRAUDIT.AU_ADDR1})=false) OR (isnull({HRAUDIT.AU_ADDR2})=false) OR (isnull({HRAUDIT.AU_CITY})=false) OR (isnull({HRAUDIT.AU_PCODE})=false) OR (isnull({HRAUDIT.AU_PROV})=false) OR (isnull({HRAUDIT.AU_PHONE})=false) ) "
            TempCri = TempCri & "AND isnull({HRAUDIT.AU_EMP}) "
        Else
            TempCri = "((isnull({HRAUDIT.AU_FNAME})=false) OR (isnull({HRAUDIT.AU_SURNAME})=false) OR (isnull({HRAUDIT.AU_ADDR1})=false) OR (isnull({HRAUDIT.AU_ADDR2})=false) OR (isnull({HRAUDIT.AU_CITY})=false) OR (isnull({HRAUDIT.AU_PCODE})=false) OR (isnull({HRAUDIT.AU_PROV})=false) OR (isnull({HRAUDIT.AU_PHONE})=false) ) "
            TempCri = TempCri & "AND isnull({HRAUDIT.AU_EMP}) "
        End If
    End If
        
    'Ticket #22334 - Begin
    If chkStatus.Value = vbChecked Then
        If Len(TempCri) >= 1 Then
            TempCri = TempCri & "AND (isnull({HRAUDIT.AU_EMP})=false ) "
        Else
            TempCri = "(isnull({HRAUDIT.AU_EMP})=false ) "
        End If
    
    End If
    'Ticket #22334 - End
        
    'Ticket #21181 Franks 11/09/2011 - begin
    If glbCompSerial = "S/N - 2382W" Then
        If chkBank.Value = vbChecked Then
            If Len(TempCri) >= 1 Then
                TempCri = TempCri & "AND ((isnull({HRAUDIT.AU_DOT})=false) OR (isnull({HRAUDIT.AU_BANK2})=false) OR (isnull({HRAUDIT.AU_BANK3})=false) ) "
                TempCri = TempCri & "AND isnull({HRAUDIT.AU_SURNAME}) "
            Else
                TempCri = "((isnull({HRAUDIT.AU_BANK})=false) OR (isnull({HRAUDIT.AU_BANK2})=false) OR (isnull({HRAUDIT.AU_BANK3})=false) ) "
                TempCri = TempCri & "AND isnull({HRAUDIT.AU_SURNAME}) "
            End If
        End If
        If chkTerm.Value = vbChecked Then
            If Len(TempCri) >= 1 Then
                TempCri = TempCri & "AND (((isnull({HRAUDIT.AU_DOT})=false) AND {HRAUDIT.AU_TYPE} = 'T') OR {HRAUDIT.AU_TYPE} = 'L') "
            Else
                TempCri = "(((isnull({HRAUDIT.AU_DOT})=false) AND {HRAUDIT.AU_TYPE} = 'T') OR {HRAUDIT.AU_TYPE} = 'L') "
            End If
        End If
    End If
    'Ticket #21181 Franks 11/09/2011 - end


    If chkEmergContact.Value = vbChecked Then
        If Len(TempCri) >= 1 Then
            TempCri = TempCri & "AND ((isnull({HRAUDIT2.AU_ECONT})=false) OR (isnull({HRAUDIT2.AU_ENBR})=false) OR (isnull({HRAUDIT2.AU_RELATE})=false) ) "
        Else
            TempCri = "((isnull({HRAUDIT2.AU_ECONT})=false) OR (isnull({HRAUDIT2.AU_ENBR})=false) OR (isnull({HRAUDIT2.AU_RELATE})=false) ) "
        End If
    End If
    
    'Ticket #28635 - Employment Dates
    If chkEmploymentDates.Value = vbChecked Then
        If Len(TempCri) >= 1 Then
            TempCri = TempCri & "AND ((isnull({HRAUDIT.AU_SFDATE})=false) OR (isnull({HRAUDIT.AU_STDATE})=false) OR (isnull({HRAUDIT.AU_PTEDATE})=false) "
            TempCri = TempCri & "OR (isnull({HRAUDIT.AU_DOH})=false) OR (isnull({HRAUDIT.AU_SENDTE})=false) OR (isnull({HRAUDIT.AU_LTHIRE})=false) "
            TempCri = TempCri & "OR (isnull({HRAUDIT.AU_UNION})=false) OR (isnull({HRAUDIT.AU_FDAY})=false) OR (isnull({HRAUDIT.AU_LDAY})=false) "
            TempCri = TempCri & "OR (isnull({HRAUDIT.AU_OMDAY})=false) OR (isnull({HRAUDIT.AU_USRDAT1})=false) OR (isnull({HRAUDIT.AU_ELIGIBLE})=false) "
            TempCri = TempCri & "OR (isnull({HRAUDIT.AU_EARLYR})=false) OR (isnull({HRAUDIT.AU_NORMALR})=false) OR (isnull({HRAUDIT.AU_LATESTR})=false) ) "
            TempCri = TempCri & "OR (isnull({HRAUDIT.AU_FMLA})=false) ) "
        Else
            TempCri = "((isnull({HRAUDIT.AU_SFDATE})=false) OR (isnull({HRAUDIT.AU_STDATE})=false) OR (isnull({HRAUDIT.AU_PTEDATE})=false) "
            TempCri = TempCri & "OR (isnull({HRAUDIT.AU_DOH})=false) OR (isnull({HRAUDIT.AU_SENDTE})=false) OR (isnull({HRAUDIT.AU_LTHIRE})=false) "
            TempCri = TempCri & "OR (isnull({HRAUDIT.AU_UNION})=false) OR (isnull({HRAUDIT.AU_FDAY})=false) OR (isnull({HRAUDIT.AU_LDAY})=false) "
            TempCri = TempCri & "OR (isnull({HRAUDIT.AU_OMDAY})=false) OR (isnull({HRAUDIT.AU_USRDAT1})=false) OR (isnull({HRAUDIT.AU_ELIGIBLE})=false) "
            TempCri = TempCri & "OR (isnull({HRAUDIT.AU_EARLYR})=false) OR (isnull({HRAUDIT.AU_NORMALR})=false) OR (isnull({HRAUDIT.AU_LATESTR})=false) "
            TempCri = TempCri & "OR (isnull({HRAUDIT.AU_FMLA})=false) ) "
        End If
    End If
    

  If Len(glbstrSelCri) > 3 And Len(TempCri) >= 1 Then glbstrSelCri = glbstrSelCri & " AND "
  glbstrSelCri = glbstrSelCri & TempCri

    
End Sub

Private Sub Cri_Sorts()
'Added by Bryan on Sep 7, 2005 Ticket#9279
    Dim grpField As String
    Dim grpCond As String
    'If optAT(0) <> 0 Then  'Ticket #18668
        Select Case comGroup.ListIndex
            Case 0:
                grpField = "{HRAUDIT.AU_LDATE}"
                grpCond = "GROUP" & CStr(1) & ";{HRAUDIT.AU_LDATE};ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(0) = grpCond
                grpCond = "GROUP" & CStr(2) & ";{@EFullName};ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(1) = grpCond
                Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = 'Date of Change:'"
                Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = {HRAUDIT.AU_LDATE}"
                Me.vbxCrystal.Formulas(2) = "DESCGROUP3 = ''"
                Me.vbxCrystal.Formulas(3) = "lblEMPNO = ''"
            Case 1:
                grpCond = "GROUP" & CStr(1) & ";{HRAUDIT.AU_EMPNBR};ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(0) = grpCond
'                grpCond = "GROUP" & CStr(2) & ";{HRAUDIT.AU_LDATE};ANYCHANGE;A"
'                Me.vbxCrystal.GroupCondition(1) = grpCond
                Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = 'Employee:'"
                Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = {@EFullName}"
                Me.vbxCrystal.Formulas(2) = "DESCGROUP3 = 'Number:'"
            Case 2:
                grpCond = "GROUP" & CStr(1) & ";{@EFullName};ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(0) = grpCond
'                grpCond = "GROUP" & CStr(2) & ";{HRAUDIT.AU_LDATE};ANYCHANGE;A"
'                Me.vbxCrystal.GroupCondition(1) = grpCond
                grpField = "{@EFullName}"
                Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = 'Employee:'"
                Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = {@EFullName}"
                Me.vbxCrystal.Formulas(2) = "DESCGROUP3 = 'Number:'"
            Case 3:
                grpCond = "GROUP" & CStr(1) & ";{HRAUDIT.AU_LUSER};ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(0) = grpCond
'                grpCond = "GROUP" & CStr(2) & ";{HRAUDIT.AU_LDATE};ANYCHANGE;A"
'                Me.vbxCrystal.GroupCondition(1) = grpCond
                Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = 'Employee:'"
                Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = {@EFullName}"
                Me.vbxCrystal.Formulas(2) = "DESCGROUP3 = 'Number:'"
            Case Else: grpField = "(none)"
        End Select
    'End If

End Sub

Private Sub Cri_User()
Dim EECri As String

If Len(elpUser.Text) > 0 Then
  EECri = "LowerCase({HRAUDIT.AU_LUSER}) ='" & LCase(elpUser.Text) & "' "
  If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
  glbstrSelCri = glbstrSelCri & EECri
End If

End Sub


Private Sub Cri_Region() 'Ticket #22423
Dim RegionCri As String
Dim countr   As Integer

If Len(clpCode(4).Text) > 0 Then
      RegionCri = " {HREMP.ED_REGION} IN ['" & getCodes(clpCode(4).Text) & "'] "
End If

If Len(RegionCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = RegionCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & RegionCri
    End If
    glbiOneWhere = True
End If
End Sub

Private Sub Cri_Loc() 'Ticket #22423
Dim LocCri As String
Dim countr   As Integer

If Len(clpCode(3).Text) > 0 Then
      LocCri = " {HREMP.ED_LOC} IN ['" & getCodes(clpCode(3).Text) & "'] "
End If

If Len(LocCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = LocCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & LocCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_Section() 'Ticket #19437
Dim SectionCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level

If Len(clpCode(2).Text) > 0 Then
      SectionCri = " {HREMP.ED_SECTION} IN ['" & getCodes(clpCode(2).Text) & "'] "
End If

If Len(SectionCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = SectionCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & SectionCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_Div1()

Dim DivCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level


If Len(clpDiv1.Text) > 0 Then
    'Hemu 06/02/2004 Begin
    'DivCri = "({HREMP.ED_DIV} = '" & clpDiv.Text & "')"
    'If glbOracle Then
    '    DivCri = "({HREMP.ED_DIV} IN ['" & getCodes(clpDiv1.Text) & "'])"
    'Else
    '    DivCri = "({HRAUDIT.AU_DIV} IN ('" & getCodes(clpDiv1.Text) & "'))"
    'End If
    'Hemu 06/02/2004 End
    
    'Ticket #12843
    'DivCri = "({HRAUDIT.AU_DIVUPL} IN ('" & getCodes(clpDiv1.Text) & "'))"
    'Ticket #13540 Frank, come AU_DIVUPL values were null or blank, but still showup on the report
    'DivCri = "(Length({HRAUDIT.AU_DIVUPL})>0  AND ({HRAUDIT.AU_DIVUPL} IN ('" & getCodes(clpDiv1.Text) & "')))"
    DivCri = "(Length({HRAUDIT.AU_DIVUPL})>0  AND ({HRAUDIT.AU_DIVUPL} IN ['" & getCodes(clpDiv1.Text) & "']))"
End If

If Len(DivCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = DivCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & DivCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub optAT_Click(Index As Integer)
    'Ticket #15483
    If Index = 1 Then
        elpEEID.LookupType = TERM
    Else
        elpEEID.LookupType = 0  '0 = ACTIVE. I cannot put as ACTIVE because it's changing to "Active" and that does not switch the lookup to ACTIVE employees
    End If
End Sub

Private Sub WFCExRPTWRK() 'Ticket #27605 Franks 10/16/2015
Dim rsLAudit As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim SQLQ, xNum, xRecNum
Dim xFieldList
Dim xFlag As Boolean
Dim xRow, xRows
Dim xUptDesc As String, xNew As String, xOld As String, xOrder As Integer

On Error GoTo Err_Line

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(2).Caption = ""
    gdbAdoIhr001.CommandTimeout = 600
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodPercent = 0
    
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute "DELETE FROM HREMPHIS_WRK WHERE EE_WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001W.CommitTrans
    Call Pause(1)
    
    SQLQ = "SELECT * FROM HRAUDIT LEFT JOIN HREMP ON HRAUDIT.AU_EMPNBR = HREMP.ED_EMPNBR WHERE (1=1) "
    If Len(elpEEID.Text) > 0 Then
        SQLQ = SQLQ & "AND AU_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
    End If
    If Len(clpDiv1.Text) > 0 Then
        SQLQ = SQLQ & "AND ED_DIV IN ('" & Replace(clpDiv1.Text, ",", "','") & "') "
    End If
    If Len(clpCode(2).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_SECTION IN ('" & Replace(clpCode(2).Text, ",", "','") & "') "
    End If
    If Len(clpCode(3).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_LOC IN ('" & Replace(clpCode(3).Text, ",", "','") & "') "
    End If
    If Len(clpCode(4).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_REGION IN ('" & Replace(clpCode(4).Text, ",", "','") & "') "
    End If
    If IsDate(dlpFrom.Text) Then
        SQLQ = SQLQ & "AND AU_LDATE >= " & Date_SQL(dlpFrom.Text) & " "
    End If
    If IsDate(dlpTo.Text) Then
        SQLQ = SQLQ & "AND AU_LDATE <= " & Date_SQL(dlpTo.Text) & " "
    End If
    
    If chkSalary.Value = vbChecked Then
        SQLQ = SQLQ & " AND NOT (AU_SALARY IS NULL) "
    End If
    If chkNewHires.Value = vbChecked Then
        SQLQ = SQLQ & " AND (AU_NEWEMP = 'Y') "
    End If
    If chkBenefits.Value = vbChecked Then
        SQLQ = SQLQ & " AND NOT (AU_BCODE IS NULL) "
    End If
    If chkNameAdd.Value = vbChecked Then
        SQLQ = SQLQ & " AND (NOT (AU_SURNAME IS NULL) OR NOT (AU_ADDR1 IS NULL)) "
    End If
    If chkStatus.Value = vbChecked Then
        SQLQ = SQLQ & " AND NOT (AU_EMP IS NULL) "
    End If
    
    If Len(elpUser.Text) > 0 Then SQLQ = SQLQ & "AND Lower(AU_LUSER) = '" & LCase(elpUser.Text) & "' "
    If cmbUpload.ListIndex > 0 Then
      If cmbUpload.ListIndex = 1 Then SQLQ = SQLQ & "  AND AU_UPLOAD = 'Y' "
      If cmbUpload.ListIndex = 2 Then SQLQ = SQLQ & " AND AU_UPLOAD = 'N' "
    End If
    If Len(clpPP.Text) > 0 Then
        SQLQ = SQLQ & "AND AU_EMPNBR IN (SELECT SH_EMPNBR FROM HR_SALARY_HISTORY "
        SQLQ = SQLQ & "WHERE SH_CURRENT<>0 AND SH_PAYP='" & clpPP.Text & "') "
    End If
    
    ''If chkEmergContact.Value = vbChecked Then
    ''SQLQ = Replace(SQLQ, "HRAUDIT", "HRAUDIT2")
    '''Ticket #22682 - Release 8.0: Deleting from HRAUDIT2 the Emergency Contact if checked
    ''If chkEmergContact.Value = vbChecked Then
    ''    SQLQ = SQLQ & " AND NOT (AU_ECONT IS NULL)"
    ''End If

    SQLQ = SQLQ & "ORDER BY AU_EMPNBR, AU_LDATE, AU_LTIME "
    If rsLAudit.State <> 0 Then rsLAudit.Close
    rsLAudit.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsLAudit.EOF Then
        MsgBox "No record found."
        Exit Sub
    Else
        xRows = rsLAudit.RecordCount
    End If
    
    SQLQ = "SELECT * FROM HREMPHIS_WRK WHERE EE_WRKEMP='" & glbUserID & "'"
    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    Do While Not rsLAudit.EOF
        MDIMain.panHelp(0).FloodPercent = (xRow / xRows) * 100
        DoEvents
        xRow = xRow + 1
        'xFlag = False: xNew = "": xOld = "": xOrder = 0
        If Not IsNull(rsLAudit("AU_TITLE")) Then
            xUptDesc = "Salutation": xOrder = 1
            xNew = rsLAudit("AU_TITLE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_BADGEID")) Then
            xUptDesc = "Badge ID": xOrder = 2
            xNew = rsLAudit("AU_BADGEID"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SURNAME")) Then
            xUptDesc = "Last Name": xOrder = 3
            xNew = rsLAudit("AU_SURNAME"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_FNAME")) Then
            xUptDesc = "First Name": xOrder = 4
            xNew = rsLAudit("AU_FNAME"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_MIDNAME")) Then
            xUptDesc = "Middle Name": xOrder = 5
            xNew = rsLAudit("AU_MIDNAME"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_ALIAS")) Then
            xUptDesc = "Alias": xOrder = 6
            xNew = rsLAudit("AU_ALIAS"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_ADDR1")) Then
            xUptDesc = "Address Line 1": xOrder = 7
            xNew = rsLAudit("AU_ADDR1"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_ADDR2")) Then
            xUptDesc = "Address Line 2": xOrder = 8
            xNew = rsLAudit("AU_ADDR2"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_CITY")) Then
            xUptDesc = "City": xOrder = 9
            xNew = rsLAudit("AU_CITY"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PROV")) Then
            xUptDesc = "Province": xOrder = 10
            xNew = rsLAudit("AU_PROV"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If

        If Not IsNull(rsLAudit("AU_COUNTRY")) Then
            xUptDesc = "Country": xOrder = 11
            xNew = rsLAudit("AU_COUNTRY"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PCODE")) Then
            xUptDesc = "Postal Code": xOrder = 12
            xNew = rsLAudit("AU_PCODE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PHONE")) Then
            xUptDesc = "Telephone": xOrder = 13
            xNew = rsLAudit("AU_PHONE"): xOld = ""
            If Len(xNew) > 0 Then
                xNew = getLocPhoneFormat(xNew)
            End If
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_BUSNBR")) Then
            xUptDesc = "Business Phone": xOrder = 14
            xNew = rsLAudit("AU_BUSNBR"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PROVRES")) Then
            xUptDesc = "Province of Residence": xOrder = 15
            xNew = rsLAudit("AU_PROVRES"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DOB")) Then
            xUptDesc = "Date of Birth": xOrder = 16
            xNew = rsLAudit("AU_DOB"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SEX")) Then
            xUptDesc = "Gender": xOrder = 17
            xNew = rsLAudit("AU_SEX"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        
        If Not IsNull(rsLAudit("AU_MSTAT")) Then
            xUptDesc = "Marital Status": xOrder = 18
            xNew = rsLAudit("AU_MSTAT"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SMOKER")) Then
            xUptDesc = "Smoker": xOrder = 19
            xNew = rsLAudit("AU_SMOKER"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SIN")) Then
            xUptDesc = "S.I.N.": xOrder = 20
            xNew = rsLAudit("AU_SIN"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SSN")) Then
            xUptDesc = "S.S.N.": xOrder = 21
            xNew = rsLAudit("AU_SSN"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DEPTNO")) Then
            xUptDesc = lStr("Department"): xOrder = 22
            xNew = rsLAudit("AU_DEPTNO"): xOld = ""
            If Not IsNull(rsLAudit("AU_OLDDEPT")) Then xOld = rsLAudit("AU_OLDDEPT")
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DEPT_GL")) Then
            xUptDesc = lStr("G/L #") & " Code": xOrder = 23
            xNew = rsLAudit("AU_DEPT_GL"): xOld = ""
            If Not IsNull(rsLAudit("AU_OLD_GL")) Then xOld = rsLAudit("AU_OLD_GL")
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        
        If Not IsNull(rsLAudit("AU_DIV")) Then
            xUptDesc = lStr("Division"): xOrder = 24
            xNew = rsLAudit("AU_DIV"): xOld = ""
            If Not IsNull(rsLAudit("AU_OLDDIV")) Then xOld = rsLAudit("AU_OLDDIV")
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_LOC")) Then
            xUptDesc = lStr("Location"): xOrder = 25
            xNew = rsLAudit("AU_LOC"): xOld = ""
            If Not IsNull(rsLAudit("AU_OLDLOC")) Then xOld = rsLAudit("AU_OLDLOC")
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        
        If Not IsNull(rsLAudit("AU_EMP")) Then
            xUptDesc = "Employee Status": xOrder = 26
            xNew = rsLAudit("AU_EMP"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PT")) Then
            xUptDesc = lStr("Category"): xOrder = 27
            xNew = rsLAudit("AU_PT"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_EMPTYPE")) Then
            xUptDesc = "Employment Type": xOrder = 28
            xNew = rsLAudit("AU_EMPTYPE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_EEPT")) Then
            xUptDesc = "Employment EE Type": xOrder = 29
            xNew = rsLAudit("AU_EEPT"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_ORG")) Then
            xUptDesc = lStr("Union"): xOrder = 30
            xNew = rsLAudit("AU_ORG"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DOH")) Then
            xUptDesc = "Original Hire Date": xOrder = 31
            xNew = rsLAudit("AU_DOH"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SENDTE")) Then
            xUptDesc = lStr("Seniority Date"): xOrder = 32
            xNew = rsLAudit("AU_SENDTE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_LTHIRE")) Then
            xUptDesc = lStr("Last Hire Date"): xOrder = 33
            xNew = rsLAudit("AU_LTHIRE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_UNION")) Then
            xUptDesc = lStr("Union Date"): xOrder = 34
            xNew = rsLAudit("AU_UNION"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_FDAY")) Then
            xUptDesc = lStr("First Day"): xOrder = 35
            xNew = rsLAudit("AU_FDAY"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_LDAY")) Then
            xUptDesc = lStr("Last Day"): xOrder = 36
            xNew = rsLAudit("AU_LDAY"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_OMDAY")) Then
            xUptDesc = lStr("OMERS Date"): xOrder = 37
            xNew = rsLAudit("AU_OMDAY"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DDI")) Then
            xUptDesc = "Direct Deposit Indicator": xOrder = 38
            xNew = rsLAudit("AU_DDI"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_ELIGIBLE")) Then
            xUptDesc = lStr("Eligibility"): xOrder = 39
            xNew = rsLAudit("AU_ELIGIBLE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_EARLYR")) Then
            xUptDesc = lStr("Earliest Retirement"): xOrder = 40
            xNew = rsLAudit("AU"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_NORMALR")) Then
            xUptDesc = lStr("Normal Retirement"): xOrder = 41
            xNew = rsLAudit("AU_NORMALR"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_LATESTR")) Then
            xUptDesc = lStr("Latest Retirement"): xOrder = 42
            xNew = rsLAudit("AU_LATESTR"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DEPOSIT")) Then
            xUptDesc = "Deposit Code": xOrder = 43
            xNew = rsLAudit("AU_DEPOSIT"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_BANK")) And Not IsNull(rsLAudit("AU_BRANCH")) Then
            xUptDesc = "Bank/Branch": xOrder = 44
            xNew = rsLAudit("AU_BANK") & "-" & rsLAudit("AU_BRANCH"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_ACCOUNT")) Then
            xUptDesc = "Bank Account": xOrder = 45
            xNew = rsLAudit("AU_ACCOUNT"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_AMTDEPOSIT")) Then
            xUptDesc = "Bank Amount Deposit": xOrder = 46
            xNew = rsLAudit("AU_AMTDEPOSIT"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_TRANSITABA")) Then
            xUptDesc = "Transit/ABA": xOrder = 47
            xNew = rsLAudit("AU_TRANSITABA"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PCDEPOSIT")) Then
            xUptDesc = "Bank Percent Deposit": xOrder = 48
            xNew = rsLAudit("AU_PCDEPOSIT"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DEPOSIT2")) Then
            xUptDesc = "Deposit Code 2": xOrder = 49
            xNew = rsLAudit("AU_DEPOSIT2"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_BANK2")) And Not IsNull(rsLAudit("AU_BRANCH2")) Then
            xUptDesc = "Bank/Branch 2": xOrder = 50
            xNew = rsLAudit("AU_BANK2") & "-" & rsLAudit("AU_BRANCH2"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        
        
        If Not IsNull(rsLAudit("AU_ACCOUNT2")) Then
            xUptDesc = "Bank Account 2": xOrder = 51
            xNew = rsLAudit("AU_ACCOUNT2"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_AMTDEPOSIT2")) Then
            xUptDesc = "Bank Amount Deposit 2": xOrder = 52
            xNew = rsLAudit("AU_AMTDEPOSIT2"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_TRANSITABA2")) Then
            xUptDesc = "Transit/ABA 2": xOrder = 53
            xNew = rsLAudit("AU_TRANSITABA2"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PCDEPOSIT2")) Then
            xUptDesc = "Bank Percent Deposit 2": xOrder = 54
            xNew = rsLAudit("AU_PCDEPOSIT2"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DEPOSIT3")) Then
            xUptDesc = "Deposit Code 3": xOrder = 55
            xNew = rsLAudit("AU_DEPOSIT3"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        
        If Not IsNull(rsLAudit("AU_BANK3")) And Not IsNull(rsLAudit("AU_BRANCH3")) Then
            xUptDesc = "Bank/Branch 3": xOrder = 56
            xNew = rsLAudit("AU_BANK3") & "-" & rsLAudit("AU_BRANCH3"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        
        If Not IsNull(rsLAudit("AU_ACCOUNT3")) Then
            xUptDesc = "Bank Account 3": xOrder = 57
            xNew = rsLAudit("AU_ACCOUNT3"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_AMTDEPOSIT3")) Then
            xUptDesc = "Bank Amount Deposit 3": xOrder = 58
            xNew = rsLAudit("AU_AMTDEPOSIT3"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_TRANSITABA3")) Then
            xUptDesc = "Transit/ABA 3": xOrder = 59
            xNew = rsLAudit("AU_TRANSITABA3"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PCDEPOSIT3")) Then
            xUptDesc = "Bank Percent Deposit 3": xOrder = 60
            xNew = rsLAudit("AU_PCDEPOSIT3"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_TD1DOL")) Then ' Or Not IsNull(rsLAudit("AU_TD1")) Then
            xUptDesc = "TD1 : Code / Dollars": xOrder = 61
            xNew = rsLAudit("AU_TD1DOL"): xOld = ""
            If Not IsNull(rsLAudit("AU_TD1")) Then
                xNew = xNew & " " & rsLAudit("AU_TD1")
            End If
            If Not IsNull(rsLAudit("AU_OLDTD1")) Then
                xOld = rsLAudit("AU_OLDTD1")
            End If
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_TD3")) Then
            xUptDesc = "TD3 Dollars": xOrder = 62
            xNew = rsLAudit("AU_TD3"): xOld = ""
            If Not IsNull(rsLAudit("AU_OLDTD3")) Then
                xOld = rsLAudit("AU_OLDTD3")
            End If
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_TD3PC")) Then
            xUptDesc = "Extra Tax %": xOrder = 63
            xNew = rsLAudit("AU_TD3PC"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_VACPC")) Then
            xUptDesc = "Vacation Pay Percent": xOrder = 64
            xNew = rsLAudit("AU_VACPC"): xOld = ""
            If Not IsNull(rsLAudit("AU_OLDVAC")) Then
                xOld = rsLAudit("AU_OLDVAC")
            End If
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_UIC")) Then
            xUptDesc = "UIC Code": xOrder = 65
            xNew = rsLAudit("AU_UIC"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_GROSSCD")) Then
            xUptDesc = "Gross Calculation Code": xOrder = 66
            xNew = rsLAudit("AU_GROSSCD"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_GARN")) Then
            xUptDesc = "Garnishee": xOrder = 67
            xNew = rsLAudit("AU_GARN"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PENSION")) Then
            xUptDesc = "Pension": xOrder = 68
            xNew = rsLAudit("AU_PENSION"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_CPP")) Then
            xUptDesc = "CPP Code": xOrder = 69
            xNew = rsLAudit("AU_CPP"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_WCB")) Then
            xUptDesc = "WCB Code": xOrder = 70
            xNew = rsLAudit("AU_WCB"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SUPCODE")) Then
            xUptDesc = "Supervisor Code": xOrder = 71
            xNew = rsLAudit("AU_SUPCODE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_JOB")) Then
            xUptDesc = "Position Code": xOrder = 72
            xNew = rsLAudit("AU_JOB"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SJDATE")) Then
            xUptDesc = "Position Starting Date": xOrder = 73
            xNew = rsLAudit("AU_SJDATE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DHRS")) Then
            xUptDesc = "Hours per Day": xOrder = 74
            xNew = rsLAudit("AU_DHRS"): xOld = ""
            If Not IsNull(rsLAudit("AU_OLDDHRS")) Then
                 xOld = rsLAudit("AU_OLDDHRS")
            End If
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_WHRS")) Then
            xUptDesc = "Hours per Week": xOrder = 75
            xNew = rsLAudit("AU_WHRS"): xOld = ""
            If Not IsNull(rsLAudit("AU_OLDWHRS")) Then
                 xOld = rsLAudit("AU_OLDWHRS")
            End If
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PHRS")) Then
            xUptDesc = "Hours per Pay Period": xOrder = 76
            xNew = rsLAudit("AU_PHRS"): xOld = ""
            If Not IsNull(rsLAudit("AU_OLDPHRS")) Then
                 xOld = rsLAudit("AU_OLDPHRS")
            End If
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SALARY")) Then
            xUptDesc = "Salary": xOrder = 77
            xNew = rsLAudit("AU_SALARY"): xOld = ""
            If Not IsNull(rsLAudit("AU_OLDSAL")) Then
                 xOld = rsLAudit("AU_OLDSAL")
            End If
            xFlag = True
            If glbUNION = "EXEC" Or glbUNION = "NONE" Then
                If glbNoEXEC Or glbNoNONE Then
                    xFlag = False
                End If
            End If
            If xFlag Then
                Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
            End If
        End If
        If Not IsNull(rsLAudit("AU_SALCD")) Then
            xUptDesc = "Salary Code": xOrder = 78
            xNew = rsLAudit("AU_SALCD"): xOld = ""
            xFlag = True
            If glbUNION = "EXEC" Or glbUNION = "NONE" Then
                If glbNoEXEC Or glbNoNONE Then
                    xFlag = False
                End If
            End If
            If xFlag Then
                Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
            End If
        End If
        If Not IsNull(rsLAudit("AU_PAYP")) Then
            xUptDesc = "Pay Period": xOrder = 79
            xNew = rsLAudit("AU_PAYP"): xOld = ""
            If Not IsNull(rsLAudit("AU_OLDPAYP")) Then
                xOld = rsLAudit("AU_OLDPAYP")
            End If
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        
        If Not IsNull(rsLAudit("AU_EDATE")) Then
            xUptDesc = "Effective Date": xOrder = 80
            xNew = rsLAudit("AU_EDATE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SEDATE")) Then
            xUptDesc = "Salary Effective Date": xOrder = 81
            xNew = rsLAudit("AU_SEDATE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SNDATE")) Then
            xUptDesc = "Next Review Date": xOrder = 82
            xNew = rsLAudit("AU_SNDATE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_BCODE")) Then
            xUptDesc = "Benefit Code": xOrder = 83
            xNew = rsLAudit("AU_BCODE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_COVER")) Then
            xUptDesc = "Benefit Coverage": xOrder = 84
            xNew = rsLAudit("AU_COVER"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_BAMT")) Then
            xUptDesc = "Benefit Amount": xOrder = 85
            xNew = rsLAudit("AU_BAMT"): xOld = ""
            xFlag = True
            If glbUNION = "EXEC" Or glbUNION = "NONE" Then
                If glbNoEXEC Or glbNoNONE Then
                    xFlag = False
                End If
            End If
            If xFlag Then
                Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
            End If
        End If
        If Not IsNull(rsLAudit("AU_PER")) Then
            xUptDesc = "Benefit - Per Unit": xOrder = 86
            xNew = rsLAudit("AU_PER"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PPAMT")) Then
            xUptDesc = "Benefit - Per Pay Period": xOrder = 87
            xNew = rsLAudit("AU_PPAMT"): xOld = ""
            If Not IsNull(rsLAudit("AU_OLDPPMT")) Then
                xOld = rsLAudit("AU_OLDPPMT")
            End If
            xFlag = True
            If glbUNION = "EXEC" Or glbUNION = "NONE" Then
                If glbNoEXEC Or glbNoNONE Then
                    xFlag = False
                End If
            End If
            If xFlag Then
                Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
            End If
        End If
        If Not IsNull(rsLAudit("AU_MAXDOL")) Then
            xUptDesc = "Maximum Dollars": xOrder = 88
            xNew = rsLAudit("AU_MAXDOL"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_UNITCOST")) Then
            xUptDesc = "Benefit Unit Cost": xOrder = 89
            xNew = rsLAudit("AU_UNITCOST"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PCE")) Then
            xUptDesc = "Company Percentage": xOrder = 90
            xNew = rsLAudit("AU_PCE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PCC")) Then
            xUptDesc = "Employee Percentage": xOrder = 91
            xNew = rsLAudit("AU_PCC"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_TCOST")) Then
            xUptDesc = "Benefit Total Cost": xOrder = 92
            xNew = rsLAudit("AU_TCOST"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_MTHCCOST")) Then
            xUptDesc = "Benefit Monthly Company Cost": xOrder = 93
            xNew = rsLAudit("AU_MTHCCOST"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_MTHECOST")) Then
            xUptDesc = "Benefit Monthly Employee Cost": xOrder = 94
            xNew = rsLAudit("AU_MTHECOST"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_TAXBEN")) Then
            xUptDesc = "Taxable Benefit": xOrder = 95
            xNew = rsLAudit("AU_TAXBEN"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_CEASEDATE")) Then
            xUptDesc = "Benefit End Date": xOrder = 96
            xNew = rsLAudit("AU_CEASEDATE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_BNAME")) Then
            xUptDesc = "Benefit - Beneficiary Data": xOrder = 97
            xNew = rsLAudit("AU_BNAME"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_BRELATE")) Then
            xUptDesc = "Benefit - Beneficiary Relationship": xOrder = 98
            xNew = rsLAudit("AU_BRELATE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_BDOB")) Then
            xUptDesc = "Benefit - Beneficiary Date of Birth": xOrder = 99
            xNew = rsLAudit("AU_BDOB"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SERVICE")) Then
            xUptDesc = "Service Code": xOrder = 100
            xNew = rsLAudit("AU_SERVICE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SAFETY")) Then
            xUptDesc = "Safety Shoes": xOrder = 101
            xNew = rsLAudit("AU_SAFETY"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_UNIFORM")) Then
            xUptDesc = "Uniform": xOrder = 102
            xNew = rsLAudit("AU_UNIFORM"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_EQUIP")) Then
            xUptDesc = "Equipment": xOrder = 103
            xNew = rsLAudit("AU_EQUIP"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_CLEAN")) Then
            xUptDesc = "Cleaning": xOrder = 104
            xNew = rsLAudit("AU_CLEAN"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DOLENT")) Then
            xUptDesc = "Dollar Entitlements": xOrder = 105
            xNew = rsLAudit("AU_DOLENT"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_EARN")) Then
            xUptDesc = "Other Earnings Code": xOrder = 106
            xNew = rsLAudit("AU_EARN"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_ADOLLAR")) Then
            xUptDesc = "Amount": xOrder = 107
            xNew = rsLAudit("AU_ADOLLAR"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DOFDATE")) Then
            xUptDesc = "From Date": xOrder = 108
            xNew = rsLAudit("AU_DOFDATE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DOTDATE")) Then
            xUptDesc = "To Date": xOrder = 109
            xNew = rsLAudit("AU_DOTDATE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        
        If Not IsNull(rsLAudit("AU_COEFLAG")) Then
            xUptDesc = "Cost of Employment Flag": xOrder = 110
            xNew = rsLAudit("AU_COEFLAG"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_EARNPCE")) Then
            xUptDesc = "Earning %": xOrder = 111
            xNew = rsLAudit("AU_EARNPCE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_REFNBR")) Then
            xUptDesc = "Reference Number": xOrder = 112
            xNew = rsLAudit("AU_REFNBR"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PAIDTO")) Then
            xUptDesc = "Paid To": xOrder = 113
            xNew = rsLAudit("AU_PAIDTO"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PAIDDATE")) Then
            xUptDesc = "Paid Date": xOrder = 114
            xNew = rsLAudit("AU_PAIDDATE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DOT")) Then
            xUptDesc = "Termination Date": xOrder = 115
            xNew = rsLAudit("AU_DOT"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
                
        If Not IsNull(rsLAudit("AU_TREAS")) Then
            xUptDesc = "Termination Code": xOrder = 116
            xNew = rsLAudit("AU_TREAS"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_WCBCODE")) Then
            xUptDesc = "WSIB Code": xOrder = 117
            xNew = rsLAudit("AU_WCBCODE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        'If Not IsNull(rsLAudit("AU_DIVUPL")) Then
        '    xUptDesc = lStr("Division") & " Number ": xOrder = 118
        '    xNew = rsLAudit("AU_DIVUPL"): xOld = ""
        '    Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        'End If
        If Not IsNull(rsLAudit("AU_PROVEMP")) Then
            xUptDesc = "Province of Employment": xOrder = 119
            xNew = rsLAudit("AU_PROVEMP"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        'If Not IsNull(rsLAudit("AU_PTUPL")) Then
        '    xUptDesc = "FT/PT/SE/TR/OT Flag": xOrder = 120
        '    xNew = rsLAudit("AU_PTUPL"): xOld = ""
        '    Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        'End If
        If Not IsNull(rsLAudit("AU_BSTATUS")) Then
            xUptDesc = "Benifits - Benificiary Status": xOrder = 121
            xNew = rsLAudit("AU_BSTATUS"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
                
        If Not IsNull(rsLAudit("AU_JOBCOMMENT")) Then
            xUptDesc = "Job Comment": xOrder = 122
            xNew = rsLAudit("AU_JOBCOMMENT"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SALARYCOMMENT")) Then
            xUptDesc = "Salary Comment": xOrder = 123
            xNew = rsLAudit("AU_SALARYCOMMENT"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_FMLA")) Then
            xUptDesc = "FMLA Date": xOrder = 124
            xNew = rsLAudit("AU_FMLA"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_POLICY")) Then
            xUptDesc = "Policy": xOrder = 125
            xNew = rsLAudit("AU_POLICY"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DEPFNAME")) Then
            xUptDesc = "Dependent First Name": xOrder = 126
            xNew = rsLAudit("AU_DEPFNAME"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DEPSNAME")) Then
            xUptDesc = "Dependent Surname": xOrder = 127
            xNew = rsLAudit("AU_DEPSNAME"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DEPSEX")) Then
            xUptDesc = "Dependent Gender": xOrder = 128
            xNew = rsLAudit("AU_DEPSEX"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DEPDOB")) Then
            xUptDesc = "Dependent Date of Birth": xOrder = 129
            xNew = rsLAudit("AU_DEPDOB"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        
        If Not IsNull(rsLAudit("AU_DEPRELATE")) Then
            xUptDesc = "Dependent Relate": xOrder = 130
            xNew = rsLAudit("AU_DEPRELATE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DEPSMOKER")) Then
            xUptDesc = "Dependent Smoker": xOrder = 131
            xNew = rsLAudit("AU_DEPSMOKER"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DEPSTATUS")) Then
            xUptDesc = "Dependent Status": xOrder = 132
            xNew = rsLAudit("AU_DEPSTATUS"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DEPSIN")) Then
            xUptDesc = "Dependent SIN": xOrder = 133
            xNew = rsLAudit("AU_DEPSIN"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DEPSDATE")) Then
            xUptDesc = "Dependent Eligibility Date": xOrder = 134
            xNew = rsLAudit("AU_DEPSDATE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DEPEDATE")) Then
            xUptDesc = "Dependent End Date": xOrder = 135
            xNew = rsLAudit("AU_DEPEDATE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DEPNO")) Then
            If Not rsLAudit("AU_DEPNO") = 0 Then
                xUptDesc = "Dependent Number": xOrder = 136
                xNew = rsLAudit("AU_DEPNO"): xOld = ""
                Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
            End If
        End If
        If Not IsNull(rsLAudit("AU_DENTAL")) Then
            xUptDesc = "Dental": xOrder = 137
            xNew = rsLAudit("AU_DENTAL"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_MEDICAL")) Then
            xUptDesc = "Medical": xOrder = 138
            xNew = rsLAudit("AU_MEDICAL"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_OTHER")) Then
            xUptDesc = "Other": xOrder = 139
            xNew = rsLAudit("AU_OTHER"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SECTION")) Then
            xUptDesc = lStr("Section"): xOrder = 140
            xNew = rsLAudit("AU_SECTION"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_EXTRATAX")) Then
            xUptDesc = "Federal Tax Method": xOrder = 141
            xNew = rsLAudit("AU_EXTRATAX"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_EXTAMT")) Then
            xUptDesc = "Excempt Amount": xOrder = 142
            xNew = rsLAudit("AU_EXTAMT"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PROVFORM")) Then
            xUptDesc = "Provincial Form": xOrder = 143
            xNew = rsLAudit("AU_PROVFORM"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PROVAMT")) Then
            xUptDesc = "Provincial Amount": xOrder = 144
            xNew = rsLAudit("AU_PROVAMT"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PROVCODE")) Then
            xUptDesc = "Provincial Code": xOrder = 145
            xNew = rsLAudit("AU_PROVCODE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_EXTRATAX")) Then
            xUptDesc = "Extra Tax": xOrder = 146
            xNew = rsLAudit("AU_EXTRATAX"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_EXTRATAXPC")) Then
            xUptDesc = "Extra Tax %": xOrder = 147
            xNew = rsLAudit("AU_EXTRATAXPC"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DIVEDATE")) Then
            xUptDesc = lStr("Division") & " Effective Date": xOrder = 148
            xNew = rsLAudit("AU_DIVEDATE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DEPTEDATE")) Then
            xUptDesc = lStr("Department") & " Effective Date": xOrder = 149
            xNew = rsLAudit("AU_DEPTEDATE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PAYROLL_ID")) Then
            xUptDesc = "Payroll ID": xOrder = 150
            xNew = rsLAudit("AU_PAYROLL_ID"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_REGION")) Then
            xUptDesc = lStr("Region"): xOrder = 151
            xNew = rsLAudit("AU_REGION"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_ADMINBY")) Then
            xUptDesc = lStr("Administered By"): xOrder = 152
            xNew = rsLAudit("AU_ADMINBY"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_INTEL")) Then
            xUptDesc = "Internal Phone Extension": xOrder = 153
            xNew = rsLAudit("AU_INTEL"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_LANG1")) Then
            xUptDesc = "Language 1": xOrder = 154
            xNew = rsLAudit("AU_LANG1"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_LANG2")) Then
            xUptDesc = "Language 2": xOrder = 155
            xNew = rsLAudit("AU_LANG2"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_USRDAT1")) Then
            xUptDesc = lStr("User Defined Date"): xOrder = 156
            xNew = rsLAudit("AU_USRDAT1"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_EXTRANN")) Then
            xUptDesc = "Extra Annual": xOrder = 157
            xNew = rsLAudit("AU_EXTRANN"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_CELLPHONE")) Then
            xUptDesc = "Cellular Telephone": xOrder = 158
            xNew = rsLAudit("AU_CELLPHONE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PAGENBR")) Then
            xUptDesc = "Pager Number": xOrder = 159
            xNew = rsLAudit("AU_PAGENBR"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_EMAIL")) Then
            xUptDesc = "Email": xOrder = 160
            xNew = rsLAudit("AU_EMAIL"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_DRIVERLIC")) Then
            xUptDesc = "Driver License": xOrder = 161
            xNew = rsLAudit("AU_DRIVERLIC"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_LICPLATE1")) Then
            xUptDesc = "License Plate #1": xOrder = 162
            xNew = rsLAudit("AU_LICPLATE1"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_LICPLATE2")) Then
            xUptDesc = "License Plate #2": xOrder = 163
            xNew = rsLAudit("AU_LICPLATE2"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_LOCKER")) Then
            xUptDesc = "Locker #": xOrder = 164
            xNew = rsLAudit("AU_LOCKER"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_COMBINATION")) Then
            xUptDesc = "Combination": xOrder = 165
            xNew = rsLAudit("AU_COMBINATION"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_HEALTHCARD")) Then
            xUptDesc = "Health Card #": xOrder = 166
            xNew = rsLAudit("AU_HEALTHCARD"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_EXPIRYDATE")) Then
            xUptDesc = "Health Card Expiry Date": xOrder = 167
            xNew = rsLAudit("AU_EXPIRYDATE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_VERSION")) Then
            xUptDesc = "Health Card Version": xOrder = 168
            xNew = rsLAudit("AU_VERSION"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_TYPEVEHICLE")) Then
            xUptDesc = "Type of Vehicle": xOrder = 169
            xNew = rsLAudit("AU_TYPEVEHICLE"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        
        If Not IsNull(rsLAudit("AU_PARKPERMIT1")) Then
            xUptDesc = "Parking Permit #1": xOrder = 170
            xNew = rsLAudit("AU_PARKPERMIT1"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_PARKPERMIT2")) Then
            xUptDesc = "Parking Permit #2": xOrder = 171
            xNew = rsLAudit("AU_PARKPERMIT2"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SREASON")) Then
            xUptDesc = "Salary Change Reason": xOrder = 172
            xNew = rsLAudit("AU_SREASON"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_JREASON")) Then
            xUptDesc = "Position Change Reason": xOrder = 173
            xNew = rsLAudit("AU_JREASON"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_SALDIST")) Then
            xUptDesc = "Salary Distribution": xOrder = 174
            xNew = rsLAudit("AU_SALDIST"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_VADIM1")) Then
            xUptDesc = "Vadim Field 1": xOrder = 175
            xNew = rsLAudit("AU_VADIM1"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        If Not IsNull(rsLAudit("AU_VADIM2")) Then
            xUptDesc = "Vadim Field 2": xOrder = 176
            xNew = rsLAudit("AU_VADIM2"): xOld = ""
            Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        End If
        
        ''If Not IsNull(rsLAudit("AU")) Then
        ''    xUptDesc = "Pro": xOrder = 17
        ''    xNew = rsLAudit("AU"): xOld = ""
        ''    Call WFCWrkAuditUpt(rsLAudit, rsWRK, xUptDesc, xNew, xOld, xOrder)
        ''End If

        rsLAudit.MoveNext
    Loop
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    MDIMain.panHelp(2).Caption = ""
    Exit Sub
Err_Line:
    MsgBox Err.Description
    'Resume Next
End Sub

Private Sub WFCExcelRpt() 'Ticket #27605 Franks 10/16/2015
Dim exApp As Object, exBook As Object, exSheet As Object
Dim rsWRK As New ADODB.Recordset
Dim xlsFileTmp As String, xlsFileMat As String
Dim SQLQ As String
Dim K As Long
Dim xRow, xRows
Dim xNewEmp As Boolean
Dim xCurName, xNextName
Dim xCurGroup, xNextGroup

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(2).Caption = ""
    

    SQLQ = "SELECT HREMPHIS_WRK.*,ED_SURNAME,ED_FNAME,ED_VADIM2 FROM HREMPHIS_WRK "
    'SQLQ = SQLQ & "LEFT JOIN qry_HREMP ON HREMPHIS_WRK.KEY_EMPNBR = qry_HREMP.KEY_EMPNBR "
    SQLQ = SQLQ & "LEFT JOIN HREMP ON HREMPHIS_WRK.EE_EMPNBR = HREMP.ED_EMPNBR "
    SQLQ = SQLQ & "WHERE EE_WRKEMP='" & glbUserID & "' "
    'SQLQ = SQLQ & "ORDER BY ED_SURNAME,ED_FNAME,EE_LDATE DESC,TERM_SEQ "
    SQLQ = SQLQ & "ORDER BY ED_SURNAME,ED_FNAME,EE_LDATE "

    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsWRK.EOF Then
        'MsgBox "There is no any record in this Selection Criteria"
        rsWRK.Close
        Exit Sub
    End If
    
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFCAuditTmp.xls"
    
    xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFCAudit(" & glbUserID & ").xls"
    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    FileCopy xlsFileTmp, xlsFileMat
    
    Set exApp = CreateObject("Excel.Application") 'New Excel.Application
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)

    exSheet.Cells(2, 1).Font.Bold = True
    exSheet.Cells(2, 1) = "Date: " & Date
    exSheet.Cells(3, 1).Font.Bold = True
    exSheet.Cells(3, 1) = "Time:" & Time$

    xRows = rsWRK.RecordCount
    xRow = 0

    K = 5
    xNewEmp = True
    xCurGroup = "**"
    xCurName = "***"
    Do While Not rsWRK.EOF
        MDIMain.panHelp(0).FloodPercent = (xRow / xRows) * 100
        DoEvents
        xRow = xRow + 1

        ''If Not comGroup(0).Text = "(none)" Then
        ''    If IsNull(rsWRK("EE_GROUP_DESC")) Then xNextGroup = "" Else xNextGroup = rsWRK("EE_GROUP_DESC")
        ''    If Not (xCurGroup = xNextGroup) Then
        ''        K = K + 1
        ''        xCurGroup = xNextGroup
        ''        exSheet.Cells(K, 1).Font.Bold = True
        ''        exSheet.Cells(K, 1) = xNextGroup
        ''        K = K + 1
        ''    End If
        ''End If

        xNextName = rsWRK("ED_SURNAME") & ", " & rsWRK("ED_FNAME")
        If Not (xCurName = xNextName) Then
            'K = K + 1
            'exSheet.Cells(K, 1).Font.Bold = True
            'exSheet.Cells(K, 1) = "Employee Number and Name: " & rsWRK("EE_EMPNBR") & " " & xNextName
            xCurName = xNextName
            K = K + 1
            exSheet.Cells(K, 1) = "Employee #": exSheet.Cells(K, 1).Font.Bold = True
            exSheet.Cells(K, 2) = "Name": exSheet.Cells(K, 2).Font.Bold = True
            exSheet.Cells(K, 3) = "Pay Group": exSheet.Cells(K, 3).Font.Bold = True
            exSheet.Cells(K, 4) = "Trans.": exSheet.Cells(K, 4).Font.Bold = True
            exSheet.Cells(K, 5) = "Data Being Updated": exSheet.Cells(K, 5).Font.Bold = True
            exSheet.Cells(K, 6) = "New Data": exSheet.Cells(K, 6).Font.Bold = True
            exSheet.Cells(K, 7) = "Previous Data": exSheet.Cells(K, 7).Font.Bold = True
            exSheet.Cells(K, 8) = "By Whom": exSheet.Cells(K, 8).Font.Bold = True
            exSheet.Cells(K, 9) = "Upload": exSheet.Cells(K, 9).Font.Bold = True
            exSheet.Cells(K, 10) = "Date of Change": exSheet.Cells(K, 10).Font.Bold = True
            K = K + 1
        End If
        exSheet.Cells(K, 1) = rsWRK("EE_EMPNBR")
        exSheet.Cells(K, 2) = xNextName
        If Not IsNull(rsWRK("ED_VADIM2")) Then exSheet.Cells(K, 3) = rsWRK("ED_VADIM2")
        exSheet.Cells(K, 4) = rsWRK("EE_SALCD") 'Trans
        exSheet.Cells(K, 5) = rsWRK("EE_HISTYPE") 'Data Being Updated
        exSheet.Cells(K, 6) = rsWRK("EE_NEWVALUE")
        exSheet.Cells(K, 7) = rsWRK("EE_OLDVALUE")
        exSheet.Cells(K, 8) = rsWRK("EE_LUSER") 'By Whom
        exSheet.Cells(K, 9) = rsWRK("EE_LTIME") 'Upload
        exSheet.Cells(K, 10) = rsWRK("EE_LDATE") 'Date of Change
        
        K = K + 1

        rsWRK.MoveNext
    Loop
    rsWRK.Close
    
    If Not exBook Is Nothing Then
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
    End If

    Call Pause(1)
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    MDIMain.panHelp(2).Caption = ""
    
    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If
    
End Sub

Private Sub WFCWrkAuditUpt(rsLAudit As ADODB.Recordset, rsWRK As ADODB.Recordset, xUptDesc As String, xNew As String, xOld As String, xOrder As Integer)
    If Len(xNew) > 0 Then
        rsWRK.AddNew
        rsWRK("EE_EMPNBR") = rsLAudit("AU_EMPNBR")
        rsWRK("EE_SALCD") = rsLAudit("AU_TYPE") 'Trans
        rsWRK("EE_HISTYPE") = Left(xUptDesc, 50) 'Data Being Updated
        rsWRK("EE_NEWVALUE") = Left(xNew, 50)
        rsWRK("EE_OLDVALUE") = Left(xOld, 50)
        rsWRK("EE_LUSER") = rsLAudit("AU_LUSER") 'By Whom
        rsWRK("EE_LTIME") = rsLAudit("AU_UPLOAD")  'Upload
        rsWRK("EE_LDATE") = rsLAudit("AU_LDATE") 'Date of Change
        rsWRK("EE_WRKEMP") = glbUserID
        rsWRK("TERM_SEQ") = xOrder
        rsWRK.Update
    End If
End Sub
Private Function getLocPhoneFormat(xStr)
Dim xTmp, xTmp1, xTmp2, xTmp3
Dim retval
    retval = xStr
    xTmp = xStr
    xTmp = Replace(xTmp, " ", "")
    xTmp = Replace(xTmp, "-", "")
    xTmp = Trim(xTmp)
    xTmp1 = Left(xTmp, 3)
    xTmp2 = Mid(xTmp, 4, 3)
    xTmp3 = Mid(xTmp, 7, 10)
    retval = "(" & xTmp1 & ")" & xTmp2 & "-" & xTmp3
    getLocPhoneFormat = retval
End Function

'Ticket #28815 - Opened for all so copied the WFC routine and making this general
Private Sub AllExcelRpt() 'Ticket #27605 Franks 10/16/2015
Dim exApp As Object, exBook As Object, exSheet As Object
Dim rsWRK As New ADODB.Recordset
Dim xlsFileTmp As String, xlsFileMat As String
Dim SQLQ As String
Dim K As Long
Dim xRow, xRows
Dim xNewEmp As Boolean
Dim xCurName, xNextName
Dim xCurGroup, xNextGroup

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(2).Caption = ""
    

    SQLQ = "SELECT HREMPHIS_WRK.*,ED_SURNAME,ED_FNAME,ED_VADIM2 FROM HREMPHIS_WRK "
    'SQLQ = SQLQ & "LEFT JOIN qry_HREMP ON HREMPHIS_WRK.KEY_EMPNBR = qry_HREMP.KEY_EMPNBR "
    SQLQ = SQLQ & "LEFT JOIN HREMP ON HREMPHIS_WRK.EE_EMPNBR = HREMP.ED_EMPNBR "
    SQLQ = SQLQ & "WHERE EE_WRKEMP='" & glbUserID & "' "
    'SQLQ = SQLQ & "ORDER BY ED_SURNAME,ED_FNAME,EE_LDATE DESC,TERM_SEQ "
    SQLQ = SQLQ & "ORDER BY ED_SURNAME,ED_FNAME,EE_LDATE "

    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsWRK.EOF Then
        'MsgBox "There is no any record in this Selection Criteria"
        rsWRK.Close
        Exit Sub
    End If
    
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "AuditTmp.xls"
    
    'xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "Audit" & glbUserID & ".xls"
    'Ticket# 8293
    If glbLinamar Then 'Or glbCompSerial = "S/N - 2336W" Then
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "Reports\Audit" & glbUserID & ".xls"
    Else
        xlsFileMat = xTrainMatrixPath & IIf(Right(xTrainMatrixPath, 1) = "\", "", "\") & "Audit" & glbUserID & ".xls"
    End If
    
    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    FileCopy xlsFileTmp, xlsFileMat
    
    Set exApp = CreateObject("Excel.Application") 'New Excel.Application
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)

    exSheet.Cells(2, 1).Font.Bold = True
    exSheet.Cells(2, 1) = "Date: " & Date
    exSheet.Cells(3, 1).Font.Bold = True
    exSheet.Cells(3, 1) = "Time:" & Time$
    exSheet.Cells(2, 3) = glbCompName

    xRows = rsWRK.RecordCount
    xRow = 0

    K = 5
    xNewEmp = True
    xCurGroup = "**"
    xCurName = "***"
    Do While Not rsWRK.EOF
        MDIMain.panHelp(0).FloodPercent = (xRow / xRows) * 100
        DoEvents
        xRow = xRow + 1

        ''If Not comGroup(0).Text = "(none)" Then
        ''    If IsNull(rsWRK("EE_GROUP_DESC")) Then xNextGroup = "" Else xNextGroup = rsWRK("EE_GROUP_DESC")
        ''    If Not (xCurGroup = xNextGroup) Then
        ''        K = K + 1
        ''        xCurGroup = xNextGroup
        ''        exSheet.Cells(K, 1).Font.Bold = True
        ''        exSheet.Cells(K, 1) = xNextGroup
        ''        K = K + 1
        ''    End If
        ''End If

        xNextName = rsWRK("ED_SURNAME") & ", " & rsWRK("ED_FNAME")
        If Not (xCurName = xNextName) Then
            'K = K + 1
            'exSheet.Cells(K, 1).Font.Bold = True
            'exSheet.Cells(K, 1) = "Employee Number and Name: " & rsWRK("EE_EMPNBR") & " " & xNextName
            xCurName = xNextName
            K = K + 1
            exSheet.Cells(K, 1) = "Employee #": exSheet.Cells(K, 1).Font.Bold = True
            exSheet.Cells(K, 2) = "Name": exSheet.Cells(K, 2).Font.Bold = True
            exSheet.Cells(K, 3) = "Pay Group": exSheet.Cells(K, 3).Font.Bold = True
            exSheet.Cells(K, 4) = "Trans.": exSheet.Cells(K, 4).Font.Bold = True
            exSheet.Cells(K, 5) = "Data Being Updated": exSheet.Cells(K, 5).Font.Bold = True
            exSheet.Cells(K, 6) = "New Data": exSheet.Cells(K, 6).Font.Bold = True
            exSheet.Cells(K, 7) = "Previous Data": exSheet.Cells(K, 7).Font.Bold = True
            exSheet.Cells(K, 8) = "By Whom": exSheet.Cells(K, 8).Font.Bold = True
            exSheet.Cells(K, 9) = "Upload": exSheet.Cells(K, 9).Font.Bold = True
            exSheet.Cells(K, 10) = "Date of Change": exSheet.Cells(K, 10).Font.Bold = True
            K = K + 1
        End If
        exSheet.Cells(K, 1) = rsWRK("EE_EMPNBR")
        exSheet.Cells(K, 2) = xNextName
        If Not IsNull(rsWRK("ED_VADIM2")) Then exSheet.Cells(K, 3) = rsWRK("ED_VADIM2")
        exSheet.Cells(K, 4) = rsWRK("EE_SALCD") 'Trans
        exSheet.Cells(K, 5) = rsWRK("EE_HISTYPE") 'Data Being Updated
        exSheet.Cells(K, 6) = rsWRK("EE_NEWVALUE")
        exSheet.Cells(K, 7) = rsWRK("EE_OLDVALUE")
        exSheet.Cells(K, 8) = rsWRK("EE_LUSER") 'By Whom
        exSheet.Cells(K, 9) = rsWRK("EE_LTIME") 'Upload
        exSheet.Cells(K, 10) = rsWRK("EE_LDATE") 'Date of Change
        
        K = K + 1

        rsWRK.MoveNext
    Loop
    rsWRK.Close
    
    If Not exBook Is Nothing Then
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
    End If

    Call Pause(1)
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    MDIMain.panHelp(2).Caption = ""
    
    'Ticket #28815 - Having trouble opening the file.
    ShellExecute Me.hwnd, "open", xlsFileMat, vbNullString, vbNullString, SW_SHOW
    
    'If Not LanchXlsW98(xlsFileMat) Then
    '    Shell "cmd /c " & GetShortName(xlsFileMat)
    'End If
    
End Sub

