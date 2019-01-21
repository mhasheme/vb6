VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmDIVISIONS 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Division"
   ClientHeight    =   8520
   ClientLeft      =   1125
   ClientTop       =   795
   ClientWidth     =   12450
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8520
   ScaleWidth      =   12450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkWIA 
      Alignment       =   1  'Right Justify
      Caption         =   "WAI"
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
      Left            =   5880
      TabIndex        =   44
      Top             =   6120
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txtDeptBonusCtr 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      DataField       =   "DV_BONUSDEPT"
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
      Left            =   5280
      MaxLength       =   8
      TabIndex        =   40
      Tag             =   "00-Bonus Reporting #"
      Top             =   5400
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CheckBox chkInactiveCode 
      Alignment       =   1  'Right Justify
      Caption         =   "Inactive Code"
      DataField       =   "DV_INACTIVE"
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
      Left            =   120
      TabIndex        =   39
      Top             =   7200
      Width           =   1395
   End
   Begin VB.TextBox txtCountry 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "DV_COUNTRY"
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
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   36
      Tag             =   "01-Country"
      Top             =   5040
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox cmbCountry 
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
      Left            =   1880
      TabIndex        =   8
      Tag             =   "40-Country Code"
      Top             =   5040
      Width           =   1575
   End
   Begin INFOHR_Controls.CodeLookup clpLGroup 
      DataField       =   "LOCGROUP"
      Height          =   285
      Left            =   6240
      TabIndex        =   3
      Top             =   3240
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBLC"
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      DataField       =   "Division_Name"
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
      Left            =   3000
      MaxLength       =   25
      TabIndex        =   2
      Tag             =   "01-Description of Code"
      Top             =   3240
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7320
      Top             =   6960
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
      TabIndex        =   19
      Top             =   7860
      Width           =   12450
      _Version        =   65536
      _ExtentX        =   21960
      _ExtentY        =   1164
      _StockProps     =   15
      ForeColor       =   0
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
      Begin VB.CommandButton CmdRRGT1 
         Appearance      =   0  'Flat
         Caption         =   "Update Organization 1"
         Height          =   375
         Left            =   9840
         TabIndex        =   42
         Tag             =   "Recalculate for all employees"
         Top             =   105
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.CommandButton CmdRecalcRA4 
         Appearance      =   0  'Flat
         Caption         =   "Update Rept. Authority 4"
         Height          =   375
         Left            =   7200
         TabIndex        =   41
         Tag             =   "Recalculate for all employees"
         Top             =   105
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
         Height          =   375
         Left            =   15
         TabIndex        =   20
         Tag             =   "Select this Division"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   855
         TabIndex        =   21
         Tag             =   "Close and exit this screen"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Tag             =   "Edit the information "
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   23
         Tag             =   "Save changes made"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   24
         Tag             =   "Cancel changes made"
         Top             =   105
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   4260
         TabIndex        =   25
         Tag             =   "Create a new Division"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5070
         TabIndex        =   26
         Tag             =   "Delete Division listed"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5895
         TabIndex        =   27
         Tag             =   "Print Division Listing"
         Top             =   105
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   1935
         Top             =   30
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowTitle     =   "Department Codes"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   6255
      TabIndex        =   15
      Tag             =   "Find specific record"
      Top             =   6765
      Width           =   720
   End
   Begin VB.TextBox txtFindDesc 
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
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Tag             =   "00-Search Description"
      Top             =   6810
      Width           =   3165
   End
   Begin VB.TextBox txtFindKey 
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
      Height          =   285
      Left            =   1875
      MaxLength       =   4
      TabIndex        =   13
      Tag             =   "00-Search Division"
      Top             =   6810
      Width           =   1080
   End
   Begin VB.TextBox txtDiv 
      Appearance      =   0  'Flat
      DataField       =   "DIV"
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
      Height          =   285
      Left            =   1880
      MaxLength       =   4
      TabIndex        =   1
      Tag             =   "01-Division's Code"
      Top             =   3240
      Width           =   1065
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LDate"
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
      Left            =   3000
      MaxLength       =   25
      TabIndex        =   16
      Text            =   "Ldate"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LTime"
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
      Left            =   4680
      MaxLength       =   25
      TabIndex        =   17
      Text            =   "LTime"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LUser"
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
      Left            =   6360
      MaxLength       =   25
      TabIndex        =   18
      Text            =   "LUser"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1590
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxdiv.frx":0000
      Height          =   3075
      Left            =   120
      OleObjectBlob   =   "fxdiv.frx":0014
      TabIndex        =   0
      Tag             =   "Division Listings"
      Top             =   0
      Width           =   10575
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "DV_ADMINBY"
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   7
      Tag             =   "00-Administered By"
      Top             =   4680
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "DV_REGION"
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   5
      Tag             =   "00-Region - Code"
      Top             =   3960
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "DV_SECTION"
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   6
      Tag             =   "00-Section - Code"
      Top             =   4320
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "DV_LOC"
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Tag             =   "00-Location - Code"
      Top             =   3600
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "DV_MARKETLINE"
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   10
      Tag             =   "00-Market Line - Code"
      Top             =   5760
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFML"
   End
   Begin INFOHR_Controls.CodeLookup clpProv 
      DataField       =   "DV_PROV"
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Tag             =   "31-Province - Code"
      Top             =   6100
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   4
   End
   Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   9
      Tag             =   "10-Employee Number of individual's supervisor"
      Top             =   5400
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   8115
      TabIndex        =   11
      Tag             =   "00-Orgranization - Code"
      Top             =   6120
      Visible         =   0   'False
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ORGN"
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization 1"
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
      Left            =   6720
      TabIndex        =   43
      Top             =   6165
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Province"
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
      Left            =   120
      TabIndex        =   38
      Top             =   6160
      Width           =   630
   End
   Begin VB.Label lblCountryTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
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
      TabIndex        =   35
      Top             =   5040
      Width           =   1140
   End
   Begin VB.Label lblMarketLine 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Market Line"
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
      TabIndex        =   34
      Top             =   5760
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label lblDivSearch 
      Caption         =   "Search Division"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   6840
      Width           =   1695
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
      Left            =   120
      TabIndex        =   32
      Top             =   3240
      Width           =   1275
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
      Left            =   120
      TabIndex        =   31
      Top             =   3600
      Width           =   1335
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
      Left            =   120
      TabIndex        =   30
      Top             =   4680
      Width           =   1125
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
      Left            =   120
      TabIndex        =   29
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   4320
      Width           =   1260
   End
   Begin VB.Label lblRptNo 
      Caption         =   "Rept. Authority 4 "
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
      Left            =   120
      TabIndex        =   37
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmDIVISIONS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRSOld As String, glbEmptyNew  As Integer
Dim fglbNewRec% ' new record
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim Ctrl As Control 'Sam add July 2002 * Remove ADO
Dim FRS As ADODB.Recordset

Private Function chkDiv()
Dim Div As String, SQLQ As String, Msg$
Dim snapDivs As New ADODB.Recordset
Dim X
chkDiv = False
On Error GoTo chkDiv_Err

If Len(txtDiv) < 1 Then
    MsgBox lStr("Division Code is a required field")
    txtDiv.SetFocus
    Exit Function
End If

If Len(txtName) < 1 Then
    MsgBox lStr("Division Description is a required field")
    txtName.SetFocus
    Exit Function
End If
If glbLinamar And (Len(txtDiv) <> 3 Or Not IsNumeric(txtDiv)) Then
    MsgBox lStr("Invalid Division")
    If txtDiv.Enabled Then txtDiv.SetFocus
    Exit Function
End If

If Len(clpProv.Text) > 0 And clpProv.Caption = "Unassigned" Then
    MsgBox "Province must be valid"
    clpProv.Text = ""
    clpProv.SetFocus
    Exit Function
End If

If fglbNewRec% Then
    Div = CStr(txtDiv)
    SQLQ = "SELECT DIV from HR_DIVISION "
    SQLQ = SQLQ & "WHERE DIV = '" & Div & "'"
    
    If snapDivs.State <> 0 Then snapDivs.Close
    snapDivs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If snapDivs.BOF And snapDivs.EOF Then
        snapDivs.Close
    Else
        Msg$ = lStr("This Division number already exists")
        MsgBox Msg$
        snapDivs.Close
        Exit Function
    End If
End If

For X = 1 To 5
    If Len(clpCode(X).Text) > 0 And clpCode(X).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(X).SetFocus
        Exit Function
    End If
Next X

If glbWFC Then
    ''If lblDeptBonusDesc = "Unassigned" Then
    ''    MsgBox "Invalid Bonus Reporting # entered."
    ''    Exit Function
    ''End If
    '''Ticket #25884 Franks 08/19/2014
    txtDeptBonusCtr.Text = elpReptAuthShow(3).Text
End If

chkDiv = True

Exit Function

chkDiv_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkDiv", "HR_Div", "Cancel")
Resume Next

End Function

Private Sub clpCode_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbCountry_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbCountry_LostFocus()
txtCountry = cmbCountry
End Sub

Private Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err

rsDATA.CancelUpdate
Call Set_Control("R", Me, rsDATA)


Call modSTUPD(False)  ' reset screen's attributes
fglbNewRec% = False
cmdClose.SetFocus


Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRPROv", "Cancel")
Resume Next

End Sub

Private Sub cmdCancel_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()

glbDiv = ""
glbDivDesc = ""
fglbNewRec% = False

Unload Me

End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdDelete_Click()
Dim Div As String, SQLQ As String, Msg$, a%
Dim snapEEDivs As New ADODB.Recordset

On Error GoTo DelErr

If Len(txtDiv) < 1 Then Exit Sub

Div$ = CStr(txtDiv)

If Data1.Recordset.RecordCount = 1 Then
    MsgBox lStr("You can not delete the last Division.")
    Exit Sub
End If

SQLQ = "SELECT ED_EMPNBR, ED_SURNAME, ED_DEPTNO FROM HREMP "
SQLQ = SQLQ & "WHERE ED_DIV = '" & Div & "'"

If snapEEDivs.State <> 0 Then snapEEDivs.Close
snapEEDivs.Open SQLQ, gdbAdoIhr001, adOpenStatic

If snapEEDivs.BOF And snapEEDivs.EOF Then
    GoTo Lok
Else
    Msg$ = lStr("Employee presently assigned to this Division")
    Msg$ = Msg$ & Chr(10) & ShowEmpnbr(snapEEDivs("ED_EMPNBR"))
    Msg$ = Msg$ & Chr(10) & snapEEDivs("ED_SURNAME")
    Msg$ = Msg$ & Chr(10) & "Delete aborted."
    MsgBox Msg$
    snapEEDivs.Close
    Exit Sub
End If

Lok:    'looks ok to me
snapEEDivs.Close

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True


Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRPROV", "Delete")
Resume Next

End Sub

Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim SQLQ As String

If Len(txtFindKey) > 0 Then
    SQLQ = "DIV = '" & txtFindKey.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
        
        Set FRS = Data1.Recordset.Clone
        vbxTrueGrid.FetchRowStyle = True
        
    Else
        txtFindKey = ""
    End If
    Exit Sub
End If

If Len(txtFindDesc) > 0 Then
    SQLQ = "Division_Name >= '" & txtFindDesc.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
    
        Set FRS = Data1.Recordset.Clone
        vbxTrueGrid.FetchRowStyle = True
    
    Else
        txtFindDesc = ""
    End If
    Exit Sub
End If

End Sub

Private Sub cmdFind_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()

On Error GoTo Mod_Err

Call modSTUPD(True)
txtDiv.Enabled = False
txtName.Enabled = True
txtName.SetFocus
fglbNewRec% = False
'Data1.Recordset.Edit

Exit Sub
Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '08June99 js

End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()

glbCodeRef = True

On Error GoTo NewErr

Call modSTUPD(True)

chkInactiveCode.Value = 0

fglbNewRec% = True

'data1.Recordset.AddNew

''' Sam add July 2002 * Remove ADO
Call Set_Control("B", Me)
rsDATA.AddNew


txtDiv.Enabled = True
txtDiv.SetFocus


Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HRPROV", "AddNew")
Resume Next

End Sub

Private Sub CmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim DivCode, ctylist
Dim xMsg As String
On Error GoTo OK_Err

If Not chkDiv() Then Exit Sub

Call UpdUStats(Me)
DivCode = txtDiv

Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans

If glbWFC Then 'Ticket #21409 Franks 01/11/2012
    If fglbNewRec% Then
        If cmbCountry.Text = "CANADA" Then
            xMsg = "Don’t forget to go into the 'Benefit Group Matrix' to add the Manulife" & Chr(10)
            xMsg = xMsg & "certificate information before assigning employees to this Division."
            MsgBox xMsg
        End If
        If cmbCountry.Text = "U.S.A." Then
            xMsg = "Don’t forget to go into the 'NGS Matrix' in Custom Feature to add the " & Chr(10)
            xMsg = xMsg & "NGS information before assigning employees to this Division."
            MsgBox xMsg
        End If
    End If
End If

Data1.RecordSource = "SELECT * FROM HR_DIVISION ORDER BY DV_INACTIVE, Division_Name"
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Data1.Recordset.Find "DIV='" & DivCode & "'"

ctylist = CountryList

fglbNewRec% = False

Call modSTUPD(False)

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRPROV", "Update")
Resume Next
Unload Me

End Sub

Private Sub cmdOK_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPrint_Click()
Dim RHeading As String, xReport

'RHeading = lStr("Divisions")
'Me.vbxCrystal.WindowTitle = RHeading & " Report"
'Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Action = 1

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup


    RHeading = lStr("Division") & " Listing Report"
    Me.vbxCrystal.WindowTitle = RHeading
    Me.vbxCrystal.BoundReportHeading = RHeading

    xReport = glbIHRREPORTS & "rgdiv.rpt"

    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.Formulas(0) = "lblDivision='" & lStr("Division") & "'"
    'Ticket #23964 Franks 06/25/2013 - begin
    Me.vbxCrystal.Formulas(1) = "lblLocation='" & lStr("Location") & "'"
    Me.vbxCrystal.Formulas(2) = "lblSection='" & lStr("Section") & "'"
    Me.vbxCrystal.Formulas(3) = "lblAdmin='" & lStr("Administered By") & "'"
    Me.vbxCrystal.Formulas(4) = "lblRegion='" & lStr("Region") & "'"
    'Ticket #23964 Franks 06/25/2013 - end
    
    'If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    'Else
    '    Me.vbxCrystal.Connect = "PWD=petman;"
    '    Me.vbxCrystal.DataFiles(0) = glbIHRDB
    'End If

    Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub CmdRecalcRA4_Click()
Dim Msg, Response, DgDef, SQLQ As String

Msg = "This function will update Employee's " & lStr("Rept. Authority 4")
Msg = Msg & " based on the Division Master "
Msg = Msg & Chr(10) & Chr(10) & "Are you sure you want to do this?"
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2

Response = MsgBox(Msg, DgDef, "Update")
If Response = IDNO Then Exit Sub

Msg = "Do you want to update the selected Division or All Divisions?"
Msg = Msg & Chr(10) & Chr(10) & "Click Yes to update the Selected Division"
Msg = Msg & Chr(10) & Chr(10) & "Click No to update All Divisions"
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2

Response = MsgBox(Msg, DgDef, "Update")

Screen.MousePointer = HOURGLASS
If Response = IDNO Then
    Call WFCUptEmployeeRA4("ALL", elpReptAuthShow(3).Text)
Else
    Call WFCUptEmployeeRA4(txtDiv.Text, elpReptAuthShow(3).Text)
End If
Screen.MousePointer = DEFAULT


End Sub

Private Sub CmdRRGT1_Click() 'Ticket #28970 Franks 07/25/2016
Dim Msg, Response, DgDef, SQLQ As String

Msg = "This function will update Employee's " & lStr("Organization 1")
Msg = Msg & " based on the Division Master "
Msg = Msg & Chr(10) & Chr(10) & "Are you sure you want to do this?"
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2

Response = MsgBox(Msg, DgDef, "Update")
If Response = IDNO Then Exit Sub

Msg = "Do you want to update the selected Division or All Divisions?"
Msg = Msg & Chr(10) & Chr(10) & "Click Yes to update the Selected Division"
Msg = Msg & Chr(10) & Chr(10) & "Click No to update All Divisions"
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2

Response = MsgBox(Msg, DgDef, "Update")

Screen.MousePointer = HOURGLASS
If Response = IDNO Then
    Call WFCUptEmployeeOrganization1("ALL", clpCode(6).Text)
Else
    Call WFCUptEmployeeOrganization1(txtDiv.Text, clpCode(6).Text)
End If
Screen.MousePointer = DEFAULT

End Sub

Private Sub cmdSelect_Click()

glbDiv = Data1.Recordset("DIV")
glbDivDesc = Data1.Recordset("Division_Name")
Unload Me

End Sub

Private Sub cmdSelect_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRDiv", "SELECT")

End Sub

Private Sub Form_Activate()
Data1.RecordSource = "SELECT * FROM HR_DIVISION WHERE " & glbSeleDiv & " order by DV_INACTIVE, Division_Name"
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

End Sub

Private Sub Form_Load()
Dim SQLQ, I, ctylist, X

'glbOnTop = "FRMDIVISIONS"
'Data1.DatabaseName = glbIHRDB

Data1.ConnectionString = glbAdoIHRDB

SQLQ = "UPDATE HR_DIVISION SET DV_INACTIVE = 0 WHERE DV_INACTIVE IS NULL"
gdbAdoIhr001.Execute SQLQ

SQLQ = "SELECT * FROM HR_DIVISION WHERE " & glbSeleDiv & " order by DV_INACTIVE, Division_Name"
Data1.RecordSource = SQLQ
Data1.LockType = adLockReadOnly
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Screen.MousePointer = HOURGLASS

'Me.vbxTrueGrid.Refresh

Screen.MousePointer = DEFAULT

Call modSTUPD(False)

If Not gSec_Upd_Divisions Then     'May99 js
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
    'Ticket #25884 Franks 08/20/2014
    CmdRecalcRA4.Enabled = False
End If                          '

ctylist = CountryList
X = 1
Do While X > 0
    X = InStr(ctylist, "&")
    If X > 0 Then
        cmbCountry.AddItem Left(ctylist, X - 1)
        'cmbCountryOfEmp.AddItem Left(ctylist, x - 1)
        ctylist = Mid(ctylist, X + 1)
    Else
        cmbCountry.AddItem ctylist
        'cmbCountryOfEmp.AddItem ctylist
    End If
Loop

cmbCountry.ListIndex = 0        '

Call setCaption(frmDIVISIONS)
Call setCaption(lblDiv)
Call setCaption(lblLocation)
Call setCaption(lblRegion)
Call setCaption(lblSection)
Call setCaption(lblAdmin)
Call setCaption(lblDivSearch)

For I = 0 To 8
    Call setCaption(frmDIVISIONS.vbxTrueGrid.Columns.Item(I))
Next I

If glbWFC Then
    WFCScreenSetup 'Ticket #25884 Franks 08/19/2014
Else
    vbxTrueGrid.Columns(8).Visible = False
    vbxTrueGrid.Columns(9).Visible = False
    vbxTrueGrid.Columns(11).Visible = False
End If

If glbLinamar Then
    clpLGroup.Visible = True
    clpLGroup.TABLTitle = "Location Group"
    vbxTrueGrid.Columns(2).Visible = True
    vbxTrueGrid.Columns(10).Visible = True
    lblTitle(1).Visible = True
    clpProv.Visible = True
Else
    clpLGroup.Visible = False
    vbxTrueGrid.Columns(2).Visible = False
    vbxTrueGrid.Columns(10).Visible = False
    lblTitle(1).Visible = False
    clpProv.Visible = False
End If

Call INI_Controls(Me)

End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub modSTUPD(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

cmdModify.Enabled = FT          'May99 js
cmdFind.Enabled = FT            '
cmdDelete.Enabled = FT          '
cmdNew.Enabled = FT             '
cmdCancel.Enabled = TF          '
cmdOK.Enabled = TF              '
txtDiv.Enabled = TF
clpLGroup.Enabled = TF
vbxTrueGrid.Enabled = FT 'Jaddy 11/12/99
txtFindDesc.Enabled = FT        '
txtFindKey.Enabled = FT         '
txtName.Enabled = TF            '
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
clpCode(4).Enabled = TF
clpCode(5).Enabled = TF
clpCode(6).Enabled = TF
'txtDeptBonusCtr.Enabled = TF
elpReptAuthShow(3).Enabled = TF
cmbCountry.Enabled = TF
clpProv.Enabled = TF
chkInactiveCode.Enabled = TF
chkWIA.Enabled = TF 'Ticket #29069 Franks 08/18/2016

cmdClose.Enabled = FT           '
cmdSelect.Enabled = FT          '
cmdPrint.Enabled = FT           '
        
If glbDivInhSel Then
    cmdSelect.Enabled = False
End If
End Sub

'Private Sub imgIcon_Click()
'Call txtDeptBonusCtr_DblClick
'End Sub

Private Sub txtCountry_Change()
cmbCountry = txtCountry
End Sub

Private Sub txtDeptBonusCtr_Change()
    If IsNumeric(txtDeptBonusCtr.Text) Then
        elpReptAuthShow(3).Text = txtDeptBonusCtr.Text
    Else
        elpReptAuthShow(3).Text = ""
    End If
End Sub

'Private Sub txtDeptBonusCtr_Change()
'    lblDeptBonusDesc = GetBonusRptDesc(txtDeptBonusCtr)
'    If Len(txtDeptBonusCtr) > 0 And Len(txtDeptBonusCtr) > 0 Then
'        If lblDeptBonusDesc = "" Then lblDeptBonusDesc = "Unassigned"
'    End If
'End Sub

'Private Sub txtDeptBonusCtr_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub txtDiv_Change()
If glbLinamar Then
    clpCode(2).TransDiv = txtDiv
    clpCode(3).TransDiv = txtDiv
End If
End Sub

Private Sub txtDiv_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtDIV_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtFindDesc_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtName_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub txtDeptBonusCtr_DblClick()
'''    frmDEPTSBonus.cmdSelect.Enabled = True
'''    glbDeptInhSel% = False
'''    frmDEPTSBonus.Show 1
'''    If Len(frmDEPTSBonus.DeptNbr) > 0 Then
'''        txtDeptBonusCtr = frmDEPTSBonus.DeptNbr
'''        lblDeptBonusDesc = frmDEPTSBonus.DeptDesc
'''    End If
'End Sub

Private Sub vbxTrueGrid_DblClick()
    
If Not Me.vbxTrueGrid.EditActive Then
    glbDiv = Data1.Recordset("DIV")
    glbDivDesc = Data1.Recordset("Division_Name")
    Unload Me
Else
    MsgBox "Save/cancel changes first"
End If

End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If Not fglbNewRec% Then
        FRS.Requery
        FRS.Bookmark = Bookmark
        If FRS("DV_INACTIVE") Then
            RowStyle.ForeColor = vbRed
        End If
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
    
    SQLQ = "select * from HR_DIVISION WHERE " & glbSeleDiv
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
    
    Set FRS = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the enter key was struck
    KeyAscii = 0
    If Me.vbxTrueGrid.EditActive Then
        cmdOK.SetFocus
    Else
        cmdClose.SetFocus
    End If
End If

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

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'''Sam add July 02 * Remove ADO
Call Display_Value
End Sub

''' Sam add July 2002 * Remove ADO
Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        Exit Sub
    End If
  
    SQLQ = "select * from HR_DIVISION WHERE DIV='" & Data1.Recordset!Div & "'" & " order by Division_Name"
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
    
End Sub

Private Sub IsNew74()
    Dim SQLQ
 
    SQLQ = "select * from HR_DIVISION WHERE DIV='" & Data1.Recordset!Div & "'" & " order by Division_Name"
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
End Sub

Function GetBonusRptDesc(TablKey)
    Dim rsTABL As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT * FROM WFC_Bonus_Loc_Department WHERE Dept_No = '" & TablKey & "' "
    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsTABL.EOF And rsTABL.BOF Then
        GetBonusRptDesc = ""
    Else
        GetBonusRptDesc = rsTABL("Dept_Name")
    End If
    rsTABL.Close
End Function

Private Function CountryList() As String
Dim xCountryList As String, ctyFile
xCountryList = ""
ctyFile = glbIHRREPORTS & "CountryList.MTF"

On Error GoTo ErrorHandler

If File(ctyFile) Then
    Open ctyFile For Input As #1
    Input #1, xCountryList
    Close #1
End If

ResumeHere:
'If InStr(xCountryList, BasicCountry) = 0 Then
'    xCountryList = BasicCountry
'End If
If InStr(xCountryList, cmbCountry) = 0 And cmbCountry <> "" Then
    xCountryList = xCountryList & "&" & cmbCountry
    cmbCountry.AddItem cmbCountry
'    comCountryOfEmp.AddItem cmbCountry
End If
Open ctyFile For Output As #1
Print #1, xCountryList
Close #1
CountryList = xCountryList
Exit Function

ErrorHandler:
If Err.Number = 62 Then
    ' Corrupted CountryList.MTF, kill it and regenerate
    Close #1
    MsgBox "Found corrupt CountryList.MTF.  info:HR will re-create this file.", vbInformation + vbOKOnly, "Corrupted Country List"
    Kill ctyFile
    Resume ResumeHere
Else
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number & " in CountryList"
    Resume Next
End If
End Function

Private Sub WFCScreenSetup()
    lblRptNo.Visible = True
    lblRptNo.Caption = lStr("Rept. Authority 4")
    vbxTrueGrid.Columns(8).Caption = lStr("Rept. Authority 4")
    elpReptAuthShow(3).Visible = True
    'imgIcon.Visible = True
    'txtDeptBonusCtr.Visible = True
    'lblDeptBonusDesc.Visible = True
    lblMarketLine.Visible = True
    clpCode(5).Visible = True
    'Ticket #28637 Franks 05/18/2016 - begin
    lblTitle(2).Left = 120
    lblTitle(2).Top = 6160
    lblTitle(2).Visible = True
    clpCode(6).Left = 1560
    clpCode(6).Top = 6100
    clpCode(6).DataField = "DV_ORGT1"
    clpCode(6).Visible = True
    lblTitle(2).Caption = lStr("Organization 1")
    vbxTrueGrid.Columns(11).Caption = lStr("Organization 1")
    'Ticket #28637 Franks 05/18/2016 - end
    If gSec_Upd_Divisions Then
        CmdRecalcRA4.Visible = True
        CmdRecalcRA4.Caption = "Update " & lStr("Rept. Authority 4")
        'Ticket #28970 Franks 07/25/2016
        CmdRRGT1.Visible = True
        CmdRRGT1.Caption = "Update " & lStr("Organization 1")
    End If
    
    chkWIA.Top = clpCode(6).Top
    chkWIA.Visible = True
    chkWIA.DataField = "DV_WIA"
    
End Sub

Private Sub WFCUptEmployeeOrganization1(xCode, xORG1) 'Ticket #28970 Franks 07/25/2016
Dim rsDiv As New ADODB.Recordset
Dim SQLQ As String
Dim xDiv, xUnion, xPayGroup, xStatus, xEmpNo
Dim I, totRec

    If Len(xCode) = 0 Then
        Exit Sub
    End If

    
    If xCode = "ALL" Then
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = "Please Wait..."
        MDIMain.panHelp(2).Caption = ""
        MDIMain.panHelp(0).FloodPercent = 0
        
        SQLQ = "SELECT * FROM HR_DIVISION ORDER BY DIV "
        rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsDiv.EOF Then
            totRec = rsDiv.RecordCount
        End If
        Do While Not rsDiv.EOF
            MDIMain.panHelp(0).FloodPercent = (I / totRec) * 100
            I = I + 1
            xDiv = rsDiv("DIV")
            'MDIMain.panHelp(2).Caption = xCode
            If Not IsNull(rsDiv("DV_ORGT1")) Then
                If Len(rsDiv("DV_ORGT1")) > 0 Then
                    xORG1 = rsDiv("DV_ORGT1")
                    Call WFCUptEmpORG1ForSingleDiv(xDiv, xORG1)
                End If
            End If
            rsDiv.MoveNext
        Loop
        rsDiv.Close
        Screen.MousePointer = vbDefault
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = ""
        MDIMain.panHelp(2).Caption = ""
    Else
        xDiv = xCode
        If Len(xORG1) > 0 Then
            Call WFCUptEmpORG1ForSingleDiv(xDiv, xORG1)
        End If
    End If

    Screen.MousePointer = vbDefault
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    
    MsgBox "   Finished!   "

End Sub

Private Sub WFCUptEmployeeRA4(xCode, xlocEmpNo)
Dim rsDiv As New ADODB.Recordset
Dim SQLQ As String
Dim xDiv, xUnion, xPayGroup, xStatus, xEmpNo
Dim I, totRec

    If Len(xCode) = 0 Then
        Exit Sub
    End If

    
    If xCode = "ALL" Then
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = "Please Wait..."
        MDIMain.panHelp(2).Caption = ""
        MDIMain.panHelp(0).FloodPercent = 0
        
        SQLQ = "SELECT * FROM HR_DIVISION ORDER BY DIV "
        rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsDiv.EOF Then
            totRec = rsDiv.RecordCount
        End If
        Do While Not rsDiv.EOF
            MDIMain.panHelp(0).FloodPercent = (I / totRec) * 100
            I = I + 1
            xDiv = rsDiv("DIV")
            'MDIMain.panHelp(2).Caption = xCode
            If Not IsNull(rsDiv("DV_BONUSDEPT")) Then
                If IsNumeric(rsDiv("DV_BONUSDEPT")) Then
                    xEmpNo = rsDiv("DV_BONUSDEPT")
                    Call WFCUptEmpRA4ForSingleDiv(xDiv, xEmpNo)
                End If
            End If
            rsDiv.MoveNext
        Loop
        rsDiv.Close
        Screen.MousePointer = vbDefault
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = ""
        MDIMain.panHelp(2).Caption = ""
    Else
        xDiv = xCode
        If IsNumeric(xlocEmpNo) Then
            xEmpNo = xlocEmpNo
            Call WFCUptEmpRA4ForSingleDiv(xDiv, xEmpNo)
        End If
    End If
    Screen.MousePointer = vbDefault
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    
    MsgBox "   Finished!   "
End Sub

Private Sub WFCUptEmpORG1ForSingleDiv(xDiv, xORG1)
Dim SQLQ As String
Dim I As Integer
    SQLQ = "UPDATE HREMP SET ED_ORGT1 = '" & xORG1 & "' WHERE ED_DIV = '" & xDiv & "' "
    gdbAdoIhr001.Execute SQLQ, I
    
End Sub

Private Sub WFCUptEmpRA4ForSingleDiv(xDiv, xEmpNo)
Dim SQLQ As String

    SQLQ = "UPDATE HR_JOB_HISTORY SET JH_REPTAU4 = " & xEmpNo & " WHERE NOT JH_CURRENT = 0 "
    SQLQ = SQLQ & "AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_DIV = '" & xDiv & "') "
    gdbAdoIhr001.Execute SQLQ

End Sub
