VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmOvtBankO 
   Appearance      =   0  'Flat
   Caption         =   "Overtime Bank Overview"
   ClientHeight    =   8490
   ClientLeft      =   -45
   ClientTop       =   1365
   ClientWidth     =   11880
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
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid1 
      Bindings        =   "fovtbanko.frx":0000
      Height          =   3195
      Left            =   0
      OleObjectBlob   =   "fovtbanko.frx":0014
      TabIndex        =   32
      Tag             =   "Employee Listing "
      Top             =   120
      Width           =   9870
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   25
      Top             =   7830
      Width           =   11880
      _Version        =   65536
      _ExtentX        =   20955
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
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7080
         TabIndex        =   38
         Tag             =   "Save changes made"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdModify1 
         Appearance      =   0  'Flat
         Caption         =   "&Edit Max. Bank"
         Height          =   375
         Left            =   5280
         TabIndex        =   35
         Tag             =   "Edit information on this screen"
         Top             =   60
         Width           =   1695
      End
      Begin VB.CommandButton CmdRecalc 
         Appearance      =   0  'Flat
         Caption         =   "&Recalculate 1 Employee"
         Height          =   375
         Index           =   0
         Left            =   2760
         TabIndex        =   11
         Tag             =   "Recalculate for the employee"
         Top             =   60
         Width           =   2415
      End
      Begin VB.CommandButton CmdRecalc 
         Appearance      =   0  'Flat
         Caption         =   "R&ecalculate All Employees"
         Height          =   375
         Index           =   1
         Left            =   5280
         TabIndex        =   12
         Tag             =   "Recalculate for all employees"
         Top             =   60
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.CommandButton cmdDays 
         Appearance      =   0  'Flat
         Caption         =   "Da&ys"
         Height          =   375
         Left            =   300
         TabIndex        =   9
         Tag             =   "Display Vacation and Sick Overview in Days"
         Top             =   60
         Width           =   875
      End
      Begin VB.CommandButton cmdHours 
         Appearance      =   0  'Flat
         Caption         =   "&Hours"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1260
         TabIndex        =   10
         Tag             =   "Display Vacation and Sick Overview in Hours"
         Top             =   60
         Width           =   855
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6000
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         BoundReportFooter=   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         GridSource      =   "vbxTrueGrid"
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   10320
      Top             =   8280
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Tag             =   "Find Employee"
      Top             =   5700
      Width           =   735
   End
   Begin VB.TextBox txtEESearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Tag             =   "00-Search for Surname"
      Top             =   5730
      Width           =   1935
   End
   Begin VB.CommandButton cmdEESort 
      Appearance      =   0  'Flat
      Caption         =   "&Sort by Surname"
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   7
      Top             =   5700
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.CommandButton cmdEESort 
      Appearance      =   0  'Flat
      Caption         =   "&Sort by Emp #"
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   8
      Tag             =   "Change the sorting method of the Employee List"
      Top             =   5700
      Width           =   2475
   End
   Begin VB.CommandButton cmdCancel1 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10200
      TabIndex        =   27
      Tag             =   "Cancel changes made"
      Top             =   4800
      Visible         =   0   'False
      Width           =   915
   End
   Begin Threed.SSPanel panDetails 
      Height          =   2055
      Left            =   0
      TabIndex        =   15
      Top             =   3360
      Width           =   9795
      _Version        =   65536
      _ExtentX        =   17277
      _ExtentY        =   3625
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
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin MSMask.MaskEdBox medMaxBank 
         DataField       =   "WK_MBANKDAY"
         Height          =   285
         Index           =   1
         Left            =   6120
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1605
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOvtNotBanked 
         DataField       =   "WK_BANKNTDAY"
         Height          =   285
         Index           =   1
         Left            =   4920
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1605
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medBankT 
         DataField       =   "WK_BANKTDAY"
         Height          =   285
         Index           =   1
         Left            =   2760
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1605
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPBank 
         DataField       =   "WK_PBANKDAY"
         Height          =   285
         Index           =   1
         Left            =   570
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "11-Banked hours of vacation from previous year"
         Top             =   1605
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medCBankDay 
         DataField       =   "WK_BANKDAY"
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   1605
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medBankT 
         DataField       =   "OT_BANKT"
         Height          =   285
         Index           =   0
         Left            =   2760
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1605
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPBank 
         DataField       =   "OT_PBANK"
         Height          =   285
         Index           =   0
         Left            =   570
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "11-Banked hours of vacation from previous year"
         Top             =   1605
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medCBank 
         DataField       =   "OT_BANK"
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "11-Total number of hours vacation time entitled"
         Top             =   1605
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medBankO 
         DataField       =   "WK_BANKODAY"
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1605
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medBankO 
         DataField       =   "WK_BANKO"
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1605
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOvtNotBanked 
         DataField       =   "WK_BANKNT"
         Height          =   285
         Index           =   0
         Left            =   4920
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1605
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMaxBank 
         DataField       =   "OT_MBANK"
         Height          =   285
         Index           =   0
         Left            =   6120
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1605
         Width           =   1005
         _ExtentX        =   1773
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
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Max. Bank"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   6150
         TabIndex        =   34
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Earned"
         Height          =   195
         Left            =   1320
         TabIndex        =   33
         Top             =   1155
         Width           =   615
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Yr Available"
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   5
         Left            =   4905
         TabIndex        =   26
         Top             =   1200
         Width           =   1035
         WordWrap        =   -1  'True
      End
      Begin VB.Label DateSeleV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "OT_EFDATE"
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
         Height          =   285
         Index           =   1
         Left            =   7440
         TabIndex        =   29
         Top             =   1605
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Range"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   7800
         TabIndex        =   23
         Top             =   1380
         Width           =   1515
      End
      Begin VB.Label lblDayHrs 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Overtime Bank Hours"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3960
         TabIndex        =   22
         Top             =   600
         Width           =   2565
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Outstanding"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3840
         TabIndex        =   19
         Top             =   1380
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Taken"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2880
         TabIndex        =   18
         Top             =   1380
         Width           =   555
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Previous"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   14
         Top             =   1380
         Width           =   750
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   17
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label DateSeleV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "OT_ETDATE"
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
         Height          =   285
         Index           =   2
         Left            =   8520
         TabIndex        =   28
         Top             =   1605
         Width           =   1035
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fovtbanko.frx":6755
      Height          =   3195
      Left            =   30
      OleObjectBlob   =   "fovtbanko.frx":6769
      TabIndex        =   31
      Tag             =   "Employee Listing "
      Top             =   120
      Visible         =   0   'False
      Width           =   9810
   End
   Begin VB.Label lblSearchBy 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search by Surname"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   24
      Top             =   5760
      Width           =   1665
   End
End
Attribute VB_Name = "frmOvtBankO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Add a New True DBGRid
Option Explicit
Dim EESNameSort As Integer
Dim OSN As Double, OSCh As String     ' last search items
Dim fglbWDate$, SavEntOpt, SavFdate, SavTdate
Dim fglbWDateS$, SavEntOptS, SAVFDATES, SAVTDATES
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim fglbNew As Integer


'Private Sub cmdCancel1_Click()
'Dim x, xID
'On Error GoTo Can_Err
'
'xID = Data1.Recordset("ED_EMPNBR")
'
'rsDATA.CancelUpdate
'Call Display_Value
'Data1.Refresh
'
'Data1.Recordset.Find "ED_EMPNBR=" & xID
'
'panDetails.Enabled = False
'cmdOK1.Enabled = False
'cmdCancel1.Enabled = False
'cmdModify1.Enabled = True
'dlpFDate1.Visible = False
'dlpTDate1.Visible = False
'
'Exit Sub
'
'Can_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_OVERTIME_BANK", "Cancel")
'End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdCancel_Click()
    cmdModify1.Enabled = True
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdDays_Click()
cmdDays.Enabled = False
cmdHours.Enabled = True
medCBank.Visible = False
medCBankDay.Visible = True
medPBank(0).Visible = False
medPBank(1).Visible = True
medBankT(0).Visible = False
medBankT(1).Visible = True
medBankO(0).Visible = False
medBankO(1).Visible = True
medMaxBank(0).Visible = False
medMaxBank(1).Visible = True
medOvtNotBanked(0).Visible = False
medOvtNotBanked(1).Visible = True
If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
    lblDayHrs.Caption = "Extra Time Bank Days"
Else
    lblDayHrs.Caption = "Overtime Bank Days"
End If
vbxTrueGrid.Visible = True
vbxTrueGrid1.Visible = False

If glbLambton Then
    lblTitle(3).Caption = "Total OT Banked"
    lblTitle(5).Visible = False
    medOvtNotBanked(0).Visible = False
    medOvtNotBanked(1).Visible = False
    lblTitle(2).Left = 2880 + 250
    medBankT(0).Left = 2760 + 200
    medBankT(1).Left = 2760 + 200
    
    lblTitle(3).Left = 3840 + 250
    medBankO(0).Left = 3840 + 500
    medBankO(1).Left = 3840 + 500
    
    lblTitle(4).Left = 6150 - 100
    medMaxBank(0).Left = 6120 - 100
    medMaxBank(1).Left = 6120 - 100
End If

End Sub

Private Sub cmdDays_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdEESort_Click(Index As Integer)

txtEESearch.Text = ""
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Refreshing Employee List - Stand by"
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "

If EESNameSort = True Then  ' was sorted by surname
    EESNameSort = False
    lblSearchBy.Caption = "Search by Emp. #"
    cmdEESort(0).Visible = False
    cmdEESort(1).Visible = True
Else
    EESNameSort = True
    lblSearchBy.Caption = "Search by Surname"
    cmdEESort(0).Visible = True
    cmdEESort(1).Visible = False
End If

If EERetrieve() = 0 Then     ' get the info for this person
    Exit Sub
End If          ' dpartment specific and populate the list

Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).Caption = " "
txtEESearch.SetFocus

End Sub

Private Sub cmdEESort_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim Sch As String, SQLQ As String
Dim bkmark

On Error GoTo Srch_Err

If Not Len(txtEESearch) > 0 Then
   MsgBox "To search you must enter something to search for."
   Exit Sub
End If
Data1.Refresh
If Not Data1.Recordset.EOF Then
    Sch = Replace(txtEESearch.Text, "'", "''")
    If EESNameSort = True Then
        SQLQ = "ED_SURNAME  >= '" & Sch & "'"
    Else
        If Not IsNumeric(txtEESearch.Text) And Not glbLinamar Then
            Beep
            MsgBox "Employee Identification must be numeric"
            Exit Sub
        End If
        If glbLinamar Then
            SQLQ = "EMPNBR >= '" & Sch & "'"
        Else
            SQLQ = "ED_EMPNBR >= '" & Sch & "'"
        End If

    End If
    Data1.Recordset.Find SQLQ
End If
If Data1.Recordset.EOF Then
    MsgBox "Employee not found"
    Data1.Refresh
End If

Exit Sub

Srch_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HR_OVERTIME_BANK", "Find Next")
Call RollBack '28July99 jsEnd Sub
End Sub

Private Sub cmdFind_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdHours_Click()

cmdDays.Enabled = True
cmdHours.Enabled = False
medCBank.Visible = True
medCBankDay.Visible = False
medPBank(0).Visible = True
medPBank(1).Visible = False
medBankT(0).Visible = True
medBankT(1).Visible = False
medBankO(0).Visible = True
medBankO(1).Visible = False
medMaxBank(0).Visible = True
medMaxBank(1).Visible = False
medOvtNotBanked(0).Visible = True
medOvtNotBanked(1).Visible = False
If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
    lblDayHrs.Caption = "Extra Time Bank Hours"
Else
    lblDayHrs.Caption = "Overtime Bank Hours"
End If
vbxTrueGrid.Visible = False
vbxTrueGrid1.Visible = True

If glbLambton Then
    lblTitle(3).Caption = "Total OT Banked"
    lblTitle(5).Visible = False
    medOvtNotBanked(0).Visible = False
    medOvtNotBanked(1).Visible = False
    lblTitle(2).Left = 2880 + 250
    medBankT(0).Left = 2760 + 200
    medBankT(1).Left = 2760 + 200
    
    lblTitle(3).Left = 3840 + 250
    medBankO(0).Left = 3840 + 500
    medBankO(1).Left = 3840 + 500
    
    lblTitle(4).Left = 6150 - 100
    medMaxBank(0).Left = 6120 - 100
    medMaxBank(1).Left = 6120 - 100
End If

End Sub

Private Sub cmdHours_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub cmdModify1_Click()
'Dim xMsg
'    If Not (Data1.Recordset("ED_PT") = "PT" And Data1.Recordset("ED_ORG") = "CUPE") Then
'        xMsg = "You only can edit the Vacation Date Range" & Chr(10)
'        xMsg = xMsg & "when the employee type is Part Time and union is 'CUPE' "
'        MsgBox xMsg
'        Exit Sub
'    End If
'    panDetails.Enabled = True
'    cmdOK1.Enabled = True
'    cmdModify1.Enabled = False
'    cmdCancel1.Enabled = True
'    dlpFDate1.Visible = True
'    dlpFDate1.Top = DateSeleV(1).Top
'    dlpTDate1.Visible = True
'    dlpTDate1.Top = DateSeleV(2).Top
'    dlpFDate1.SetFocus
'End Sub

'Private Sub cmdOK1_Click()
'Dim xID, SQLQ
'Dim rsTA As New ADODB.Recordset
'    If Len(dlpFDate1) > 0 Then
'       If Not IsDate(dlpFDate1) Then
'           MsgBox "Not a valid date"
'           dlpFDate1 = ""
'           dlpFDate1.SetFocus
'           Exit Sub
'       End If
'    Else
'        MsgBox "Vacation From Date is required"
'        dlpFDate1 = ""
'        dlpFDate1.SetFocus
'        Exit Sub
'    End If
'    If Len(dlpTDate1) > 0 Then
'       If Not IsDate(dlpTDate1) Then
'           MsgBox "Not a valid date"
'           dlpTDate1 = ""
'           dlpTDate1.SetFocus
'           Exit Sub
'       End If
'    Else
'        MsgBox "Vacation To Date is required"
'        dlpTDate1 = ""
'        dlpTDate1.SetFocus
'        Exit Sub
'    End If
'
'    xID = Data1.Recordset("ED_EMPNBR")
'    gdbAdoIhr001.Execute "UPDATE HREMP SET ED_VACT=" & 0 & " WHERE ED_EMPNBR=" & Data1.Recordset("ED_EMPNBR")
'
'    SQLQ = "SELECT ED_EMPNBR, Sum(AD_HRS) AS SumHRS"
'    SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
'    SQLQ = SQLQ & " WHERE LEFT(AD_REASON,3)='VAC' "
'    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpFDate1)
'    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpTDate1)
'    SQLQ = SQLQ & " AND HREMP.ED_EMPNBR = " & Data1.Recordset("ED_EMPNBR")
'    SQLQ = SQLQ & " GROUP BY ED_EMPNBR "
'    rsTA.Open SQLQ, gdbAdoIhr001, adOpenKeyset
'    Do Until rsTA.EOF
'        gdbAdoIhr001.Execute "UPDATE HREMP SET ED_VACT=" & rsTA("SUMHRS") & " WHERE ED_EMPNBR=" & rsTA("ED_EMPNBR")
'        rsTA.MoveNext
'    Loop
'    rsTA.Close
'
'    Call Set_Control("U", Me, rsDATA)
'
'    gdbAdoIhr001.BeginTrans
'    rsDATA.Update
'    gdbAdoIhr001.CommitTrans
'    rsDATA.Resync
'    xID = rsDATA!ED_EMPNBR
'
'    Data1.Refresh
'
'    Data1.Recordset.Find "ED_EMPNBR=" & xID
'    panDetails.Enabled = False
'    cmdOK1.Enabled = False
'    cmdCancel1.Enabled = False
'    cmdModify1.Enabled = True
'    dlpFDate1.Visible = False
'    dlpTDate1.Visible = False
'
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String, xReport

'----------\\
    If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
        RHeading = "Employee Extra Time Listing Report"
    Else
        RHeading = "Employee Overtime Listing Report"
    End If
    Me.vbxCrystal.Reset
    Me.vbxCrystal.WindowTitle = RHeading
    Me.vbxCrystal.BoundReportHeading = RHeading
    'Me.vbxCrystal(1).Action = 1
    If cmdDays.Enabled = False Then
        xReport = glbIHRREPORTS & "rgovtbnk1.rpt"
    Else
        xReport = glbIHRREPORTS & "rgovtbnk.rpt"
    End If
    
    Me.vbxCrystal.ReportFileName = xReport
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
    End If
    If EESNameSort = True Then  ' was sorted by surname
        Me.vbxCrystal.SortFields(0) = "+{HREMP.ED_SURNAME}"
        Me.vbxCrystal.SortFields(1) = "+{HREMP.ED_FNAME}"
    Else
        Me.vbxCrystal.SortFields(0) = "+{HREMP.ED_EMPNBR}"
    End If
    
    ' dkostka - 10/18/2001 - Added check for security, used to print for all facilities.
    glbiOneWhere = False
    glbstrSelCri = ""
    glbCri_DeptUN ""
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
    Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
    Me.vbxCrystal.Action = 1
End Sub

Sub cmdView_Click()
Dim RHeading As String, xReport

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

'----------\\
    If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
        RHeading = "Employee Extra Time Listing Report"
    Else
        RHeading = "Employee Overtime Listing Report"
    End If
    Me.vbxCrystal.WindowTitle = RHeading
    Me.vbxCrystal.BoundReportHeading = RHeading
    'Me.vbxCrystal(1).Action = 1
    If cmdDays.Enabled = False Then
        xReport = glbIHRREPORTS & "rgovtbnk1.rpt"
    Else
        xReport = glbIHRREPORTS & "rgovtbnk.rpt"
    End If
    
    Me.vbxCrystal.ReportFileName = xReport
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
    End If
    If EESNameSort = True Then  ' was sorted by surname
        Me.vbxCrystal.SortFields(0) = "+{HREMP.ED_SURNAME}"
        Me.vbxCrystal.SortFields(1) = "+{HREMP.ED_FNAME}"
    Else
        Me.vbxCrystal.SortFields(0) = "+{HREMP.ED_EMPNBR}"
    End If
    
    ' dkostka - 10/18/2001 - Added check for security, used to print for all facilities.
    glbiOneWhere = False
    glbstrSelCri = ""
    glbCri_DeptUN ""
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
    Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.Action = 1
End Sub

Private Sub cmdModify1_Click()
    'Change the to hours before allowing to edit Max Bank hours
    Call cmdHours_Click
    
    panDetails.Enabled = True
    cmdDays.Enabled = False
    medCBank.Enabled = False
    medCBankDay.Enabled = False
    medPBank(0).Enabled = False
    medPBank(1).Enabled = False
    medBankT(0).Enabled = False
    medBankT(1).Enabled = False
    medBankO(0).Enabled = False
    medBankO(1).Enabled = False
    medOvtNotBanked(0).Enabled = False
    medOvtNotBanked(1).Enabled = False
    medMaxBank(0).Enabled = True
    medMaxBank(1).Enabled = False
    DateSeleV(1).Enabled = False
    DateSeleV(2).Enabled = False
    cmdModify1.Enabled = False
    medMaxBank(0).SetFocus
    cmdOK.Enabled = True
    
    vbxTrueGrid1.Enabled = False
    vbxTrueGrid.Enabled = False
    txtEESearch.Enabled = False
    cmdFind.Enabled = False
    cmdEESort(0).Enabled = False
    cmdEESort(1).Enabled = False
    CmdRecalc(0).Enabled = False
End Sub

Public Sub cmdOK_Click()
    Dim rsOvtBank As New ADODB.Recordset
    Dim SQLQ As String
    Dim xEmpnbr
    
    xEmpnbr = Data1.Recordset("OT_EMPNBR")
    SQLQ = "SELECT OT_EMPNBR, OT_MBANK FROM HR_OVERTIME_BANK WHERE OT_EMPNBR = " & Data1.Recordset("OT_EMPNBR")
    rsOvtBank.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsOvtBank.EOF Then
        rsOvtBank("OT_MBANK") = medMaxBank(0).Text
        rsOvtBank.Update
        Data1.Refresh
        
        SQLQ = "OT_EMPNBR = " & xEmpnbr
        Data1.Recordset.Find SQLQ
    End If
    cmdOK.Enabled = False
    cmdModify1.Enabled = True
    
    panDetails.Enabled = False
    cmdDays.Enabled = True
    medCBank.Enabled = True
    medCBankDay.Enabled = True
    medPBank(0).Enabled = True
    medPBank(1).Enabled = True
    medBankT(0).Enabled = True
    medBankT(1).Enabled = True
    medBankO(0).Enabled = True
    medBankO(1).Enabled = True
    medOvtNotBanked(0).Enabled = True
    medOvtNotBanked(1).Enabled = True
    medMaxBank(0).Enabled = True
    medMaxBank(1).Enabled = True
    DateSeleV(1).Enabled = True
    DateSeleV(2).Enabled = True
    
    vbxTrueGrid1.Enabled = True
    vbxTrueGrid.Enabled = True
    txtEESearch.Enabled = True
    cmdFind.Enabled = True
    cmdEESort(0).Enabled = True
    cmdEESort(1).Enabled = True
    CmdRecalc(0).Enabled = True
End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdRecalc_Click(Index As Integer)
Dim Msg, Response, DgDef, SQLQ As String

Msg = "Do you wish to proceed and recalculate "
If Index = 1 Then
    Msg = Msg & "all Employees' "
Else
    Msg = Msg & "the Employee's "
End If
If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
    Msg = Msg & "outstanding Extra Time Bank ?"
Else
    Msg = Msg & "outstanding Overtime Bank ?"
End If
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2

Response = MsgBox(Msg, DgDef, "ReCalculate")
If Response = IDNO Then Exit Sub

Screen.MousePointer = HOURGLASS
If Index = 1 Then
    Call ReCalcOvt("")
Else
    'If glbGuelph Then   ' FOR Guelph-Willington
    '    Call AddFTE(Data1.Recordset("ED_EMPNBR"), "NEW")
    'End If
    SQLQ = "OT_EMPNBR = " & Data1.Recordset("ED_EMPNBR")
    Call ReCalcOvt(SQLQ)
    
    If glbCompSerial = "S/N - 2173W" Then 'Town of Ajax 'Ticket #30402 Franks 08/02/2017
        Call Recalculate_OTBANK_Ajax_AllEmployees(Data1.Recordset("ED_EMPNBR"))
    End If
End If

If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
Screen.MousePointer = DEFAULT

glbENTScreen = True
If Index = 1 Then
    Call Form_Activate
Else
    Data1.Recordset.Find SQLQ
End If

End Sub

Function EERetrieve()
Dim SQLQ As String, Q As QueryDef
'Dim db As Database
Dim countr   As Integer  ' EERetrieve_Snap is definded at form level

EERetrieve = False         ' if not found - no depts

'Hemu - Overtime
'SavEntOpt = glbEntOutStanding$
'
'Select Case glbEntOutStanding$ ' sets field reference for basic 'which date'
'    Case "2": fglbWDate$ = "ED_DOH"
'    Case "3": fglbWDate$ = "ED_SENDTE"
'    Case "4": fglbWDate$ = "ED_LTHIRE"
'    Case "5": fglbWDate$ = "ED_USRDAT1"
'    Case "6": fglbWDate$ = "ED_UNION"
'End Select
'
'SavEntOptS = glbEntOutStandingS$
'
'Select Case glbEntOutStandingS$ ' sets field reference for basic 'which date'
'    Case "2": fglbWDateS$ = "ED_DOH"
'    Case "3": fglbWDateS$ = "ED_SENDTE"
'    Case "4": fglbWDateS$ = "ED_LTHIRE"
'    Case "5": fglbWDateS$ = "ED_USRDAT1"
'    Case "6": fglbWDateS$ = "ED_UNION"
'End Select
'
'If glbEntOutStanding$ > "0" And glbEntOutStanding$ < "7" Then
'    If glbEntOutStanding$ = "1" Then DateSeleV(0).Caption = "Entitlements Date"
'    If glbEntOutStanding$ = "2" Then DateSeleV(0).Caption = lStr("Original Hire Date")
'    If glbEntOutStanding$ = "3" Then DateSeleV(0).Caption = lStr("Seniority Date")
'    If glbEntOutStanding$ = "4" Then DateSeleV(0).Caption = lStr("Last Hire Date")
'    If glbEntOutStanding$ = "5" Then DateSeleV(0).Caption = lStr("User Defined Date")
'    If glbEntOutStanding$ = "6" Then DateSeleV(0).Caption = lStr("Union Date")
'End If
'
'If glbEntOutStandingS$ > "0" And glbEntOutStandingS$ < "7" Then
'    If glbEntOutStandingS$ = "1" Then DateSeleS(0).Caption = "Entitlements Date"
'    If glbEntOutStandingS$ = "2" Then DateSeleS(0).Caption = lStr("Original Hire Date")
'    If glbEntOutStandingS$ = "3" Then DateSeleS(0).Caption = lStr("Seniority Date")
'    If glbEntOutStandingS$ = "4" Then DateSeleS(0).Caption = lStr("Last Hire Date")
'    If glbEntOutStandingS$ = "5" Then DateSeleS(0).Caption = lStr("User Defined Date")
'    If glbEntOutStandingS$ = "6" Then DateSeleS(0).Caption = lStr("Union Date")
'End If
'Hemu - Overtime

SQLQ = "SELECT ED_SURNAME,ED_FNAME,"
If glbLinamar Then
    SQLQ = SQLQ & "ED_REGION AS PROD_LINE,"     'Ticket #14775
    SQLQ = SQLQ & "right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3) AS EMPNBR,"
Else
    If glbOracle Then
        SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
    Else
        SQLQ = SQLQ & "LTRIM(STR(ED_EMPNBR)) AS EMPNBR,"
    End If
End If

SQLQ = SQLQ & "ED_EMPNBR,"
SQLQ = SQLQ & "ED_LDATE,ED_LTIME,ED_LUSER,ED_PT,ED_ORG,"
SQLQ = SQLQ & "OT_EMPNBR, OT_PBANK, OT_BANK, OT_BANKT, OT_EFDATE, OT_ETDATE, OT_MBANK, " 'OM_MAX_BANK_HRS, "
If glbOracle Or glbSQL Then
    SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE OT_PBANK/ED_DHRS END) AS WK_PBANKDAY, "
    SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE OT_BANK/ED_DHRS END) AS WK_BANKDAY, "
    SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE OT_BANKT/ED_DHRS END) AS WK_BANKTDAY, "
    SQLQ = SQLQ & "OT_BANK+OT_PBANK-OT_BANKT AS WK_BANKO, "
    SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(OT_BANK/ED_DHRS,2)+ROUND(OT_PBANK/ED_DHRS,2)-ROUND(OT_BANKT/ED_DHRS,2) END) AS WK_BANKODAY, "
    'SQLQ = SQLQ & "OM_MAX_BANK_HRS-(OT_BANK) AS WK_BANKNT, "
    'SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(OM_MAX_BANK_HRS/ED_DHRS,2)-(ROUND(OT_BANK/ED_DHRS,2)) END) AS WK_BANKNTDAY "
    If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
        SQLQ = SQLQ & "OT_MBANK-(OT_BANK+OT_PBANK-OT_BANKT) AS WK_BANKNT, "
    Else
        SQLQ = SQLQ & "OT_MBANK-(OT_BANK) AS WK_BANKNT, "
    End If
    If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(OT_MBANK/ED_DHRS,2)-(ROUND((OT_BANK+OT_PBANK-OT_BANKT)/ED_DHRS,2)) END) AS WK_BANKNTDAY, "
    Else
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(OT_MBANK/ED_DHRS,2)-(ROUND(OT_BANK/ED_DHRS,2)) END) AS WK_BANKNTDAY, "
    End If
    SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(OT_MBANK/ED_DHRS,2) END) AS WK_MBANKDAY"
Else
    SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[OT_PBANK]/[ED_DHRS]) AS WK_PBANKDAY, "
    SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[OT_BANK]/[ED_DHRS]) AS WK_BANKDAY, "
    SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[OT_BANKT]/[ED_DHRS]) AS WK_BANKTDAY, "
    SQLQ = SQLQ & "[OT_BANK]+[OT_PBANK]-[OT_BANKT] AS WK_BANKO, "
    SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([OT_BANK]+[OT_PBANK]-[OT_BANKT])/[ED_DHRS]) AS WK_BANKODAY, "
    'SQLQ = SQLQ & "[OM_MAX_BANK_HRS]-([OT_BANK]) AS WK_BANKNT, "
    'SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([OM_MAX_BANK_HRS]-([OT_BANK]))/[ED_DHRS]) AS WK_BANKNTDAY "
    SQLQ = SQLQ & "[OT_MBANK]-([OT_BANK]) AS WK_BANKNT, "
    'SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([OT_MBANKS]-([OT_BANK]))/[ED_DHRS]) AS WK_BANKNTDAY, "
    SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([OT_MBANK]-([OT_BANK]))/[ED_DHRS]) AS WK_BANKNTDAY, "
    'commented by Sam as Round was not supported in ACCESS
    'SQLQ = SQLQ & "ROUND([OT_MBANK]/[ED_DHRS],2) AS WK_MBANKDAY"
    SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[OT_MBANK]/[ED_DHRS]) AS WK_MBANKDAY"
End If

'If glbtermopen Then
'    SQLQ = SQLQ & ",TERM_SEQ "
'    SQLQ = SQLQ & " From Term_HREMP "
'Else
    'SQLQ = SQLQ & " From HREMP, HR_OVERTIME_BANK, HR_OVERTIME_MASTER "
    SQLQ = SQLQ & " From HREMP, HR_OVERTIME_BANK "
'End If
'SQLQ = SQLQ & "Where " & glbSeleDeptUn & " AND ED_EMPNBR = OT_EMPNBR AND ED_ORG = OM_ORG"
SQLQ = SQLQ & "Where " & glbSeleDeptUn & " AND ED_EMPNBR = OT_EMPNBR"

If EESNameSort = True Then
    SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME "
Else
    SQLQ = SQLQ & " ORDER BY " & IIf(glbLinamar, "EMPNBR", "ED_EMPNBR")
End If
    
Data1.RecordSource = SQLQ
Data1.Refresh
'If glbtermopen Then
'    If glbTERM_Seq > 0 Then
'        SQLQ = "TERM_SEQ = " & glbTERM_Seq
'        Data1.Recordset.Find SQLQ
'    End If
'Else
    If glbLEE_ID > 0 Then
        SQLQ = "OT_EMPNBR = " & glbLEE_ID
        Data1.Recordset.Find SQLQ
    End If
'End If

If Data1.Recordset.EOF Then
    cmdDays.Enabled = False
    cmdHours.Enabled = False
    CmdRecalc(0).Enabled = False
    cmdModify1.Enabled = False
End If

EERetrieve = True
Exit Function

EERetrieve_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "OvertimeList", "HR_OVERTIME_BANK", "Select")
Call RollBack '28July99 js

End Function

Private Sub CmdRecalc_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
Dim SQLQ

glbOnTop = "FRMOVTBANKO"

If glbENTScreen = True Then
    glbENTScreen = False
    If EERetrieve() = False Then     ' get the info for this person
        Exit Sub
    End If          ' dpartment specific and populate the list
End If
Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMOVTBANKO"
End Sub

Private Sub Form_Load()
Dim SQLQ As String, EEID&, CompNo&
Dim x%

glbOnTop = "FRMOVTBANKO"
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Retrieving Employee List - Stand by"
If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
    Me.Caption = "Extra Time Bank Overview"
End If
'If glbtermopen Then         'Lucy July 5, 2000
'    Data1.ConnectionString = glbAdoIHRAUDIT
'Else
    Data1.ConnectionString = glbAdoIHRDB
'End If

EESNameSort = True  'first sort is by surname
glbENTScreen = True     'Refresh DATA1 ... FORM.ACTIVATE

'If glbCompSerial = "S/N - 2235W" Then   'laura 03/05/98
'    lblDayHrs.Caption = "DAYS"
'    cmdDays.Visible = False
'    cmdHours.Visible = False
'ElseIf glbCompSerial = "S/N - 2236W" Then
'    lblDayHrs.Caption = "DAYS"
'    cmdDays.Visible = False
'    cmdHours.Visible = False
'Else
    If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
        lblDayHrs.Caption = "Extra Time Bank Hours"
    Else
        lblDayHrs.Caption = "Overtime Bank Hours"
    End If
'End If

If glbLambton Then
    lblTitle(3).Caption = "Total OT Banked"
    lblTitle(5).Visible = False
    medOvtNotBanked(0).Visible = False
    medOvtNotBanked(1).Visible = False
    lblTitle(2).Left = 2880 + 250
    medBankT(0).Left = 2760 + 200
    medBankT(1).Left = 2760 + 200
    lblTitle(3).Left = 3840 + 250
    medBankO(0).Left = 3840 + 500
    medBankO(1).Left = 3840 + 500
    lblTitle(4).Left = 6150 - 100
    medMaxBank(0).Left = 6120 - 100
    medMaxBank(1).Left = 6120 - 100
End If

If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
    vbxTrueGrid.Columns(3).Caption = "Prv. Ext. Time"
    vbxTrueGrid.Columns(4).Caption = "Extra Time"
    vbxTrueGrid.Columns(5).Caption = "Ext. Time Taken"
    vbxTrueGrid.Columns(6).Caption = "Ext. Time Outstd."
    vbxTrueGrid1.Columns(3).Caption = "Prv. Ext. Time"
    vbxTrueGrid1.Columns(4).Caption = "Extra Time"
    vbxTrueGrid1.Columns(5).Caption = "Ext. Time Taken"
    vbxTrueGrid1.Columns(6).Caption = "Ext. Time Outstd."
End If

'If glbCompSerial = "S/N - 2262W" Then
'    medPVac(0).Format = "#,##0.0000"
'    medPVac(1).Format = "#,##0.0000"
'    medCVacDay.Format = "#,##0.0000"
'    medCVac.Format = "#,##0.0000"
'    medVacC(0).Format = "#,##0.0000"
'    medVacC(1).Format = "#,##0.0000"
'    medVacR(0).Format = "#,##0.0000"
'    medVacR(1).Format = "#,##0.0000"
'    vbxTrueGrid.Columns(3).NumberFormat = "0.0000"
'    vbxTrueGrid.Columns(4).NumberFormat = "0.0000"
'    vbxTrueGrid.Columns(5).NumberFormat = "0.0000"
'    vbxTrueGrid.Columns(6).NumberFormat = "0.0000"
'    vbxTrueGrid1.Columns(3).NumberFormat = "0.0000"
'    vbxTrueGrid1.Columns(4).NumberFormat = "0.0000"
'    vbxTrueGrid1.Columns(5).NumberFormat = "0.0000"
'    vbxTrueGrid1.Columns(6).NumberFormat = "0.0000"
'End If

'If glbCompSerial = "S/N - 2296W" Then  'For Essex Library
'    cmdModify1.Visible = True
'    cmdOK1.Visible = True
'    cmdCancel1.Visible = True
'End If
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Me.Show
Screen.MousePointer = DEFAULT

If Not gSec_Upd_Entitlements Then  'js
    CmdRecalc(0).Enabled = False
    CmdRecalc(1).Enabled = False
End If
'If glbtermopen Then
'    CmdRecalc(0).Visible = False
'    CmdRecalc(1).Visible = False
'End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)

End Sub

Private Sub medCBank_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPBank_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtEESearch_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_DblClick()

glbLEE_ID = Data1.Recordset("ED_EMPNBR")
glbLEE_FName = Data1.Recordset("ED_FNAME")
glbLEE_SName = Data1.Recordset("ED_SURNAME")
 
If glbLinamar Then
    glbLEE_ProdLine = Mid(Data1.Recordset("PROD_LINE"), 4) & " - " & GetTABLDesc("EDRG", Data1.Recordset("PROD_LINE")) 'Ticket #14775
End If
 
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
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


'Sub AddFTE(xEmpNo, xFLAG)
'    Dim OldFTE, NewFTE, xEFDATE, xETDATE, xNumVac
'    Dim fNewFTE, fOldFTE, FlagOldFTE
'    Dim RsFTEHis As New ADODB.Recordset
'    Dim xDays1, xDays2, xVacDays, xDate1, xDate2, xFDate, xTDate, xHrsDay, xHrsDayN
'    Dim xVacHours, xYear, xNum As Integer, II, J
'    Dim xArray(100, 2)
'    Dim tNewFTE, xNumVacINS, VAC_First
'    Dim RsTempEmp As New ADODB.Recordset
'    Dim RsJobEmp As New ADODB.Recordset
'    Dim SQLQ, xTxtJOB
'    Dim FlagLoop As Boolean
'
'    SQLQ = "Select ED_EMPNBR,ED_VAC,ED_EFDATE,ED_ETDATE from HREMP Where ED_EMPNBR = " & xEmpNo
'    RsTempEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    xEFDATE = ""
'    xETDATE = ""
'    xNumVac = 0
'    If Not RsTempEmp.EOF Then
'        xNumVac = RsTempEmp("ED_VAC")
'        xNumVacINS = RsTempEmp("ED_VAC")
'        xEFDATE = RsTempEmp("ED_EFDATE")
'        xETDATE = RsTempEmp("ED_ETDATE")
'    End If
'    RsTempEmp.Close
'
'    If Len(xEFDATE) = 0 Or Len(xETDATE) = 0 Then
'        Exit Sub
'    End If
'
'    SQLQ = "Select * from HR_JOB_HISTORY Where JH_EMPNBR = " & xEmpNo
'    SQLQ = SQLQ & " ORDER BY JH_SDATE DESC"
'    RsJobEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    If RsJobEmp.EOF Then
'        Exit Sub
'    End If
'
'    SQLQ = "SELECT * FROM FTE_HISTORY WHERE CP_EMPNBR = " & xEmpNo & " "
'    If IsDate(xEFDATE) Then
'    SQLQ = SQLQ & "AND CP_FDATE = " & Date_SQL(xEFDATE)
'    End If
'    If IsDate(xETDATE) Then
'    SQLQ = SQLQ & "AND CP_TDATE = " & Date_SQL(xETDATE)
'    End If
'    SQLQ = SQLQ & "ORDER BY CP_FDATE DESC"
'    RsFTEHis.Open SQLQ, gdbAdoSN2322, adOpenKeyset, adLockOptimistic
'    If RsFTEHis.EOF And xFLAG <> "NEW" Then
'        Exit Sub
'    End If
'
'    If xFLAG = "NEW" Then
'        If xNumVac = 0 Then
'            Exit Sub
'        End If
'        If Not RsFTEHis.EOF Then ' IF CP_VACORIGION EXIST AND CHANGE IN THE SAME YEAR
'            If RsFTEHis("CP_FDATE") = xEFDATE Then
'                xNumVac = RsFTEHis("CP_VACORIGION")
'                GoTo MAIN_DEAL
'            End If
'        End If
'        '' The following shows how to calculate the VAC days at the end of last year
'        '' We always suppose the FTE# is 1.00 at the end of last year
'        ' X is VAC days when FTE# = 1
'        ' VAC_First is the first VAC days before FTE# change
'        ' days1,days2, ... daysn are date range when FTE# change within this year
'        ' VAC_First = X/365 * FTE#1 * days1 + X/365 * FTE#2 * days2 + ... + X/365 * FTE#n * daysn
'        ' X = (VAC_First * 365)/(FTE#1 * days1 + FTE#2 * days2 + ... + FTE#n * daysn)
'        VAC_First = xNumVac
'
'        xDate1 = "**"
'        xFDate = xEFDATE
'        xTDate = xETDATE
'        FlagLoop = True
'        xHrsDayN = 0
'        If RsJobEmp("JH_DHRS") = 0 Then
'            xHrsDayN = 0
'        Else
'            If IsNull(RsJobEmp("JH_DHRS")) Then
'                xHrsDayN = 0
'            Else
'                xHrsDayN = RsJobEmp("JH_DHRS")
'            End If
'        End If
'        If IsNull(RsJobEmp("JH_FTENUM")) Then
'            fNewFTE = 0
'        Else
'            fNewFTE = RsJobEmp("JH_FTENUM")
'        End If
'        RsJobEmp.MoveNext
'        fOldFTE = 0
'        FlagOldFTE = True
'        II = 0
'        Do While (Not RsJobEmp.EOF) And FlagLoop
'            xDate1 = RsJobEmp("JH_SDATE")
'            If FlagOldFTE Then
'                If Not IsNull(RsJobEmp("JH_FTENUM")) Then
'                    fOldFTE = RsJobEmp("JH_FTENUM")
'                End If
'                FlagOldFTE = False
'            End If
'            If CVDate(xDate1) > CVDate(xETDATE) Then
'                GoTo Next_Rec00
'            End If
'            If RsJobEmp("JH_FTENUM") = 0 Then
'                GoTo Next_Rec00
'            End If
'            If IsNull(RsJobEmp("JH_FTENUM")) Then
'                GoTo Next_Rec00
'            End If
'            OldFTE = RsJobEmp("JH_FTENUM")
'
'            If RsJobEmp("JH_DHRS") = 0 Then
'                GoTo Next_Rec00
'            End If
'            If IsNull(RsJobEmp("JH_DHRS")) Then
'                GoTo Next_Rec00
'            End If
'            xHrsDay = RsJobEmp("JH_DHRS")
'
'            If CVDate(xDate1) < CVDate(xEFDATE) Then
'                II = II + 1
'                xArray(II, 1) = DateDiff("d", CVDate(xFDate), CVDate(xTDate)) * OldFTE
'                FlagLoop = False
'            Else
'                II = II + 1
'                xArray(II, 1) = DateDiff("d", CVDate(xDate1), CVDate(xTDate)) * OldFTE
'                xTDate = xDate1 'DateAdd("d", -1, CVDate(xDate1))
'            End If
'
'Next_Rec00:
'            RsJobEmp.MoveNext
'        Loop
'        If IsDate(xDate1) Then
'            If CVDate(xDate1) > CVDate(xEFDATE) Then
'                II = II + 1
'                xArray(II, 1) = DateDiff("d", CVDate(xDate1), CVDate(xTDate)) * OldFTE
'            End If
'        End If
'
'        xVacDays = 0
'        For J = 1 To II
'            xVacDays = xVacDays + xArray(J, 1)
'        Next
'        If xVacDays = 0 Then
'            Exit Sub
'        End If
'        If xHrsDay = 0 Then
'            Exit Sub
'        End If
'        xNumVac = Round((((VAC_First * 365) / (xVacDays)) / xHrsDayN), 0) * xHrsDayN
'
'    End If
'
'
'    '--- Above Got vacation days per year when FTE = 1 (xNumVac)
'MAIN_DEAL:
'    II = 0
'    xDate1 = "**"
'    xFDate = xEFDATE
'    xTDate = xETDATE
'    FlagLoop = True
'    RsJobEmp.MoveFirst
'    Do While (Not RsJobEmp.EOF) And FlagLoop
'        xDate1 = RsJobEmp("JH_SDATE")
'        If CVDate(xDate1) > CVDate(xETDATE) Then
'            GoTo Next_Rec01
'        End If
'        If RsJobEmp("JH_FTENUM") = 0 Then
'            GoTo Next_Rec01
'        End If
'        If IsNull(RsJobEmp("JH_FTENUM")) Then
'            GoTo Next_Rec01
'        End If
'        OldFTE = RsJobEmp("JH_FTENUM")
'
'        If RsJobEmp("JH_DHRS") = 0 Then
'            GoTo Next_Rec01
'        End If
'        If IsNull(RsJobEmp("JH_DHRS")) Then
'            GoTo Next_Rec01
'        End If
'        xHrsDay = RsJobEmp("JH_DHRS")
'
'        If CVDate(xDate1) < CVDate(xEFDATE) Then
'            II = II + 1
'            xArray(II, 1) = DateDiff("d", CVDate(xFDate), CVDate(xTDate))
'            xArray(II, 2) = xArray(II, 1) * Round(((xNumVac * OldFTE) / (365 * xHrsDay)), 3)
'            FlagLoop = False
'        Else
'            II = II + 1
'            xArray(II, 1) = DateDiff("d", CVDate(xDate1), CVDate(xTDate))
'            xArray(II, 2) = xArray(II, 1) * Round(((xNumVac * OldFTE) / (365 * xHrsDay)), 3)
'            xTDate = xDate1 'DateAdd("d", -1, CVDate(xDate1))
'
'        End If
'
'Next_Rec01:
'        RsJobEmp.MoveNext
'    Loop
'
'    xVacDays = 0
'    For J = 1 To II
'        xVacDays = xVacDays + xArray(J, 2)
'    Next
'
'    If xVacDays = 0 Then
'        Exit Sub
'    End If
'    xVacHours = Round(xVacDays, 0) * xHrsDay
'
'    If xVacHours <> xNumVacINS Then
'        gdbAdoIhr001.BeginTrans
'        'Dim RsTempEmp As New ADODB.Recordset
'        SQLQ = "Select ED_EMPNBR,ED_VAC,ED_EFDATE,ED_ETDATE from HREMP Where ED_EMPNBR = " & xEmpNo
'        RsTempEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'
'        If Not RsTempEmp.EOF Then
'            RsTempEmp("ED_VAC") = xVacHours
'            RsTempEmp.Update
'        End If
'        RsTempEmp.Close
'        gdbAdoIhr001.CommitTrans
'
'        If RsFTEHis.EOF Then
'            RsFTEHis.AddNew
'            RsFTEHis("CP_EMPNBR") = xEmpNo
'            RsFTEHis("CP_VACORIGION") = xNumVac
'            RsFTEHis("CP_VACO") = xNumVacINS
'            RsFTEHis("CP_VACN") = xVacHours
'            If fOldFTE > 0 Then
'            RsFTEHis("CP_FTENUMO") = fOldFTE
'            End If
'            If fNewFTE > 0 Then
'            RsFTEHis("CP_FTENUMN") = fNewFTE
'            End If
'            RsFTEHis("CP_FDATE") = CVDate(xEFDATE)
'            RsFTEHis("CP_TDATE") = CVDate(xETDATE)
'            RsFTEHis("CP_LDATE") = Date
'            RsFTEHis("CP_LTIME") = Time$
'            RsFTEHis("CP_LUSER") = glbUserID
'            RsFTEHis.Update
'        End If
'    End If
'    RsFTEHis.Close
'
'    Exit Sub
'
'ExitLin1:
'End Sub

''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
Dim SQLQ
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    'If glbtermopen Then
    '    rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    'Else
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    'End If
    Exit Sub
End If
    
    
SQLQ = Data1.RecordSource
SQLQ = Left(SQLQ, InStr(SQLQ, "ORDER BY") - 1)
SQLQ = SQLQ & " AND ED_EMPNBR= " & Data1.Recordset!ED_EMPNBR
'If glbtermopen Then
'    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
'    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
'Else
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'End If

If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
Call Set_Control("R", Me, rsDATA)

End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT ED_SURNAME,ED_FNAME,"
        If glbLinamar Then
            SQLQ = SQLQ & "ED_REGION AS PROD_LINE,"     'Ticket #14775
            SQLQ = SQLQ & "right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3) AS EMPNBR,"
        Else
            If glbOracle Then
                SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
            Else
                SQLQ = SQLQ & "LTRIM(STR(ED_EMPNBR)) AS EMPNBR,"
            End If
        End If
        
        SQLQ = SQLQ & "ED_EMPNBR,"
        SQLQ = SQLQ & "ED_LDATE,ED_LTIME,ED_LUSER,ED_PT,ED_ORG,"
        SQLQ = SQLQ & "OT_EMPNBR, OT_PBANK, OT_BANK, OT_BANKT, OT_EFDATE, OT_ETDATE, OT_MBANK, " 'OM_MAX_BANK_HRS, "
        If glbOracle Or glbSQL Then
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE OT_PBANK/ED_DHRS END) AS WK_PBANKDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE OT_BANK/ED_DHRS END) AS WK_BANKDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE OT_BANKT/ED_DHRS END) AS WK_BANKTDAY, "
            SQLQ = SQLQ & "OT_BANK+OT_PBANK-OT_BANKT AS WK_BANKO, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(OT_BANK/ED_DHRS,2)+ROUND(OT_PBANK/ED_DHRS,2)-ROUND(OT_BANKT/ED_DHRS,2) END) AS WK_BANKODAY, "
            'SQLQ = SQLQ & "OM_MAX_BANK_HRS-(OT_BANK) AS WK_BANKNT, "
            'SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(OM_MAX_BANK_HRS/ED_DHRS,2)-(ROUND(OT_BANK/ED_DHRS,2)) END) AS WK_BANKNTDAY "
            If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
                SQLQ = SQLQ & "OT_MBANK-(OT_BANK+OT_PBANK-OT_BANKT) AS WK_BANKNT, "
            Else
                SQLQ = SQLQ & "OT_MBANK-(OT_BANK) AS WK_BANKNT, "
            End If
            If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(OT_MBANK/ED_DHRS,2)-(ROUND((OT_BANK+OT_PBANK-OT_BANKT)/ED_DHRS,2)) END) AS WK_BANKNTDAY, "
            Else
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(OT_MBANK/ED_DHRS,2)-(ROUND(OT_BANK/ED_DHRS,2)) END) AS WK_BANKNTDAY, "
            End If
            SQLQ = SQLQ & "ROUND(OT_MBANK/ED_DHRS,2) AS WK_MBANKDAY"
        Else
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[OT_PBANK]/[ED_DHRS]) AS WK_PBANKDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[OT_BANK]/[ED_DHRS]) AS WK_BANKDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[OT_BANKT]/[ED_DHRS]) AS WK_BANKTDAY, "
            SQLQ = SQLQ & "[OT_BANK]+[OT_PBANK]-[OT_BANKT] AS WK_BANKO, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([OT_BANK]+[OT_PBANK]-[OT_BANKT])/[ED_DHRS]) AS WK_BANKODAY, "
            'SQLQ = SQLQ & "[OM_MAX_BANK_HRS]-([OT_BANK]) AS WK_BANKNT, "
            'SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([OM_MAX_BANK_HRS]-([OT_BANK]))/[ED_DHRS]) AS WK_BANKNTDAY "
            SQLQ = SQLQ & "[OT_MBANK]-([OT_BANK]) AS WK_BANKNT, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([OT_MBANK]-([OT_BANK]))/[ED_DHRS]) AS WK_BANKNTDAY "
            SQLQ = SQLQ & "ROUND([OT_MBANK]/[ED_DHRS],2) AS WK_MBANKDAY"
        End If
        SQLQ = SQLQ & " From HREMP, HR_OVERTIME_BANK, HR_OVERTIME_MASTER "
        SQLQ = SQLQ & "Where " & glbSeleDeptUn & " AND ED_EMPNBR = OT_EMPNBR AND ED_ORG = OM_ORG"
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value

End Sub

Private Sub vbxTrueGrid1_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
       
        If vbxTrueGrid1.Tag = "ASC" Then
            vbxTrueGrid1.Tag = "DESC"
        Else
            vbxTrueGrid1.Tag = "ASC"
        End If
        
        SQLQ = "SELECT ED_SURNAME,ED_FNAME,"
        If glbLinamar Then
            SQLQ = SQLQ & "ED_REGION AS PROD_LINE,"     'Ticket #14775
            SQLQ = SQLQ & "right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3) AS EMPNBR,"
        Else
            If glbOracle Then
                SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
            Else
                SQLQ = SQLQ & "LTRIM(STR(ED_EMPNBR)) AS EMPNBR,"
            End If
        End If
        
        SQLQ = SQLQ & "ED_EMPNBR,"
        SQLQ = SQLQ & "ED_LDATE,ED_LTIME,ED_LUSER,ED_PT,ED_ORG,"
        SQLQ = SQLQ & "OT_EMPNBR, OT_PBANK, OT_BANK, OT_BANKT, OT_EFDATE, OT_ETDATE, OT_MBANK, "  'OM_MAX_BANK_HRS, "
        If glbOracle Or glbSQL Then
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE OT_PBANK/ED_DHRS END) AS WK_PBANKDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE OT_BANK/ED_DHRS END) AS WK_BANKDAY, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE OT_BANKT/ED_DHRS END) AS WK_BANKTDAY, "
            SQLQ = SQLQ & "OT_BANK+OT_PBANK-OT_BANKT AS WK_BANKO, "
            SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(OT_BANK/ED_DHRS,2)+ROUND(OT_PBANK/ED_DHRS,2)-ROUND(OT_BANKT/ED_DHRS,2) END) AS WK_BANKODAY, "
            'SQLQ = SQLQ & "OM_MAX_BANK_HRS-(OT_BANK) AS WK_BANKNT, "
            'SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(OM_MAX_BANK_HRS/ED_DHRS,2)-(ROUND(OT_BANK/ED_DHRS,2)) END) AS WK_BANKNTDAY "
            If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
                SQLQ = SQLQ & "OT_MBANK-(OT_BANK+OT_PBANK-OT_BANKT) AS WK_BANKNT, "
            Else
                SQLQ = SQLQ & "OT_MBANK-(OT_BANK) AS WK_BANKNT, "
            End If
            If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(OT_MBANK/ED_DHRS,2)-(ROUND((OT_BANK+OT_PBANK-OT_BANKT)/ED_DHRS,2)) END) AS WK_BANKNTDAY, "
            Else
                SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND(OT_MBANK/ED_DHRS,2)-(ROUND(OT_BANK/ED_DHRS,2)) END) AS WK_BANKNTDAY, "
            End If
            SQLQ = SQLQ & "ROUND(OT_MBANK/ED_DHRS,2) AS WK_MBANKDAY"
        Else
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[OT_PBANK]/[ED_DHRS]) AS WK_PBANKDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[OT_BANK]/[ED_DHRS]) AS WK_BANKDAY, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,[OT_BANKT]/[ED_DHRS]) AS WK_BANKTDAY, "
            SQLQ = SQLQ & "[OT_BANK]+[OT_PBANK]-[OT_BANKT] AS WK_BANKO, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([OT_BANK]+[OT_PBANK]-[OT_BANKT])/[ED_DHRS]) AS WK_BANKODAY, "
            'SQLQ = SQLQ & "[OM_MAX_BANK_HRS]-([OT_BANK]) AS WK_BANKNT, "
            'SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([OM_MAX_BANK_HRS]-([OT_BANK]))/[ED_DHRS]) AS WK_BANKNTDAY "
            SQLQ = SQLQ & "[OT_MBANK]-([OT_BANK]) AS WK_BANKNT, "
            SQLQ = SQLQ & "iif([ED_DHRS]=0,0,([OT_MBANK]-([OT_BANK]))/[ED_DHRS]) AS WK_BANKNTDAY "
            SQLQ = SQLQ & "ROUND([OT_MBANK]/[ED_DHRS],2) AS WK_MBANKDAY"
        End If
        'SQLQ = SQLQ & " From HREMP, HR_OVERTIME_BANK, HR_OVERTIME_MASTER "
        'SQLQ = SQLQ & "Where " & glbSeleDeptUn & " AND ED_EMPNBR = OT_EMPNBR AND ED_ORG = OM_ORG"
        SQLQ = SQLQ & " From HREMP, HR_OVERTIME_BANK "
        SQLQ = SQLQ & "Where " & glbSeleDeptUn & " AND ED_EMPNBR = OT_EMPNBR "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid1.Columns(ColIndex).DataField & " " & vbxTrueGrid1.Tag
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Sub vbxTrueGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
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
UpdateRight = gSec_Upd_Ovt_Overview
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = True
End Property
Public Property Get Deleteble() As Boolean
Deleteble = False
End Property
Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
    UpdateState = OPENING
    TF = True
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
End Sub


