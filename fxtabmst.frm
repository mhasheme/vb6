VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTABLMASTER 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Table Master"
   ClientHeight    =   7890
   ClientLeft      =   1080
   ClientTop       =   1050
   ClientWidth     =   9405
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
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7890
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkWorkSchedule 
      Alignment       =   1  'Right Justify
      Caption         =   "Work Schedule "
      DataField       =   "TB_WORKSCHED"
      DataSource      =   "Data1"
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
      Left            =   2640
      TabIndex        =   43
      Top             =   5040
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.TextBox txtWarnPrd 
      Appearance      =   0  'Flat
      DataField       =   "TB_USR1"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   6480
      MaxLength       =   64
      TabIndex        =   20
      Tag             =   "00-Warning Period in Days"
      Top             =   5200
      Visible         =   0   'False
      Width           =   850
   End
   Begin VB.TextBox txtUSR1 
      Appearance      =   0  'Flat
      DataSource      =   "Data1"
      Height          =   285
      Left            =   7800
      MaxLength       =   10
      TabIndex        =   39
      Tag             =   "00-Point"
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox chkUnion 
      Caption         =   "Union"
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
      Left            =   5400
      TabIndex        =   37
      Top             =   4275
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CheckBox chkEDEM 
      Alignment       =   1  'Right Justify
      Caption         =   "Employment Status Only"
      DataSource      =   "Data1"
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
      Left            =   3960
      TabIndex        =   36
      Top             =   6240
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.CheckBox chkWFCWPS 
      Alignment       =   1  'Right Justify
      Caption         =   "Course Codes Only"
      DataSource      =   "Data1"
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
      Left            =   3960
      TabIndex        =   35
      Top             =   5925
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.TextBox txtLEPoint 
      Appearance      =   0  'Flat
      DataSource      =   "Data1"
      Height          =   285
      Left            =   8400
      MaxLength       =   64
      TabIndex        =   19
      Tag             =   "00-L/LE Point"
      Top             =   4860
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox chkLE 
      Alignment       =   1  'Right Justify
      Caption         =   "L/LE Flag"
      DataSource      =   "Data1"
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
      Left            =   7680
      TabIndex        =   17
      Top             =   4530
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CheckBox chkAbsence 
      Alignment       =   1  'Right Justify
      Caption         =   "Absent"
      DataField       =   "TB_ABSENCE"
      DataSource      =   "Data1"
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
      Left            =   7680
      TabIndex        =   15
      Top             =   4200
      Width           =   1035
   End
   Begin VB.TextBox txtDWM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "TB_USR1"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3570
      MaxLength       =   7
      TabIndex        =   31
      Tag             =   "01-Department - Code"
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.ComboBox cmbDWM 
      Height          =   315
      ItemData        =   "fxtabmst.frx":0000
      Left            =   2190
      List            =   "fxtabmst.frx":000D
      TabIndex        =   12
      Tag             =   "40-Select Day, Week or Month"
      Top             =   4620
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtWaitPeriod 
      Appearance      =   0  'Flat
      DataField       =   "TB_USR2"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1320
      MaxLength       =   7
      TabIndex        =   11
      Tag             =   "00-Waiting Period"
      Top             =   4650
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CheckBox ChkInc 
      Alignment       =   1  'Right Justify
      Caption         =   "Incentive"
      DataField       =   "TB_INDICATOR"
      DataSource      =   "Data1"
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
      Left            =   5250
      TabIndex        =   13
      Top             =   4170
      Width           =   1215
   End
   Begin VB.CheckBox chkSen 
      Alignment       =   1  'Right Justify
      Caption         =   "Seniority"
      DataField       =   "TB_SEN"
      DataSource      =   "Data1"
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
      Left            =   6600
      TabIndex        =   14
      Top             =   4200
      Width           =   915
   End
   Begin VB.CheckBox chkEMELEA 
      Alignment       =   1  'Right Justify
      Caption         =   "Emergency Leave  "
      DataField       =   "TB_USR3"
      DataSource      =   "Data1"
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
      Left            =   5280
      TabIndex        =   16
      Top             =   4530
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.TextBox txtPoint 
      Appearance      =   0  'Flat
      DataField       =   "TB_USR2"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   6480
      MaxLength       =   64
      TabIndex        =   18
      Tag             =   "00-Point"
      Top             =   4860
      Width           =   855
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxtabmst.frx":002C
      Height          =   3975
      Left            =   120
      OleObjectBlob   =   "fxtabmst.frx":0040
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7320
      Top             =   6720
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
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
      TabIndex        =   22
      Top             =   7230
      Width           =   9405
      _Version        =   65536
      _ExtentX        =   16589
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
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Tag             =   "Select the Code listed above"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   6300
         TabIndex        =   29
         Tag             =   "Print Code Listing Report"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5445
         TabIndex        =   28
         Tag             =   "Delete code listed above"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   4575
         TabIndex        =   27
         Tag             =   "Add a new Code"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3540
         TabIndex        =   26
         Tag             =   "Cancel the changes made"
         Top             =   135
         Width           =   915
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2700
         TabIndex        =   25
         Tag             =   "Save the changes made"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1860
         TabIndex        =   24
         Tag             =   "Edit the Information"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   1020
         TabIndex        =   23
         Tag             =   "Close and exit this screen"
         Top             =   135
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6210
         Top             =   15
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   1
         DiscardSavedData=   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      DataField       =   "TB_COMPNO"
      DataSource      =   "Data1"
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
      Left            =   7920
      MaxLength       =   3
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5790
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      DataField       =   "TB_DESC"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   10
      Tag             =   "01-Description of the Code"
      Top             =   4290
      Width           =   3375
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   6450
      TabIndex        =   9
      Tag             =   "Find specific record"
      Top             =   5565
      Width           =   735
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
      Left            =   1800
      TabIndex        =   8
      Tag             =   "00-Search Description"
      Top             =   5610
      Width           =   4425
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
      Left            =   900
      MaxLength       =   8
      TabIndex        =   7
      Tag             =   "00-Search Code"
      Top             =   5610
      Width           =   855
   End
   Begin VB.TextBox txtFindTabl 
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
      Left            =   120
      MaxLength       =   4
      TabIndex        =   6
      Tag             =   "00-Search Table"
      Top             =   5610
      Width           =   750
   End
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      DataField       =   "TB_KEY"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   900
      MaxLength       =   8
      TabIndex        =   5
      Tag             =   "01-Code"
      Top             =   4290
      Width           =   855
   End
   Begin VB.TextBox txtTable 
      Appearance      =   0  'Flat
      DataField       =   "TB_NAME"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   4
      Tag             =   "01-Table / Category of Code"
      Top             =   4290
      Width           =   735
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TB_LUSER"
      DataSource      =   "Data1"
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
      Left            =   3240
      MaxLength       =   25
      TabIndex        =   3
      Text            =   "LUser"
      Top             =   6705
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TB_LTIME"
      DataSource      =   "Data1"
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
      Left            =   1680
      MaxLength       =   25
      TabIndex        =   2
      Text            =   "LTime"
      Top             =   6705
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TB_LDATE"
      DataSource      =   "Data1"
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
      Left            =   120
      MaxLength       =   25
      TabIndex        =   1
      Text            =   "Ldate"
      Top             =   6705
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CheckBox chkInactiveCode 
      Alignment       =   1  'Right Justify
      Caption         =   "Inactive Code"
      DataField       =   "TB_INACTIVE"
      DataSource      =   "Data1"
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
      TabIndex        =   34
      Top             =   5925
      Width           =   1395
   End
   Begin VB.Label lblDays 
      AutoSize        =   -1  'True
      Caption         =   "Days"
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
      Left            =   7380
      TabIndex        =   42
      Top             =   5245
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblWarnPrd 
      AutoSize        =   -1  'True
      Caption         =   "Warning Period"
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
      Left            =   5280
      TabIndex        =   41
      Top             =   5245
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblUsr1 
      Caption         =   "User 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6600
      TabIndex        =   40
      Top             =   6270
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblLEPoint 
      Caption         =   "L/LE Point"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7440
      TabIndex        =   33
      Top             =   4890
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting Period"
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
      Index           =   0
      Left            =   120
      TabIndex        =   32
      Top             =   4680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblPoint 
      Caption         =   "Point"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5280
      TabIndex        =   30
      Top             =   4890
      Width           =   1125
   End
End
Attribute VB_Name = "frmTABLMASTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNewRec%
Dim fglbUDMode As Integer ', glbEmptyNew As Integer
Dim fglbRSOld As String
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim fRS As ADODB.Recordset
Dim xLinSDLBFlag As Boolean 'Ticket #28846 Franks 08/19/2016

Private Function chkMastTable()
Dim SQLQ As String, Msg$, Tabl As String, Ky As String
Dim snapTabs As New ADODB.Recordset


On Error GoTo chkMastTable_Err

chkMastTable = False
If Len(txtTable) < 1 Then
    MsgBox "Table is a required field"
    txtTable.SetFocus
    Exit Function
End If

If Len(txtKey) < 1 Then
    MsgBox "Key (or Code) is a required field"
    txtKey.SetFocus
    Exit Function
End If

If Len(txtDesc) < 1 Then
    MsgBox "Description is a required field"
    txtDesc.SetFocus
    Exit Function
End If
If Len(txtWaitPeriod) > 0 And txtWaitPeriod.Visible Then
    If IsNumeric(txtWaitPeriod) Then
        If cmbDWM.ListIndex = -1 Then
            MsgBox "Please select Day/Week/Month"
            cmbDWM.SetFocus
            Exit Function
        End If
    Else
        MsgBox "Waiting Period must be numeric"
        txtWaitPeriod.SetFocus
        Exit Function
    End If
End If
If fglbNewRec Then
    Tabl = txtTable
    Ky = txtKey
    SQLQ = "SELECT TB_NAME, TB_KEY from HRTABL "
    SQLQ = SQLQ & "WHERE TB_NAME = '" & Tabl & "' "
    SQLQ = SQLQ & " AND TB_KEY = '" & Ky & "'"
    
    snapTabs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If snapTabs.BOF And snapTabs.EOF Then
        snapTabs.Close
    Else
        Msg$ = "This Table Reference already exists"
        MsgBox Msg$
        snapTabs.Close
        Exit Function
    End If
End If

If txtPoint.Visible Then 'Ticket #19955
    If Len(txtPoint.Text) = 0 Then
        txtPoint.Text = 0
    End If
End If

'Ticket #23636 - Warning Period for ADRE - WHSC
If glbCompSerial = "S/N - 2448W" Then
    If txtTable = "ADRE" Then
        If Len(txtWarnPrd.Text) > 0 Then
            If Not IsNumeric(txtWarnPrd.Text) Then
                MsgBox "Warning Period must be numeric"
                txtWarnPrd.SetFocus
                Exit Function
            ElseIf Val(txtWarnPrd.Text) < 0 Then
                MsgBox "Warning Period must be whole positive number"
                txtWarnPrd.SetFocus
                Exit Function
            ElseIf Val(txtWarnPrd.Text) - Int(txtWarnPrd.Text) > 0 Then
                MsgBox "Warning Period must be whole positive number"
                txtWarnPrd.SetFocus
                Exit Function
            End If
        End If
    End If
End If

chkMastTable = True

Exit Function

chkMastTable_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HRTABLE", "HRTABL", "Cancel")
Call RollBack '10June99 js

End Function

Private Sub chkEDEM_Click()
Dim SQLQ As String
    If chkEDEM.Value Then
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'EDEM' AND TB_NAME IN (SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) ORDER BY TB_NAME,TB_KEY"
        chkWFCWPS.Value = False
    Else
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME IN (SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) ORDER BY TB_NAME,TB_KEY"
    End If
    Data1.RecordSource = SQLQ
    Data1.Refresh
    
    Set fRS = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True
    
End Sub

Private Sub chkWFCWPS_Click()
Dim SQLQ As String
    If chkWFCWPS.Value Then
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ESCD' AND TB_NAME IN (SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) ORDER BY TB_NAME,TB_KEY"
        chkEDEM.Value = False
    Else
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME IN (SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) ORDER BY TB_NAME,TB_KEY"
    
    End If
    Data1.RecordSource = SQLQ
    Data1.Refresh
    
    Set fRS = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True
    
End Sub

Private Sub cmbDWM_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdCancel_Click()

On Error GoTo Can_Err
Dim bk

'Data1.UpdateControls
Data1.Recordset.CancelUpdate

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

Data1.Refresh

Set fRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Call ST_UPD_MODE(False)  ' reset screen's attributes

fglbNewRec% = False
If xLinSDLBFlag Then 'Ticket #28846 Franks 08/19/2016
    'do not show this field
Else
    txtFindTabl.Visible = True
End If
txtFindKey.Visible = True
txtFindDesc.Visible = True
cmdFind.Visible = True

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRTABL", "Cancel")
Call RollBack '10June99 js

End Sub

Private Sub cmdCancel_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdDelete_Click()
    Dim a%, Msg
    Dim SQLQ As String

    On Error GoTo DelErr

    Msg = "Are You Sure You Want To Delete "
    Msg = Msg & "This Record?"
    a% = MsgBox(Msg, 36, "Confirm Delete")
    If a% <> 6 Then Exit Sub
    
    If Data1.Recordset.RecordCount < 2 Then
        MsgBox "You can not delete the last reference for this code"
    Else
        Call Codes_Master_Integration(txtTable, txtKey, , True)
        
        'For All - Ticket #19104
        'Delete the Attendance Code from HRATT_MATRIX table as well - used for Custom reports
        If txtTable = "ADRE" Then
            SQLQ = "DELETE FROM HRATT_MATRIX WHERE AM_REASON = '" & txtKey & "'"
            gdbAdoIhr001.Execute SQLQ
        End If
        
        'Ticket #19358 - the Course Code Master record remains hanging - delete it as well.
        If txtTable = "ESCD" Then
            SQLQ = "DELETE FROM HR_COURSECODE_MASTER WHERE ES_CRSCODE = '" & txtKey & "'"
            gdbAdoIhr001.Execute SQLQ
        End If
        
        Data1.Recordset.Delete
        
        If Not glbSQL And Not glbOracle Then Call Pause(0.5)
        Data1.Refresh
        
        Set fRS = Data1.Recordset.Clone
        vbxTrueGrid.FetchRowStyle = True
        
    End If


Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRTable", "Delete")
Call RollBack '10June99 js

End Sub

Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim SQLQ As String

    'Hemu - 05/29/2003 Begin - Ticket # 4204
    If glbCompSerial = "S/N - 2161W" Then
        'Since the txtKey maxlength is changing on Find to 4
        txtKey.MaxLength = 8
    End If
    'Hemu - 05/29/2003 End
    
If Len(txtFindTabl) > 0 Then
    SQLQ = "TB_NAME >= '" & txtFindTabl.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
        
        Set fRS = Data1.Recordset.Clone
        vbxTrueGrid.FetchRowStyle = True
        
    Else
        txtFindTabl = ""
    End If
    Exit Sub
End If

If Len(txtFindKey) > 0 Then
    SQLQ = "TB_KEY >= '" & txtFindKey.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
    
        Set fRS = Data1.Recordset.Clone
        vbxTrueGrid.FetchRowStyle = True
    
    Else
        txtFindKey = ""
    End If
    Exit Sub
End If

If Len(txtFindDesc) > 0 Then
    SQLQ = "TB_DESC like '" & txtFindDesc.Text & "%'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
    
        Set fRS = Data1.Recordset.Clone
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

Call ST_UPD_MODE(True)

txtTable.Enabled = False
txtKey.Enabled = False
txtDesc.Enabled = True
txtDesc.SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '10June99 js

End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()

On Error GoTo NewErr

glbCodeRef = True   'global call for code refresh
fglbNewRec% = True

Call ST_UPD_MODE(True)

Data1.Recordset.AddNew

chkEMELEA.Value = False
txtComp.Text = glbCompNo
chkInactiveCode.Value = 0
'fglbNewRec% = True

If xLinSDLBFlag Then 'Ticket #28846 Franks 08/19/2016
    txtKey.SetFocus
Else
    txtTable.SetFocus
End If

txtFindTabl.Visible = False
txtFindKey.Visible = False
txtFindDesc.Visible = False
cmdFind.Visible = False
If glbBurlTech Then
    chkLE.Value = False
End If

If glbLinamar Then 'Ticket #28846 Franks 08/17/2016
    If Len(glbTabNam) > 0 Then
        txtTable.Text = glbTabNam
    End If
End If

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HRTABLE", "HRTABL", "add new")
Call RollBack '10June99 js

End Sub

Private Sub CmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim strT, strK, SQLQ  As String
Dim bk
Dim rsCrsCodeMst As New ADODB.Recordset

On Error GoTo OK_Err

If Not chkMastTable() Then Exit Sub

glbCodeRef = True   'table entrie modified/added - forces refresh
                    ' at form level of codes/descriptions.
Call UpdUStats(Me)

strT = txtTable
strK = txtKey

If cmbDWM.Visible Then
    txtDWM = Left(cmbDWM, 1)
    If cmbDWM.ListIndex <> -1 And Len(txtWaitPeriod) = 0 Then
        txtWaitPeriod = 0
    End If
    If txtWaitPeriod = "" And txtWaitPeriod.DataChanged Then txtWaitPeriod.DataChanged = False: Data1.Recordset("TB_USR2") = Null
End If
Data1.Recordset("TB_NAME") = txtTable & ""

'Ticket #20038 - Mainly for WSIB Form 7 to indicate which Union Code is actually a Union
If txtTable = "EDOR" And chkUnion.Visible = True Then
    Data1.Recordset("TB_USR1") = IIf(chkUnion.Value = 1, "1", "0")
End If

'Ticket #23636 - Warning Period for ADRE - WHSC
If glbCompSerial = "S/N - 2448W" Then
    If txtTable = "ADRE" And txtWarnPrd.Visible = True Then
        'If IsNumeric(txtWarnPrd.Text) Then
            Data1.Recordset("TB_USR1") = Trim(txtWarnPrd.Text)
        'End If
    End If
End If

Data1.Recordset.UpdateBatch

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

Data1.Refresh
Set fRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Data1.Recordset.Filter = "TB_NAME = '" & strT & "' AND TB_KEY = '" & strK & "'"
If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
    bk = Data1.Recordset.Bookmark
    Data1.Recordset.Filter = ""
    Data1.Recordset.Bookmark = bk
Else
    Data1.Recordset.Filter = ""
End If

Call Codes_Master_Integration(txtTable, txtKey)

'Ticket #20840 - Add default Course Code Master record as well if Course Code Master table is there
If strT = "ESCD" And fglbNewRec% = True Then
    SQLQ = "SELECT COUNT(ES_CRSCODE) AS TOT_RECS FROM HR_COURSECODE_MASTER"
    rsCrsCodeMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsCrsCodeMst.EOF Then
        If rsCrsCodeMst("TOT_RECS") > 0 Then
            'Add new default Course Code Master record
            rsCrsCodeMst.Close
            Set rsCrsCodeMst = Nothing
            
            SQLQ = "SELECT * FROM HR_COURSECODE_MASTER WHERE ES_CRSCODE = '" & strK & "' "
            If rsCrsCodeMst.State <> 0 Then rsCrsCodeMst.Close
            rsCrsCodeMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsCrsCodeMst.EOF Then
                rsCrsCodeMst.AddNew
                rsCrsCodeMst("ES_COMPNO") = "001"
                rsCrsCodeMst("ES_CRSCODE") = strK
                'rsCrsCodeMst("ES_CTYPE") = rsCRSCodes("TB_USR1")
                rsCrsCodeMst("ES_STATUS") = 1
                rsCrsCodeMst("ES_CORPONLY") = 0
                
                'Ticket #20840
                rsCrsCodeMst("ES_UNIQUE_FOR_POS") = 0
                rsCrsCodeMst("ES_RENEW_FOLLOWUP") = 99
                rsCrsCodeMst("ES_FLWUP_PRD_DWMY") = "Y"
    
                rsCrsCodeMst("ES_LDATE") = Date
                rsCrsCodeMst("ES_LTIME") = Time$
                rsCrsCodeMst("ES_LUSER") = glbUserID
                rsCrsCodeMst.Update
            
                rsCrsCodeMst.Close
                Set rsCrsCodeMst = Nothing
            End If
        Else
            rsCrsCodeMst.Close
            Set rsCrsCodeMst = Nothing
        End If
    Else
        rsCrsCodeMst.Close
        Set rsCrsCodeMst = Nothing
    End If
End If

fglbNewRec% = False


Call ST_UPD_MODE(False)

cmdClose.SetFocus
txtFindTabl.Visible = True
txtFindKey.Visible = True
txtFindDesc.Visible = True
cmdFind.Visible = True

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRTABL", "Update")
Call RollBack '10June99 js

End Sub

Private Sub cmdOK_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPrint_Click()
'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

Me.vbxCrystal.ReportTitle = "All Table Codes"
Me.vbxCrystal.BoundReportHeading = Me.Caption
Me.vbxCrystal.WindowTitle = Me.Caption & " Report"
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdSelect_Click()

Dim x
If Data1.Recordset.EOF And Data1.Recordset.BOF Then
  Exit Sub
End If

If vbxTrueGrid.SelBookmarks.count <> 0 Then
    If vbxTrueGrid.SelBookmarks.count > 1000 Then
        MsgBox vbxTrueGrid.SelBookmarks.count & " codes are selected" + Chr(10) + " Please make that less than 1000 codes"
        Exit Sub
    End If
    glbCode = ""
    For x = 0 To vbxTrueGrid.SelBookmarks.count - 1
        vbxTrueGrid.Bookmark = vbxTrueGrid.SelBookmarks(x)
        glbCode = glbCode & Data1.Recordset!TB_KEY & ","
    Next
    glbCode = Left(glbCode, Len(glbCode) - 1)
    Unload frmTABLMASTER
Else
    If Len(Data1.Recordset("TB_KEY")) > 0 Then
        glbCode = Data1.Recordset("TB_KEY")
        If IsNull(Data1.Recordset("TB_DESC")) Then
            glbCodeDesc = ""
        Else
            glbCodeDesc = Data1.Recordset("TB_DESC")
        End If
        Unload frmTABLMASTER
    Else
        Exit Sub
    End If
End If


End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "data1 error", "HRTABL", "SELECT")
Call RollBack '10June99 js

End Sub

Private Sub Form_Activate()
Dim SQLQ As String
If glbOracle Then
    Data1.RecordSource = "SELECT * FROM HRTABL WHERE TB_NAME IN(SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) ORDER BY TB_NAME, TB_INACTIVE, TB_KEY"
Else
    'Data1.RecordSource = "SELECT * FROM HRTABL WHERE TB_NAME IN (SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & glbUserID & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) ORDER BY TB_NAME, TB_INACTIVE, TB_KEY"
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME IN (SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) "
    ' Ticket #21670 Franks 03/07/2012 - begin
    If Len(glbTabNam) > 0 Then
        SQLQ = SQLQ & " AND TB_NAME = '" & glbTabNam & "'"
    End If
    If Len(glbTransDiv) > 0 Then 'Ticket #15248
        SQLQ = SQLQ & " AND TB_KEY IN (" & glbTransDiv & ") "
    End If
    ' Ticket #21670 Franks 03/07/2012 - end
    SQLQ = SQLQ & " ORDER BY TB_NAME, TB_INACTIVE, TB_KEY"
    Data1.RecordSource = SQLQ
End If
Data1.Refresh
Set fRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

End Sub

Private Sub Form_Load()
Dim SQLQ As String
Dim rsSR As New ADODB.Recordset

xLinSDLBFlag = False
If glbLinamar Then 'Ticket #28846 Franks 08/19/2016
    If Len(glbTabNam) > 0 Then
        xLinSDLBFlag = True
        Call LinamarCodeLookupForOneCode
    End If
End If

SQLQ = "UPDATE HRTABL SET TB_INACTIVE = 0 WHERE TB_INACTIVE IS NULL"
gdbAdoIhr001.Execute SQLQ

glbCodeRef = False 'table entrie modified/added false
                   'forces refresh at form level of codes/descriptions
Data1.ConnectionString = glbAdoIHRDB
If glbOracle Then
    Data1.RecordSource = "SELECT * FROM HRTABL WHERE TB_NAME IN(SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) ORDER BY TB_NAME, TB_INACTIVE, TB_KEY"
Else
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME IN (SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) "
    ' Ticket #21670 Franks 03/07/2012 - begin
    If Len(glbTabNam) > 0 Then
        SQLQ = SQLQ & " AND TB_NAME = '" & glbTabNam & "'"
    End If
    If Len(glbTransDiv) > 0 Then 'Ticket #15248
        SQLQ = SQLQ & " AND TB_KEY IN (" & glbTransDiv & ") "
    End If
    ' Ticket #21670 Franks 03/07/2012 - end
    SQLQ = SQLQ & " ORDER BY TB_NAME, TB_INACTIVE, TB_KEY"
    Data1.RecordSource = SQLQ
End If
Data1.Refresh
Set fRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

If glbCompSerial = "S/N - 2347W" Then 'Surrey Place
    ChkInc.Caption = "LTD"
    ChkInc.Tag = "LTD"
End If
If glbBurlTech Then
    ChkInc.Caption = "Unexcused"
    chkSen.Caption = "Excused"
    lblPoint.Caption = "Absence Point"
    chkLE.DataField = "TB_LEFLAG"
    txtLEPoint.DataField = "TB_LEPOINT"
End If

If glbCompSerial = "S/N - 2214W" Then   'Casey House - Ticket #15276
    ChkInc.Caption = "HOOPP"
    ChkInc.Tag = "HOOPP"
End If

If glbCompSerial = "S/N - 2376W" And txtTable = "ADRE" Then    'Assembly of First Nations - Ticket #16181
    txtKey.MaxLength = 10
    txtFindKey.MaxLength = 10
End If
'Ticket 19375 - Mostafa
If glbCompSerial = "S/N - 2241W" And txtTable = "EDSE" Then    'Assembly of First Nations - Ticket #16181
    txtKey.MaxLength = 10
    txtFindKey.MaxLength = 10
    
End If
'Ticket #19137
If glbWFC Then
    Call WFC_Setup
End If
chkWFCWPS.Visible = True 'Jerry asked to show this checkbox for all customers
'Ticket #19137

If txtTable = "EDOR" Then
    chkUnion.Visible = True
    chkUnion.DataField = "TB_USR1"
Else
    chkUnion.Visible = False
    chkUnion.DataField = ""
    chkUnion.Value = 0
End If

'Ticket #23636 - Warning Period for ADRE
If txtTable = "ADRE" Then
    'Ticket #23636 - Warning Period for ADRE - WHSC
    If glbCompSerial = "S/N - 2448W" Then
        lblWarnPrd.Visible = True
        txtWarnPrd.Visible = True
        lblDays.Visible = True
        txtWarnPrd.DataField = "TB_USR1"
    Else
        lblWarnPrd.Visible = False
        txtWarnPrd.Visible = False
        lblDays.Visible = False
        txtWarnPrd.DataField = ""
    End If
Else
    lblWarnPrd.Visible = False
    txtWarnPrd.Visible = False
    lblDays.Visible = False
    txtWarnPrd.DataField = ""
End If


glbCode = ""    'set to null - implies none found/cancel
glbCodeDesc = ""

Screen.MousePointer = HOURGLASS

Me.vbxTrueGrid.Refresh

Call ST_UPD_MODE(False)

'rsSR.Open "SELECT * FROM HRTABL WHERE TB_NAME NOT IN (SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & glbUserID & "' AND CODENAME IS NOT NULL AND Maintainable<>0) ", gdbAdoIhr001, adOpenKeyset
'If Not rsSR.EOF Then
'    cmdNew.Enabled = False
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
'End If

' Ticket #21670 Franks 03/07/2012 - for selection codes
If Len(glbTabNam) > 0 Then
    chkWFCWPS.Visible = False
    chkEDEM.Visible = False
End If
                                
Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

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

fglbUDMode = TF     'in update/new mode

txtWaitPeriod.Enabled = TF
cmbDWM.Enabled = TF

cmdOK.Enabled = TF
cmdCancel.Enabled = TF
cmdModify.Enabled = FT
cmdClose.Enabled = FT
cmdFind.Enabled = FT
cmdNew.Enabled = FT
cmdDelete.Enabled = FT
cmdPrint.Enabled = FT
ChkInc.Enabled = TF
chkSen.Enabled = TF
chkAbsence.Enabled = TF
chkLE.Enabled = TF
chkInactiveCode.Enabled = TF
txtLEPoint.Enabled = TF
txtPoint.Enabled = TF   'Jaddy changed for WFC Aug 30
chkEMELEA.Enabled = TF
chkWorkSchedule.Enabled = TF
txtDesc.Enabled = TF
txtKey.Enabled = TF
txtTable.Enabled = TF
If glbWFC Then 'Ticket #21597 Franks 04/24/2012
    txtUSR1.Enabled = TF
End If
txtFindDesc.Enabled = FT
txtFindKey.Enabled = FT
txtFindTabl.Enabled = FT
vbxTrueGrid.Enabled = FT

'Ticket #20038 - Mainly for WSIB Form 7 to identify which Union Code is actually a Union.
chkUnion.Enabled = TF

'Ticket #23636 - Warning Period for ADRE - WHSC
If glbCompSerial = "S/N - 2448W" Then
    If txtTable = "ADRE" Then
        txtWarnPrd.Enabled = TF
    End If
End If

End Sub

Private Sub txtDesc_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtDWM_Change()
cmbDWM.ListIndex = -1
Select Case txtDWM
Case "D"
    cmbDWM.ListIndex = 0
Case "W"
    cmbDWM.ListIndex = 1
Case "M"
    cmbDWM.ListIndex = 2
End Select
End Sub

Private Sub txtFindDesc_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_KeyPress(KeyAscii As Integer)
If glbCompSerial = "S/N - 2241W" And txtTable = "EDSE" Then
Else
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End If
End Sub

Private Sub txtFindTabl_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindTabl_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtKey_Change()
    'Hemu - 05/29/2003 Begin - Ticket # 4204
    If glbCompSerial = "S/N - 2161W" Then
        If txtTable = "ESCT" Or txtTable = "ESCD" Then
            txtKey.MaxLength = 8
        End If
    End If
    'Hemu - 05/29/2003 End
    'Bryan 15/Sep/05 Ticket #9327
    If txtTable = "BNCD" Then
        txtKey.MaxLength = 10
    End If
    
    'Assembly of First Nations - Ticket #16181
    If glbCompSerial = "S/N - 2376W" And txtTable = "ADRE" Then
        txtKey.MaxLength = 10
        txtFindKey.MaxLength = 10
    End If
    If glbCompSerial = "S/N - 2241W" And txtTable = "EDSE" Then
        txtKey.MaxLength = 10
        txtFindKey.MaxLength = 10
    End If
    
    If glbLinamar Then 'Ticket #28846 Franks 08/17/2016
        If txtTable = "SDLB" Then
            txtKey.MaxLength = 10
        End If
    End If
End Sub

Private Sub txtKey_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtKey_KeyPress(KeyAscii As Integer)
If glbCompSerial = "S/N - 2241W" And txtTable = "EDSE" Then
        
Else
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End If
End Sub


Private Sub txtTable_Change()

If txtTable <> "EDEM" Then
    'Ticket #22220
    chkWorkSchedule.Visible = False
End If

If txtTable = "ADRE" Then
    ChkInc.Visible = True
    chkSen.Visible = True
    chkAbsence.Visible = True
    chkEMELEA.Visible = True
    lblPoint.Visible = True
    txtPoint.Visible = True
    If glbBurlTech Then
        chkLE.Visible = True
        lblLEPoint.Visible = True
        txtLEPoint.Visible = True
    End If
ElseIf txtTable = "EDEM" And Not glbLinamar Then
    ChkInc.Visible = False
    chkSen.Visible = False
    chkAbsence.Visible = False
    chkEMELEA.Visible = True
    chkEMELEA.Caption = "Leave of Absence"
    lblPoint.Visible = False
    txtPoint.Visible = False
    If glbBurlTech Then
        chkLE.Visible = False
        lblLEPoint.Visible = False
        txtLEPoint.Visible = False
    End If
    
    'Ticket #22220
    chkWorkSchedule.Visible = True
    chkWorkSchedule.Left = 5280
    chkWorkSchedule.Top = 4890
ElseIf txtTable = "EDEM" Then
    'Ticket #22220
    chkWorkSchedule.Visible = True
    chkWorkSchedule.Left = 5280
    chkWorkSchedule.Top = 4890
Else
    ChkInc.Visible = False
    chkSen.Visible = False
    chkAbsence.Visible = False
    chkEMELEA.Visible = False
    lblPoint.Visible = False
    txtPoint.Visible = False
    If glbBurlTech Then
        chkLE.Visible = False
        lblLEPoint.Visible = False
        txtLEPoint.Visible = False
    End If
End If

If txtTable = "SDLB" Then txtKey.MaxLength = 1 Else txtKey.MaxLength = 8 '4 'Ticket #23166 Franks 01/29/2013

'Ticket #29769
If txtTable = "ECOM" Or txtTable = "EDEM" Or txtTable = "EDHC" Or txtTable = "EDLC" Or txtTable = "EDPT" Or _
    txtTable = "FURE" Or txtTable = "SDRC" Or txtTable = "TERM" Then
    
    txtKey.MaxLength = 4
Else
    txtKey.MaxLength = 8
End If


    'Hemu - 05/29/2003 Begin - Ticket # 4204
    If glbCompSerial = "S/N - 2161W" Then
        If txtTable = "ESCT" Or txtTable = "ESCD" Then
            txtKey.MaxLength = 8
        End If
    End If
    'Hemu - 05/29/2003 End

'Bryan 15/Sep/05 Ticket #9327
    If txtTable = "BNCD" Then
        txtKey.MaxLength = 10
    End If
'Bryan

If glbCompSerial = "S/N - 2376W" And txtTable = "ADRE" Then    'Assembly of First Nations - Ticket #16181
    txtKey.MaxLength = 10
    txtFindKey.MaxLength = 10
End If
If glbCompSerial = "S/N - 2241W" And txtTable = "EDSE" Then
    txtKey.MaxLength = 10
    txtFindKey.MaxLength = 10
End If
If txtTable = "BNCD" And glbLinamar Then
    lblTitle(0).Visible = True
    txtWaitPeriod.Visible = True
    cmbDWM.Visible = True
Else
    lblTitle(0).Visible = False
    txtWaitPeriod.Visible = False
    cmbDWM.Visible = False
End If

'Ticket #20038 - Mainly for WSIB Form 7 to indicate which Union Code is actually a Union.
If txtTable = "EDOR" Then
    chkUnion.DataField = "TB_USR1"
    chkUnion.Visible = True
    If Data1.Recordset("TB_USR1") = "1" Then
        chkUnion.Value = 1
    Else
        chkUnion.Value = 0
    End If
Else
    chkUnion.Visible = False
    chkUnion.DataField = ""
    chkUnion.Value = 0
End If

'Ticket #19137
If glbWFC Then
    Call WFCScreenSetup
End If

'Ticket #23636 - Warning Period for ADRE
If txtTable = "ADRE" Then
    'Ticket #23636 - Warning Period for ADRE - WHSC
    If glbCompSerial = "S/N - 2448W" Then
        lblWarnPrd.Visible = True
        txtWarnPrd.Visible = True
        lblDays.Visible = True
        txtWarnPrd.DataField = "TB_USR1"
    Else
        lblWarnPrd.Visible = False
        txtWarnPrd.Visible = False
        lblDays.Visible = False
        txtWarnPrd.DataField = ""
    End If
Else
    lblWarnPrd.Visible = False
    txtWarnPrd.Visible = False
    lblDays.Visible = False
    txtWarnPrd.DataField = ""
End If
End Sub

Private Sub WFCScreenSetup()
    If txtTable = "ESCD" Then
        chkLE.Top = chkWFCWPS.Top
        chkLE.Caption = "WPS"
        chkLE.Visible = True
    Else
        chkLE.Top = chkEMELEA.Top
        chkLE.Caption = "L/LE Flag"
        chkLE.Visible = False
    End If
    
    chkEDEM.Visible = True
    If txtTable = "EDEM" Then
        lblPoint.Caption = "Position #411"
        lblPoint.Visible = True
        txtPoint.Visible = True
    Else
        lblPoint.Caption = "Point"
    End If
End Sub

Private Sub txtTable_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtTable_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtUSR1_KeyPress(KeyAscii As Integer)
    If glbWFC Then
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    End If
End Sub

Private Sub txtWaitPeriod_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtWarnPrd_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_DblClick()

If Not Me.vbxTrueGrid.EditActive Then
    glbCode = Data1.Recordset("TB_KEY")
    glbCodeDesc = Data1.Recordset("TB_DESC")
    Unload frmTABLMASTER
Else
    MsgBox "Save/cancel changes first"
End If

End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    'added by Frank 16/Apr/07 Ticket #12913
    If Not fglbNewRec% Then
        fRS.Requery
        fRS.Bookmark = Bookmark
        If fRS("TB_INACTIVE") Then
            RowStyle.ForeColor = vbRed
        End If
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

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
    Dim SQLQ As String
           
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    If glbOracle Then
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME IN(SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) "
    Else
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME IN (SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) "
        ' Ticket #21670 Franks 03/07/2012 - begin
        If Len(glbTabNam) > 0 Then
            SQLQ = SQLQ & " AND TB_NAME = '" & glbTabNam & "'"
        End If
        If Len(glbTransDiv) > 0 Then 'Ticket #15248
            SQLQ = SQLQ & " AND TB_KEY IN (" & glbTransDiv & ") "
        End If
        ' Ticket #21670 Franks 03/07/2012 - end
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh

    Set fRS = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rsSR As New ADODB.Recordset

If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
    If Data1.Recordset("TB_NAME") = "SDLB" Then txtKey.MaxLength = 1 Else txtKey.MaxLength = 8 '4 'Ticket #23166 Franks 01/29/2013
    
    'Hemu - 05/29/2003 Begin - Ticket # 4204
    If glbCompSerial = "S/N - 2161W" Then
        If txtTable = "ESCT" Or txtTable = "ESCD" Then
            txtKey.MaxLength = 8
        End If
    End If
    'Hemu - 05/29/2003 End

    'Assembly of First Nations - Ticket #16181
    If glbCompSerial = "S/N - 2376W" And txtTable = "ADRE" Then
        txtKey.MaxLength = 10
        txtFindKey.MaxLength = 10
    End If
    If glbCompSerial = "S/N - 2241W" And txtTable = "EDSE" Then
        txtKey.MaxLength = 10
        txtFindKey.MaxLength = 10
    End If
    If glbVadim Then
        If txtTable = "EDOR" Then txtKey.MaxLength = 1
        If txtTable = Vadim_PayType_TABLName Then txtKey.MaxLength = 1
        If txtTable = Vadim_EmpType_TABLName Then txtKey.MaxLength = 2
    End If
    
    If glbWFC Then 'Ticket #21597 Franks 04/24/2012
        If txtTable = "EDEM" Then
            lblUsr1.Visible = True
            txtUSR1.Visible = True
        Else
            lblUsr1.Visible = False
            txtUSR1.Visible = False
        End If
    End If
    
    'Ticket # 20038 - Mainly for WSIB Form 7 - to indicate which Union Code is actually a union
    If txtTable = "EDOR" Then
        If Data1.Recordset("TB_USR1") = "1" Then
            chkUnion.Value = 1
        Else
            chkUnion.Value = 0
        End If
    End If
    
    'Ticket #23636 - Warning Period for ADRE
    If txtTable = "ADRE" Then
        'Ticket #23636 - Warning Period for ADRE - WHSC
        If glbCompSerial = "S/N - 2448W" Then
            lblWarnPrd.Visible = True
            txtWarnPrd.Visible = True
            lblDays.Visible = True
            txtWarnPrd.DataField = "TB_USR1"
        Else
            lblWarnPrd.Visible = False
            txtWarnPrd.Visible = False
            lblDays.Visible = False
            txtWarnPrd.DataField = ""
        End If
    Else
        lblWarnPrd.Visible = False
        txtWarnPrd.Visible = False
        lblDays.Visible = False
        txtWarnPrd.DataField = ""
    End If
    
    rsSR.Open "SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME='" & Data1.Recordset("TB_NAME") & "' AND Maintainable<>0 ", gdbAdoIhr001, adOpenKeyset
    If rsSR.EOF Then
        cmdNew.Enabled = False
        cmdModify.Enabled = False
        cmdDelete.Enabled = False
    Else
        cmdNew.Enabled = True
        cmdModify.Enabled = True
        cmdDelete.Enabled = True
    End If
Else
    cmdModify.Enabled = False
End If

End Sub

Private Sub WFC_Setup()
    chkLE.DataField = "TB_LEFLAG"
    'Ticket #21597 Franks 04/24/2012 - begin
    lblUsr1.Caption = "Pension Status"
    txtUSR1.MaxLength = 1
    lblUsr1.Top = txtKey.Top + 20
    txtUSR1.Top = txtKey.Top
    txtUSR1.DataField = "TB_USR1"
    'lblUsr1.Visible = True
    'txtUSR1.Visible = True
    'Ticket #21597 Franks 04/24/2012 - end
End Sub

Private Sub LinamarCodeLookupForOneCode() 'Ticket #28846 Franks 08/19/2016
    If xLinSDLBFlag Then
        txtTable.Visible = False
        txtFindTabl.Visible = False
        vbxTrueGrid.Columns(0).Visible = False
    End If
End Sub
