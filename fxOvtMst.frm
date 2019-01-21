VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmSOvertimeMst 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Overtime Master"
   ClientHeight    =   8925
   ClientLeft      =   150
   ClientTop       =   915
   ClientWidth     =   13140
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
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8925
   ScaleWidth      =   13140
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtMaxBankHrs 
      Appearance      =   0  'Flat
      DataField       =   "OM_MAX_BANK_HRS"
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
      TabIndex        =   9
      Tag             =   "10-Maximum Hours that can be Banked"
      Top             =   6480
      Width           =   885
   End
   Begin VB.TextBox txtMultiplier 
      Appearance      =   0  'Flat
      DataField       =   "OM_MULTIPLIER"
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
      TabIndex        =   10
      Tag             =   "00-"
      Top             =   6840
      Width           =   885
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      DataField       =   "OM_EMAIL"
      DataSource      =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3000
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Tag             =   "00-Email Address"
      Top             =   7200
      Width           =   6660
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5700
      Top             =   3090
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "OM_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   5340
      MaxLength       =   25
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2460
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "OM_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   5820
      MaxLength       =   25
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2460
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "OM_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   6300
      MaxLength       =   25
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2460
      Visible         =   0   'False
      Width           =   420
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   13140
      _Version        =   65536
      _ExtentX        =   23177
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
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6300
      Top             =   2580
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
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Align           =   1  'Align Top
      Bindings        =   "fxOvtMst.frx":0000
      Height          =   2625
      Left            =   0
      OleObjectBlob   =   "fxOvtMst.frx":0014
      TabIndex        =   12
      Tag             =   "Overtime Master"
      Top             =   495
      Width           =   13140
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   24
      Top             =   8340
      Width           =   13140
      _Version        =   65536
      _ExtentX        =   23177
      _ExtentY        =   1032
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
      Begin VB.CommandButton cmdForfeitHrs 
         Appearance      =   0  'Flat
         Caption         =   "Forfeit Hours"
         Height          =   375
         Left            =   3840
         TabIndex        =   35
         Tag             =   "Forfeit Monthly Hours"
         Top             =   120
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton cmdUpdVadim 
         Appearance      =   0  'Flat
         Caption         =   "Update in Vadim"
         Height          =   375
         Left            =   9240
         TabIndex        =   29
         Tag             =   "Update Max Bank Hrs in Vadim"
         Top             =   120
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton CmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "Update All Employees"
         Height          =   375
         Left            =   6360
         TabIndex        =   28
         Tag             =   "Update all employees bank time"
         Top             =   120
         Width           =   2505
      End
      Begin VB.CommandButton CmdRecalc 
         Appearance      =   0  'Flat
         Caption         =   "R&ecalculate All Employees"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   25
         Tag             =   "Recalculate for all employees"
         Top             =   120
         Width           =   2505
      End
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "OM_ORG"
      Height          =   285
      Index           =   1
      Left            =   2685
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   4320
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      DataField       =   "OM_PT"
      DataSource      =   " "
      Height          =   285
      Left            =   2685
      TabIndex        =   5
      Tag             =   "00-Category Codes"
      Top             =   5040
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "OM_EMP"
      Height          =   285
      Index           =   0
      Left            =   2685
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   4680
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "OM_LOC"
      Height          =   285
      Index           =   2
      Left            =   2685
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   3960
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "OM_SECTION"
      Height          =   285
      Index           =   5
      Left            =   2685
      TabIndex        =   8
      Tag             =   "00-Enter Section Code"
      Top             =   6120
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "OM_ADMINBY"
      Height          =   285
      Index           =   4
      Left            =   2685
      TabIndex        =   7
      Tag             =   "00-Enter Administered By Code"
      Top             =   5760
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "OM_REGION"
      Height          =   285
      Index           =   3
      Left            =   2685
      TabIndex        =   6
      Tag             =   "00-Enter Region Code"
      Top             =   5400
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      DataField       =   "OM_EFDATE"
      Height          =   285
      Index           =   0
      Left            =   2685
      TabIndex        =   0
      Tag             =   "40-From Date"
      Top             =   3600
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1210
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      DataField       =   "OM_ETDATE"
      Height          =   285
      Index           =   1
      Left            =   4455
      TabIndex        =   1
      Tag             =   "40-To Date"
      Top             =   3600
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1210
   End
   Begin VB.Label lblPeriod 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Overtime Entitlement Period"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   3645
      Width           =   2370
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
      TabIndex        =   33
      Top             =   5805
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
      TabIndex        =   32
      Top             =   5445
      Width           =   510
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
      TabIndex        =   31
      Top             =   6165
      Width           =   540
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      TabIndex        =   30
      Top             =   4005
      Width           =   615
   End
   Begin VB.Label lblEEStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status"
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
      TabIndex        =   27
      Top             =   4725
      Width           =   1350
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      TabIndex        =   26
      Top             =   5085
      Width           =   630
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Per Year"
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
      Index           =   5
      Left            =   3960
      TabIndex        =   23
      Top             =   6525
      Width           =   615
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "(Email notification will be sent when Overtime Bank Outstanding is Negative and when exceeded the Maximum Bank Hours.)"
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
      Index           =   4
      Left            =   3000
      TabIndex        =   22
      Top             =   7800
      Width           =   8775
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Multiplier"
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
      TabIndex        =   21
      Top             =   6885
      Width           =   615
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Email Addresses"
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
      Left            =   120
      TabIndex        =   20
      Top             =   7200
      Width           =   1155
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Maximum Bank Hours"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   6525
      Width           =   1830
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   4365
      Width           =   735
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CompNo"
      DataField       =   "OM_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5700
      TabIndex        =   17
      Top             =   2865
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "frmSOvertimeMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean
Dim rsDATA As New ADODB.Recordset

Private Function chkOvtMst()
Dim SQLQ As String, Msg As String, dd#, PID&, Expr#, Skill$

chkOvtMst = False

On Error GoTo chkOvtMst_Err

If Len(dlpDateRange(0).Text) > 0 Then
    If Not IsDate(dlpDateRange(0).Text) Then
        If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
            MsgBox "Invalid Extra Time Entitlement Period From Date"
        Else
            MsgBox "Invalid Overtime Entitlement Period From Date"
        End If
        dlpDateRange(0).SetFocus
        Exit Function
    End If
Else
    If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
        MsgBox "Extra Time Entitlement Period From Date is mandatory"
    Else
        MsgBox "Overtime Entitlement Period From Date is mandatory"
    End If
    dlpDateRange(0).SetFocus
    Exit Function
End If

If Len(dlpDateRange(1).Text) > 0 Then
    If Not IsDate(dlpDateRange(1).Text) Then
        If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
            MsgBox "Invalid Extra Time Entitlement Period To Date"
        Else
            MsgBox "Invalid Overtime Entitlement Period To Date"
        End If
        dlpDateRange(1).SetFocus
        Exit Function
    End If
Else
    If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
        MsgBox "Extra Time Entitlement Period To Date is mandatory"
    Else
        MsgBox "Overtime Entitlement Period To Date is mandatory"
    End If
    dlpDateRange(1).SetFocus
    Exit Function
End If

If CVDate(dlpDateRange(0)) > CVDate(dlpDateRange(1)) Then
    If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
        MsgBox "Extra Time Entitlement Period To Date cannot be prior to From Date"
    Else
        MsgBox "Overtime Entitlement Period To Date cannot be prior to From Date"
    End If
    dlpDateRange(1).SetFocus
    Exit Function
End If

If Not clpCode(2).ListChecker Then Exit Function    'Location

'If glbCompSerial <> "S/N - 2425W" Then   'Ticket #18223 - Four Villages CHC
'    If Len(clpCode(1)) < 1 Then
'        MsgBox lStr("Union Code is a required field")
'        clpCode(1).SetFocus
'        Exit Function
'    End If
'    If clpCode(1).Caption = "Unassigned" Then
'        MsgBox lStr("Union code must be valid")
'        clpCode(1).SetFocus
'        Exit Function
'    End If
'End If

If Not clpCode(1).ListChecker Then Exit Function    'Union
If Not clpCode(0).ListChecker Then Exit Function    'Employment Status
If Not clpPT.ListChecker Then Exit Function
If Not clpCode(3).ListChecker Then Exit Function    'Region
If Not clpCode(4).ListChecker Then Exit Function    'Admin By
If Not clpCode(5).ListChecker Then Exit Function    'Section

If Len(txtMaxBankHrs) < 1 Then
    MsgBox "Maximum Bank Hours is a required field"
    txtMaxBankHrs.SetFocus
    Exit Function
End If
If Not IsNumeric(txtMaxBankHrs) Then
    MsgBox "Maximum Bank Hours must be numeric."
    txtMaxBankHrs.SetFocus
    Exit Function
End If

If modIsDupUnion() Then
    MsgBox "Duplicate entry."
    clpCode(1).SetFocus
    Exit Function
End If

If txtMultiplier = "" Then
    txtMultiplier = 1
Else
    If Not IsNumeric(txtMultiplier) Then
        MsgBox "Multiplier must be numeric"
        txtMultiplier.SetFocus
        Exit Function
    End If
End If
If Len(txtAddress.Text) > 0 Then
If Not IsEmail(txtAddress.Text) Then
    MsgBox "Email address must be in xxx@yyy.zzz format.", vbExclamation + vbOKOnly, "Invalid Email Address"
    Exit Function
End If
End If
chkOvtMst = True

Exit Function

chkOvtMst_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HRJOBSKL", "edit/Add")
Call RollBack

End Function

Private Function IsEmail(Address As String) As Boolean
    IsEmail = True
    ' Make sure there's an @ in the address
    If InStr(Address, "@") = 0 Then IsEmail = False: Exit Function
    ' Make sure they have at least one period after the @
    If InStr(InStr(Address, "@"), Address, ".") = 0 Then IsEmail = False: Exit Function
    ' Make sure they have text before the period
    If Mid(Address, InStr(Address, "@") + 1, 1) = "." Then IsEmail = False: Exit Function
    ' Make sure they have text after the period
    If Right(Address, 1) = "." Then IsEmail = False: Exit Function
End Function

Public Sub cmdCancel_Click()
On Error GoTo Can_Err

fglbNew = False
Call Display_Value

If Data1.Recordset.EOF Then
    CmdRecalc(1).Enabled = False
    CmdUpdate.Enabled = False
    
    If glbCompSerial = "S/N - 2276W" Then
        cmdUpdVadim.Enabled = False
    End If
    
    'Ticket #19998 - Four Villages
    If glbCompSerial = "S/N - 2425W" Then
        cmdForfeitHrs.Enabled = False
    End If
Else
    CmdRecalc(1).Enabled = True
    CmdUpdate.Enabled = True
    
    If glbCompSerial = "S/N - 2276W" Then
        cmdUpdVadim.Enabled = True
    End If
    
    'Ticket #19998 - Four Villages
    If glbCompSerial = "S/N - 2425W" Then
        cmdForfeitHrs.Enabled = True
    End If
End If

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_OVERTIME_MASTER", "Cancel")
Call RollBack

End Sub

Public Sub cmdClose_Click()
glbUserUploadMode = SwitchForm: Unload Me
End Sub

Public Sub cmdDelete_Click()
Call clkDelete
End Sub

Public Sub cmdNew_Click()
Dim SQLQ As String

On Error GoTo AddN_Err

fglbNew = True
Call Set_Control("B", Me)
Call SET_UP_MODE

lblCNum.Caption = "001"

If glbCompSerial <> "S/N - 2425W" Then   'Ticket #18223 - Four Villages CHC
    dlpDateRange(0).SetFocus
Else
    dlpDateRange(0).SetFocus
End If

CmdRecalc(1).Enabled = False
CmdUpdate.Enabled = False
If glbCompSerial = "S/N - 2276W" Then
    cmdUpdVadim.Enabled = False
End If

'Ticket #19998 - Four Villages
If glbCompSerial = "S/N - 2425W" Then
    cmdForfeitHrs.Enabled = False
End If

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "clkNew", "HR_OVERTIME_MASTER", "Add")
Call RollBack
End Sub

Public Sub cmdOK_Click()
Call clkOK
End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String

RHeading = Me.Caption
'RHeading = Mid(RHeading, 1, InStr(RHeading, "-"))
'RHeading = RHeading & " " & lblPosDesc.Caption
'RHeading = Me.Caption & lblPosDesc.Caption

Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1
End Sub

Public Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = Me.Caption
'RHeading = Mid(RHeading, 1, InStr(RHeading, "-"))
'RHeading = RHeading & " " & lblPosDesc.Caption
'RHeading = Me.Caption & lblPosDesc.Caption

Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub

Private Sub cmdForfeitHrs_Click()
     frmForfeitHrs.Show 1
End Sub

Private Sub cmdRecalc_Click(Index As Integer)
Dim Msg, Response, DgDef, SQLQ As String

Msg = "Do you wish to proceed and recalculate "
If Index = 1 Then
    Msg = Msg & "all Employees' "
Else
    Msg = Msg & "the Employee's "
End If
If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
    Msg = Msg & "Extra Time Bank ?"
Else
    Msg = Msg & "Overtime Bank ?"
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
    'SQLQ = "OT_EMPNBR = " & Data1.Recordset("ED_EMPNBR")
    'Call ReCalcOvt(SQLQ)
End If

If Not glbSQL And Not glbOracle Then Call Pause(0.5)
'Data1.Refresh
Screen.MousePointer = DEFAULT

'If Index = 1 Then
'    Call Form_Activate
'Else
'    Data1.Recordset.Find SQLQ
'End If

End Sub

Private Sub cmdUpdate_Click()
    Dim Msg, Response, DgDef, SQLQ As String
    Dim rsEmp As New ADODB.Recordset
    Dim rsOvtEmp As New ADODB.Recordset
    Dim L1, L2 As Integer
    Dim L11, L21 As Integer
    Dim recCount As Integer
    
    'Since this option will change the Overtime Bank Period, inform the user and get confirmation for the update
    If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
        Msg = "This option will change the Extra Time Bank Period to Current Year. " & vbCrLf
        Msg = Msg & "Make sure the Rollover and Zero Out Year End tasks for Extra Time are completed depending upon your policies." & vbCrLf & vbCrLf
    Else
        Msg = "This option will change the Overtime Bank Period to Current Year. " & vbCrLf
        Msg = Msg & "Make sure the Rollover and Zero Out Year End tasks for Overtime are completed depending upon your policies." & vbCrLf & vbCrLf
    End If
    Msg = Msg & "Do you wish to proceed with this Update? "
    
    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
    
    If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
        Response = MsgBox(Msg, DgDef, "Update Extra Time Bank Period")
    Else
        Response = MsgBox(Msg, DgDef, "Update Overtime Bank Period")
    End If
    If Response = IDNO Then Exit Sub
    
       
    Screen.MousePointer = HOURGLASS
    
    'Proceed with Overtime Bank Period change.
    If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
        Data1.Recordset.MoveFirst
        
        L1 = Data1.Recordset.RecordCount
        L2 = 0
                
        MDIMain.panHelp(0).FloodType = 1
        
        Do
            Call Display_Value
            
            MDIMain.panHelp(0).FloodPercent = 100 * (L2 / L1)
            
            Set rsEmp = Nothing
            SQLQ = "SELECT ED_EMPNBR, ED_EMP,ED_PT,ED_ORG FROM HREMP WHERE 1=1 "
            If Len(clpCode(1).Text) > 0 Then
                SQLQ = SQLQ & " AND ED_ORG = '" & Data1.Recordset("OM_ORG") & "'"
            End If
            If Len(clpCode(0).Text) > 0 Then
                SQLQ = SQLQ & " AND ED_EMP = '" & Data1.Recordset("OM_EMP") & "'"
            End If
            If Len(clpPT.Text) > 0 Then
                SQLQ = SQLQ & " AND ED_PT = '" & Data1.Recordset("OM_PT") & "'"
            End If
                        
            'Ticket #15753
            If Len(clpCode(2).Text) > 0 Then
                SQLQ = SQLQ & " AND ED_LOC = '" & Data1.Recordset("OM_LOC") & "'"
            End If
            If Len(clpCode(3).Text) > 0 Then
                SQLQ = SQLQ & " AND ED_REGION = '" & Data1.Recordset("OM_REGION") & "'"
            End If
            If Len(clpCode(4).Text) > 0 Then
                SQLQ = SQLQ & " AND ED_ADMINBY = '" & Data1.Recordset("OM_ADMINBY") & "'"
            End If
            If Len(clpCode(5).Text) > 0 Then
                SQLQ = SQLQ & " AND ED_SECTION = '" & Data1.Recordset("OM_SECTION") & "'"
            End If
            
            'SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT OT_EMPNBR FROM HR_OVERTIME_BANK)"
            rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsEmp.EOF Then
                rsEmp.MoveFirst
                
                L11 = rsEmp.RecordCount
                L21 = 0
                
                If L11 > 0 Then
                    Msg = Str(L11)
                    If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
                        If L11 = 1 Then Msg = Msg & " Extra Time Bank Period " Else Msg = Msg & " Extra Time Bank Periods "
                    Else
                        If L11 = 1 Then Msg = Msg & " Overtime Bank Period " Else Msg = Msg & " Overtime Bank Periods "
                    End If
                    Msg = Msg & "for the selected record will be Updated. " & vbCrLf & vbCrLf & "Do you want to proceed?"
                    If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
                        Response = MsgBox(Msg, DgDef, "Update Extra Time Bank Period")    ' Get user response.
                    Else
                        Response = MsgBox(Msg, DgDef, "Update Overtime Bank Period")    ' Get user response.
                    End If
                    If Response = IDNO Then
                        rsEmp.Close
                        Set rsEmp = Nothing
                        GoTo Next_OvertimePeriod
                    End If
                Else
                    If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
                        MsgBox "No employee found to update Overtime Bank Period."
                    Else
                        MsgBox "No employee found to update Extra Time Bank Period."
                    End If
                    rsEmp.Close
                    Set rsEmp = Nothing
                    GoTo Next_OvertimePeriod
                End If
                
                MDIMain.panHelp(1).FloodType = 1
                
                Do While Not rsEmp.EOF
                    
                    MDIMain.panHelp(1).FloodPercent = 100 * (L21 / L11)
                    
                    Set rsOvtEmp = Nothing
                    rsOvtEmp.Open "SELECT * FROM HR_OVERTIME_BANK WHERE OT_EMPNBR=" & rsEmp("ED_EMPNBR"), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If rsOvtEmp.EOF Then
                        rsOvtEmp.AddNew
                        rsOvtEmp("OT_PBANK") = 0
                    End If
                    rsOvtEmp("OT_COMPNO") = "001"
                    rsOvtEmp("OT_EMPNBR") = rsEmp("ED_EMPNBR")
                    rsOvtEmp("OT_BANK") = Get_OvertimeBank(rsEmp("ED_EMPNBR"), Data1.Recordset("OM_EFDATE"), Data1.Recordset("OM_ETDATE")) * Val(Data1.Recordset("OM_MULTIPLIER"))
                    rsOvtEmp("OT_BANKT") = Get_OvertimeTaken(rsEmp("ED_EMPNBR"), Data1.Recordset("OM_EFDATE"), Data1.Recordset("OM_ETDATE"))
                    rsOvtEmp("OT_MBANK") = Data1.Recordset("OM_MAX_BANK_HRS")
                    rsOvtEmp("OT_EFDATE") = Data1.Recordset("OM_EFDATE")  'Format("1/1/" & Year(Now()), "mm/dd/yyyy")
                    rsOvtEmp("OT_ETDATE") = Data1.Recordset("OM_ETDATE")  'Format("12/31/" & Year(Now()), "mm/dd/yyyy")
                    rsOvtEmp("OT_LDATE") = Date
                    rsOvtEmp("OT_LTIME") = Time$
                    rsOvtEmp("OT_LUSER") = glbUserID
                    rsOvtEmp.Update
                    
                    rsOvtEmp.Close
                    
                    L21 = L21 + 1
                    
                    rsEmp.MoveNext
                Loop
            Else
                If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
                    MsgBox "No employee found to update Extra Time Bank Period."
                Else
                    MsgBox "No employee found to update Overtime Bank Period."
                End If
            End If
            rsEmp.Close
    
Next_OvertimePeriod:
            L2 = L2 + 1

            Data1.Recordset.MoveNext
            
        Loop Until Data1.Recordset.EOF
        
        MDIMain.panHelp(0).FloodPercent = 100
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).FloodPercent = 100
        MDIMain.panHelp(1).FloodType = 0

    End If
        
    Call ReCalcOvt("")

    If Not glbSQL And Not glbOracle Then Call Pause(0.5)

    Data1.Refresh
    Me.vbxTrueGrid.Refresh
    Call Display_Value
    
    If glbCompSerial = "S/N - 2425W" Then   'Ticket #19998 - Four Villages CHC
        MsgBox "Extra Time Bank Period Updated Successfully."
    Else
        MsgBox "Overtime Bank Period Updated Successfully."
    End If
    
    Screen.MousePointer = DEFAULT
End Sub

Private Sub cmdUpdVadim_Click()
    Dim Msg, Response, DgDef, SQLQ As String
    Dim rsEmp As New ADODB.Recordset
    Dim L1, L2 As Integer
    Dim EMPMAXBANKBatchID
    
    'Since this update will update Max Bank Hrs in Vadim for all the employees falling in this rule, it is important to prompt them
    Msg = "This option will update Max Bank Hrs in Vadim for all the employee who are part of this Overtime Bank rule. " & vbCrLf & vbCrLf
    Msg = Msg & "Do you wish to proceed with this Update? "
    
    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
    
    Response = MsgBox(Msg, DgDef, "Update Max Bank Hrs in Vadim")
    If Response = IDNO Then Exit Sub
    
    Screen.MousePointer = HOURGLASS
    
    'Proceed with Max Overtime Bank update into Vadim files.
    If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
        Set rsEmp = Nothing
        SQLQ = "SELECT ED_EMPNBR, ED_PAYROLL_ID, ED_EMP,ED_PT,ED_ORG FROM HREMP WHERE 1=1"
        If Len(clpCode(1).Text) > 0 Then
            SQLQ = SQLQ & " AND ED_ORG = '" & clpCode(1).Text & "'"
        End If
        If Len(clpCode(0).Text) > 0 Then
            SQLQ = SQLQ & " AND ED_EMP = '" & clpCode(0).Text & "'"
        End If
        If Len(clpPT.Text) > 0 Then
            SQLQ = SQLQ & " AND ED_PT = '" & clpPT.Text & "'"
        End If
        
        'Ticket #15753
        If Len(clpCode(2).Text) > 0 Then
            SQLQ = SQLQ & " AND ED_LOC = '" & clpCode(2).Text & "'"
        End If
        If Len(clpCode(3).Text) > 0 Then
            SQLQ = SQLQ & " AND ED_REGION = '" & clpCode(3).Text & "'"
        End If
        If Len(clpCode(4).Text) > 0 Then
            SQLQ = SQLQ & " AND ED_ADMINBY = '" & clpCode(4).Text & "'"
        End If
        If Len(clpCode(5).Text) > 0 Then
            SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(5).Text & "'"
        End If
        
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmp.EOF Then
        
            L1 = rsEmp.RecordCount
            L2 = 0
            
            If L1 > 0 Then
                Msg = Str(L1)
                If L1 = 1 Then Msg = Msg & " employee's Max Bank Hrs in Vadim " Else Msg = Msg & " employees Max Bank Hrs in Vadim "
                Msg = Msg & "for the selected rule will be Updated. " & vbCrLf & vbCrLf & "Do you want to proceed?"
                Response = MsgBox(Msg, DgDef, "Update Max Bank Hrs in Vadim")    ' Get user response.
                If Response = IDNO Then
                    rsEmp.Close
                    Set rsEmp = Nothing
                    Exit Sub
                End If
            Else
                MsgBox "No employee record found to update."
                rsEmp.Close
                Set rsEmp = Nothing
                Exit Sub
            End If
            
            MDIMain.panHelp(0).FloodType = 1
            
            rsEmp.MoveFirst
            Do While Not rsEmp.EOF
                MDIMain.panHelp(0).FloodPercent = 100 * (L2 / L1)
                                
                'Update Vadim interface tables
                If Not IsNull(rsEmp("ED_PAYROLL_ID")) And rsEmp("ED_PAYROLL_ID") <> "" Then
                    EMPMAXBANKBatchID = AddBatchVadim("M")
                    Call VadimInterface(EMPMAXBANKBatchID, rsEmp("ED_PAYROLL_ID"), "ED_SUPCODE", 0, Val(txtMaxBankHrs.Text))
                    Call CloseBatchVadim(EMPMAXBANKBatchID)
                End If
                                
                L2 = L2 + 1
                
                rsEmp.MoveNext
            Loop
        Else
            MsgBox "No employee record found to update."
        End If
        rsEmp.Close
    End If
    
    MDIMain.panHelp(0).FloodPercent = 100
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""

    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    
    Screen.MousePointer = DEFAULT

    MsgBox "Update complete.", vbInformation, "Update Max Bank Hrs in Vadim"

End Sub

Private Sub Form_Activate()

glbOnTop = "FRMSOVERTIMEMST"

Call SET_UP_MODE

If Data1.Recordset.EOF Then
    CmdRecalc(1).Enabled = False
    CmdUpdate.Enabled = False
    
    If glbCompSerial = "S/N - 2276W" Then
        cmdUpdVadim.Enabled = False
    End If

    'Ticket #19998 - Four Villages
    If glbCompSerial = "S/N - 2425W" Then
        cmdForfeitHrs.Enabled = False
    End If
Else
    CmdRecalc(1).Enabled = True
    CmdUpdate.Enabled = True
    
    If glbCompSerial = "S/N - 2276W" Then
        cmdUpdVadim.Enabled = True
    End If
    
    'Ticket #19998 - Four Villages
    If glbCompSerial = "S/N - 2425W" Then
        cmdForfeitHrs.Enabled = True
    End If
End If

End Sub

Private Sub Form_Deactivate()
glbUserUploadMode = SwitchForm: Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub

Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)

End Sub

Private Sub Form_Load()
Dim SQLQ
On Error GoTo FLErr

glbOnTop = "FRMSOVERTIMEMST"
Me.Height = 6000
Me.Width = 7000

Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "SELECT * FROM HR_OVERTIME_MASTER ORDER BY OM_ORG"
Data1.Refresh
Screen.MousePointer = HOURGLASS
Me.vbxTrueGrid.Refresh
Screen.MousePointer = DEFAULT

Call setCaption(lblLocation)
Call setCaption(lblTitle(1))
Call setCaption(lblPT)
Call setCaption(lblRegion)
Call setCaption(lblAdmin)
Call setCaption(lblSection)
Call setCaption(vbxTrueGrid.Columns(0))
Call setCaption(vbxTrueGrid.Columns(1))
Call setCaption(vbxTrueGrid.Columns(2))
Call setCaption(vbxTrueGrid.Columns(3))
Call setCaption(vbxTrueGrid.Columns(5))
Call setCaption(vbxTrueGrid.Columns(6))
Call setCaption(vbxTrueGrid.Columns(7))
Call setCaption(vbxTrueGrid.Columns(8))

'If EERetrieve() = False Then
'    MsgBox "Sorry, Position can not be found"
'    frmJOBS.Show 1
'Else
'    Me.Show
'End If

'Screen.MousePointer = DEFAULT

'Call Display_Value
Call INI_Controls(Me)

If glbCompSerial = "S/N - 2276W" Then
    cmdUpdVadim.Visible = True
Else
    cmdUpdVadim.Visible = False
End If

'Ticket #19998 - Four Villages
If glbCompSerial = "S/N - 2425W" Then
    cmdForfeitHrs.Visible = True
    lblPeriod.Caption = "Extra Time Entitlement Period"
Else
    cmdForfeitHrs.Visible = False
End If

Screen.MousePointer = DEFAULT

Exit Sub

FLErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form load Error", "Overtime Master", "Select")
Call RollBack

End Sub

Private Sub Form_LostFocus()
'MDIMain.MainToolBar.ButtonS(2).Enabled = True
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub


Public Function EERetrieve()
Dim SQLQ$

EERetrieve = False
Screen.MousePointer = HOURGLASS

On Error GoTo EERetrieveErr

SQLQ$ = "SELECT * FROM HR_OVERTIME_MASTER "
'SQLQ$ = SQLQ$ & "WHERE JS_CODE = '" & glbPos & "'"
SQLQ$ = SQLQ$ & "ORDER BY OM_ORG"

Data1.RecordSource = SQLQ$
Data1.Refresh


EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERetrieveErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Overtime Master", "HR_OVERTIME_MASTER", "SELECT")
Call RollBack

End Function

Private Function modIsDupUnion()
Dim SQLQ$
Dim snapOvtMst As New ADODB.Recordset

modIsDupUnion = True

On Error GoTo modIsDupUnion_Err

Screen.MousePointer = HOURGLASS
SQLQ$ = "SELECT * FROM HR_OVERTIME_MASTER WHERE "
If Len(clpCode(1).Text) > 0 Then
    SQLQ$ = SQLQ$ & " OM_ORG = '" & clpCode(1).Text & "'"
Else
    SQLQ$ = SQLQ$ & " OM_ORG IS NULL"
End If
If Len(clpCode(0).Text) > 0 Then
    SQLQ$ = SQLQ$ & " AND OM_EMP = '" & clpCode(0).Text & "'"
Else
    SQLQ$ = SQLQ$ & " AND OM_EMP IS NULL"
End If

If Len(clpPT.Text) > 0 Then
    SQLQ$ = SQLQ$ & " AND OM_PT = '" & clpPT.Text & "'"
Else
    SQLQ$ = SQLQ$ & " AND OM_PT IS NULL"
End If

'Ticket #15753
If Len(clpCode(2).Text) > 0 Then
    SQLQ$ = SQLQ$ & " AND OM_LOC = '" & clpCode(2).Text & "'"
Else
    SQLQ$ = SQLQ$ & " AND OM_LOC IS NULL"
End If
If Len(clpCode(3).Text) > 0 Then
    SQLQ$ = SQLQ$ & " AND OM_REGION = '" & clpCode(3).Text & "'"
Else
    SQLQ$ = SQLQ$ & " AND OM_REGION IS NULL"
End If
If Len(clpCode(4).Text) > 0 Then
    SQLQ$ = SQLQ$ & " AND OM_ADMINBY = '" & clpCode(4).Text & "'"
Else
    SQLQ$ = SQLQ$ & " AND OM_ADMINBY IS NULL"
End If
If Len(clpCode(5).Text) > 0 Then
    SQLQ$ = SQLQ$ & " AND OM_SECTION = '" & clpCode(5).Text & "'"
Else
    SQLQ$ = SQLQ$ & " AND OM_SECTION IS NULL"
End If

If Not fglbNew Then SQLQ$ = SQLQ$ & " AND OM_ID <> " & Data1.Recordset("OM_ID")

snapOvtMst.Open SQLQ$, gdbAdoIhr001, adOpenStatic

If snapOvtMst.BOF And snapOvtMst.EOF Then
    modIsDupUnion = False
End If

Screen.MousePointer = DEFAULT
snapOvtMst.Close

Exit Function

modIsDupUnion_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Code Snap", "TABL", "SELECT")
Call RollBack

End Function

Private Sub txtAddress_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtMaxBankHrs_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub txtMaxBankHrs_LostFocus()
'
'If Len(txtMaxBankHrs) > 0 Then
'    If IsNumeric(txtMaxBankHrs) Then
'        txtMaxBankHrs = (Int(txtMaxBankHrs * 100) / 100)
'    End If
'Else
'    txtExperience = 0
'End If
'
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
    
    SQLQ = "SELECT * FROM HR_OVERTIME_MASTER "
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Display_Value
End Sub

Private Function RollBack()
On Error GoTo rr
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    glbUserUploadMode = UploadFormWithoutCheck: Unload Me
End If
rr:
End Function

Public Function clkOK()
Dim SQLQ
Dim xID
On Error GoTo OK_Err

clkOK = False
If Not chkOvtMst() Then Exit Function
If fglbNew Then rsDATA.AddNew: fglbNew = False

Screen.MousePointer = HOURGLASS

Call UpdUStats(Me)
Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
xID = rsDATA("OM_ID")
gdbAdoIhr001.CommitTrans
Data1.Refresh
Data1.Recordset.Find "OM_ID=" & xID
'Call Display_Value

'Recalculate the Overtime Bank
Dim rsOvtEmp As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset

SQLQ = "SELECT * FROM HR_OVERTIME_BANK WHERE OT_EMPNBR IN "
SQLQ = SQLQ & "(SELECT ED_EMPNBR FROM HREMP WHERE 1=1 "
If Len(clpCode(1).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_ORG = '" & clpCode(1).Text & "'"
End If
If Len(clpCode(0).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_EMP = '" & clpCode(0).Text & "'"
End If
If Len(clpPT.Text) > 0 Then
    SQLQ = SQLQ & " AND ED_PT = '" & clpPT.Text & "'"
End If
'Ticket #15753
If Len(clpCode(2).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_LOC = '" & clpCode(2).Text & "'"
End If
If Len(clpCode(3).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_REGION = '" & clpCode(3).Text & "'"
End If
If Len(clpCode(4).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_ADMINBY = '" & clpCode(4).Text & "'"
End If
If Len(clpCode(5).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(5).Text & "'"
End If

SQLQ = SQLQ & " )"

rsOvtEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not rsOvtEmp.EOF Then
    rsOvtEmp.MoveFirst
    Do While Not rsOvtEmp.EOF
        Call ReCalcOvt("OT_EMPNBR = " & rsOvtEmp("OT_EMPNBR"))
        rsOvtEmp.MoveNext
    Loop
End If
rsOvtEmp.Close

'Not to be done in this routine - just save the rule. On Update All Employees will update the
'Employee's records.
'Add record Overtime Bank records
'SQLQ = "SELECT * FROM HREMP WHERE ED_ORG = '" & clpCode(1).Text & "'"
'If Len(clpCode(0).Text) > 0 Then
'    SQLQ = SQLQ & " AND ED_EMP = '" & clpCode(0).Text & "'"
'End If
'If Len(clpPT.Text) > 0 Then
'    SQLQ = SQLQ & " AND ED_PT = '" & clpPT.Text & "'"
'End If
'
'SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (SELECT OT_EMPNBR FROM HR_OVERTIME_BANK)"
'rsEMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'If Not rsEMP.EOF Then
'    rsOvtEmp.Open "SELECT * FROM HR_OVERTIME_BANK WHERE 1=2", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    rsEMP.MoveFirst
'    Do While Not rsEMP.EOF
'        rsOvtEmp.AddNew
'        rsOvtEmp("OT_COMPNO") = "001"
'        rsOvtEmp("OT_EMPNBR") = rsEMP("ED_EMPNBR")
'        rsOvtEmp("OT_PBANK") = 0
'        rsOvtEmp("OT_BANK") = Get_OvertimeBank(rsEMP("ED_EMPNBR")) * Val(txtMultiplier.Text)
'        rsOvtEmp("OT_BANKT") = Get_OvertimeTaken(rsEMP("ED_EMPNBR"))
'        rsOvtEmp("OT_MBANK") = txtMaxBankHrs.Text
'        rsOvtEmp("OT_EFDATE") = Format("1/1/" & Year(Now()), "mm/dd/yyyy")
'        rsOvtEmp("OT_ETDATE") = Format("12/31/" & Year(Now()), "mm/dd/yyyy")
'        rsOvtEmp("OT_LDATE") = Date
'        rsOvtEmp("OT_LTIME") = Time$
'        rsOvtEmp("OT_LUSER") = glbUserID
'        rsOvtEmp.Update
'        rsEMP.MoveNext
'    Loop
'    rsOvtEmp.Close
'End If
'rsEMP.Close

clkOK = True

fglbNew = False

If Data1.Recordset.EOF Then
    CmdRecalc(1).Enabled = False
    CmdUpdate.Enabled = False
    
    If glbCompSerial = "S/N - 2276W" Then
        cmdUpdVadim.Enabled = False
    End If
    
    'Ticket #19998 - Four Villages
    If glbCompSerial = "S/N - 2425W" Then
        cmdForfeitHrs.Enabled = False
    End If
Else
    CmdRecalc(1).Enabled = True
    CmdUpdate.Enabled = True
    
    If glbCompSerial = "S/N - 2276W" Then
        cmdUpdVadim.Enabled = True
    End If
    
    'Ticket #19998 - Four Villages
    If glbCompSerial = "S/N - 2425W" Then
        cmdForfeitHrs.Enabled = True
    End If
End If

Screen.MousePointer = vbDefault

Exit Function

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OVERTIME_MASTER", "Update")
Call RollBack   '1
End Function

Public Sub clkCancel()
On Error GoTo Can_Err

fglbNew = False
Call Display_Value

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_OVERTIME_MASTER", "Cancel")
Call RollBack

End Sub

Public Sub clkNew()
Dim SQLQ As String


On Error GoTo AddN_Err

fglbNew = True
Call Set_Control("B", Me)
Call SET_UP_MODE


lblCNum.Caption = "001"
'lblPOSID.Caption = glbPos$
If glbCompSerial <> "S/N - 2425W" Then   'Ticket #18223 - Four Villages CHC
    dlpDateRange(0).SetFocus
Else
    dlpDateRange(0).SetFocus
End If


Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "clkNew", "HR_OVERTIME_MASTER", "Add")
Call RollBack

End Sub

Public Function clkDelete()
Dim a As Integer, Msg As String, INo&

If rsDATA.BOF And rsDATA.EOF Then
    MsgBox "Nothing to Delete"
    Exit Function
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Record?  "
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Function

'Delete Employee overtime records as well matching the criteria.
Dim rsEmp As New ADODB.Recordset
Dim SQLQ
SQLQ = "DELETE FROM HR_OVERTIME_BANK WHERE OT_EMPNBR IN ("
SQLQ = SQLQ & "SELECT ED_EMPNBR FROM HREMP WHERE 1=1"
If Len(clpCode(1).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_ORG = '" & clpCode(1).Text & "'"
End If
If Len(clpCode(0).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_EMP = '" & clpCode(0).Text & "'"
End If
If Len(clpPT.Text) > 0 Then
    SQLQ = SQLQ & " AND ED_PT = '" & clpPT.Text & "'"
End If
'Ticket #15753
If Len(clpCode(2).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_LOC = '" & clpCode(2).Text & "'"
End If
If Len(clpCode(3).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_REGION = '" & clpCode(3).Text & "'"
End If
If Len(clpCode(4).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_ADMINBY = '" & clpCode(4).Text & "'"
End If
If Len(clpCode(5).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(5).Text & "'"
End If

SQLQ = SQLQ & ")"
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

'Delete the Overtime Rule
gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

'If Data1.Recordset.EOF And Data1.Recordset.BOF Then
'    Call Display_Value
'End If

fglbNew = False
Call SET_UP_MODE

If Data1.Recordset.EOF Then
    CmdRecalc(1).Enabled = False
    CmdUpdate.Enabled = False
    
    If glbCompSerial = "S/N - 2276W" Then
        cmdUpdVadim.Enabled = False
    End If
    
    'Ticket #19998 - Four Villages
    If glbCompSerial = "S/N - 2425W" Then
        cmdForfeitHrs.Enabled = False
    End If
Else
    CmdRecalc(1).Enabled = True
    CmdUpdate.Enabled = True
    
    If glbCompSerial = "S/N - 2276W" Then
        cmdUpdVadim.Enabled = True
    End If
    
    'Ticket #19998 - Four Villages
    If glbCompSerial = "S/N - 2425W" Then
        cmdForfeitHrs.Enabled = True
    End If
End If

Exit Function

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "clkDelete", "HR_OVERTIME_MASTER", "Delete")
Call RollBack
End Function

Public Sub clkReport(Destination As DestinationConstants)
Dim RHeading As String

RHeading = Me.Caption
RHeading = Mid(RHeading, 1, InStr(RHeading, "-"))
'RHeading = RHeading & " " & lblPosDesc.Caption

Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = Destination
Me.vbxCrystal.Action = 1
End Sub

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum

If fglbNew Then
    UpdateState = NewRecord
    TF = True
ElseIf Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False

dlpDateRange(0).Enabled = TF
dlpDateRange(1).Enabled = TF
clpCode(1).Enabled = TF
clpCode(0).Enabled = TF
clpPT.Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
clpCode(4).Enabled = TF
clpCode(5).Enabled = TF
txtMaxBankHrs.Enabled = TF
txtMultiplier.Enabled = TF
txtAddress.Enabled = TF

End Sub

Public Sub Display_Value()
Dim SQLQ

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then rsDATA.Close
    rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "SELECT * FROM HR_OVERTIME_MASTER "
    SQLQ = SQLQ & " WHERE OM_ID = " & Data1.Recordset!OM_ID
    If rsDATA.State <> 0 Then rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
End If
Call SET_UP_MODE
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelatePOS
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = GetMassUpdateSecurities("OvertimeMaster_MassUpdate", glbUserID)
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

'Private Sub WeeklyOTCalculation_LeedsnGrenville()
'    Dim SQLQ As String
'    Dim rsAttend As New ADODB.Recordset
'
'    'Create the EXPO code - Expired Overtime, in the HRTABL
'
'
'    'Forfeit the expired hours based on the Expiry Date as of the date entered
'    SQLQ = "UPDATE HR_ATTENDANCE SET AD_REASON = 'EXPO' WHERE AD_BANKHRS_EXP <= " & Date_SQL(dlpAsOfDate.Text)
'    gdbAdoIhr001.Execute SQLQ
'
'    'Calculate the OT at the end of the week (As of Date)
'    SQLQ = "SELECT SUM(AD_HRS)"
'    SQLQ = SQLQ & " WHERE AD_BANKHRS_EXP > " & Date_SQL(dlpAsOfDate.Text)
'    SQLQ = SQLQ & " AD_REASON = 'OT'"
'
'    'AD_SOURCE update


'                SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsEmp("ED_EMPNBR")
'                SQLQ = SQLQ & " AND AD_DOA = " & Date_SQL(MonthLastDate(dlpMnthEndDate))
'                SQLQ = SQLQ & " AND AD_REASON = 'CTF'"
'                rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                If rsAttend.EOF Then
'                    'No CTF record found for the month end
'
'                    'Make sure CTF code exists if not then add the code
'                    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_KEY = 'CTF' "
'                    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                    If rsTABL.EOF Then
'                        rsTABL.AddNew
'                        rsTABL("TB_COMPNO") = "001"
'                        rsTABL("TB_NAME") = "ADRE"
'                        rsTABL("TB_KEY") = "CTF"
'                        rsTABL("TB_DESC") = "FORFEITED HOURS"
'                        rsTABL("TB_LDATE") = Date
'                        rsTABL("TB_LTIME") = Time$
'                        rsTABL("TB_LUSER") = glbUserID
'                        rsTABL.Update
'                    End If
'                    rsTABL.Close
'
'                    'Add a new record with CTF hours
'                    rsAttend.AddNew
'                    rsAttend("AD_COMPNO") = "001"
'                    rsAttend("AD_EMPNBR") = rsEmp("ED_EMPNBR")
'
'                    'Update with Salary info.
'                    SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & rsEmp("ED_EMPNBR")
'                    rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
'                    If Not rsCurSal.BOF Then
'                        If rsCurSal("SH_SALARY") > 0 Then
'                            rsAttend("AD_SALARY") = rsCurSal("SH_SALARY")
'                            rsAttend("AD_SALCD") = rsCurSal("SH_SALCD")
'                        End If
'                    End If
'                    rsCurSal.Close
'                    Set rsCurSal = Nothing
'
'                    'Update with Position info.
'                    SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS,JH_REPTAU,JH_PAYROLL_ID,JH_SHIFT,JH_GLNO,JH_ORG FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & rsEmp("ED_EMPNBR")
'                    rsCurPos.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
'                    If Not rsCurPos.EOF Then
'                        rsAttend("AD_JOB") = rsCurPos("JH_JOB")
'                        rsAttend("AD_DHRS") = rsCurPos("JH_DHRS")
'                        rsAttend("AD_WHRS") = rsCurPos("JH_WHRS")
'                        rsAttend("AD_SUPER") = rsCurPos("JH_REPTAU")
'                        rsAttend("AD_PAYROLL_ID") = rsCurPos("JH_PAYROLL_ID")
'                        rsAttend("AD_SHIFT") = rsCurPos("JH_SHIFT")
'                        rsAttend("AD_GLNO") = rsCurPos("JH_GLNO")
'                        rsAttend("AD_ORG") = rsCurPos("JH_ORG")
'                    End If
'                    rsCurPos.Close
'                    Set rsCurPos = Nothing
'                End If
'
'                'Update with the CTF hours
'                rsAttend("AD_DOA") = CVDate(MonthLastDate(dlpMnthEndDate))
'                rsAttend("AD_REASON") = "CTF"
'                rsAttend("AD_HRS") = xForfeitHrs
'                rsAttend("AD_COMM") = "Weekly Expired Hours."
'                rsAttend("AD_LUSER") = glbUserID
'                rsAttend("AD_LDATE") = Date
'                rsAttend("AD_LTIME") = Time$
'
'                'Weekly Expired Hours
'                rsAttend("AD_SOURCE") = "IHREXP"
'
'                rsAttend.Update
'
'                rsAttend.Close
'                Set rsAttend = Nothing
'            End If
'End Sub
