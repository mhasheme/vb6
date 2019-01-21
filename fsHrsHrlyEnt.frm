VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSHrsHrlyEnt 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Hours Based Entitlement Master"
   ClientHeight    =   6990
   ClientLeft      =   2565
   ClientTop       =   525
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
   ForeColor       =   &H00C0C0C0&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6990
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin VB.Frame VacFram03 
      BorderStyle     =   0  'None
      Height          =   5475
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   11415
      Begin INFOHR_Controls.CodeLookup clpProv 
         Height          =   285
         Left            =   2160
         TabIndex        =   36
         Tag             =   "31-Province - Code"
         Top             =   3135
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
      End
      Begin Threed.SSCheck chkManual 
         Height          =   255
         Left            =   5535
         TabIndex        =   9
         Top             =   3960
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Exclude from Update All"
         ForeColor       =   -2147483640
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
      Begin INFOHR_Controls.DateLookup dlpAsOf 
         Height          =   285
         Left            =   10440
         TabIndex        =   10
         Tag             =   "40-As of Date"
         Top             =   5220
         Visible         =   0   'False
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   945
         TabIndex        =   2
         Tag             =   "00-Enter Union Code"
         Top             =   2805
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOR"
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   945
         TabIndex        =   1
         Tag             =   "00-Specific Department Desired"
         Top             =   2460
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   7
         LookupType      =   2
      End
      Begin INFOHR_Controls.CodeLookup clpPT 
         Height          =   285
         Left            =   6450
         TabIndex        =   4
         Tag             =   "EDPT-Category"
         Top             =   2460
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDPT"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   6450
         TabIndex        =   6
         Tag             =   "00-Section - Code"
         Top             =   3135
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   945
         TabIndex        =   3
         Tag             =   "00-Enter Location Code"
         Top             =   3135
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   0
         Left            =   2100
         TabIndex        =   7
         Tag             =   "40-From Date"
         Top             =   3945
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1210
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   1
         Left            =   3870
         TabIndex        =   8
         Tag             =   "40-To Date"
         Top             =   3945
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1210
      End
      Begin INFOHR_Controls.DateLookup dlpAccDateRange 
         Height          =   285
         Index           =   0
         Left            =   2100
         TabIndex        =   11
         Tag             =   "40-From Date"
         Top             =   4320
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1210
      End
      Begin INFOHR_Controls.DateLookup dlpAccDateRange 
         Height          =   285
         Index           =   1
         Left            =   3870
         TabIndex        =   12
         Tag             =   "40-To Date"
         Top             =   4320
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1210
      End
      Begin Threed.SSFrame frmType 
         Height          =   525
         Left            =   2100
         TabIndex        =   13
         Top             =   4560
         Width           =   3310
         _Version        =   65536
         _ExtentX        =   5856
         _ExtentY        =   926
         _StockProps     =   14
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin VB.TextBox txtUpdMethod 
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
            Height          =   195
            Left            =   2880
            MaxLength       =   1
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin Threed.SSOption Replace 
            Height          =   195
            Left            =   1935
            TabIndex        =   32
            Tag             =   "Replace Entitlement Amount"
            Top             =   255
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Replace"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption Accum 
            Height          =   195
            Left            =   300
            TabIndex        =   33
            TabStop         =   0   'False
            Tag             =   "Add to Exist Entitlements"
            Top             =   255
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Accumulate"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
         Bindings        =   "fsHrsHrlyEnt.frx":0000
         Height          =   1695
         Left            =   0
         OleObjectBlob   =   "fsHrsHrlyEnt.frx":0014
         TabIndex        =   0
         Top             =   0
         Width           =   9135
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   6450
         TabIndex        =   5
         Tag             =   "01-Entitlement Code"
         Top             =   2805
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ADRE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   6450
         TabIndex        =   37
         Tag             =   "00-Specific Employment Status Desired"
         Top             =   2130
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDEM"
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         Height          =   285
         Left            =   945
         TabIndex        =   38
         Tag             =   "00-Specific Division Desired"
         Top             =   2130
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin VB.Label lblCriteria 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Entitlement Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   4920
         TabIndex        =   35
         Top             =   2865
         Width           =   1455
      End
      Begin VB.Label lblCriteria 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Update Method"
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
         Left            =   30
         TabIndex        =   34
         Top             =   4785
         Width           =   1110
      End
      Begin VB.Label lblAccPeriod 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Accrual Date Range"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   30
         Top             =   4365
         Width           =   1740
      End
      Begin VB.Label lblPeriod 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Entitlement Period"
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
         Left            =   30
         TabIndex        =   29
         Top             =   3990
         Width           =   1275
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
         Left            =   30
         TabIndex        =   28
         Top             =   2175
         Width           =   555
      End
      Begin VB.Label lblDept 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         Left            =   30
         TabIndex        =   27
         Top             =   2505
         Width           =   825
      End
      Begin VB.Label lblUnion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
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
         Left            =   30
         TabIndex        =   26
         Top             =   2850
         Width           =   420
      End
      Begin VB.Label lblCriteria 
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
         Index           =   3
         Left            =   4920
         TabIndex        =   25
         Top             =   2175
         Width           =   1350
      End
      Begin VB.Label lblAsOf 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9030
         TabIndex        =   24
         Top             =   5250
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblSelCri 
         Caption         =   "Selection Criteria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   1800
         Width           =   1575
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
         Left            =   4920
         TabIndex        =   22
         Top             =   2505
         Width           =   630
      End
      Begin VB.Label lblSection 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Left            =   4920
         TabIndex        =   21
         Top             =   3180
         Width           =   1260
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
         Left            =   30
         TabIndex        =   20
         Top             =   3180
         Width           =   615
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   645
      Left            =   0
      TabIndex        =   14
      Top             =   6345
      Width           =   11760
      _Version        =   65536
      _ExtentX        =   20743
      _ExtentY        =   1138
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
      Begin VB.CommandButton cmdUpdateAll 
         Caption         =   "Update All"
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   120
         Width           =   1665
      End
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&Update Entitlement"
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Tag             =   "Change all matching records to the above"
         Top             =   120
         Width           =   1905
      End
      Begin VB.CommandButton CmdRecalc 
         Appearance      =   0  'Flat
         Caption         =   "R&ecalculate"
         Height          =   375
         Left            =   6360
         TabIndex        =   18
         Tag             =   "Recalculation"
         Top             =   120
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton cmdPrintAll 
         Appearance      =   0  'Flat
         Caption         =   "Print &All"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Tag             =   "Print all Vacation Entitlement Report"
         Top             =   120
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   405
         Left            =   7800
         Top             =   0
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   714
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   1
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         BoundReportHeading=   "RGELIST"
         BoundReportFooter=   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
End
Attribute VB_Name = "frmSHrsHrlyEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fTablHREMP As New ADODB.Recordset         ' table view of HREMP
Dim snapEntitle As New ADODB.Recordset     'user vier
Dim fglbWDate$, fglbWDateS$
Dim fglbAsOf As Date
Dim Actn

Dim fglbSDate As Variant
Dim fglbMaxRange%
Dim fglbCompMonthly%

Dim fglbMaxRanges%
Dim glbFrmCaption$, glbErrNum&

Dim ControlsShown As Boolean
Dim ODIV, ODept, oOrg, oAsOf, oEMP, oEmpMode, oGRPCE, oFDate, OTDate
Dim OSection, OLoc, oHETYPE
Dim OFromDate, OToDate
Dim OAccFromDate, OAccToDate
Dim FlagRefresh As Boolean

Dim fglbESQLQ, fglbVSQLQ, fglbWSQLQ
Dim fglbNew As Boolean
Dim fglbRunTimes

Private Function chkMUEntitle()
Dim x%, Y%

chkMUEntitle = False

On Error GoTo chkMUEntitle_Err
For x% = 0 To 4
If Len(clpCode(x%).Text) > 0 And clpCode(x%).Caption = "Unassigned" Then
    MsgBox "If Code entered it must be known"
    clpCode(x%).SetFocus
    Exit Function
End If
Next x%


If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox lStr("Invalid Department")
     clpDept.SetFocus
    Exit Function
End If
If Len(clpDiv.Text) < 1 Then
    If glbDIVCount = 1 And glbLinamar Then
        MsgBox lStr("Division is required field")
         clpDiv.SetFocus
        Exit Function
    End If
Else
    If clpDiv.Caption = "Unassigned" Then
        MsgBox lStr("Invalid Division")
         clpDiv.SetFocus
        Exit Function
    End If
End If
'Hemu - 05/13/2003 Begin
If clpPT.Caption = "Unassigned" Then
    MsgBox "Invalid " & lblPT.Caption
    clpPT.SetFocus
    Exit Function
End If

'Entitlement Period
'Sam 02/02/2006
'Ticket #15276 - Commented
If Len(dlpDateRange(0).Text) > 0 Then
    If Not IsDate(dlpDateRange(0).Text) Then
        MsgBox "Invalid Entitlement Period From Date"
        dlpDateRange(0).SetFocus
        Exit Function
    End If
Else
    MsgBox "Entitlement Period From Date is mandatory field"
    dlpDateRange(0).SetFocus
    Exit Function
End If

If Len(dlpDateRange(1).Text) > 0 Then
    If Not IsDate(dlpDateRange(1).Text) Then
        MsgBox "Invalid Entitlement Period To Date"
        dlpDateRange(1).SetFocus
        Exit Function
    End If
Else
    MsgBox "Entitlement Period To Date is mandatory field"
    dlpDateRange(1).SetFocus
    Exit Function
End If

If IsDate(dlpDateRange(0).Text) And IsDate(dlpDateRange(1).Text) Then
If CVDate(dlpDateRange(0).Text) > CVDate(dlpDateRange(1).Text) Then
    MsgBox "Entitlement Period From Date cannot be greater than Entitlement Period To Date"
    dlpDateRange(0).SetFocus
    Exit Function
End If
End If

'Accrual Period
If Len(dlpAccDateRange(0).Text) > 0 Then
    If Not IsDate(dlpAccDateRange(0).Text) Then
        MsgBox "Invalid Accrual Date Range - From Date"
        dlpAccDateRange(0).SetFocus
        Exit Function
    End If
Else
    MsgBox "Accrual Date Range - From Date is mandatory field"
    dlpAccDateRange(0).SetFocus
    Exit Function
End If

If Len(dlpAccDateRange(1).Text) > 0 Then
    If Not IsDate(dlpAccDateRange(1).Text) Then
        MsgBox "Invalid Accrual Date Range - To Date"
        dlpAccDateRange(1).SetFocus
        Exit Function
    End If
Else
    MsgBox "Accrual Date Range - To Date is mandatory field"
    dlpAccDateRange(1).SetFocus
    Exit Function
End If

If IsDate(dlpAccDateRange(0).Text) And IsDate(dlpAccDateRange(1).Text) Then
If CVDate(dlpAccDateRange(0).Text) > CVDate(dlpAccDateRange(1).Text) Then
    MsgBox "Accrual Date Range - From Date cannot be greater than Accrual Date Range - To Date"
    dlpAccDateRange(0).SetFocus
    Exit Function
End If
End If

'If Len(dlpAsOf.Text) > 0 Then
'  If Not IsDate(dlpAsOf.Text) Then
'    MsgBox "Invalid Effective Date"
'    dlpAsOf.SetFocus
'    Exit Function
'  End If
'Else
'    'If UCase(glbCompEntSick$) = "A" Then
'    '    If glbLinamar Then
'            MsgBox "Effective Date is required field"
'            dlpAsOf.SetFocus
'            Exit Function
'    '    End If
'    'End If
'End If

'Frank 05/13/2004 Ticket#
If glbWFC Then
    If Len(clpCode(3).Text) = 0 Then
        MsgBox lStr("Section is required field")
        clpCode(3).SetFocus
        Exit Function
    End If
End If

fglbMaxRanges% = 0  ' 0 is first range

chkMUEntitle = True

Exit Function

chkMUEntitle_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkMUEntitle", "HRHRSVACENT", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Sub cmdCancel_Click()
    fglbNew = False
    Data1.Refresh
    
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    
    Call Display_Value
    
    vbxTrueGrid.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Sub cmdDelete_Click()
    Dim SQLQ, Msg, A%
    
    If Data1.Recordset.BOF And Data1.Recordset.EOF Then
        MsgBox "Nothing to Delete"
        Exit Sub
    End If
    Msg = "Are You Sure You Want To Delete "
    Msg = Msg & Chr(10) & "The Hours Based Entitlement Rules?  "
    
    A% = MsgBox(Msg, 36, "Confirm Delete")
    If A% <> 6 Then Exit Sub
    
    Call getWSQLQ("C")
    SQLQ = "DELETE FROM HRHRSVACENT WHERE " & fglbVSQLQ
    
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    
    Data1.Refresh
    
    Call Display_Value

End Sub

Sub cmdModify_Click()
    ODIV = clpDiv.Text
    ODept = clpDept.Text
    oOrg = clpCode(0).Text
    
    'Franks 04/08/03 Ticket# 3943
    'Fix the problem: enter or change Effective Date first, click Edit and then Save,
    'it create another record
    oAsOf = ""
'    If Not Data1.Recordset.EOF Then
'        If Not IsNull(Data1.Recordset("VE_EDATE")) Then
'            oAsOf = Data1.Recordset("VE_EDATE")
'        End If
'    End If

    'Sam 02/02/2006
    OFromDate = dlpDateRange(0).Text
    OToDate = dlpDateRange(1).Text
    'Sam 02/02/2006
    
    OAccFromDate = dlpAccDateRange(0).Text
    OAccToDate = dlpAccDateRange(1).Text
    
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
        OLoc = clpProv.Text
    Else
        OLoc = clpCode(4).Text
    End If
    OSection = clpCode(3).Text
    oEMP = clpCode(1).Text
    oEmpMode = clpPT.Text
    oGRPCE = clpCode(2).Text
    Actn = "M"
End Sub

Sub cmdNew_Click()
    Dim x
    
    'Sam 02/2/2006
    dlpDateRange(0).Text = ""
    dlpDateRange(1).Text = ""
    'Sam 02/2/2006
    
    dlpAccDateRange(0).Text = ""
    dlpAccDateRange(1).Text = ""
    
    clpDiv.Text = ""
    clpDept.Text = ""
    clpCode(0).Text = ""
    dlpAsOf.Text = ""
    clpCode(1).Text = ""
    clpCode(2).Text = ""
    clpCode(3).Text = ""
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
        clpProv.Text = ""
    Else
        clpCode(4).Text = ""
    End If
    clpPT.Text = ""
    Actn = "A"
    fglbNew = True
      
    Call SET_UP_MODE
    
    clpDiv.SetFocus

End Sub

Sub cmdOK_Click()
    Dim x%, Y%, xUnion, xPT, SQLQ, SQLQW
    Dim xStr
    Dim rsVE As New ADODB.Recordset
    Dim rsVT As New ADODB.Recordset
    Dim glbiOneWhere As Boolean
    Dim bmk As Variant
    
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        bmk = 0 'Ticket #11885 Frank Oct 11th, 2006
    Else
        bmk = Data1.Recordset.Bookmark
    End If
    
    If Not chkMUEntitle() Then Exit Sub
    

'    If Actn = "M" Then
'        Call getWSQLQ("O")
'        SQLQ = "DELETE FROM HRHRSVACENT WHERE " & fglbVSQLQ
'        gdbAdoIhr001.BeginTrans
'        gdbAdoIhr001.Execute SQLQ
'        gdbAdoIhr001.CommitTrans
'    Else
'        Call getWSQLQ("C")
'        SQLQ = "SELECT * FROM HRHRSVACENT WHERE " & fglbVSQLQ
'        rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
'        If Not rsVT.EOF Then
'            MsgBox "You can not add duplicate record"
'            clpDiv.SetFocus
'            Exit Sub
'        End If
'    End If
'    gdbAdoIhr001.BeginTrans
'    SQLQ = "SELECT * FROM HRHRSVACENT"
'    rsVE.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
'        rsVE.AddNew
'        rsVE("VH_ORDER") = x + 1
'        rsVE("VH_ORG_TABL") = "EDOR"
'        rsVE("VH_ORG") = clpCode(0).Text
'        rsVE("VH_PT") = clpPT.Text
'        rsVE("VH_DIV") = clpDiv.Text
'        rsVE("VH_DEPT") = clpDept.Text
'        rsVE("VH_EMP_TABL") = "EDEM"
'        rsVE("VH_EMP") = clpCode(1).Text
'        rsVE("VH_SECTION") = clpCode(3).Text
'        If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
'            rsVE("EH_LOC") = clpProv.Text
'        Else
'            rsVE("EH_LOC") = clpCode(4).Text
'        End If
'        'rsVE("VE_EDATE") = dlpAsOf.Text
'
'        If Len(dlpDateRange(0).Text) > 0 Then
'            rsVE("VH_FRDATE") = dlpDateRange(0).Text
'        End If
'        If Len(dlpDateRange(1).Text) > 0 Then
'            rsVE("VH_TODATE") = dlpDateRange(1).Text
'        End If
'
'        If Len(dlpAccDateRange(0).Text) > 0 Then
'            rsVE("VH_ACCFRDATE") = dlpAccDateRange(0).Text
'        End If
'        If Len(dlpAccDateRange(1).Text) > 0 Then
'            rsVE("VH_ACCTODATE") = dlpAccDateRange(1).Text
'        End If
'
'        rsVE("VH_GRPCD_TABL") = "JBGC"
'        rsVE("VH_GRPCD") = clpCode(2).Text
''        If optD(X%) Then rsVE("VE_TYPE") = "D"
''        If optH(X%) Then rsVE("VE_TYPE") = "H"
''        If optF(X%) Then rsVE("VE_TYPE") = "F"
''        rsVE("VE_MAX") = medMax(X%)
'        rsVE("VH_MANUAL") = chkManual.Value
'        rsVE("VH_UPDMETHOD") = txtUpdMethod.Text
'        rsVE.Update
'    rsVE.Close
'    gdbAdoIhr001.CommitTrans
    
    'If Not glbSQL and not glboracle Then Call Pause(0.5)
    
    Data1.Refresh
    
    If Not bmk = 0 Then
        Data1.Recordset.Bookmark = bmk
    End If
    
    fglbNew = False
    
    Call Display_Value

End Sub

Sub cmdPrint_Click()
    Dim RHeading As String, xReport, x%
    Dim SQLQ
    Dim dtYYY%, dtMM%, dtDD%
    'cmdPrint.Enabled = False
    
    Me.vbxCrystal.Reset
    Me.vbxCrystal.WindowTitle = "Hours Based Entitlement Master Report"
    
    Call setRptLabel(Me, 0) '1)
    
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 5
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rghrsvacentmst.rpt"
    
    SQLQ = "(1=1) "
    If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND {HRHRSVACENT.VH_DIV} = '" & clpDiv.Text & "'"
    If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND {HRHRSVACENT.VH_DEPT} = '" & clpDept.Text & "'"
    If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND {HRHRSVACENT.VH_ORG} = '" & clpCode(0).Text & "'"
'    If Len(dlpAsOf.Text) > 0 Then
'        dtYYY% = Year(dlpAsOf.Text)
'        dtMM% = Month(dlpAsOf.Text)
'        dtDD% = Day(dlpAsOf.Text)
'        SQLQ = SQLQ & " AND {HR_VACATION_INCR.VC_EDATE} = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
'    End If
    If Len(clpCode(1).Text) > 0 Then SQLQ = SQLQ & " AND {HRHRSVACENT.VH_EMP} = '" & clpCode(1).Text & "'"
    If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND {HRHRSVACENT.VH_PT} = '" & clpPT.Text & "' "
    If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND {HRHRSVACENT.VH_GRPCD} = '" & clpCode(2).Text & "'"
    If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND {HRHRSVACENT.VH_SECTION} = '" & clpCode(3).Text & "'"
    If Len(clpCode(4).Text) > 0 Then 'SQLQ = SQLQ & " AND {HRHRSVACENT.VH_LOC} = '" & clpCode(4).Text & "'"
        If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
            SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_LOC} = '" & clpProv.Text & "'"
        Else
            SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_LOC} = '" & clpCode(4).Text & "'"
        End If
    End If
    If Len(dlpDateRange(0).Text) > 0 Then
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
        SQLQ = SQLQ & " AND {HRHRSVACENT.VH_FRDATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    End If
    If Len(dlpDateRange(1).Text) > 0 Then
        dtYYY% = Year(dlpDateRange(1).Text)
        dtMM% = month(dlpDateRange(1).Text)
        dtDD% = Day(dlpDateRange(1).Text)
        SQLQ = SQLQ & " AND {HRHRSVACENT.VH_TODATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    End If
  
    If Len(dlpAccDateRange(0).Text) > 0 Then
        dtYYY% = Year(dlpAccDateRange(0).Text)
        dtMM% = month(dlpAccDateRange(0).Text)
        dtDD% = Day(dlpAccDateRange(0).Text)
        SQLQ = SQLQ & " AND {HRHRSVACENT.VH_ACCFRDATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    End If
    If Len(dlpAccDateRange(1).Text) > 0 Then
        dtYYY% = Year(dlpAccDateRange(1).Text)
        dtMM% = month(dlpAccDateRange(1).Text)
        dtDD% = Day(dlpAccDateRange(1).Text)
        SQLQ = SQLQ & " AND {HRHRSVACENT.VH_ACCTODATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    End If
  
    Me.vbxCrystal.SelectionFormula = SQLQ
    Me.vbxCrystal.Destination = 1
    Me.vbxCrystal.Action = 1
    
    'cmdPrint.Enabled = True

End Sub

Sub cmdView_Click()
    Dim RHeading As String, xReport, x%
    Dim SQLQ
    Dim dtYYY%, dtMM%, dtDD%
    'cmdPrint.Enabled = False
    
    Me.vbxCrystal.Reset
    Me.vbxCrystal.WindowTitle = "Hours Based Entitlement Master Report"
    
    Call setRptLabel(Me, 0) '1)
    
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 5
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rghrsvacentmst.rpt"
    
    SQLQ = "(1=1) "
    If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND {HRHRSVACENT.VH_DIV} = '" & clpDiv.Text & "'"
    If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND {HRHRSVACENT.VH_DEPT} = '" & clpDept.Text & "'"
    If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND {HRHRSVACENT.VH_ORG} = '" & clpCode(0).Text & "'"
'    If Len(dlpAsOf.Text) > 0 Then
'        dtYYY% = Year(dlpAsOf.Text)
'        dtMM% = Month(dlpAsOf.Text)
'        dtDD% = Day(dlpAsOf.Text)
'        SQLQ = SQLQ & " AND {HR_VACATION_INCR.VC_EDATE} = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
'    End If
    If Len(clpCode(1).Text) > 0 Then SQLQ = SQLQ & " AND {HRHRSVACENT.VH_EMP} = '" & clpCode(1).Text & "'"
    If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND {HRHRSVACENT.VH_PT} = '" & clpPT.Text & "' "
    If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND {HRHRSVACENT.VH_GRPCD} = '" & clpCode(2).Text & "'"
    If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND {HRHRSVACENT.VH_SECTION} = '" & clpCode(3).Text & "'"
    If Len(clpCode(4).Text) > 0 Then 'SQLQ = SQLQ & " AND {HRHRSVACENT.VH_LOC} = '" & clpCode(4).Text & "'"
        If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
            SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_LOC} = '" & clpProv.Text & "'"
        Else
            SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_LOC} = '" & clpCode(4).Text & "'"
        End If
    End If
    If Len(dlpDateRange(0).Text) > 0 Then
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
        SQLQ = SQLQ & " AND {HRHRSVACENT.VH_FRDATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    End If
    If Len(dlpDateRange(1).Text) > 0 Then
        dtYYY% = Year(dlpDateRange(1).Text)
        dtMM% = month(dlpDateRange(1).Text)
        dtDD% = Day(dlpDateRange(1).Text)
        SQLQ = SQLQ & " AND {HRHRSVACENT.VH_TODATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    End If
            
    If Len(dlpAccDateRange(0).Text) > 0 Then
        dtYYY% = Year(dlpAccDateRange(0).Text)
        dtMM% = month(dlpAccDateRange(0).Text)
        dtDD% = Day(dlpAccDateRange(0).Text)
        SQLQ = SQLQ & " AND {HRHRSVACENT.VH_ACCFRDATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    End If
    If Len(dlpAccDateRange(1).Text) > 0 Then
        dtYYY% = Year(dlpAccDateRange(1).Text)
        dtMM% = month(dlpAccDateRange(1).Text)
        dtDD% = Day(dlpAccDateRange(1).Text)
        SQLQ = SQLQ & " AND {HRHRSVACENT.VH_ACCTODATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    End If
            
    Me.vbxCrystal.SelectionFormula = SQLQ
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.Action = 1
    'cmdPrint.Enabled = True
End Sub

Private Sub Accum_Click(Value As Integer)
    If Accum.Value = True Then
        txtUpdMethod.Text = "A"
    ElseIf Replace.Value = True Then
        txtUpdMethod.Text = "R"
    End If
End Sub

Private Sub cmdPrintAll_Click()
Dim RHeading As String, xReport, x%
Dim SQLQ
Dim dtYYY%, dtMM%, dtDD%
cmdPrintAll.Enabled = False

Me.vbxCrystal.Reset

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

Me.vbxCrystal.WindowTitle = "Hours Based Entitlement Master Report"
Call setRptLabel(Me, 0) '1)
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For x% = 0 To 5
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next
End If
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rghrsvacentmst.rpt"
Me.vbxCrystal.Action = 1

cmdPrintAll.Enabled = True
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo Mod_Err
Dim sFlag As Boolean

If Not gSec_Upd_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

If Not chkMUEntitle() Then Exit Sub

'Added by Bryan 25/Oct/05 Ticket#9560
'made the code a separate sub because it's being used in two places
sFlag = DoWork

Data1.Refresh

Call Display_Value

If sFlag Then
    MsgBox "Update Completed Successfully", vbInformation + vbOKOnly, "Hours Based Entitlement Master"
End If

Screen.MousePointer = DEFAULT

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdateAll", "Single", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
     RollBack
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub cmdUpdate_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Function CR_SnapEntitle()
Dim SQLQ As String
Dim SQLQ2 As String
Dim snapMultiEmp As New ADODB.Recordset

CR_SnapEntitle = False
On Error GoTo CR_SnapEntitle_Err


Call getWSQLQ("")

SQLQ = "SELECT ED_EMPNBR,ED_VACPC,ED_PVAC,ED_VAC,ED_VACT,ED_PSICK,ED_SICK,ED_SICKT,ED_EFDATES,ED_ETDATES, HREMP.ED_ANNVAC, HREMP.ED_ANNSICK, "
SQLQ = SQLQ & " ED_DIV,ED_PT, ED_SECTION, ED_LOC, ED_EMP,"
SQLQ = SQLQ & " ED_HIRECODE," 'County of Brant Ticket #12525
SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME "
SQLQ = SQLQ & " FROM HREMP WHERE " & fglbESQLQ
If Len(clpCode(2).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_EMPNBR IN "
    SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
    SQLQ = SQLQ & " WHERE JB_GRPCD = '" & clpCode(2).Text & "') "
End If

If snapEntitle.State <> 0 Then snapEntitle.Close
If glbOracle Then
    snapEntitle.CursorLocation = adUseServer
End If
snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic

CR_SnapEntitle = True

Exit Function

CR_SnapEntitle_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapEntitle", "HrsBasedEntitlement/EMP", "Select")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub cmdUpdateAll_Click()
On Error GoTo Mod_Err

Dim c As Long
Dim failed As String

If Not gSec_Upd_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

failed = ""
c = 1
If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
    Data1.Recordset.MoveFirst
    Do
        Call Display_Value

        'made the DoWork a separate sub because it's being used in two places
        If chkManual.Value = False Then
            If chkMUEntitle() Then
                If DoWork = False Then
                    failed = failed & "Rule " & CStr(c) & ": "
                    If Not IsNull(Data1.Recordset("VH_DIV")) Then failed = failed & Data1.Recordset("VH_DIV") & ", "
                    If Not IsNull(Data1.Recordset("VH_DEPT")) Then failed = failed & Data1.Recordset("VH_DEPT") & ", "
                    If Not IsNull(Data1.Recordset("VH_ORG")) Then failed = failed & Data1.Recordset("VH_ORG") & ", "
                    'If Not IsNull(Data1.Recordset("VE_EDATE")) Then failed = failed & Data1.Recordset("VE_EDATE") & ", "
                    If Not IsNull(Data1.Recordset("VH_EMP")) Then failed = failed & Data1.Recordset("VH_EMP") & ", "
                    If Not IsNull(Data1.Recordset("VH_PT")) Then failed = failed & Data1.Recordset("VH_PT") & ", "
                    If Not IsNull(Data1.Recordset("VH_GRPCD")) Then failed = failed & Data1.Recordset("VH_GRPCD") & ", "
                    If Not IsNull(Data1.Recordset("VH_LOC")) Then failed = failed & Data1.Recordset("VH_LOC") & ", "
                    If Not IsNull(Data1.Recordset("VH_SECTION")) Then failed = failed & Data1.Recordset("VH_SECTION") & ", "
                    If Not IsNull(Data1.Recordset("VH_FRDATE")) Then failed = failed & Data1.Recordset("VH_FRDATE") & ", "
                    If Not IsNull(Data1.Recordset("VH_TODATE")) Then failed = failed & Data1.Recordset("VH_TODATE") & ", "
                    If Not IsNull(Data1.Recordset("VH_ACCFRDATE")) Then failed = failed & Data1.Recordset("VH_ACCFRDATE") & ", "
                    If Not IsNull(Data1.Recordset("VH_ACCTODATE")) Then failed = failed & Data1.Recordset("VH_ACCTODATE") & ", "
                    failed = Left(failed, Len(failed) - 2) & vbCrLf
                End If
            Else
                failed = failed & "Rule " & CStr(c) & ": "
                If Not IsNull(Data1.Recordset("VH_DIV")) Then failed = failed & Data1.Recordset("VH_DIV") & ", "
                If Not IsNull(Data1.Recordset("VH_DEPT")) Then failed = failed & Data1.Recordset("VH_DEPT") & ", "
                If Not IsNull(Data1.Recordset("VH_ORG")) Then failed = failed & Data1.Recordset("VH_ORG") & ", "
                'If Not IsNull(Data1.Recordset("VE_EDATE")) Then failed = failed & Data1.Recordset("VE_EDATE") & ", "
                If Not IsNull(Data1.Recordset("VH_EMP")) Then failed = failed & Data1.Recordset("VH_EMP") & ", "
                If Not IsNull(Data1.Recordset("VH_PT")) Then failed = failed & Data1.Recordset("VH_PT") & ", "
                If Not IsNull(Data1.Recordset("VH_GRPCD")) Then failed = failed & Data1.Recordset("VH_GRPCD") & ", "
                If Not IsNull(Data1.Recordset("VH_LOC")) Then failed = failed & Data1.Recordset("VH_LOC") & ", "
                If Not IsNull(Data1.Recordset("VH_SECTION")) Then failed = failed & Data1.Recordset("VH_SECTION") & ", "
                If Not IsNull(Data1.Recordset("VH_FRDATE")) Then failed = failed & Data1.Recordset("VH_FRDATE") & ", "
                If Not IsNull(Data1.Recordset("VH_TODATE")) Then failed = failed & Data1.Recordset("VH_TODATE") & ", "
                If Not IsNull(Data1.Recordset("VH_ACCFRDATE")) Then failed = failed & Data1.Recordset("VH_ACCFRDATE") & ", "
                If Not IsNull(Data1.Recordset("VH_ACCTODATE")) Then failed = failed & Data1.Recordset("VH_ACCTODATE") & ", "
                failed = Left(failed, Len(failed) - 2) & vbCrLf
            End If
        End If
        c = c + 1
        Data1.Recordset.MoveNext
    Loop Until Data1.Recordset.EOF
End If

Data1.Refresh

Call Display_Value

Screen.MousePointer = DEFAULT

If Len(failed) = 0 Then
    MsgBox "All Rules applied", vbInformation + vbOKOnly, "Hours Based Entitlement Master"
Else
    MsgBox "The Following Rules failed:" & vbCrLf & failed, vbInformation + vbOKOnly, "Hours Based Entitlement Master"
End If

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdateAll", "Single", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
     RollBack
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub Form_Activate()

Call SET_UP_MODE
Call INI_Controls(Me)

glbOnTop = "frmSHrsHrlyEnt"

End Sub

Private Sub Form_Load()
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    Dim Answer, DefVal, Msg, Title  ' Declare variables.
    Dim RFound As Integer ' records found
    Dim x%
    Dim SQLQ
    
    glbOnTop = "FRMSHRSHRLYENT"
    
    FlagRefresh = False
    
    Data1.ConnectionString = glbAdoIHRDB
    SQLQ = "SELECT DISTINCT EH_DIV,EH_DEPT,EH_ORG,EH_FDATE,EH_TDATE,EH_EMP,EH_SECTION,EH_LOC,EH_PT,EH_HETYPE,EH_MANUAL,EH_EDATE,EH_UPDMETHOD FROM HR_HOURLYENT "
    If glbDIVCount = 1 And glbLinamar Then
        SQLQ = SQLQ & " WHERE EH_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
    End If
    If glbWFC Then 'Ticket #28553 Franks 05/03/2016
        SQLQ = SQLQ & " WHERE " & getWFCPlantSecurity("EH_SECTION")
    End If
    Data1.RecordSource = SQLQ
    Data1.Refresh
    
    'If UCase(glbCompEntSick$) = "M" Or UCase(glbCompEntSick$) = "N" Then
    '    vbxTrueGrid.Columns(5).Visible = False
    'End If
    
    Screen.MousePointer = HOURGLASS
    
    Select Case glbCompWDate$ ' sets field reference for basic 'which date'
        Case "O": fglbWDate$ = "ED_DOH"
        Case "S": fglbWDate$ = "ED_SENDTE"
        Case "U": fglbWDate$ = "ED_UNION"
        Case "L": fglbWDate$ = "ED_LTHIRE"
        Case "D": fglbWDate$ = "ED_USRDAT1"
    End Select
    
    vbxTrueGrid.Columns(0).Caption = lStr(vbxTrueGrid.Columns(0).Caption)
    vbxTrueGrid.Columns(1).Caption = lStr(vbxTrueGrid.Columns(1).Caption)
    vbxTrueGrid.Columns(2).Caption = lStr(vbxTrueGrid.Columns(2).Caption)
    
    Call setRptCaption(Me)
    
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
        clpCode(4).Visible = False
        clpProv.Left = clpCode(4).Left
        clpProv.Top = clpCode(4).Top
        clpProv.Visible = True
        lblLocation.Caption = "Province"
        vbxTrueGrid.Columns(8).Caption = "Province"
    End If
    
    If glbSyndesis Then
        lblCriteria(5).Caption = "Position Grade"
        vbxTrueGrid.Columns(8).Caption = "Position Grade"
        clpCode(2).Tag = "00-Enter Position Grade"
    End If
    
    If glbWFC Then
        lblSection.FontBold = True
    End If
    
    'Ticket #29617 - Mississaugas of Scugog Island First Nation
    If glbCompSerial = "S/N - 2485W" Then
        lblAccPeriod.Caption = "Pay Period"
    End If
    
    Screen.MousePointer = DEFAULT
    
    Call INI_Controls(Me)
        
    ST_UPD_MODE (False)
    
    Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
    MDIMain.panHelp(0).Caption = " "
    MDIMain.panHelp(1).Caption = " "
    MDIMain.panHelp(2).Caption = " "
    MDIMain.panHelp(3).Caption = " "
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Dim Keepfocus As Boolean
    'If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
    'Keepfocus = Not isUpdated(Me)
    'Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select FROM the menu the appropriate function."
    
    Set frmUEntitle = Nothing  'carmen apr 2000
End Sub

Private Function modUpdateSelection()
Dim EmpNo As Long
Dim SQLQ As String
Dim snapHREnt As New ADODB.Recordset
Dim rzAttend As New ADODB.Recordset
Dim rsHREmp As New ADODB.Recordset
Dim rsCurJobSal As New ADODB.Recordset
Dim dblServiceHours#, dblNewEntitle#, dblEntitleUpd#, dblEntitle#
Dim lngRecs&
Dim intWhereFit&, x%, Y%, z%
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec%
Dim xComments
Dim PayPerc
Dim varStartDate As Variant
Dim oldEntitleUpd
Dim dblServiceYears#
Dim dblNewMax#
Dim NumRec As Integer
Dim xKey

On Error GoTo modUpdateSelection_Err

modUpdateSelection = False

If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)

Screen.MousePointer = DEFAULT

If snapEntitle.BOF And snapEntitle.EOF Then
    MsgBox "Employees for this selection do not exist!"
    Exit Function
Else
    lngRecs& = snapEntitle.RecordCount
    
    Msg$ = lngRecs& & " Records to process" & Chr(10) & "Would You Like To Proceed?"
    Title$ = "Update Hours Based Entitlement"
    DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then    ' Evaluate response
        Exit Function
    End If
    
    Screen.MousePointer = HOURGLASS
End If

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 5


While Not snapEntitle.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / lngRecs&))
    MDIMain.panHelp(0).FloodPercent = pct%
    
    'Initialise
    dblNewEntitle = 0
    dblEntitleUpd = 0

    'If snapEntitle("ED_EMPNBR") = 3190 Then
    '    EmpNo& = snapEntitle("ED_EMPNBR")
    'End If
    
    If IsNull(snapEntitle(fglbWDate$)) Then
        GoTo lblNextRec
    End If


    EmpNo& = snapEntitle("ED_EMPNBR")

    'Get the existing entitlement
    SQLQ = "SELECT HE_EMPNBR,HE_TYPE,HE_ID ,"
    SQLQ = SQLQ & " HE_ENTITLE, HE_TDATE FROM HRENTHRS "
    SQLQ = SQLQ & " WHERE HE_EMPNBR = " & snapEntitle("ED_EMPNBR")
    SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
    SQLQ = SQLQ & " AND HE_FDATE = " & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND HE_TDATE = " & Date_SQL(dlpDateRange(1).Text)
    snapHREnt.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    If Not snapHREnt.EOF And Not snapHREnt.BOF Then
        snapHREnt.MoveLast
    End If

    NumRec = snapHREnt.RecordCount
    If snapHREnt.EOF Then
        oldEntitleUpd = 0
    Else
        oldEntitleUpd = snapHREnt("HE_ENTITLE")
    End If
    
    If Accum = True Then
        If NumRec > 0 Then
            dblEntitle# = snapHREnt("HE_ENTITLE")
        Else
            dblEntitle# = 0
        End If
    Else
        dblEntitle# = 0
    End If
    snapHREnt.Close
    Set snapHREnt = Nothing


    'Get Total Worked Hours from Attendance and Attendance History
    dblServiceHours# = 0
    dblServiceHours# = Total_Worked_Hours(snapEntitle("ED_EMPNBR"))
        
    'Compute the Worked Hours based on the Vacation / Pay Percentage based on the type of Entitlement Code
    'Retrieve the HR Attendance Matrix based on Entitlement Code entitled
    PayPerc = 0
    Select Case clpCode(2).Text
        Case "VC"
            If Not IsNull(snapEntitle("ED_VACPC")) Then
                PayPerc = snapEntitle("ED_VACPC")
            End If
        Case "SK"
            PayPerc = 4 / 100
        Case "PD"
            PayPerc = 4 / 100
        Case "PTO"
            PayPerc = 4 / 100
        Case "IW"
            PayPerc = 11.11 / 100
    End Select
    
    
    'Compute the new entitlement based on the Pay %
    dblNewEntitle# = dblServiceHours# * PayPerc

    'Check against Maximum before updating the Hourly Entitlement record
    'Get Start Date to compute the Service Months
    varStartDate = snapEntitle(fglbWDate$)
    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(dlpAccDateRange(1).Text))    'Ticket #17924
    
    'Check if valid Service Months
    If dblServiceYears# < 0 Then GoTo lblNextRec
    
    'Get the Maximum based on the Service Months and Entitlement Code
    dblNewMax# = Get_HourlyEntitle_Maximum(dblServiceYears#)
    
    If dblNewMax# > 0 Then
        If dblNewEntitle# > dblNewMax# Then
            dblNewEntitle# = dblNewMax#
        End If
    End If

    'Accumulate to the existing entitlement or replace the existing entitlement
    If Accum = True Then
        dblEntitleUpd = dblEntitle# + dblNewEntitle
    Else
        dblEntitleUpd = dblNewEntitle
    End If
    
    'Update respective Hourly Entitlement records
    If Accum = True Then
        'ACCUMULATE METHOD
        
        If NumRec > 0 Then  'if accumulate and found duplicate record
            SQLQ = "UPDATE HRENTHRS "
            SQLQ = SQLQ & " SET HE_ENTITLE = " & dblEntitleUpd & " "
            SQLQ = SQLQ & " WHERE HE_EMPNBR = " & snapEntitle("ED_EMPNBR")
            SQLQ = SQLQ & " AND HRENTHRS.HE_TYPE = '" & clpCode(2).Text & "' "
            SQLQ = SQLQ & " AND HRENTHRS.HE_FDATE = " & Date_SQL(dlpDateRange(0).Text)
            SQLQ = SQLQ & " AND HRENTHRS.HE_TDATE = " & Date_SQL(dlpDateRange(1).Text)
            gdbAdoIhr001.Execute (SQLQ)
            
            Call Append_Accrual(snapEntitle("ED_EMPNBR"), clpCode(2).Text, dlpAccDateRange(1).Text, dblEntitleUpd - oldEntitleUpd, "U", "Mass changed the existing Hours Based Entitlement") 'Ticket #17924
        Else
            'Ticket #17924 - If Flex logic (+) then update the existing Flex code hourly entitlement record instead
            'of adding a new record.
            If Right(clpCode(2).Text, 1) = "+" Then
                If fTablHREMP.State <> adStateClosed Then fTablHREMP.Close
                SQLQ = "SELECT * FROM HRENTHRS "
                SQLQ = SQLQ & " WHERE HE_EMPNBR = " & snapEntitle("ED_EMPNBR")
                SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
                SQLQ = SQLQ & " ORDER BY HE_FDATE DESC"
                fTablHREMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not fTablHREMP.EOF Then
                    fTablHREMP.MoveFirst
                Else
                    fTablHREMP.AddNew
                    fTablHREMP("HE_PREV") = 0
                End If
            Else
                fTablHREMP.AddNew     'if accumulate and no duplicate record
                fTablHREMP("HE_PREV") = 0
            End If
            
            fTablHREMP("HE_EMPNBR") = snapEntitle("ED_EMPNBR")
            fTablHREMP("HE_COMPNO") = "001"
            fTablHREMP("HE_TYPE_TABL") = "ADRE"
            fTablHREMP("HE_TYPE") = clpCode(2).Text
            fTablHREMP("HE_FDATE") = dlpDateRange(0).Text
            fTablHREMP("HE_TDATE") = dlpDateRange(1).Text
            fTablHREMP("HE_ENTITLE") = dblEntitleUpd
            fTablHREMP("HE_COE") = True
            fTablHREMP("HE_DHRS") = snapEntitle("ED_DHRS")
            fTablHREMP("HE_LDATE") = Now
            fTablHREMP("HE_LTIME") = Time$
            fTablHREMP("HE_LUSER") = glbUserID
            fTablHREMP.Update
                        
            Call Append_Accrual(snapEntitle("ED_EMPNBR"), clpCode(2).Text, dlpAccDateRange(1).Text, dblEntitleUpd, "A", "Mass added the Hours Based Entitlement")
        End If
    Else
        'REPLACE METHOD
        
        'Ticket #17924 - If Flex logic (+) then update the existing Flex code hourly entitlement record instead
        'of adding a new record.
        If Right(clpCode(2).Text, 1) = "+" Then
            If fTablHREMP.State <> adStateClosed Then fTablHREMP.Close
            SQLQ = "SELECT * FROM HRENTHRS "
            SQLQ = SQLQ & " WHERE HE_EMPNBR = " & snapEntitle("ED_EMPNBR")
            SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
            SQLQ = SQLQ & " ORDER BY HE_FDATE DESC"
            fTablHREMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not fTablHREMP.EOF Then
                fTablHREMP.MoveFirst
            Else
                fTablHREMP.AddNew
                fTablHREMP("HE_PREV") = 0
            End If
        Else
            'Ticket #18559 - Jerry does not want the Previous to be replaced with 0 after the rollover which
            'creates a new record on the Hourly Entitlement screen. In which case we cannot delete an existing
            'Hourly Entitlement record but instead update the values.
            'SQLQ$ = "DELETE FROM HRENTHRS "
            'SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
            'SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
            'SQLQ = SQLQ & " AND HE_TDATE = " & Date_SQL(dlpTo.Text)
            'gdbAdoIhr001.Execute SQLQ
            
            If fTablHREMP.State <> adStateClosed Then fTablHREMP.Close
            SQLQ = "SELECT * FROM HRENTHRS "
            SQLQ = SQLQ & " WHERE HE_EMPNBR = " & snapEntitle("ED_EMPNBR")
            SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
            SQLQ = SQLQ & " AND HE_FDATE = " & Date_SQL(dlpDateRange(0).Text)
            SQLQ = SQLQ & " AND HE_TDATE = " & Date_SQL(dlpDateRange(1).Text)
            fTablHREMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If fTablHREMP.EOF Then
                fTablHREMP.AddNew
                fTablHREMP("HE_PREV") = 0
            End If
        End If
        
        'fTablHREMP.AddNew
        
        fTablHREMP("HE_EMPNBR") = snapEntitle("ED_EMPNBR")
        fTablHREMP("HE_COMPNO") = "001"
        fTablHREMP("HE_TYPE_TABL") = "ADRE"
        fTablHREMP("HE_TYPE") = clpCode(2).Text
        fTablHREMP("HE_FDATE") = dlpDateRange(0).Text
        fTablHREMP("HE_TDATE") = dlpDateRange(1).Text
        fTablHREMP("HE_ENTITLE") = dblEntitleUpd
        fTablHREMP("HE_COE") = True
        fTablHREMP("HE_DHRS") = snapEntitle("ED_DHRS")
        fTablHREMP("HE_LDATE") = Now
        fTablHREMP("HE_LTIME") = Time$
        fTablHREMP("HE_LUSER") = glbUserID
        fTablHREMP.Update
        
        If NumRec > 0 Then  'if accumulate and found duplicate record
            'Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpFrom.Text, dblEntitleUpd - oldEntitleUpd, "U", "Mass modified the Hourly Entitlement")
            Call Append_Accrual(snapEntitle("ED_EMPNBR"), clpCode(2).Text, dlpAccDateRange(1).Text, dblEntitleUpd - oldEntitleUpd, "U", "Mass modified the Hours Based Entitlement")
        Else
            'Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpFrom.Text, dblEntitleUpd, "A", "Mass added the Hourly Entitlement")
            Call Append_Accrual(snapEntitle("ED_EMPNBR"), clpCode(2).Text, dlpAccDateRange(1).Text, dblEntitleUpd, "A", "Mass added the Hours Based Entitlement")
        End If
    End If

    'Ticket #22682 - Release 8.0: Jerry said not to check for duplicate, simply add new Attendance record, even
    'though it is a duplicate record.
'    'Ticket #17924 - Begin
'    'If the Entitlement Code is suffixed with + then insert an Attendance record
'    'for the Hourly Entitlement earned - helps in the Recalculate function
    If Right(clpCode(2).Text, 1) = "+" Then
        'Add Record in Attendance screen
        'Ticket #22682 - Release 8.0: Do not check for duplicates
        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE 1 = 2"
        'SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR =" & SnapAddEntitle("ED_EMPNBR")
        'SQLQ = SQLQ & " AND AD_REASON = '" & clpCode(2).Text & "'"
        'Ticket #18550 - Attendance record date cannot be prior to hire date
        'If CVDate(SnapAddEntitle("ED_DOH")) > CVDate(dlpFrom.Text) Then
        '    SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(SnapAddEntitle("ED_DOH"))
        'Else
        '    SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(dlpFrom.Text)
        'End If
        rzAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        'If rzAttend.EOF Then
            rzAttend.AddNew
        'End If
        rzAttend("AD_COMPNO") = "001"
        rzAttend("AD_EMPNBR") = snapEntitle("ED_EMPNBR")
        rzAttend("AD_DOA") = dlpAccDateRange(1).Text
        rzAttend("AD_REASON") = clpCode(2).Text
        rzAttend("AD_HRS") = dblEntitleUpd

        SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_ORG,ED_GLNO,ED_DOH FROM HREMP WHERE ED_EMPNBR = " & snapEntitle("ED_EMPNBR")
        rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsHREmp.EOF Then
            rzAttend("AD_PAYROLL_ID") = rsHREmp("ED_PAYROLL_ID")
            rzAttend("AD_GLNO") = rsHREmp("ED_GLNO")
            rzAttend("AD_ORG") = rsHREmp("ED_ORG")
            
            'Ticket #18550 - Attendance record date cannot be prior to hire date
            If CVDate(rsHREmp("ED_DOH")) > CVDate(dlpAccDateRange(1).Text) Then
                rzAttend("AD_DOA") = rsHREmp("ED_DOH")
            End If
        End If
        rsHREmp.Close

        SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & snapEntitle("ED_EMPNBR")
        rsCurJobSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsCurJobSal.BOF Then
            If rsCurJobSal("SH_SALARY") > 0 Then
                rzAttend("AD_SALARY") = rsCurJobSal("SH_SALARY")
                rzAttend("AD_SALCD") = rsCurJobSal("SH_SALCD")
            End If
        End If
        rsCurJobSal.Close
        Set rsCurJobSal = Nothing

        SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & snapEntitle("ED_EMPNBR")
        rsCurJobSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsCurJobSal.EOF Then
            rzAttend("AD_JOB") = rsCurJobSal("JH_JOB")
            rzAttend("AD_DHRS") = rsCurJobSal("JH_DHRS")
            rzAttend("AD_WHRS") = rsCurJobSal("JH_WHRS")
        End If
        rsCurJobSal.Close
        Set rsCurJobSal = Nothing

        'Ticket #18550
        'rzAttend("AD_COMM") = "Entitlement earned for the period: " & dlpFrom.Text & " to " & dlpTo.Text & "."
        rzAttend("AD_COMM") = "Entitlement earned for the period: " & dlpAccDateRange(0).Text & " to " & dlpAccDateRange(1).Text & "."
        rzAttend("AD_LDATE") = Date
        rzAttend("AD_LUSER") = glbUserID
        rzAttend("AD_LTIME") = Time$
        rzAttend.Update
        rzAttend.Close
        Set rzAttend = Nothing
    End If
'    'Ticket #17924 - End
    
    DoEvents
    xKey = snapEntitle("ED_EMPNBR")
    xKey = xKey & "|" & Format(dlpAccDateRange(0).Text, "dd-mmm-yyyy")
    xKey = xKey & "|" & Format(dlpAccDateRange(1).Text, "dd-mmm-yyyy")
    xKey = xKey & "|" & clpCode(2).Text
    xKey = xKey & "|" & dblEntitleUpd
    xKey = xKey & "|" & Format(Date, "dd-mmm-yyyy") 'Transaction Date
    Call Entitlements_Master_Integration(xKey, 0)
    
    DoEvents

lblNextRec:
    DoEvents
    
    snapEntitle.MoveNext
Wend

modUpdateSelection = True
MDIMain.panHelp(0).FloodType = 0

snapEntitle.Close
Set snapEntitle = Nothing

Screen.MousePointer = DEFAULT

Exit Function

modUpdateSelection_Err:
'These errors are:
'13=type mismatch
'94=invalid use of null
'3018=couln't find field 'item'
If Err = 13 Or Err = 94 Or Err = 3018 Then
   ' MsgBox "Err:" & Str(Err) & Chr(10) & Error$ & Chr(10) & " modUpdateSelection" & Chr(10) & "FORM:FUENTITL.FRM"
    'commented out by RAUBREY 5/20/97
    Err = 0
    Resume Next
End If

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateHrsEntitle", "HREMP", "edit/Add")

Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    'Rollback
    Resume Next
Else
    Unload Me
End If
End Function

Sub ST_UPD_MODE(TF As Boolean)
    Dim x, FT
    FT = Not TF
    
    clpDiv.Enabled = TF
    clpDept.Enabled = TF
    clpCode(0).Enabled = TF
    'If Not TF Or glbLinamar Then
    '    lblAsOf.FontBold = True
    'Else
    '    lblAsOf.FontBold = False
    'End If
    'If glbCompEntSick$ = "M" Or glbCompEntSick$ = "N" Or glbCompEntSick$ = "A" Then
    '    dlpAsOf.Enabled = True 'FT
    'Else
    '    dlpAsOf.Enabled = True 'Ticket #3419
    'End If
    'If sick Entitlement Outstanding based on "1" then ok, otherwise disenable
    'If glbEntOutStandingS$ = "1" Then
    '    CmdRecalc.Enabled = True
    'Else
    '    CmdRecalc.Enabled = False
    'End If
    If Not glbWHSCC Then
        clpCode(1).Enabled = TF
    Else
        clpCode(1).Enabled = False
    End If
    clpCode(2).Enabled = TF
    clpCode(3).Enabled = TF
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
        clpProv.Enabled = TF
    Else
        clpCode(4).Enabled = TF
    End If
    clpPT.Enabled = TF
    dlpDateRange(0).Enabled = TF
    dlpDateRange(1).Enabled = TF
    
    dlpAccDateRange(0).Enabled = TF
    dlpAccDateRange(1).Enabled = TF
    frmType.Enabled = TF
    
    'cmdClose.Enabled = FT
    'cmdModify.Enabled = FT
    'cmdDelete.Enabled = FT
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    '    cmdModify.Enabled = False
    '    cmdDelete.Enabled = False
    End If
    'cmdOK.Enabled = TF
    'cmdCancel.Enabled = TF
    'cmdNew.Enabled = FT
    'cmdPrint.Enabled = FT
    ''cmdPrintAll.Enabled = FT
    'cmdUpdate.Enabled = FT
    'vbxTrueGrid.Enabled = FT
    'Call modSetFGlobals("SICK")
End Sub

Sub Display_Value()
    Dim SQLQ, xOrder, nOrder, aa, SQLQW, glbiOneWhere
    Dim rsVE As New ADODB.Recordset
    Dim x
    clpDiv.Text = ""
    clpDept.Text = ""
    clpCode(0).Text = ""
    'If Not (glbCompEntSick$ = "M" Or glbCompEntSick$ = "N") Then
    '    dlpAsOf.Text = ""
    'End If
    clpCode(1).Text = ""
    clpCode(2).Text = ""
    clpCode(3).Text = ""
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
        clpProv.Text = ""
    Else
        clpCode(4).Text = ""
    End If
    clpPT.Text = ""
    dlpDateRange(0).Text = ""
    dlpDateRange(1).Text = ""
    dlpAccDateRange(0).Text = ""
    dlpAccDateRange(1).Text = ""
    
    If Not Data1.Recordset.EOF Then
        Call getWSQLQ("D")
        
        SQLQ = "SELECT * FROM HR_HOURLYENT WHERE " & fglbVSQLQ
        SQLQ = SQLQ & "Order By EH_DIV,EH_DEPT,EH_ORG, EH_FDATE,EH_EMP,EH_PT,EH_LOC,EH_SECTION,EH_ORDER "
        rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset
        
        If Not IsNull(Data1.Recordset("EH_DIV")) Then clpDiv.Text = Data1.Recordset("EH_DIV")
        If Not IsNull(Data1.Recordset("EH_DEPT")) Then clpDept.Text = Data1.Recordset("EH_DEPT")
        If Not IsNull(Data1.Recordset("EH_ORG")) Then clpCode(0).Text = Data1.Recordset("EH_ORG")
        If Not IsNull(Data1.Recordset("EH_FDATE")) Then dlpDateRange(0).Text = Data1.Recordset("EH_FDATE")
        If Not IsNull(Data1.Recordset("EH_TDATE")) Then dlpDateRange(1).Text = Data1.Recordset("EH_TDATE")
        'If Not IsNull(Data1.Recordset("EH_EDATE")) Then dlpAsOf.Text = Data1.Recordset("EH_EDATE")
        If Not IsNull(Data1.Recordset("EH_EMP")) Then clpCode(1).Text = Data1.Recordset("EH_EMP")
        If Not IsNull(Data1.Recordset("EH_PT")) Then clpPT.Text = Data1.Recordset("EH_PT")
        If Not IsNull(Data1.Recordset("EH_HETYPE")) Then clpCode(2).Text = Data1.Recordset("EH_HETYPE")
        If Not IsNull(Data1.Recordset("EH_SECTION")) Then clpCode(3).Text = Data1.Recordset("EH_SECTION")
        If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
            If Not IsNull(Data1.Recordset("EH_LOC")) Then clpProv.Text = Data1.Recordset("EH_LOC")
        Else
            If Not IsNull(Data1.Recordset("EH_LOC")) Then clpCode(4).Text = Data1.Recordset("EH_LOC")
        End If
        If Not IsNull(Data1.Recordset("EH_MANUAL")) Then
            chkManual.Value = Data1.Recordset("EH_MANUAL")
        End If
        If Not IsNull(Data1.Recordset("EH_UPDMETHOD")) Then
            txtUpdMethod.Text = Data1.Recordset("EH_UPDMETHOD")
        End If
        rsVE.Close
    End If
    
    Call SET_UP_MODE
    Call cmdModify_Click
End Sub

Private Sub Replace_Click(Value As Integer)
    If Accum.Value = True Then
        txtUpdMethod.Text = "A"
    ElseIf Replace.Value = True Then
        txtUpdMethod.Text = "R"
    End If
End Sub

Private Sub txtUpdMethod_Change()
    If txtUpdMethod = "A" Then
        Accum.Value = True
    ElseIf txtUpdMethod = "R" Then
        Replace.Value = True
    End If
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
    Dim SQLQ As String
    
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    SQLQ = "SELECT DISTINCT VH_DIV,VH_DEPT,VH_ORG,VH_LOC,VH_SECTION,VH_EMP,VH_PT,VH_GRPCD,VH_FRDATE,VH_TODATE,VH_ACCFRDATE,VH_ACCTODATE,VH_UPDMETHOD,VH_MANUAL FROM HRHRSVACENT"
    If glbDIVCount = 1 And glbLinamar Then
        SQLQ = SQLQ & " WHERE VH_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Display_Value
End Sub

Private Function getWSQLQ(xType)
Dim SQLQ As String
Dim xDiv, xDept, xORG, xFDate, xTDate, xEMP, xEmpMode, xHETYPE
Dim xLoc, xSection

fglbESQLQ = glbSeleDeptUn

If Len(clpDept.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DEPTNO = '" & clpDept.Text & "'"
If Len(clpDiv.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & clpDiv.Text & "' "
If Len(clpCode(0).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & clpCode(0).Text & "' "
If Len(clpCode(1).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMP = '" & clpCode(1).Text & "' "
If Len(clpCode(3).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & clpCode(3).Text & "' "
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    If Len(clpProv.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_PROV = '" & clpProv.Text & "' "
Else
    If Len(clpCode(4).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & clpCode(4).Text & "' "
End If
If Len(clpPT.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_PT = '" & clpPT.Text & "' "

fglbWSQLQ = "HE_TYPE = '" & clpCode(2).Text & "' "

'Ticket #18518 "+" and "-" need the Hourly entitlement setup before do the attendance import
'but it cannot update "VAC" and "SICK"
'Lanark Ticket #17711
'If glbCompSerial = "S/N - 2172W" Then
'Ticket #19782 Franks 02/03/2011 for Frontenac
If glbCompSerial = "S/N - 2172W" Or glbCompSerial = "S/N - 2410W" Then
    fglbWSQLQ = fglbWSQLQ & " AND NOT (HE_TYPE = 'VAC' OR HE_TYPE = 'SICK') "
End If

fglbWSQLQ = fglbWSQLQ & " AND HE_FDATE >= " & Date_SQL(dlpDateRange(0).Text)
fglbWSQLQ = fglbWSQLQ & " AND HE_TDATE <= " & Date_SQL(dlpDateRange(1).Text)

If xType = "" Then Exit Function

If xType = "O" Then
    xDiv = ODIV
    xDept = ODept
    xORG = oOrg
    xFDate = oFDate
    xTDate = OTDate
    xEMP = oEMP
    xEmpMode = oEmpMode
    xHETYPE = oHETYPE
    xLoc = OLoc
    xSection = OSection
ElseIf xType = "D" Then
    xDiv = Data1.Recordset("EH_DIV")
    xDept = Data1.Recordset("EH_DEPT")
    xORG = Data1.Recordset("EH_ORG")
    xFDate = Data1.Recordset("EH_FDATE")
    xTDate = Data1.Recordset("EH_TDATE")
    xEMP = Data1.Recordset("EH_EMP")
    xEmpMode = Data1.Recordset("EH_PT")
    xHETYPE = Data1.Recordset("EH_HETYPE")
    xLoc = Data1.Recordset("EH_LOC")
    xSection = Data1.Recordset("EH_SECTION")
Else
    xDiv = clpDiv.Text
    xDept = clpDept.Text
    xORG = clpCode(0).Text
    xFDate = dlpDateRange(0).Text
    xTDate = dlpDateRange(1).Text
    xEMP = clpCode(1).Text
    xEmpMode = clpPT.Text
    xHETYPE = clpCode(2).Text
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
        xLoc = clpProv.Text
    Else
        xLoc = clpCode(4).Text
    End If
    xSection = clpCode(3).Text
End If
    
If Len(xDiv) = 0 Or IsNull(xDiv) Then
    fglbVSQLQ = " (EH_DIV IS NULL OR EH_DIV='')"
Else
    fglbVSQLQ = "EH_DIV = '" & xDiv & "'"
End If
If Len(xDept) = 0 Or IsNull(xDept) Then
    fglbVSQLQ = fglbVSQLQ & " AND (EH_DEPT IS NULL OR EH_DEPT='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND EH_DEPT = '" & xDept & "'"
End If
If Len(xORG) = 0 Or IsNull(xORG) Then
    fglbVSQLQ = fglbVSQLQ & " AND (EH_ORG IS NULL OR EH_ORG='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND EH_ORG = '" & xORG & "'"
End If
If Len(oFDate) > 0 Or IsNull(xFDate) Then
    SQLQ = SQLQ & " AND  EH_FDATE = " & Date_SQL(oFDate)
End If
If Len(OTDate) > 0 Or IsNull(xTDate) Then
    SQLQ = SQLQ & " AND  EH_TDATE = " & Date_SQL(OTDate)
End If
If Len(xEMP) = 0 Or IsNull(xEMP) Then
    fglbVSQLQ = fglbVSQLQ & " AND (EH_EMP IS NULL OR EH_EMP='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND EH_EMP = '" & xEMP & "'"
End If
If Len(xLoc) = 0 Or IsNull(xLoc) Then
    fglbVSQLQ = fglbVSQLQ & " AND (EH_LOC IS NULL OR EH_LOC='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND EH_LOC = '" & xLoc & "'"
End If
If Len(xSection) = 0 Or IsNull(xSection) Then
    fglbVSQLQ = fglbVSQLQ & " AND (EH_SECTION IS NULL OR EH_SECTION='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND EH_SECTION = '" & xSection & "'"
End If
If Len(xEmpMode) = 0 Or IsNull(xEmpMode) Then
    fglbVSQLQ = fglbVSQLQ & " AND (EH_PT IS NULL OR EH_PT='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND EH_PT = '" & xEmpMode & "' "
End If
If Len(xHETYPE) = 0 Or IsNull(xHETYPE) Then
    fglbVSQLQ = fglbVSQLQ & " AND (EH_HETYPE IS NULL OR EH_HETYPE='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND EH_HETYPE = '" & xHETYPE & "'"
End If

End Function

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum

If fglbNew Then
    UpdateState = NewRecord
    TF = True
    cmdPrintAll.Enabled = False
    cmdUpdate.Enabled = False
    cmdUpdateAll.Enabled = False
    CmdRecalc.Enabled = False
ElseIf Me.Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
    cmdPrintAll.Enabled = True
    cmdUpdate.Enabled = False
    cmdUpdateAll.Enabled = False
    CmdRecalc.Enabled = False
Else
    UpdateState = OPENING
    TF = True
    cmdPrintAll.Enabled = True
    cmdUpdate.Enabled = True
    cmdUpdateAll.Enabled = True
    CmdRecalc.Enabled = True
End If

Call ST_UPD_MODE(TF)
Call set_Buttons(UpdateState)

If Not UpdateRight Then TF = False

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
RelateMode = nothingrelate
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Entitlements
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

Private Function DoWork() As Boolean
    Dim lastday
    Dim flglastdate As Boolean
    Dim lngRecs As Long, pct As Long, prec As Long

    Screen.MousePointer = DEFAULT
    
    DoWork = False
    
    If Not modUpdateSelection() Then Exit Function
        
    Screen.MousePointer = HOURGLASS
    Call EntReCalcHr

    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    
    DoWork = True
    
End Function

Private Function Total_Worked_Hours(xEmpNbr)
    Dim rsAttend As New ADODB.Recordset
    Dim rsHRAttMatrix As New ADODB.Recordset
    Dim SQLQ As String
    Dim SQLQMatrix As String
    Dim xTotHrs As Double
    
    'FYI: A table field value FT,PT,CA would be returned as 'FT','PT','CA' when using the following combination of REPLACE and QUOTENAME functions in TSQL
    'SELECT REPLACE(QUOTENAME(AM_PT,''''),',',''','''),* FROM HRATT_MATRIX
    
    Total_Worked_Hours = 0
    
    xTotHrs = 0
    
    'Retrieve the HR Attendance Matrix based on Entitlement Code entitled
    Select Case clpCode(2).Text
        Case "VC"
            SQLQMatrix = "SELEC * FROM HRATT_MATRIX WHERE AM_VAC_HRS <> 0"
        Case "SK"
            SQLQMatrix = "SELEC * FROM HRATT_MATRIX WHERE AM_ABSENT_HRS <> 0)"
        Case "PD"
            SQLQMatrix = "SELEC * FROM HRATT_MATRIX WHERE AM_EXTRA_HRS <> 0)"
        Case "PTO"
            SQLQMatrix = "SELEC * FROM HRATT_MATRIX WHERE AM_REG_HRS <> 0)"
        Case "IW"
            SQLQMatrix = "SELEC * FROM HRATT_MATRIX WHERE AM_INCID <> 0)"
    End Select
    
    'Attendance
    rsHRAttMatrix.Open SQLQMatrix, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRAttMatrix.EOF Then
        rsHRAttMatrix.MoveFirst
        
        'For each Attendance Matrix record's Reason Code, sum the hours from Attendance for the employee if the employee is part of the Attendance Matrix Category
        Do While Not rsHRAttMatrix.EOF
            'Get Total Hours for the Attendance Matrix Reason Code
            SQLQ = "SELECT SUM(AD_HRS) AS TOT_HRS FROM HR_ATTENDANCE"
            SQLQ = SQLQ & " WHERE AD_EMPNBR = " & xEmpNbr
            SQLQ = SQLQ & " AND (AD_DOA >= " & Date_SQL(dlpDateRange(0).Text)
            SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
            SQLQ = SQLQ & " AND AD_REASON = '" & rsHRAttMatrix("AM_REASON") & "'"
            'Also if the employee's Category matches the Attendance Matrix
            SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PT IN ('" & Replace(rsHRAttMatrix("AM_REASON"), ",", "','") & "'))"
            SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsAttend.EOF Then
                rsAttend.MoveFirst
                        
                'Sum Total Hours
                If rsAttend("TOT_HRS") > 0 Then
                    xTotHrs = xTotHrs + rsAttend("TOT_HRS")
                End If
            End If
            rsAttend.Close
            Set rsAttend = Nothing
            
            rsHRAttMatrix.MoveNext
        Loop
    End If
    rsHRAttMatrix.Close
    Set rsHRAttMatrix = Nothing
    
'    'Attendance History
'    rsHRAttMatrix.Open SQLQMatrix, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If Not rsHRAttMatrix.EOF Then
'        rsHRAttMatrix.MoveFirst
'
'        'For each Attendance Matrix record's Reason Code, sum the hours from Attendance for the employee if the employee is part of the Attendance Matrix Category
'        Do While Not rsHRAttMatrix.EOF
'            'Get Total Hours for the Attendance Matrix Reason Code
'            SQLQ = "SELECT SUM(AH_HRS) AS TOT_HRS FROM HR_ATTENDANCE_HISTORY"
'            SQLQ = SQLQ & " WHERE AH_EMPNBR = " & xEmpNbr
'            SQLQ = SQLQ & " AND (AH_DOA >= " & Date_SQL(dlpDateRange(0).Text)
'            SQLQ = SQLQ & " AND AH_DOA <= " & Date_SQL(dlpDateRange(1).Text) & ")"
'            SQLQ = SQLQ & " AND AH_REASON = '" & rsHRAttMatrix("AM_REASON") & "'"
'            'Also if the employee's Category matches the Attendance Matrix
'            SQLQ = SQLQ & " AND AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PT IN ('" & Replace(rsHRAttMatrix("AM_REASON"), ",", "','") & "'))"
'            SQLQ = SQLQ & " GROUP BY AH_EMPNBR"
'            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'            If Not rsAttend.EOF Then
'                rsAttend.MoveFirst
'
'                'Sum Total Hours
'                If rsAttend("TOT_HRS") > 0 Then
'                    xTotHrs = xTotHrs + rsAttend("TOT_HRS")
'                End If
'
'            End If
'            rsAttend.Close
'            Set rsAttend = Nothing
'            rsHRAttMatrix.MoveNext
'        Loop
'    End If
'    rsHRAttMatrix.Close
'    Set rsHRAttMatrix = Nothing

    Total_Worked_Hours = xTotHrs

End Function

Private Function Get_HourlyEntitle_Maximum(xServiceYrs)
    Dim rsHrlyEnt As New ADODB.Recordset
    Dim SQLQ As String
    
    Get_HourlyEntitle_Maximum = 0
    
    SQLQ = "SELECT * FROM HR_HOURLYENT WHERE "
    SQLQ = SQLQ & " EH_DIV = '" & clpDiv.Text & "'"
    SQLQ = SQLQ & " EH_DEPT= '" & clpDept.Text & "'"
    SQLQ = SQLQ & " EH_ORG= '" & clpCode(0).Text & "'"
    SQLQ = SQLQ & " EH_FDATE= " & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " EH_TDATE= " & Date_SQL(dlpDateRange(1).Text)
    SQLQ = SQLQ & " EH_EMP= '" & clpCode(1).Text & "'"
    SQLQ = SQLQ & " EH_PT= '" & clpPT.Text & "'"
    SQLQ = SQLQ & " EH_HETYPE= '" & clpCode(2).Text & "'"
    SQLQ = SQLQ & " EH_SECTION= '" & clpCode(3).Text & "'"
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
        SQLQ = SQLQ & " EH_LOC= '" & clpProv.Text & "'"
    Else
        SQLQ = SQLQ & " EH_LOC= '" & clpCode(4).Text & "'"
    End If
    SQLQ = SQLQ & " ORDER BY EH_DIV,EH_DEPT,EH_ORG, EH_FDATE,EH_EMP,EH_PT,EH_LOC,EH_SECTION,EH_ORDER "
    rsHrlyEnt.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    If Not rsHrlyEnt.EOF Then
        rsHrlyEnt.MoveFirst
        
        Do While Not rsHrlyEnt.EOF
            If rsHrlyEnt("EH_BMONTH") = "" And rsHrlyEnt("EH_EMONTH") = "" Then Exit Do
            
            If IsNumeric(rsHrlyEnt("EH_BMONTH")) And rsHrlyEnt("EH_EMONTH") = "" Then
                If xServiceYrs >= CDbl(rsHrlyEnt("EH_BMONTH")) Then
                    Get_HourlyEntitle_Maximum = rsHrlyEnt("EH_MAX")
                    Exit Do
                End If
            End If
            If IsNumeric(rsHrlyEnt("EH_BMONTH")) And IsNumeric(rsHrlyEnt("EH_EMONTH")) Then
                If xServiceYrs >= CDbl(rsHrlyEnt("EH_BMONTH")) And xServiceYrs <= CDbl(rsHrlyEnt("EH_EMONTH")) Then
                    Get_HourlyEntitle_Maximum = rsHrlyEnt("EH_MAX")
                    Exit Do
                End If
            End If
            
            rsHrlyEnt.MoveNext
        Loop
    End If
    rsHrlyEnt.Close
    Set rsHrlyEnt = Nothing
    
End Function
