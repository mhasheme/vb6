VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmUHrsEnt 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Rollover Hourly Entitlement"
   ClientHeight    =   9885
   ClientLeft      =   2565
   ClientTop       =   525
   ClientWidth     =   11790
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
   ScaleHeight     =   9885
   ScaleWidth      =   11790
   WindowState     =   2  'Maximized
   Begin VB.Frame VacFram03 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3075
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   11175
      Begin INFOHR_Controls.CodeLookup clpProv 
         Height          =   285
         Left            =   2640
         TabIndex        =   8
         Tag             =   "31-Province - Code"
         Top             =   1635
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Left            =   3150
         TabIndex        =   9
         Tag             =   "40-As of Date"
         Top             =   2070
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Left            =   1335
         TabIndex        =   10
         Tag             =   "40-As of Date"
         Top             =   2070
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   7200
         TabIndex        =   11
         Tag             =   "01-Entitlement Code"
         Top             =   1185
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ADRE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   1335
         TabIndex        =   12
         Tag             =   "00-Enter Location Code"
         Top             =   1185
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOR"
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   1335
         TabIndex        =   13
         Tag             =   "00-Specific Department Desired"
         Top             =   750
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   7
         LookupType      =   2
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   7200
         TabIndex        =   14
         Tag             =   "00-Section - Code"
         Top             =   1635
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   1335
         TabIndex        =   15
         Tag             =   "00-Enter Location Code"
         Top             =   1635
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin Threed.SSCheck chkManual 
         Height          =   255
         Left            =   5670
         TabIndex        =   16
         Top             =   2085
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Exclude from Update All  "
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
         Left            =   1350
         TabIndex        =   17
         Tag             =   "40-As of Date"
         Top             =   2520
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   7200
         TabIndex        =   18
         Tag             =   "00-Specific Employment Status Desired"
         Top             =   315
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDEM"
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         Height          =   285
         Left            =   1335
         TabIndex        =   19
         Tag             =   "00-Specific Division Desired"
         Top             =   315
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin INFOHR_Controls.CodeLookup clpPT 
         Height          =   285
         Left            =   7200
         TabIndex        =   20
         Tag             =   "EDPT-Category"
         Top             =   750
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDPT"
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
         Left            =   5670
         TabIndex        =   31
         Top             =   360
         Width           =   1350
      End
      Begin VB.Label lblCriteria 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Entitlement Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   5670
         TabIndex        =   30
         Top             =   1230
         Width           =   1455
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
         TabIndex        =   29
         Top             =   60
         Visible         =   0   'False
         Width           =   1575
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
         Left            =   5670
         TabIndex        =   28
         Top             =   1680
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
         TabIndex        =   27
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDtRange 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Range"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   26
         Top             =   2115
         Width           =   1035
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
         TabIndex        =   25
         Top             =   1230
         Width           =   420
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
         TabIndex        =   24
         Top             =   795
         Width           =   825
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
         TabIndex        =   23
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblAsOf 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   22
         Top             =   2565
         Width           =   1245
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
         Left            =   5670
         TabIndex        =   21
         Top             =   795
         Width           =   630
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   6
      Top             =   9120
      Width           =   11790
      _Version        =   65536
      _ExtentX        =   20796
      _ExtentY        =   1349
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
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Font3D          =   1
      Alignment       =   1
      Begin VB.CommandButton cmdRollover 
         Appearance      =   0  'Flat
         Caption         =   "&Rollover Entitlement"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Tag             =   "Rollover selected Hourly Entitlement"
         Top             =   120
         Width           =   2025
      End
      Begin VB.CommandButton cmdRolloverAll 
         Appearance      =   0  'Flat
         Caption         =   "Rollover All"
         Height          =   375
         Left            =   4875
         TabIndex        =   3
         Tag             =   "Rollover all Hourly Entitlements"
         Top             =   120
         Width           =   2025
      End
      Begin VB.CommandButton cmdZeroOut 
         Appearance      =   0  'Flat
         Caption         =   "&Zero Out Entitlement"
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Tag             =   "Zero Out selected Hourly Entitlement"
         Top             =   120
         Width           =   2025
      End
      Begin VB.CommandButton cmdZeroOutAll 
         Caption         =   "Zero Out All"
         Height          =   375
         Left            =   4875
         TabIndex        =   5
         Tag             =   "Zero Out all Hourly Entitlements"
         Top             =   120
         Width           =   2025
      End
      Begin MSAdodcLib.Adodc data1 
         Height          =   405
         Left            =   9840
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
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fuHrsEnt.frx":0000
      Height          =   2415
      Left            =   120
      OleObjectBlob   =   "fuHrsEnt.frx":0014
      TabIndex        =   32
      Top             =   0
      Width           =   11295
   End
   Begin Threed.SSFrame frmZero 
      Height          =   540
      Left            =   120
      TabIndex        =   33
      Top             =   6120
      Visible         =   0   'False
      Width           =   9120
      _Version        =   65536
      _ExtentX        =   16087
      _ExtentY        =   952
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Font3D          =   1
      ShadowStyle     =   1
      Begin Threed.SSCheck chkZeroCurrent 
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   195
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Current Year"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCheck chkZeroPrev 
         Height          =   255
         Left            =   3000
         TabIndex        =   1
         Top             =   195
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Previous Year"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSOption optDH 
      Height          =   195
      Index           =   0
      Left            =   4440
      TabIndex        =   34
      TabStop         =   0   'False
      Tag             =   "Hours to Rollover"
      Top             =   5805
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "Hours"
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
   Begin MSMask.MaskEdBox medMaxRollover 
      Height          =   285
      Left            =   3360
      TabIndex        =   35
      Tag             =   "10-Maximum Hours/Days to Rollover"
      Top             =   5760
      Width           =   900
      _ExtentX        =   1588
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
      Format          =   "##0"
      PromptChar      =   "_"
   End
   Begin Threed.SSOption optDH 
      Height          =   195
      Index           =   1
      Left            =   5280
      TabIndex        =   36
      Tag             =   "Days to Rollover"
      Top             =   5805
      Width           =   690
      _Version        =   65536
      _ExtentX        =   1217
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "Days"
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
   End
   Begin VB.Label lblMaxRollover 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum Hours or Days to Rollover"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   37
      Top             =   5805
      Width           =   3135
   End
End
Attribute VB_Name = "frmUHrsEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fTablHREMP As New ADODB.Recordset         ' table view of HREMP
Dim snapEntitle As New ADODB.Recordset     'user vier
Dim fglbWDate$, fglbWDateS$
Dim NumAddRec%
Dim fglbSick%
Dim fglbVac%
Dim NoErrorFlg  As Boolean
Dim fglbCompMonthly%

Dim ffieldEntitle$    ' ED_VAC or ED_SICK for name of field for entitlement
Dim ffieldPEntitle$     ' ED_PVAC or ED_PSICK for previous entitlement's field name
Dim fglbCode$           ' are we dealing with Vac/Sick records?"
Dim glbFrmCaption$, glbErrNum&

Dim ODIV, ODept, oOrg, oFDate, OTDate, oEMP, oEmpMode, oHETYPE
Dim OLoc, OSection

Dim FlagRefresh As Boolean

Dim SnapAddEntitle As New ADODB.Recordset
Dim fglbESQLQ, fglbWSQLQ, fglbVSQLQ
Dim fFormCalled As String

Private Function chkMUEntitle()
Dim X%, Y%

chkMUEntitle = False

On Error GoTo chkMUEntitle_Err
For X% = 0 To 4
If Len(clpCode(X%).Text) > 0 And clpCode(X%).Caption = "Unassigned" Then
    MsgBox "If Code entered it must be known"
    'clpCode(X%).SetFocus
    Exit Function
End If
Next X%

If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
    'clpDept.SetFocus
    Exit Function
End If
If Len(clpDiv.Text) < 1 Then
    If glbDIVCount = 1 And glbLinamar Then
        MsgBox lStr("Division is required field")
        'clpDiv.SetFocus
        Exit Function
    End If
Else
    If clpDiv.Caption = "Unassigned" Then
        MsgBox lStr("If Division Entered - it must be known")
        'clpDiv.SetFocus
        Exit Function
    End If
End If
If Not glbCBrant Then 'Ticket #12524
    If Not IsDate(dlpFrom.Text) Then
        MsgBox "Invalid From Date"
        'dlpFrom.SetFocus
        Exit Function
    End If
End If
If Not glbCBrant Then
    If Not IsDate(dlpTo.Text) Then
        MsgBox "Invalid To Date"
        'dlpTo.SetFocus
        Exit Function
    End If
End If

If Not IsDate(dlpAsOf.Text) Then
    MsgBox "Invalid Effective Date"
    'dlpAsOf.SetFocus
    Exit Function
End If

If Len(clpCode(2).Text) < 1 Then
    MsgBox "Entitlement Code is required field"
    'clpCode(2).SetFocus
    Exit Function
Else
    If clpCode(2).Caption = "Unassigned" Then
        MsgBox "If Code Entered - it must be known"
        'clpCode(2).SetFocus
        Exit Function
    End If
End If
If glbWFC Then
    If Len(clpCode(3).Text) = 0 Then
        MsgBox lStr("Section is required field")
        'clpCode(3).SetFocus
        Exit Function
    End If
End If


chkMUEntitle = True

Exit Function

chkMUEntitle_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEntitle", "HRBENFT", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

'Private Sub cmdAddEnt_Click()
''**************** Hourly Entitlement Master Add Procedure
'Dim SQLQ As String, Msg$, x%
'Dim Title$, DgDef As Variant, Response%
'On Error GoTo AddN_Err
'If Not gSec_Upd_Hrly_Entitlements Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If
''
'Title$ = "Mass Hourly Entitlement Records Addition"
'DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
'Msg$ = "Are you sure you want to add Records for this criteria?"
'Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
'If Response% = IDNO Then    ' Evaluate response
'    Exit Sub
'End If
''
'AddChgDel = "A"
'
'If Not chkMUEntitle() Then Exit Sub
'
'If Not modInsSelection() Then Exit Sub   'laura 03/04/98
'
'Call EntReCalcHr  'laura dec 15, 1997
'
'If Not glbSQL And Not glbOracle Then Call Pause(0.5)
'
'data1.Refresh
'
'Call Display_Value
'
'Screen.MousePointer = DEFAULT
'
'MsgBox "Records Added Successfully"
'
'Exit Sub
'
'AddN_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HOURLY ENTITLEMENTS", "Add")
'Resume Next
'
'End Sub

Public Sub cmdCancel_Click()

data1.Refresh

Call Display_Value
End Sub

Public Sub cmdClose_Click()
Unload Me

End Sub


'Public Sub cmdDelete_Click()
'Dim SQLQ, Msg, a%
'If data1.Recordset.BOF And data1.Recordset.EOF Then
'    MsgBox "Nothing to Delete"
'    Exit Sub
'End If
'Msg = "Are You Sure You Want To Delete "
'Msg = Msg & Chr(10) & "The Vacation Entitlement Rules?  "
'
'a% = MsgBox(Msg, 36, "Confirm Delete")
'If a% <> 6 Then Exit Sub
'
'Call getWSQLQ("C")
'
'SQLQ = "DELETE FROM HR_HOURLYENT WHERE " & fglbVSQLQ
'
'gdbAdoIhr001.BeginTrans
'gdbAdoIhr001.Execute SQLQ
'gdbAdoIhr001.CommitTrans
'
'data1.Refresh
'
'Call Display_Value
'
'End Sub

'Private Sub cmdDeleteEnt_Click()
'Dim a As Integer
'Dim SQLQ As String, rc%, DtTm As Variant, x%
'Dim DgDef, Title$, Msg$, Response%
'
'If Not gSec_Upd_Hrly_Entitlements Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If
'
'AddChgDel = "D"
'
'If Not chkMUEntitle() Then Exit Sub
'
'Title$ = "Mass Hourly Entitlement Records Delete"
'DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
'Msg$ = "Are You Sure You Want To Delete ALL records for this criteria?"
'Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
'If Response% = IDNO Then    ' Evaluate response
'    Exit Sub
'End If
'
'x% = modDelRecs()
'
'Call EntReCalcHr  'laura dec 15, 1997
'
'If Not glbSQL And Not glbOracle Then Call Pause(0.5)
'
'data1.Refresh
'
'Call Display_Value
'
'Screen.MousePointer = DEFAULT
'
'MsgBox "Records Deleted Successfully"
'
'Exit Sub
'
'Del_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "ATTEND", "Delete")
'Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
'    Resume Next
'Else
'    Unload Me
'End If
'End Sub
'
'Public Sub cmdModify_Click()
'
'ODIV = clpDiv.Text
'ODept = clpDept.Text
'oOrg = clpCode(0).Text
'oFDate = dlpFrom.Text
'OTDate = dlpTo.Text
'oEMP = clpCode(1).Text
'oEmpMode = clpPT.Text
'oHETYPE = clpCode(2).Text
'If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
'    OLoc = clpProv.Text
'Else
'    OLoc = clpCode(4).Text
'End If
'OSection = clpCode(3).Text
'Actn = "M"
'End Sub

'Public Sub cmdNew_Click()
'Dim x
'
'clpDiv.Text = ""
'clpDept.Text = ""
'clpCode(0).Text = ""
'dlpFrom.Text = ""
'dlpTo.Text = ""
'clpCode(1).Text = ""
'clpCode(2).Text = ""
'clpCode(3).Text = ""
'If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
'    clpProv.Text = ""
'Else
'    clpCode(4).Text = ""
'End If
'clpPT.Text = ""
'Actn = "A"
'fglbNew = True
'SET_UP_MODE
'clpDiv.SetFocus
'End Sub

'Public Sub cmdOK_Click()
'Dim x%, Y%, xUnion, xPT, SQLQ, SQLQW
'Dim xStr
'Dim rsVE As New ADODB.Recordset
'Dim rsVT As New ADODB.Recordset
'Dim glbiOneWhere As Boolean
'Dim bmk As Variant
'
'On Error GoTo AddN_Err
'
'If Data1.Recordset.EOF And Data1.Recordset.BOF Then
'    bmk = 0 'Ticket #11885 Frank Oct 11th, 2006
'Else
'    bmk = Data1.Recordset.Bookmark
'End If
'
'If Not chkMUEntitle() Then Exit Sub
'
'For x% = 0 To 19
'    If Not IsNumeric(medLTServ(x%)) Then Exit For
'    If Not IsNumeric(medGTServ(x%)) Then
'        medGTServ(x%) = 0
'    Else
'        If glbFrench Then
'            If medGTServ(x%) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
'        Else
'            If Val(medGTServ(x%)) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
'        End If
'    End If
'    If medLTServ(x%) > 0 And medGTServ(x%) = 0 Then medGTServ(x%) = 9999999
'Next
'
'If Actn = "M" Then
'    Call getWSQLQ("O")
'    SQLQ = "DELETE FROM HR_HOURLYENT WHERE " & fglbVSQLQ
'    gdbAdoIhr001.BeginTrans
'    gdbAdoIhr001.Execute SQLQ
'    gdbAdoIhr001.CommitTrans
'Else
'    Call getWSQLQ("C")
'    SQLQ = "SELECT * FROM HR_HOURLYENT WHERE " & fglbVSQLQ
'    rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    If Not rsVT.EOF Then
'        MsgBox "You can not add duplicate record"
'         clpDiv.SetFocus
'        Exit Sub
'    End If
'End If
'
'gdbAdoIhr001.BeginTrans
'SQLQ = "SELECT * FROM HR_HOURLYENT"
'rsVE.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
'
'For x% = 0 To 19
'    If Len(medLTServ(x%)) > 0 Then
'        rsVE.AddNew
'        rsVE("EH_ORDER") = x + 1
'        rsVE("EH_ORG_TABL") = "EDOR"
'        rsVE("EH_ORG") = clpCode(0).Text
'        rsVE("EH_PT") = clpPT.Text
'        rsVE("EH_DIV") = clpDiv.Text
'        rsVE("EH_DEPT") = clpDept.Text
'        rsVE("EH_EMP_TABL") = "EDEM"
'        rsVE("EH_EMP") = clpCode(1).Text
'        If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
'            rsVE("EH_LOC") = clpProv.Text
'        Else
'            rsVE("EH_LOC") = clpCode(4).Text
'        End If
'        rsVE("EH_SECTION") = clpCode(3).Text
'
'        If Len(dlpFrom.Text) > 0 Then rsVE("EH_FDATE") = dlpFrom.Text
'        If Len(dlpTo.Text) > 0 Then rsVE("EH_TDATE") = dlpTo.Text
'        If Len(dlpAsOf.Text) > 0 Then rsVE("EH_EDATE") = dlpAsOf.Text
'
'        rsVE("EH_HETYPE_TABL") = "ADRE"
'        rsVE("EH_HETYPE") = clpCode(2).Text
'        If glbFrench Then
'            rsVE("EH_BMONTH") = Replace(medLTServ(x%), ",", ".")
'        Else
'            rsVE("EH_BMONTH") = medLTServ(x%)
'        End If
'        If glbFrench Then
'            rsVE("EH_EMONTH") = Replace(medGTServ(x%), ",", ".")
'        Else
'            rsVE("EH_EMONTH") = medGTServ(x%)
'        End If
'        If glbFrench Then
'            rsVE("EH_ENTITLE") = Replace(medEntitle(x%), ",", ".")
'        Else
'            rsVE("EH_ENTITLE") = medEntitle(x%)
'        End If
'        If optD(x%) Then rsVE("EH_TYPE") = "D"
'        If optH(x%) Then rsVE("EH_TYPE") = "H"
'        If optF(x%) Then rsVE("EH_TYPE") = "F"
'        rsVE("EH_MANUAL") = chkManual.Value
'        rsVE.Update
'    End If
'Next
'rsVE.Close
'gdbAdoIhr001.CommitTrans
'Data1.Refresh
'
'If Not bmk = 0 Then
'    Data1.Recordset.Bookmark = bmk
'End If
'
'fglbNew = False
'
'Call Display_Value
'
'vbxTrueGrid.SetFocus
'
'Exit Sub
'
'AddN_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'If Err.Number = -2147217887 Then '01/01/1200 can cause this error Ticket #18227
'    MsgBox "    Invalid Date!    "
'    gdbAdoIhr001.RollbackTrans
'    Exit Sub
'Else
'    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdOK", "HOURLY ENTITLEMENTS", "UPDATE")
'    Unload Me
'End If
'
'End Sub
'
'Public Sub cmdPrint_Click()
'Dim RHeading As String, xReport, x%
'Dim SQLQ
'Dim dtYYY%, dtMM%, dtDD%
'
''mdPrint.Enabled = False
'cmdPrintAll.Enabled = False
'
'Me.vbxCrystal.Reset
'
'Me.vbxCrystal.WindowTitle = "Hourly Entitlement Master Report"
'
'Call setRptLabel(Me, 0) '1)
'
'Me.vbxCrystal.Connect = RptODBC_SQL
'Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rghrsent.rpt"
'
'SQLQ = ""
'SQLQ = SQLQ & "{HR_HOURLYENT.EH_DIV} = '" & clpDiv.Text & "'"
'SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_DEPT} = '" & clpDept.Text & "'"
'SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_ORG} = '" & clpCode(0).Text & "'"
'If Len(dlpFrom.Text) > 0 Then
'    dtYYY% = Year(dlpFrom.Text)
'    dtMM% = month(dlpFrom.Text)
'    dtDD% = Day(dlpFrom.Text)
'    SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_FDATE} in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
'End If
'If Len(dlpTo.Text) > 0 Then
'    dtYYY% = Year(dlpTo.Text)
'    dtMM% = month(dlpTo.Text)
'    dtDD% = Day(dlpTo.Text)
'    SQLQ = SQLQ & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
'End If
'SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_EMP} = '" & clpCode(1).Text & "'"
'SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_PT} = '" & clpPT.Text & "' "
'SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_HETYPE} = '" & clpCode(2).Text & "'"
'SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_SECTION} = '" & clpCode(3).Text & "'"
'If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
'    SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_LOC} = '" & clpProv.Text & "'"
'Else
'    SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_LOC} = '" & clpCode(4).Text & "'"
'End If
'Me.vbxCrystal.SelectionFormula = SQLQ
'
'Me.vbxCrystal.Destination = 1
'Me.vbxCrystal.Action = 1
'
''cmdPrint.Enabled = True
'cmdPrintAll.Enabled = True
'Call SET_UP_MODE
'End Sub
'
'Public Sub cmdView_Click()
'Dim RHeading As String, xReport, x%
'Dim SQLQ
'Dim dtYYY%, dtMM%, dtDD%
'
''cmdPrint.Enabled = False
'cmdPrintAll.Enabled = False
'
'Me.vbxCrystal.Reset
'
'Me.vbxCrystal.WindowTitle = "Hourly Entitlement Master Report"
'
'Call setRptLabel(Me, 0) '1)
'
'Me.vbxCrystal.Connect = RptODBC_SQL
'Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rghrsent.rpt"
'
'SQLQ = ""
'SQLQ = SQLQ & "{HR_HOURLYENT.EH_DIV} = '" & clpDiv.Text & "'"
'SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_DEPT} = '" & clpDept.Text & "'"
'SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_ORG} = '" & clpCode(0).Text & "'"
'If Len(dlpFrom.Text) > 0 Then
'    dtYYY% = Year(dlpFrom.Text)
'    dtMM% = month(dlpFrom.Text)
'    dtDD% = Day(dlpFrom.Text)
'    SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_FDATE} in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
'End If
'If Len(dlpTo.Text) > 0 Then
'    dtYYY% = Year(dlpTo.Text)
'    dtMM% = month(dlpTo.Text)
'    dtDD% = Day(dlpTo.Text)
'    SQLQ = SQLQ & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
'End If
'SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_EMP} = '" & clpCode(1).Text & "'"
'SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_PT} = '" & clpPT.Text & "' "
'SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_HETYPE} = '" & clpCode(2).Text & "'"
'SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_SECTION} = '" & clpCode(3).Text & "'"
'If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
'    SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_LOC} = '" & clpProv.Text & "'"
'Else
'    SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_LOC} = '" & clpCode(4).Text & "'"
'End If
'Me.vbxCrystal.SelectionFormula = SQLQ
'
'Me.vbxCrystal.Destination = 0
'Me.vbxCrystal.Action = 1
'
'Call SET_UP_MODE
''cmdPrint.Enabled = True
'cmdPrintAll.Enabled = True
'End Sub
'
'Private Sub cmdPrintAll_Click()
'Dim RHeading As String, xReport, x%
'Dim SQLQ
'Dim dtYYY%, dtMM%, dtDD%
'
'cmdPrintAll.Enabled = False
''cmdPrint.Enabled = False
'Me.vbxCrystal.Reset
'Me.vbxCrystal.WindowTitle = "Hourly Entitlement Master Report"
'
'Call setRptLabel(Me, 0) '1)
'
'Me.vbxCrystal.Connect = RptODBC_SQL
'Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rghrsent.rpt"
'Me.vbxCrystal.Action = 1
'
''cmdPrint.Enabled = True
'cmdPrintAll.Enabled = True
'End Sub
'
'Public Sub cmdUpdate_Click()
'
'
''************ UPDATE PROCEDRE*************
''ZAHOOR BUTT 01/11/2006
'Dim SQLQ As String, Msg$, x%
'Dim Title$, DgDef As Variant, Response%
'Dim sFlag As Boolean
'On Error GoTo AddN_Err
'If Not gSec_Upd_Hrly_Entitlements Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If
'
'Title$ = "Mass Hourly Entitlement Records Update"
'DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2
'Msg$ = "Are you sure you want to Update Records for this criteria?"
'Response% = MsgBox(Msg$, DgDef, Title)
'If Response% = IDNO Then
'    Exit Sub
'End If
'
'If Not chkMUEntitle() Then Exit Sub
'
'If Not glbSQL And Not glbOracle Then Call Pause(0.5)
'
' sFlag = DoWork
'
'Data1.Refresh
'
'Call Display_Value
'
'Screen.MousePointer = DEFAULT
'MsgBox "Records Updated Successfully"
'
'Exit Sub
'
'AddN_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HOURLY ENTITLEMENTS", "Add")
'Resume Next
''ZAHOOR BUTT 01/11/2006
'
'End Sub
'
'Private Function DoWork() As Boolean
'Dim sFlag As Boolean
'
'Screen.MousePointer = DEFAULT
'DoWork = False
'
'If Not gSec_Upd_Hrly_Entitlements Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Function
'End If
'
'AddChgDel = "A"
'
'If glbCBrant Then
'    If Not modInsSelectionCBrant() Then Exit Function
'Else
'    If Not modInsSelection() Then Exit Function
'    'If Not modUpdateSelection() Then Exit Function
'End If
'
'Call EntReCalcHr
'
'If Not glbSQL And Not glbOracle Then Call Pause(0.5)
'
'DoWork = True
'
'End Function

Private Function CR_SnapAddEntitle()

Dim BD As Integer
Dim SQLQ As String, countr As Integer
Dim Dat1 As Variant, Dat2 As Variant
Dim iOneWhere As Integer, NxtSQL As String, strReas$, strTm$, X%
Dim Dt As Variant

CR_SnapAddEntitle = False

On Error GoTo CR_SnapAddEntitle_Err

strTm$ = Time$

Dt = Date$

Call getWSQLQ("")

Screen.MousePointer = HOURGLASS

SQLQ = "SELECT JH_DHRS,JH_FTENUM,ED_EMPNBR,ED_DHRS,ED_DOH,ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1 "
If glbOracle Then
    SQLQ = SQLQ & "FROM HREMP, HR_JOB_HISTORY WHERE HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
    SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_CURRENT<>0"
Else
    SQLQ = SQLQ & "FROM HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
    SQLQ = SQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT<>0"
End If

SQLQ = SQLQ & " AND " & fglbESQLQ

If SnapAddEntitle.State <> 0 Then SnapAddEntitle.Close
SnapAddEntitle.Open SQLQ, gdbAdoIhr001, adOpenStatic

NumAddRec% = SnapAddEntitle.RecordCount
Screen.MousePointer = DEFAULT
CR_SnapAddEntitle = True

Exit Function

CR_SnapAddEntitle_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapAddEntitle", "Entitlements/EMP", "Select")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function CR_SnapEntitle()

Dim SQLQ As String

CR_SnapEntitle = False
On Error GoTo CR_SnapEntitle_Err

Screen.MousePointer = HOURGLASS

Call getWSQLQ("")

SQLQ = "SELECT * FROM qry_MU_Hourly "
SQLQ = SQLQ & " Where " & fglbESQLQ & " AND " & fglbWSQLQ

If snapEntitle.State <> 0 Then snapEntitle.Close
snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenStatic
Screen.MousePointer = DEFAULT
CR_SnapEntitle = True

Exit Function

CR_SnapEntitle_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapEntitle", "Hourly Rollover/ZeroOut", "Select")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub cmdRollover_Click()
    
    If Len(medMaxRollover) > 0 Then
        If Not IsNumeric(medMaxRollover) Then
            Screen.MousePointer = DEFAULT
            MsgBox "Invalid Maximum Hours/Days to Rollover."
            Exit Sub
        End If
    End If
    
    'Ticket #26612 - Recalculate the TAKEN before rolling over
    Call EntReCalcHr
    
    'Ticket #29623 - WDGPHU - Cannot roll over the Flex Time - no Year End for Flex Time as it is an ongoing balance
    If glbCompSerial = "S/N - 2411W" And UCase(Left(clpCode(2), 2)) = "FX" Then
        MsgBox "Flex Time should not be rolled over.", vbExclamation, "info:HR - Flex Time"
        Exit Sub
    End If
    
    'Rollover an Hourly Entitlement
    If Rollover_Hourly_Entitlement And NoErrorFlg = False Then
        data1.Refresh
        Call Display_Value
        Screen.MousePointer = DEFAULT
        MsgBox "Rollover completed successfully.", vbInformation + vbOKOnly, "Hourly Entitlements Rollover"
    End If

End Sub

Private Sub cmdRolloverAll_Click()
    Dim failed As String
    Dim c As Long

    If Len(medMaxRollover) > 0 Then
        If Not IsNumeric(medMaxRollover) Then
            Screen.MousePointer = DEFAULT
            MsgBox "Invalid Maximum Hours/Days to Rollover."
            Exit Sub
        End If
    End If

    failed = ""
    c = 1
    If data1.Recordset.EOF = False And data1.Recordset.BOF = False Then
        'Ticket #26612 - Recalculate the TAKEN before rolling over
        Call EntReCalcHr
    
        data1.Recordset.MoveFirst
        Do
            Call Display_Value
            
            If chkManual.Value = False Then
                If chkMUEntitle() Then
                   If Rollover_Hourly_Entitlement = False And NoErrorFlg = False Then
                        failed = failed & "Rule " & CStr(c) & ": "
                        If Not IsNull(data1.Recordset("EH_DIV")) Then failed = failed & data1.Recordset("EH_DIV") & ", "
                        If Not IsNull(data1.Recordset("EH_DEPT")) Then failed = failed & data1.Recordset("EH_DEPT") & ", "
                        If Not IsNull(data1.Recordset("EH_ORG")) Then failed = failed & data1.Recordset("EH_ORG") & ", "
                        If Not IsNull(data1.Recordset("EH_EMP")) Then failed = failed & data1.Recordset("EH_EMP") & ", "
                        If Not IsNull(data1.Recordset("EH_PT")) Then failed = failed & data1.Recordset("EH_PT") & ", "
                        If Not IsNull(data1.Recordset("EH_HETYPE")) Then failed = failed & data1.Recordset("EH_HETYPE") & ", "
                        If Not IsNull(data1.Recordset("EH_FDATE")) Then failed = failed & data1.Recordset("EH_FDATE") & ", "
                        If Not IsNull(data1.Recordset("EH_TDATE")) Then failed = failed & data1.Recordset("EH_TDATE") & ", "
                        If Not IsNull(data1.Recordset("EH_EDATE")) Then failed = failed & data1.Recordset("EH_EDATE") & ", "
                        If Not IsNull(data1.Recordset("EH_LOC")) Then failed = failed & data1.Recordset("EH_LOC") & ", "
                        If Not IsNull(data1.Recordset("EH_SECTION")) Then failed = failed & data1.Recordset("EH_SECTION") & ", "
                        failed = Left(failed, Len(failed) - 2) & vbCrLf
                   End If
                Else
                    failed = failed & "Rule " & CStr(c) & ": "
                    If Not IsNull(data1.Recordset("EH_DIV")) Then failed = failed & data1.Recordset("EH_DIV") & ", "
                    If Not IsNull(data1.Recordset("EH_DEPT")) Then failed = failed & data1.Recordset("EH_DEPT") & ", "
                    If Not IsNull(data1.Recordset("EH_ORG")) Then failed = failed & data1.Recordset("EH_ORG") & ", "
                    If Not IsNull(data1.Recordset("EH_EMP")) Then failed = failed & data1.Recordset("EH_EMP") & ", "
                    If Not IsNull(data1.Recordset("EH_PT")) Then failed = failed & data1.Recordset("EH_PT") & ", "
                    If Not IsNull(data1.Recordset("EH_HETYPE")) Then failed = failed & data1.Recordset("EH_HETYPE") & ", "
                    If Not IsNull(data1.Recordset("EH_FDATE")) Then failed = failed & data1.Recordset("EH_FDATE") & ", "
                    If Not IsNull(data1.Recordset("EH_TDATE")) Then failed = failed & data1.Recordset("EH_TDATE") & ", "
                    If Not IsNull(data1.Recordset("EH_EDATE")) Then failed = failed & data1.Recordset("EH_EDATE") & ", "
                    If Not IsNull(data1.Recordset("EH_LOC")) Then failed = failed & data1.Recordset("EH_LOC") & ", "
                    If Not IsNull(data1.Recordset("EH_SECTION")) Then failed = failed & data1.Recordset("EH_SECTION") & ", "
                    failed = Left(failed, Len(failed) - 2) & vbCrLf
                End If
            End If
            c = c + 1
            data1.Recordset.MoveNext
        Loop Until data1.Recordset.EOF
    End If
    
    data1.Refresh
    
    Call Display_Value
    
    Screen.MousePointer = DEFAULT
    
    If Len(failed) = 0 Then
        MsgBox "All Rules applied. Rollover completed successfully.", vbInformation + vbOKOnly, "Hourly Entitlements Rollover"
    Else
        MsgBox "The Rollover for following Rules failed:" & vbCrLf & failed, vbInformation + vbOKOnly, "Hourly Entitlements Rollover"
    End If
    
    Exit Sub
    
Mod_Err:
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdRolloverAll", "Hourly Rollover/ZeroOut", "Modify")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
         RollBack
        Resume Next
    Else
        Unload Me
    End If

End Sub

Private Function Rollover_Hourly_Entitlement() As Boolean
    Dim rsHEnt As New ADODB.Recordset
    Dim rsHEntNew As New ADODB.Recordset
    Dim rsHrEntMst As New ADODB.Recordset
    Dim EmpNo As Long, strJob$, spt As Variant, lngRecs&
    Dim Msg$, Title$, DgDef As Variant
    Dim Response%, pct%, prec%, xErr
    Dim SQLQ As String, dblOUTS#, dblOUTV#
    Dim xComments As String
    Dim xSkipped As String
    Dim xHrsDay
    Dim xSickOut#, xMaxHrlEnt
    
    On Error GoTo Rollover_Hourly_Entitlement_Err
    
    Rollover_Hourly_Entitlement = False
        
    'Ticket #29623 - WDGPHU - Cannot roll over the Flex Time - no Year End for Flex Time as it is an ongoing balance
    If glbCompSerial = "S/N - 2411W" And UCase(Left(clpCode(2), 2)) = "FX" Then
        MsgBox "Flex Time should not be rolled over.", vbExclamation, "info:HR - Flex Time"
        Rollover_Hourly_Entitlement = False
        NoErrorFlg = False
        Exit Function
    End If
        
    If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
    
    Screen.MousePointer = DEFAULT
    
    NoErrorFlg = False
    If snapEntitle.BOF And snapEntitle.EOF Then
        MsgBox "Employees for this selection do not exist!"
        Rollover_Hourly_Entitlement = True
        NoErrorFlg = True
        Exit Function
    Else
        lngRecs& = snapEntitle.RecordCount
        Msg$ = lngRecs& & " Records to process" & Chr(10) & "Proceed?"
        Title$ = "Rollover Hourly Entitlements"
        DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
        Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
        If Response% = IDNO Then    ' Evaluate response
            Rollover_Hourly_Entitlement = True
            NoErrorFlg = True
            Exit Function
        End If
    End If
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 5
    Screen.MousePointer = HOURGLASS
    
    xSkipped = ""
    
    While Not snapEntitle.EOF
        prec% = prec% + 1
        pct% = Int(100 * (prec% / (lngRecs&)))
        MDIMain.panHelp(0).FloodPercent = pct%
        
        EmpNo& = snapEntitle("HE_EMPNBR")
        
        dblOUTS# = 0
        
        If IsNumeric(snapEntitle("HE_PREV")) Then
            dblOUTS# = dblOUTS# + snapEntitle("HE_PREV")
        End If
        If IsNumeric(snapEntitle("HE_ENTITLE")) Then
            dblOUTS# = dblOUTS# + snapEntitle("HE_ENTITLE")
        End If
        If IsNumeric(snapEntitle("HE_TAKEN")) Then
            dblOUTS# = dblOUTS# - snapEntitle("HE_TAKEN")
        End If
            
        If glbCompSerial = "S/N - 2430W" Then  'Ticket #27729 Franks 03/15/2016 Carizon Rollover
            If snapEntitle("HE_TYPE") = "FSB" Then
                xSickOut# = getEmpSickOuting(EmpNo&)
                dblOUTS# = dblOUTS# + xSickOut#
            End If
        End If
        
        'Maximum Hours/Days to Rollover
        If IsNumeric(medMaxRollover) Then
            If optDH(0) Then    'Hours
                If Val(medMaxRollover) < Val(dblOUTS#) Then
                    dblOUTS# = medMaxRollover
                End If
            Else
                'Convert Days into Hours before comparison
                xHrsDay = GetJHData(EmpNo&, "JH_DHRS", 0)
                If xHrsDay = 0 Then
                    xSkipped = xSkipped & ", " & EmpNo&
                    GoTo Skip_Employee
                Else
                    If (Val(medMaxRollover) * Val(xHrsDay)) < Val(dblOUTS#) Then
                        dblOUTS# = Val(medMaxRollover) * Val(xHrsDay)
                    End If
                End If
            End If
        Else 'If it is Blank of "Maximum Hours or Days to Rollover"
            If glbCompSerial = "S/N - 2430W" Then  'Ticket #27729 Franks 03/15/2016 Carizon Rollover
                If snapEntitle("HE_TYPE") = "FSB" Then
                    xMaxHrlEnt = getHourlyEntMax(EmpNo&, "FSB")
                    If Val(xMaxHrlEnt) < Val(dblOUTS#) Then
                        dblOUTS# = xMaxHrlEnt
                    End If
                End If
            End If
        End If
            
        'Ticket #23141 - For Vadim clients Rolling over differently.
        'I will have to clear the balance in Vadim first, i.e. pass -ve OS Bal, so it becomes 0 balance in Vadim
        'and then pass OS to add back the OS. This will show the clear in and out in Accrual file and in Vadim.
        If glbVadim Then
            'Clear the Previous from Vadim first
            xComments = "Vadim OS: Prev. Hourly Ent. Chg from " & dblOUTS# & " to 0" '& dblOUTS#
            Call Append_Accrual(EmpNo&, snapEntitle("HE_TYPE"), snapEntitle("HE_TDATE"), 0 - dblOUTS#, "R", xComments)
        End If
                    
        If glbVadim Then
            'Ticket #23141 - For Vadim it is actually changing from 0 to OS amount
            xComments = "Prev. Hourly Ent. Chg from 0" & " to " & dblOUTS#
        Else
            xComments = "Prev. Hourly Ent. Chg from " & snapEntitle("HE_PREV") & " to " & dblOUTS#
        End If
        
        If glbVadim Then
            'Ticket #23141 - Add full OS back after clearing above
            Call Append_Accrual(EmpNo&, snapEntitle("HE_TYPE"), snapEntitle("HE_TDATE"), dblOUTS#, "R", xComments)
        Else
            Call Append_Accrual(EmpNo&, snapEntitle("HE_TYPE"), snapEntitle("HE_TDATE"), Val(dblOUTS# & "") - Val(snapEntitle("HE_PREV")), "R", xComments)
        End If
        
        'Update Hourly Entitlement - Previous
        rsHEnt.Open "SELECT * FROM HRENTHRS WHERE HE_ID= " & snapEntitle("HE_ID"), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsHEnt.EOF Then
            'Ticket #18559 - Instead add a new record of the Hourly Entitlement with new year date range
            'and Previous value
            'Add year to the above retrieved existing hourly entitlement date range
            
            'rsHEnt("HE_PREV") = dblOUTS#
            'rsHEnt("HE_LDATE") = Now
            'rsHEnt("HE_LTIME") = Time$
            'rsHEnt("HE_LUSER") = glbLEE_ID
            'rsHEnt.Update
            
            SQLQ = "SELECT * FROM HRENTHRS "
            SQLQ = SQLQ & " WHERE HE_EMPNBR = " & snapEntitle("HE_EMPNBR")
            SQLQ = SQLQ & " AND HE_TYPE = '" & rsHEnt("HE_TYPE") & "'"
            SQLQ = SQLQ & " AND HE_FDATE = " & Date_SQL(DateAdd("yyyy", 1, rsHEnt("HE_FDATE")))
            SQLQ = SQLQ & " AND HE_TDATE = " & Date_SQL(DateAdd("yyyy", 1, rsHEnt("HE_TDATE")))
            rsHEntNew.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsHEntNew.EOF Then
                rsHEntNew.AddNew
            End If
            rsHEntNew("HE_EMPNBR") = snapEntitle("HE_EMPNBR")
            rsHEntNew("HE_COMPNO") = "001"
            rsHEntNew("HE_TYPE_TABL") = "ADRE"
            rsHEntNew("HE_TYPE") = rsHEnt("HE_TYPE")
            rsHEntNew("HE_FDATE") = DateAdd("yyyy", 1, rsHEnt("HE_FDATE"))
            rsHEntNew("HE_TDATE") = DateAdd("yyyy", 1, rsHEnt("HE_TDATE"))
            rsHEntNew("HE_PREV") = dblOUTS#
            rsHEntNew("HE_ENTITLE") = 0     'For Vadim this is good since we are not doing Zero Out for Vadim except when no rollover and zeroing out Previous
            rsHEntNew("HE_COE") = rsHEnt("HE_COE")
            rsHEntNew("HE_DHRS") = rsHEnt("HE_DHRS")
            rsHEntNew("HE_LDATE") = Now
            rsHEntNew("HE_LTIME") = Time$
            rsHEntNew("HE_LUSER") = glbUserID
            rsHEntNew.Update
            
            'Update the existing Hourly Entitlement rule with new date range
            SQLQ = "SELECT * FROM HR_HOURLYENT WHERE " & fglbVSQLQ  'EH_ID = " & data1.Recordset("EH_ID")
            rsHrEntMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            Do While Not rsHrEntMst.EOF
                rsHrEntMst("EH_FDATE") = DateAdd("yyyy", 1, data1.Recordset("EH_FDATE"))
                rsHrEntMst("EH_TDATE") = DateAdd("yyyy", 1, data1.Recordset("EH_TDATE"))
                rsHrEntMst("EH_EDATE") = DateAdd("yyyy", 1, data1.Recordset("EH_EDATE"))
                rsHrEntMst.Update
                
                rsHrEntMst.MoveNext
            Loop
            
            rsHEntNew.Close
            rsHrEntMst.Close
            Set rsHEntNew = Nothing
            Set rsHrEntMst = Nothing
            
        End If
        rsHEnt.Close
        Set rsHEnt = Nothing
        
        'Release 8.0 - Ticket #22682: Function to delete all Previous Year's Exceeding Follow Up records.
        If Not IsNull(snapEntitle("HE_TDATE")) Then
            Call Delete_Exceeding_FollowUp(EmpNo&, snapEntitle("HE_TYPE"), Year(snapEntitle("HE_TDATE")))
        End If
        
Skip_Employee:
        snapEntitle.MoveNext
    Wend
    
    MDIMain.panHelp(0).FloodType = 0
    
    snapEntitle.Close
    
    Rollover_Hourly_Entitlement = True
    
    Screen.MousePointer = DEFAULT

    If Len(xSkipped) > 0 Then
        MsgBox "Employee(s) skipped due to missing Hours/Day on Position screen for Maximum Rollover: " & xSkipped, vbExclamation, "Skipped Rollover"
    End If

Exit Function

Rollover_Hourly_Entitlement_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "RollOver Entitle", "HRENTHRS", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub cmdZeroOut_Click()
    'Zero Out Hourly Entitlement
    If ZeroOut_Hourly_Entitlement And NoErrorFlg = False Then
        MsgBox "Zero Out completed successfully.", vbInformation + vbOKOnly, "Hourly Entitlements Zero Out"
    End If
End Sub

Private Sub cmdZeroOutAll_Click()
    Dim failed As String
    Dim c As Long

    failed = ""
    c = 1
    If data1.Recordset.EOF = False And data1.Recordset.BOF = False Then
        data1.Recordset.MoveFirst
        Do
            Call Display_Value
            
            If chkManual.Value = False Then
                If chkMUEntitle() Then
                   If ZeroOut_Hourly_Entitlement = False And NoErrorFlg = False Then
                        failed = failed & "Rule " & CStr(c) & ": "
                        If Not IsNull(data1.Recordset("EH_DIV")) Then failed = failed & data1.Recordset("EH_DIV") & ", "
                        If Not IsNull(data1.Recordset("EH_DEPT")) Then failed = failed & data1.Recordset("EH_DEPT") & ", "
                        If Not IsNull(data1.Recordset("EH_ORG")) Then failed = failed & data1.Recordset("EH_ORG") & ", "
                        If Not IsNull(data1.Recordset("EH_EMP")) Then failed = failed & data1.Recordset("EH_EMP") & ", "
                        If Not IsNull(data1.Recordset("EH_PT")) Then failed = failed & data1.Recordset("EH_PT") & ", "
                        If Not IsNull(data1.Recordset("EH_HETYPE")) Then failed = failed & data1.Recordset("EH_HETYPE") & ", "
                        If Not IsNull(data1.Recordset("EH_FDATE")) Then failed = failed & data1.Recordset("EH_FDATE") & ", "
                        If Not IsNull(data1.Recordset("EH_TDATE")) Then failed = failed & data1.Recordset("EH_TDATE") & ", "
                        If Not IsNull(data1.Recordset("EH_EDATE")) Then failed = failed & data1.Recordset("EH_EDATE") & ", "
                        If Not IsNull(data1.Recordset("EH_LOC")) Then failed = failed & data1.Recordset("EH_LOC") & ", "
                        If Not IsNull(data1.Recordset("EH_SECTION")) Then failed = failed & data1.Recordset("EH_SECTION") & ", "
                        failed = Left(failed, Len(failed) - 2) & vbCrLf
                   End If
                Else
                    failed = failed & "Rule " & CStr(c) & ": "
                    If Not IsNull(data1.Recordset("EH_DIV")) Then failed = failed & data1.Recordset("EH_DIV") & ", "
                    If Not IsNull(data1.Recordset("EH_DEPT")) Then failed = failed & data1.Recordset("EH_DEPT") & ", "
                    If Not IsNull(data1.Recordset("EH_ORG")) Then failed = failed & data1.Recordset("EH_ORG") & ", "
                    If Not IsNull(data1.Recordset("EH_EMP")) Then failed = failed & data1.Recordset("EH_EMP") & ", "
                    If Not IsNull(data1.Recordset("EH_PT")) Then failed = failed & data1.Recordset("EH_PT") & ", "
                    If Not IsNull(data1.Recordset("EH_HETYPE")) Then failed = failed & data1.Recordset("EH_HETYPE") & ", "
                    If Not IsNull(data1.Recordset("EH_FDATE")) Then failed = failed & data1.Recordset("EH_FDATE") & ", "
                    If Not IsNull(data1.Recordset("EH_TDATE")) Then failed = failed & data1.Recordset("EH_TDATE") & ", "
                    If Not IsNull(data1.Recordset("EH_EDATE")) Then failed = failed & data1.Recordset("EH_EDATE") & ", "
                    If Not IsNull(data1.Recordset("EH_LOC")) Then failed = failed & data1.Recordset("EH_LOC") & ", "
                    If Not IsNull(data1.Recordset("EH_SECTION")) Then failed = failed & data1.Recordset("EH_SECTION") & ", "
                    failed = Left(failed, Len(failed) - 2) & vbCrLf
                End If
            End If
            c = c + 1
            data1.Recordset.MoveNext
        Loop Until data1.Recordset.EOF
    End If
    
    data1.Refresh
    
    Call Display_Value
    
    Screen.MousePointer = DEFAULT
    
    If Len(failed) = 0 Then
        MsgBox "All Rules applied. Zero Out completed successfully.", vbInformation + vbOKOnly, "Hourly Entitlements Zero Out"
    Else
        MsgBox "The Zero Out for following Rules failed:" & vbCrLf & failed, vbInformation + vbOKOnly, "Hourly Entitlements Zero Out"
    End If
    
    Exit Sub
    
Mod_Err:
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdZeroOutAll", "Hourly Rollover/ZeroOut", "Modify")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
         RollBack
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Function ZeroOut_Hourly_Entitlement() As Boolean
    Dim rsHEnt As New ADODB.Recordset
    Dim SQLQ As String
    Dim EmpNo&
    Dim lngRecs&
    Dim Msg$, Title$, DgDef As Variant
    Dim Response%, pct%
    Dim prec%
    Dim xKey, xEntS, xfdateS, xtdateS
    Dim xComments
    
    On Error GoTo ZeroOut_Hourly_Entitlement_Err
    
    'Zero Out the current selected Hourly Entitlement
    
    ZeroOut_Hourly_Entitlement = False
    
    If Not CR_SnapEntitle() Then Exit Function ' create snapEntitle (form level recordset)
              
    Screen.MousePointer = DEFAULT
    
    NoErrorFlg = False
    If snapEntitle.BOF And snapEntitle.EOF Then
        MsgBox "Employees for this selection do not exist!"
        ZeroOut_Hourly_Entitlement = True
        NoErrorFlg = True
        Exit Function
    Else
        lngRecs& = snapEntitle.RecordCount
        Msg$ = lngRecs& & " Records to process" & Chr(10) & "Proceed?"
        Title$ = "Zero Out Hourly Entitlements"
        DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
        Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
        If Response% = IDNO Then    ' Evaluate response
            ZeroOut_Hourly_Entitlement = True
            NoErrorFlg = True
            Exit Function
        End If
    End If
    
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 5
        
    'Ticket #11992, Don't use BeginTrans because the Integration is called in the loop
    'gdbAdoIhr001.BeginTrans
    
    While Not snapEntitle.EOF
        prec% = prec% + 1
        pct% = Int(100 * (prec% / (lngRecs&)))
        MDIMain.panHelp(0).FloodPercent = pct%
        
        EmpNo& = snapEntitle("HE_EMPNBR")
                
        xEntS = 0: xfdateS = "": xtdateS = ""
        
        xfdateS = snapEntitle("HE_FDATE")
        xtdateS = snapEntitle("HE_TDATE")
        If chkZeroCurrent.Value Then
            xComments = "Current Hourly Ent. Chg from " & snapEntitle("HE_ENTITLE") & " to 0"
            Call Append_Accrual(EmpNo&, snapEntitle("HE_TYPE"), snapEntitle("HE_TDATE"), -Val(snapEntitle("HE_ENTITLE") & ""), "Z", xComments)
            
            'Update Hourly Entitlement - Previous Zero Out
            rsHEnt.Open "SELECT * FROM HRENTHRS WHERE HE_ID= " & snapEntitle("HE_ID"), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsHEnt.EOF Then
                rsHEnt("HE_ENTITLE") = 0
                rsHEnt.Update
            End If
            rsHEnt.Close
            Set rsHEnt = Nothing
            
            xEntS = 0
        End If
        If chkZeroPrev.Value Then
            xComments = "Prev. Hourly Ent. Chg from " & snapEntitle("HE_PREV") & " to 0"
            Call Append_Accrual(EmpNo&, snapEntitle("HE_TYPE"), snapEntitle("HE_TDATE"), -Val(snapEntitle("HE_PREV") & ""), "Z", xComments)
            
            'Update Hourly Entitlement - Previous Zero Out
            rsHEnt.Open "SELECT * FROM HRENTHRS WHERE HE_ID= " & snapEntitle("HE_ID"), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsHEnt.EOF Then
                rsHEnt("HE_PREV") = 0
                rsHEnt.Update
            End If
            rsHEnt.Close
            Set rsHEnt = Nothing
            
            xEntS = 0
        End If
    
        'snapEntitle.Update
        
        xKey = EmpNo&
        xKey = xKey & "|" & Format(xfdateS, "dd-mmm-yyyy")
        xKey = xKey & "|" & Format(xtdateS, "dd-mmm-yyyy")
        xKey = xKey & "|" & snapEntitle("HE_TYPE")
        xKey = xKey & "|" & xEntS
        xKey = xKey & "|" & Format(xfdateS, "dd-mmm-yyyy") 'Format(Date, "dd-mmm-yyyy") 'Transaction Date
        Call Entitlements_Master_Integration(xKey, EmpNo&) 'George added for Advance Tracker
        DoEvents
    
lblNextZRec:
        snapEntitle.MoveNext
        DoEvents
    
    Wend
    
    MDIMain.panHelp(0).FloodType = 0
    
    snapEntitle.Close
    
    ZeroOut_Hourly_Entitlement = True

    Screen.MousePointer = DEFAULT
    'gdbAdoIhr001.CommitTrans


Exit Function

ZeroOut_Hourly_Entitlement_Err:
    Screen.MousePointer = DEFAULT
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Zero Out Entitle", "HRENTHRS", "edit/Add")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        RollBack
        Resume Next
    Else
        Unload Me
    End If
End Function

'Private Sub cmdUpdateAll_Click()
'Dim failed As String
'Dim c As Long
'
'On Error GoTo Mod_Err
'If Not gSec_Upd_Hrly_Entitlements Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If
'
'failed = ""
'c = 1
'If data1.Recordset.EOF = False And data1.Recordset.BOF = False Then
'    data1.Recordset.MoveFirst
'    Do
'        Call Display_Value
'
'        If chkManual.Value = False Then
'            If chkMUEntitle() Then
'               If DoWork = False Then
'                    failed = failed & "Rule " & CStr(c) & ": "
'                    If Not IsNull(data1.Recordset("EH_DIV")) Then failed = failed & data1.Recordset("EH_DIV") & ", "
'                    If Not IsNull(data1.Recordset("EH_DEPT")) Then failed = failed & data1.Recordset("EH_DEPT") & ", "
'                    If Not IsNull(data1.Recordset("EH_ORG")) Then failed = failed & data1.Recordset("EH_ORG") & ", "
'                    If Not IsNull(data1.Recordset("EH_EMP")) Then failed = failed & data1.Recordset("EH_EMP") & ", "
'                    If Not IsNull(data1.Recordset("EH_PT")) Then failed = failed & data1.Recordset("EH_PT") & ", "
'                    If Not IsNull(data1.Recordset("EH_HETYPE")) Then failed = failed & data1.Recordset("EH_HETYPE") & ", "
'                    If Not IsNull(data1.Recordset("EH_FDATE")) Then failed = failed & data1.Recordset("EH_FDATE") & ", "
'                    If Not IsNull(data1.Recordset("EH_TDATE")) Then failed = failed & data1.Recordset("EH_TDATE") & ", "
'                    If Not IsNull(data1.Recordset("EH_EDATE")) Then failed = failed & data1.Recordset("EH_EDATE") & ", "
'                    If Not IsNull(data1.Recordset("EH_LOC")) Then failed = failed & data1.Recordset("EH_LOC") & ", "
'                    If Not IsNull(data1.Recordset("EH_SECTION")) Then failed = failed & data1.Recordset("EH_SECTION") & ", "
'                    failed = Left(failed, Len(failed) - 2) & vbCrLf
'               End If
'            Else
'                failed = failed & "Rule " & CStr(c) & ": "
'                If Not IsNull(data1.Recordset("EH_DIV")) Then failed = failed & data1.Recordset("EH_DIV") & ", "
'                If Not IsNull(data1.Recordset("EH_DEPT")) Then failed = failed & data1.Recordset("EH_DEPT") & ", "
'                If Not IsNull(data1.Recordset("EH_ORG")) Then failed = failed & data1.Recordset("EH_ORG") & ", "
'                If Not IsNull(data1.Recordset("EH_EMP")) Then failed = failed & data1.Recordset("EH_EMP") & ", "
'                If Not IsNull(data1.Recordset("EH_PT")) Then failed = failed & data1.Recordset("EH_PT") & ", "
'                If Not IsNull(data1.Recordset("EH_HETYPE")) Then failed = failed & data1.Recordset("EH_HETYPE") & ", "
'                If Not IsNull(data1.Recordset("EH_FDATE")) Then failed = failed & data1.Recordset("EH_FDATE") & ", "
'                If Not IsNull(data1.Recordset("EH_TDATE")) Then failed = failed & data1.Recordset("EH_TDATE") & ", "
'                If Not IsNull(data1.Recordset("EH_EDATE")) Then failed = failed & data1.Recordset("EH_EDATE") & ", "
'                If Not IsNull(data1.Recordset("EH_LOC")) Then failed = failed & data1.Recordset("EH_LOC") & ", "
'                If Not IsNull(data1.Recordset("EH_SECTION")) Then failed = failed & data1.Recordset("EH_SECTION") & ", "
'                failed = Left(failed, Len(failed) - 2) & vbCrLf
'            End If
'        End If
'        c = c + 1
'        data1.Recordset.MoveNext
'    Loop Until data1.Recordset.EOF
'End If
'
'data1.Refresh
'
'Call Display_Value
'
'Screen.MousePointer = DEFAULT
'
'If Len(failed) = 0 Then
'    MsgBox "All Rules applied", vbInformation + vbOKOnly, "Hourly Entitlements"
'Else
'    MsgBox "The Following Rules failed:" & vbCrLf & failed, vbInformation + vbOKOnly, "Hourly Entitlements"
'End If
'
'Exit Sub
'
'Mod_Err:
'
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdateAll", "Hourly", "Modify")
'Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
'     RollBack
'    Resume Next
'Else
'    Unload Me
'End If
'End Sub

Private Sub Form_Activate()

Call SET_UP_MODE

Call INI_Controls(Me)

glbOnTop = "FRMUHRSENT"

End Sub

Private Sub Form_Load()

Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found
Dim X%
Dim SQLQ


MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = "FRMUHRSENT"

Unload frmUEntitle

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    clpCode(4).Visible = False
    clpProv.Left = clpCode(4).Left
    clpProv.Top = clpCode(4).Top
    clpProv.Visible = True
    lblLocation.Caption = "Province"
    vbxTrueGrid.Columns(8).Caption = "Province"
End If

FlagRefresh = False

data1.ConnectionString = glbAdoIHRDB
SQLQ = "SELECT DISTINCT EH_DIV,EH_DEPT,EH_ORG,EH_FDATE,EH_TDATE,EH_EMP,EH_SECTION,EH_LOC,EH_PT,EH_HETYPE,EH_MANUAL,EH_EDATE FROM HR_HOURLYENT "
If glbDIVCount = 1 And glbLinamar Then
    SQLQ = SQLQ & " WHERE EH_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
End If
data1.RecordSource = SQLQ
data1.Refresh

Select Case glbCompWDate$ ' sets field reference for basic 'which date'
    Case "O": fglbWDate$ = "ED_DOH"
    Case "S": fglbWDate$ = "ED_SENDTE"
    Case "U": fglbWDate$ = "ED_UNION"
    Case "L": fglbWDate$ = "ED_LTHIRE"
    Case "D": fglbWDate$ = "ED_USRDAT1"
End Select
Select Case glbEntOutStandingS$
    Case "2": fglbWDateS$ = "ED_DOH"
    Case "3": fglbWDateS$ = "ED_SENDTE"
    Case "4": fglbWDateS$ = "ED_LTHIRE"
    Case "5": fglbWDateS$ = "ED_USRDAT1"
    Case "6": fglbWDateS$ = "ED_UNION"
End Select

If UCase(glbCompEntVac$) = "M" Then
    vbxTrueGrid.Columns(3).Visible = False
End If
If glbWFC Then
    lblSection.FontBold = True
End If

Screen.MousePointer = HOURGLASS
vbxTrueGrid.Columns(0).Caption = lStr(vbxTrueGrid.Columns(0).Caption)
vbxTrueGrid.Columns(1).Caption = lStr(vbxTrueGrid.Columns(1).Caption)
vbxTrueGrid.Columns(2).Caption = lStr(vbxTrueGrid.Columns(2).Caption)

Call setRptCaption(Me)

If glbCBrant Then
    lblDtRange.Visible = False
    dlpFrom.Visible = False
    dlpTo.Visible = False
End If

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

Private Sub Form_Resize()
'If Me.Height >= 6960 + VacFram.Height + panControls.Height + 230 Then
'    scrControl.Value = 0
'    VacFram.Top = 4570 '3960
'    scrControl.Visible = False
'    Exit Sub
'End If
'scrControl.Visible = True
'scrControl.Max = VacFram.Height + panControls.Height + 6960 - Me.Height
'scrControl.Left = Me.Width - scrControl.Width - 120
'If Me.Height - scrControl.Top - panControls.Height - 400 > 0 Then
'    scrControl.Height = Me.Height - scrControl.Top - panControls.Height - 400
'Else
'    scrControl.Height = 0
'End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select FROM the menu the appropriate function."

Set frmUEntitle = Nothing  'carmen apr 2000
End Sub

Private Sub medEntitle_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medGTServ_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medLTServ_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optD_Click(Index As Integer, Value As Integer)
    Call ST_OPT_VALUE
End Sub

Private Sub optD_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optF_Click(Index As Integer, Value As Integer)
    Call ST_OPT_VALUE
End Sub

Private Sub optF_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optH_Click(Index As Integer, Value As Integer)
    Call ST_OPT_VALUE
End Sub

Private Sub optH_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Sub Display_Value()
Dim SQLQ, xOrder, nOrder, aa, SQLQW, glbiOneWhere
Dim rsVE As New ADODB.Recordset
Dim X

clpDiv.Text = ""
clpDept.Text = ""
clpCode(0).Text = ""
dlpFrom.Text = ""
dlpTo.Text = ""
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

If Not data1.Recordset.EOF Then
    Call getWSQLQ("D")
    
    SQLQ = "SELECT * FROM HR_HOURLYENT WHERE " & fglbVSQLQ
    SQLQ = SQLQ & "Order By EH_DIV,EH_DEPT,EH_ORG, EH_FDATE,EH_EMP,EH_PT,EH_LOC,EH_SECTION,EH_ORDER "
    rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    
    If Not IsNull(data1.Recordset("EH_DIV")) Then clpDiv.Text = data1.Recordset("EH_DIV")
    If Not IsNull(data1.Recordset("EH_DEPT")) Then clpDept.Text = data1.Recordset("EH_DEPT")
    If Not IsNull(data1.Recordset("EH_ORG")) Then clpCode(0).Text = data1.Recordset("EH_ORG")
    If Not IsNull(data1.Recordset("EH_FDATE")) Then dlpFrom.Text = data1.Recordset("EH_FDATE")
    If Not IsNull(data1.Recordset("EH_TDATE")) Then dlpTo.Text = data1.Recordset("EH_TDATE")
    If Not IsNull(data1.Recordset("EH_EDATE")) Then dlpAsOf.Text = data1.Recordset("EH_EDATE")
    If Not IsNull(data1.Recordset("EH_EMP")) Then clpCode(1).Text = data1.Recordset("EH_EMP")
    If Not IsNull(data1.Recordset("EH_PT")) Then clpPT.Text = data1.Recordset("EH_PT")
    If Not IsNull(data1.Recordset("EH_HETYPE")) Then clpCode(2).Text = data1.Recordset("EH_HETYPE")
    If Not IsNull(data1.Recordset("EH_SECTION")) Then clpCode(3).Text = data1.Recordset("EH_SECTION")
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
        If Not IsNull(data1.Recordset("EH_LOC")) Then clpProv.Text = data1.Recordset("EH_LOC")
    Else
        If Not IsNull(data1.Recordset("EH_LOC")) Then clpCode(4).Text = data1.Recordset("EH_LOC")
    End If
    If Not IsNull(data1.Recordset("EH_MANUAL")) Then
        chkManual.Value = data1.Recordset("EH_MANUAL")
    End If
    
'    Do While Not rsVE.EOF
'        xOrder = rsVE("EH_ORDER")
'        nOrder = Format(Val(xOrder), "##0") - 1
'        If Not (nOrder < 0 Or nOrder > 19) Then
'            If Not IsNull(rsVE("EH_BMONTH")) Then medLTServ(nOrder) = rsVE("EH_BMONTH")
'            If Not IsNull(rsVE("EH_EMONTH")) Then medGTServ(nOrder) = rsVE("EH_EMONTH")
'            If Not IsNull(rsVE("EH_ENTITLE")) Then medEntitle(nOrder) = rsVE("EH_ENTITLE")
'            If rsVE("EH_TYPE") = "D" Then optD(nOrder) = True
'            If rsVE("EH_TYPE") = "H" Then optH(nOrder) = True
'            If rsVE("EH_TYPE") = "F" Then optF(nOrder) = True
'        End If
'        rsVE.MoveNext
'    Loop
    rsVE.Close
End If

Call SET_UP_MODE


End Sub

Private Sub optDH_Click(Index As Integer, Value As Integer)
    If Index = 1 Then
        If optDH(1) Then
            MsgBox "Make sure Employee's Hours/Day is specified on the Position screen otherwise the Rollover will skip for that Employee.", vbExclamation, "Maximum Rollover in Days"
        End If
    End If
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
           
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT DISTINCT EH_DIV,EH_DEPT,EH_ORG,EH_EDATE,EH_FDATE,EH_TDATE,EH_EMP,EH_SECTION,EH_LOC,EH_PT,EH_HETYPE,EH_MANUAL FROM HR_HOURLYENT "
        If glbDIVCount = 1 And glbLinamar Then
            SQLQ = SQLQ & " WHERE EH_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    
        data1.RecordSource = SQLQ
        data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
End Sub

Private Sub modSetFGlobals(strTyp$)
fglbSick% = False
fglbVac% = True
If glbCompEntVac$ = "M" Then
    fglbCompMonthly% = True
Else
    fglbCompMonthly% = False
End If
ffieldEntitle$ = "ED_VAC"
ffieldPEntitle$ = "ED_PVAC"
fglbCode$ = "VAC"

End Sub

Sub ST_OPT_VALUE()

End Sub

'Private Function modDelRecs()
'Dim BD As Integer
'Dim SQLQ As String, SQL1 As String, countr As Integer
'Dim Dat1 As Variant, Dat2 As Variant
'Dim iOneWhere As Integer, NxtSQL As String, strReas$
'Dim oldEntitleUpd
'Dim rsHRE As New ADODB.Recordset
'Dim rzAttend As New ADODB.Recordset
'Dim rsCurSal As New ADODB.Recordset
'Dim rsHREmp As New ADODB.Recordset
'Dim pct#, prec#
'Dim xKey
'
'modDelRecs = False
'
'On Error GoTo modDelRecs_Err
'
'Screen.MousePointer = HOURGLASS
'
'Call getWSQLQ("")
'
'SQLQ = "SELECT * FROM HRENTHRS WHERE " & fglbWSQLQ
'SQLQ = SQLQ & " AND HE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
'rsHRE.Open SQLQ, gdbAdoIhr001, adOpenStatic
'pct# = 0
'prec# = 0
'If NumAddRec% = 0 Then
'    prec# = 0
'Else
'    If rsHRE.RecordCount <> 0 Then
'        prec# = 90 / rsHRE.RecordCount
'    End If
'End If
'
'MDIMain.panHelp(0).FloodType = 1
'MDIMain.panHelp(0).FloodPercent = 0
'
'Do Until rsHRE.EOF
'    pct# = pct# + prec#
'    MDIMain.panHelp(0).FloodPercent = pct#
'    'In Vadim we will have to balance out to zero the entitlement
'    'Call Append_Accrual(rsHRE("HE_EMPNBR"), rsHRE("HE_TYPE"), Date, 0 - rsHRE("HE_ENTITLE"), "U", "Mass deleted the existing Hourly Entitlement")
'    If rsHRE("HE_ENTITLE") - rsHRE("HE_TAKEN") < 0 Then
'        'Used more entitlement than entitled - Jerry said to borrow it from next year
'        'To borrow, add a new record in Attendance for next year
'        If glbVadim Then
'            'Add Record in Attendance screen
'            SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR =" & rsHRE("HE_EMPNBR")
'            SQLQ = SQLQ & " AND AD_REASON = '" & rsHRE("HE_TYPE") & "'"
'            SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(DateAdd("d", 1, rsHRE("HE_TDATE")))
'            rzAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'            If rzAttend.EOF Then
'                rzAttend.AddNew
'            End If
'            rzAttend("AD_COMPNO") = "001"
'            rzAttend("AD_EMPNBR") = rsHRE("HE_EMPNBR")
'            rzAttend("AD_DOA") = DateAdd("d", 1, rsHRE("HE_TDATE")) 'Next year
'            rzAttend("AD_REASON") = rsHRE("HE_TYPE")
'            rzAttend("AD_HRS") = Abs(rsHRE("HE_ENTITLE") - rsHRE("HE_TAKEN")) 'Borrowed Hours
'
'            SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_ORG,ED_GLNO FROM HREMP WHERE ED_EMPNBR = " & rsHRE("HE_EMPNBR")
'            rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
'            If Not rsHREmp.EOF Then
'                rzAttend("AD_PAYROLL_ID") = rsHREmp("ED_PAYROLL_ID")
'                rzAttend("AD_GLNO") = rsHREmp("ED_GLNO")
'                rzAttend("AD_ORG") = rsHREmp("ED_ORG")
'            End If
'            rsHREmp.Close
'
'            SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & rsHRE("HE_EMPNBR")
'            rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
'            If Not rsCurSal.BOF Then
'                If rsCurSal("SH_SALARY") > 0 Then
'                    rzAttend("AD_SALARY") = rsCurSal("SH_SALARY")
'                    rzAttend("AD_SALCD") = rsCurSal("SH_SALCD")
'                End If
'            End If
'            rsCurSal.Close
'
'            SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & rsHRE("HE_EMPNBR")
'            rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
'            If Not rsCurSal.EOF Then
'                rzAttend("AD_JOB") = rsCurSal("JH_JOB")
'                rzAttend("AD_DHRS") = rsCurSal("JH_DHRS")
'                rzAttend("AD_WHRS") = rsCurSal("JH_WHRS")
'            End If
'            rsCurSal.Close
'
'            rzAttend("AD_COMM") = "Exceeded Hours in last year so borrowed from this year"
'            rzAttend("AD_LDATE") = Date
'            rzAttend("AD_LUSER") = "BORROWED"
'            rzAttend("AD_LTIME") = Time$
'            rzAttend.Update
'            rzAttend.Close
'        End If
'        Call Append_Accrual(rsHRE("HE_EMPNBR"), rsHRE("HE_TYPE"), dlpTo.Text, rsHRE("HE_ENTITLE") - rsHRE("HE_TAKEN"), "D", "Mass deleted existing Hourly Entitlement")
'    Else
'        Call Append_Accrual(rsHRE("HE_EMPNBR"), rsHRE("HE_TYPE"), dlpTo.Text, rsHRE("HE_TAKEN") - rsHRE("HE_ENTITLE"), "D", "Mass deleted existing Hourly Entitlement")
'    End If
'
'    xKey = rsHRE("HE_EMPNBR")
'    xKey = xKey & "|" & Format(dlpFrom.Text, "dd-mmm-yyyy")
'    xKey = xKey & "|" & Format(dlpTo.Text, "dd-mmm-yyyy")
'    xKey = xKey & "|" & clpCode(2).Text
'    xKey = xKey & "|"
'    xKey = xKey & "|" & Format(Date, "dd-mmm-yyyy") 'Transaction Date
'    DoEvents
'    Call Entitlements_Master_Integration(xKey, , True)
'    DoEvents
'
'    rsHRE.MoveNext
'    DoEvents
'Loop
'rsHRE.Close
'
'SQLQ = "DELETE FROM HRENTHRS WHERE " & fglbWSQLQ
'SQLQ = SQLQ & " AND HE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
'gdbAdoIhr001.Execute SQLQ
'
'MDIMain.panHelp(0).FloodPercent = 100
'MDIMain.panHelp(0).FloodType = 0
'
'modDelRecs = True
'
'Exit Function
'
'modDelRecs_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modDelRecs", "DeleteHrEntitlement", "Delete")
'modDelRecs = False
'Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
'    Resume Next
'Else
'    Unload Me
'End If
'
'End Function

'Private Function modInsSelectionCBrant() 'Ticket #12524
''Share the logic of Sick Entitlement
'Dim HEID&
'Dim strDivision$
'Dim strJob$, dblServiceYears#
'Dim spt As Variant, varStartDate As Variant, lngRecs&
'Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
'Dim dblFTEHours#
'Dim dblEntitleUpd#, DtTm As Variant
'Dim Msg$, Title$, DgDef As Variant
'Dim Response%, pct#
'Dim prec#, SQLQ As String, NumRec As Integer
'Dim snapDuplic As New ADODB.Recordset
'Dim oldEntitleUpd
'Dim xKey
'Dim xHEFromDate, xHEToDate, xDiffYear
'
'On Error GoTo modInsSelectionCBrant_Err
'
'modInsSelectionCBrant = False
'
'If Not CR_SnapAddEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
'
'If fTablHREMP.State <> adStateClosed Then fTablHREMP.Close
'fTablHREMP.Open "HRENTHRS", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
'
'Screen.MousePointer = HOURGLASS
'
'MDIMain.panHelp(0).FloodType = 1
'MDIMain.panHelp(0).FloodPercent = 0
'
'pct# = 0
'prec# = 0
'If NumAddRec% = 0 Then
'    prec# = 0
'Else
'    prec# = 90 / NumAddRec% 'SBH avoid divid by zero...
'End If
'
'For x% = 0 To 19
'    If IsNumeric(medGTServ(x%)) Then
'        If glbFrench Then
'            If medGTServ(x%) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
'        Else
'            If Val(medGTServ(x%)) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
'        End If
'    End If
'    If Len(medLTServ(x%)) > 0 And Len(medGTServ(x%)) = 0 Then medGTServ(x%) = 9999999
'Next
'BeginTrans
'
'While Not SnapAddEntitle.EOF
'    pct# = pct# + prec#
'    MDIMain.panHelp(0).FloodPercent = pct#
'
'
'    If IsNull(SnapAddEntitle(fglbWDateS$)) Then
'        GoTo lblNext2Rec
'    End If
'
'    varStartDate = SnapAddEntitle(fglbWDateS$)   ' set start date
'    xDiffYear = DateDiff("d", varStartDate, Now) / 365
'
'    If xDiffYear > 1 Then
'        xHEFromDate = DateAdd("YYYY", CInt(xDiffYear), varStartDate)
'    Else
'        xHEFromDate = varStartDate
'    End If
'
'    If xHEFromDate > Now Then
'        xHEFromDate = DateAdd("YYYY", -1, xHEFromDate)
'    End If
'    xHEToDate = DateAdd("YYYY", 1, xHEFromDate)
'    xHEToDate = DateAdd("d", -1, xHEToDate)
'
'    If Not IsNumeric(SnapAddEntitle("JH_DHRS")) Then
'        dblDHours# = 0
'    Else
'        dblDHours# = SnapAddEntitle("JH_DHRS")
'    End If
'    If Not IsNumeric(SnapAddEntitle("JH_FTENUM")) Then
'        dblFTEHours# = 0
'    Else
'        dblFTEHours# = SnapAddEntitle("JH_FTENUM")
'    End If
'
'    'dblServiceYears# = (DateDiff("d", varStartDate, Now) / 365) * 12
'    'dblServiceYears# = MonthDiff(CVDate(varStartDate), Date)
'    dblServiceYears# = MonthDiff(CVDate(varStartDate), dlpAsOf.Text)    'Ticket #17924
'
'    If dblServiceYears# < 0 Then GoTo lblNext2Rec     'laura 03/06/98
'
'    intWhereFit& = -1   ' first record can be just less than
'    For x% = 0 To 19
'        If medLTServ(x%) = "" And medGTServ(x%) = "" Then Exit For
'        If IsNumeric(Val(medLTServ(x%))) And medGTServ(x%) = "" Then
'            If dblServiceYears# >= CDbl(Val(medLTServ(x%))) Then
'                intWhereFit& = x%
'                Exit For
'            End If
'        End If
'        If IsNumeric(medLTServ(x%)) And IsNumeric(medGTServ(x%)) Then
'            If dblServiceYears# >= CDbl(Val(medLTServ(x%))) And dblServiceYears# <= CDbl(Val(medGTServ(x%))) Then
'                intWhereFit& = x%
'                Exit For
'            End If
'        End If
'    Next x%
'
'    If intWhereFit& = -1 Then GoTo lblNext2Rec  ' skip record if not in any of the ranges
'    dblNewEntitle# = Val(medEntitle(intWhereFit&))   'laura
'    If optD(intWhereFit&) = True Then           ' Entitlements entered in days
'        dblNewEntitle# = dblNewEntitle# * dblDHours#
'    End If
'    If optH(intWhereFit&) = True Then           ' Entitlements entered in Hours
'        dblNewEntitle# = dblNewEntitle#
'    End If
'    If optF(intWhereFit&) = True Then           ' Entitlements entered in FTE
'        dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
'    End If
'
'    SQLQ = "SELECT HE_EMPNBR,HE_TYPE,HE_ID ,"
'    SQLQ = SQLQ & " HE_ENTITLE, HE_TDATE FROM HRENTHRS "
'    SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
'    SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
'    SQLQ = SQLQ & " AND HE_TDATE = " & Date_SQL(xHEToDate)  'dlpTo.Text
'    snapDuplic.Open SQLQ, gdbAdoIhr001, adOpenKeyset
'    If Not snapDuplic.EOF And Not snapDuplic.BOF Then
''        xID = snapDuplic("HE_ID")
'        snapDuplic.MoveLast
'    End If
'
'    NumRec = snapDuplic.RecordCount
'    If snapDuplic.EOF Then
'        oldEntitleUpd = 0
'    Else
'        oldEntitleUpd = snapDuplic("HE_ENTITLE")
'    End If
'    If Accum = True Then
'        If NumRec > 0 Then
'            dblEntitleUpd = snapDuplic("HE_ENTITLE")
'        Else
'            dblEntitleUpd = 0
'        End If
'    Else
'        dblEntitleUpd = 0
'    End If
'    snapDuplic.Close
'
'    If Accum = True Then
'        dblEntitleUpd = dblEntitleUpd + dblNewEntitle
'    Else
'        dblEntitleUpd = dblNewEntitle
'    End If
'
'    DtTm = Now
'
'If Accum = True Then
'    If NumRec > 0 Then  'if accumulate and found duplicate record
'
'        SQLQ = "UPDATE HRENTHRS "
'        SQLQ = SQLQ & " SET HE_ENTITLE = " & dblEntitleUpd & " "
'        SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
'        SQLQ = SQLQ & " AND HRENTHRS.HE_TYPE = '" & clpCode(2).Text & "' "
'        SQLQ = SQLQ & " AND HRENTHRS.HE_TDATE = " & Date_SQL(xHEToDate)
'
'        gdbAdoIhr001.Execute (SQLQ)
'        Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, Date, dblEntitleUpd - oldEntitleUpd, "U", "Mass changed the existing Hourly Entitlement")
'    Else
'        fTablHREMP.AddNew     'if accumulate and no duplicate record
'        fTablHREMP("HE_EMPNBR") = SnapAddEntitle("ED_EMPNBR")
'        fTablHREMP("HE_COMPNO") = "001"
'        fTablHREMP("HE_TYPE_TABL") = "ADRE"
'        fTablHREMP("HE_TYPE") = clpCode(2).Text
'        fTablHREMP("HE_FDATE") = xHEFromDate 'dlpFrom.Text
'        fTablHREMP("HE_TDATE") = xHEToDate  'dlpTo.Text
'        fTablHREMP("HE_ENTITLE") = dblEntitleUpd
'        fTablHREMP("HE_COE") = True
'        fTablHREMP("HE_DHRS") = SnapAddEntitle("ED_DHRS")
'        fTablHREMP("HE_LDATE") = Now
'        fTablHREMP("HE_LTIME") = Time$
'        fTablHREMP("HE_LUSER") = glbUserID
'        fTablHREMP.Update
'        '    xID = fTablHREMP("HE_ID")
'        Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, xHEFromDate, dblEntitleUpd, "A", "Mass added the Hourly Entitlement")
'
'    End If
'Else
'    SQLQ$ = "DELETE FROM HRENTHRS "
'    SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
'    SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
'    SQLQ = SQLQ & " AND HE_TDATE = " & Date_SQL(xHEToDate)
'
'    gdbAdoIhr001.Execute SQLQ
'
'    fTablHREMP.AddNew
'
'    fTablHREMP("HE_EMPNBR") = SnapAddEntitle("ED_EMPNBR")
'    fTablHREMP("HE_COMPNO") = "001"
'    fTablHREMP("HE_TYPE_TABL") = "ADRE"
'    fTablHREMP("HE_TYPE") = clpCode(2).Text
'    fTablHREMP("HE_FDATE") = xHEFromDate
'    fTablHREMP("HE_TDATE") = xHEToDate
'    fTablHREMP("HE_ENTITLE") = dblEntitleUpd
'    fTablHREMP("HE_COE") = True
'    fTablHREMP("HE_DHRS") = SnapAddEntitle("ED_DHRS")
'    fTablHREMP("HE_LDATE") = Now
'    fTablHREMP("HE_LTIME") = Time$
'    fTablHREMP("HE_LUSER") = glbUserID
'    fTablHREMP.Update
'    '    xID = fTablHREMP("HE_ID")
'    If NumRec > 0 Then  'if accumulate and found duplicate record
'        Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, xHEFromDate, dblEntitleUpd - oldEntitleUpd, "U", "Mass modified the Hourly Entitlement")
'    Else
'        Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, xHEFromDate, dblEntitleUpd, "A", "Mass added the Hourly Entitlement")
'    End If
'
'End If
'    DoEvents
'    xKey = SnapAddEntitle("ED_EMPNBR")
'    xKey = xKey & "|" & Format(xHEFromDate, "dd-mmm-yyyy")
'    xKey = xKey & "|" & Format(xHEToDate, "dd-mmm-yyyy")
'    xKey = xKey & "|" & clpCode(2).Text
'    xKey = xKey & "|" & dblEntitleUpd
'    xKey = xKey & "|" & Format(Date, "dd-mmm-yyyy") 'Transaction Date
'    Call Entitlements_Master_Integration(xKey, 0)
'    DoEvents
'lblNext2Rec:
'    SnapAddEntitle.MoveNext
'Wend
'
'modInsSelectionCBrant = True
'
'MDIMain.panHelp(0).FloodPercent = 100
'MDIMain.panHelp(0).FloodType = 0
'
'CommitTrans
'
'fTablHREMP.Close
'
'SnapAddEntitle.Close
'
'Screen.MousePointer = DEFAULT
'
'Exit Function
'
'modInsSelectionCBrant_Err:
'
'If Err = 13 Or Err = 94 Or Err = 3018 Then
'    Err = 0
'    Resume Next
'   'MsgBox "Conflicting Dates"
'    Screen.MousePointer = DEFAULT
'    Exit Function
'End If
'
'Screen.MousePointer = DEFAULT
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "InsertEntitle", "HR_EMP", "edit/Add")
'Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
'    'Rollback
'    Resume Next
'Else
'    Unload Me
'End If
'
'
'End Function
'
'Private Function modInsSelection()
''laura 03/04/98
'Dim HEID&
'Dim strDivision$
'Dim strJob$, dblServiceYears#
'Dim spt As Variant, varStartDate As Variant, lngRecs&
'Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
'Dim dblFTEHours#
'Dim dblEntitleUpd#, DtTm As Variant
'Dim Msg$, Title$, DgDef As Variant
'Dim Response%, pct#
'Dim prec#, SQLQ As String, NumRec As Integer
'Dim snapDuplic As New ADODB.Recordset
'Dim oldEntitleUpd
'Dim xKey
'
'On Error GoTo modInsSelection_Err
'
'modInsSelection = False
'
'If Not CR_SnapAddEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
'
'If fTablHREMP.State <> adStateClosed Then fTablHREMP.Close
'fTablHREMP.Open "HRENTHRS", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
'
'Screen.MousePointer = HOURGLASS
'
'MDIMain.panHelp(0).FloodType = 1
'MDIMain.panHelp(0).FloodPercent = 0
'pct# = 0
'prec# = 0
'If NumAddRec% = 0 Then
'    prec# = 0
'Else
'    prec# = 90 / NumAddRec% 'SBH avoid divid by zero...
'End If
'For x% = 0 To 19
'    If IsNumeric(medGTServ(x%)) Then
'        If glbFrench Then
'            If medGTServ(x%) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
'        Else
'            If Val(medGTServ(x%)) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
'        End If
'    End If
'    If Len(medLTServ(x%)) > 0 And Len(medGTServ(x%)) = 0 Then medGTServ(x%) = 9999999
'Next
'BeginTrans
'
'While Not SnapAddEntitle.EOF
'    pct# = pct# + prec#
'    MDIMain.panHelp(0).FloodPercent = pct#
'
'
'    If IsNull(SnapAddEntitle(fglbWDate$)) Then
'        GoTo lblNext2Rec
'    End If
'
'    varStartDate = SnapAddEntitle(fglbWDate$)  ' set start date
'    If Not IsNumeric(SnapAddEntitle("JH_DHRS")) Then
'        dblDHours# = 0
'    Else
'        dblDHours# = SnapAddEntitle("JH_DHRS")
'    End If
'    If Not IsNumeric(SnapAddEntitle("JH_FTENUM")) Then
'        dblFTEHours# = 0
'    Else
'        dblFTEHours# = SnapAddEntitle("JH_FTENUM")
'    End If
'
'    'dblServiceYears# = (DateDiff("d", varStartDate, Now) / 365) * 12
'    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(dlpAsOf.Text))    'Ticket #17924
'    If dblServiceYears# < 0 Then GoTo lblNext2Rec     'laura 03/06/98
'    intWhereFit& = -1   ' first record can be just less than
'    For x% = 0 To 19
'        If medLTServ(x%) = "" And medGTServ(x%) = "" Then Exit For
'        If IsNumeric(medLTServ(x%)) And medGTServ(x%) = "" Then
'            If dblServiceYears# >= CDbl(medLTServ(x%)) Then
'                intWhereFit& = x%
'                Exit For
'            End If
'        End If
'        If IsNumeric(medLTServ(x%)) And IsNumeric(medGTServ(x%)) Then
'            If dblServiceYears# >= CDbl(medLTServ(x%)) And dblServiceYears# <= CDbl(medGTServ(x%)) Then
'                intWhereFit& = x%
'                Exit For
'            End If
'        End If
'    Next x%
'
'    If intWhereFit& = -1 Then GoTo lblNext2Rec  ' skip record if not in any of the ranges
'    If glbFrench Then
'        dblNewEntitle# = CDbl(medEntitle(intWhereFit&))   'laura
'    Else
'        dblNewEntitle# = Val(medEntitle(intWhereFit&))   'laura
'    End If
'    If optD(intWhereFit&) = True Then           ' Entitlements entered in days
'        dblNewEntitle# = dblNewEntitle# * dblDHours#
'    End If
'    If optH(intWhereFit&) = True Then           ' Entitlements entered in Hours
'        dblNewEntitle# = dblNewEntitle#
'    End If
'    If optF(intWhereFit&) = True Then           ' Entitlements entered in FTE
'        dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
'    End If
'
'    SQLQ = "SELECT HE_EMPNBR,HE_TYPE,HE_ID ,"
'    SQLQ = SQLQ & " HE_ENTITLE, HE_TDATE FROM HRENTHRS "
'    SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
'    SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
'    SQLQ = SQLQ & " AND HE_TDATE = " & Date_SQL(dlpTo.Text)
'
'    snapDuplic.Open SQLQ, gdbAdoIhr001, adOpenKeyset
'    If Not snapDuplic.EOF And Not snapDuplic.BOF Then
''        xID = snapDuplic("HE_ID")
'        snapDuplic.MoveLast
'    End If
'
'    NumRec = snapDuplic.RecordCount
'    If snapDuplic.EOF Then
'        oldEntitleUpd = 0
'    Else
'        oldEntitleUpd = snapDuplic("HE_ENTITLE")
'    End If
'    If Accum = True Then
'        If NumRec > 0 Then
'            dblEntitleUpd = snapDuplic("HE_ENTITLE")
'        Else
'            dblEntitleUpd = 0
'        End If
'    Else
'        dblEntitleUpd = 0
'    End If
'
'    snapDuplic.Close
'    If Accum = True Then
'        dblEntitleUpd = dblEntitleUpd + dblNewEntitle
'    Else
'        dblEntitleUpd = dblNewEntitle
'    End If
'
'    DtTm = Now
'
'If Accum = True Then
'    If NumRec > 0 Then  'if accumulate and found duplicate record
'
'        SQLQ = "UPDATE HRENTHRS "
'        SQLQ = SQLQ & " SET HE_ENTITLE = " & dblEntitleUpd & " "
'        SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
'        SQLQ = SQLQ & " AND HRENTHRS.HE_TYPE = '" & clpCode(2).Text & "' "
'        SQLQ = SQLQ & " AND HRENTHRS.HE_TDATE = " & Date_SQL(dlpTo.Text)
'
'        gdbAdoIhr001.Execute (SQLQ)
'        Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpAsOf.Text, dblEntitleUpd - oldEntitleUpd, "U", "Mass changed the existing Hourly Entitlement") 'Ticket #17924
'    Else
'        fTablHREMP.AddNew     'if accumulate and no duplicate record
'        fTablHREMP("HE_EMPNBR") = SnapAddEntitle("ED_EMPNBR")
'        fTablHREMP("HE_COMPNO") = "001"
'        fTablHREMP("HE_TYPE_TABL") = "ADRE"
'        fTablHREMP("HE_TYPE") = clpCode(2).Text
'        fTablHREMP("HE_FDATE") = dlpFrom.Text
'        fTablHREMP("HE_TDATE") = dlpTo.Text
'        fTablHREMP("HE_ENTITLE") = dblEntitleUpd
'        fTablHREMP("HE_COE") = True
'        fTablHREMP("HE_DHRS") = SnapAddEntitle("ED_DHRS")
'        fTablHREMP("HE_LDATE") = Now
'        fTablHREMP("HE_LTIME") = Time$
'        fTablHREMP("HE_LUSER") = glbUserID
'        fTablHREMP.Update
'        '    xID = fTablHREMP("HE_ID")
'        'Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpFrom.Text, dblEntitleUpd, "A", "Mass added the Hourly Entitlement")
'        Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpAsOf.Text, dblEntitleUpd, "A", "Mass added the Hourly Entitlement")
'
'    End If
'Else
'    SQLQ$ = "DELETE FROM HRENTHRS "
'    SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
'    SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
'    SQLQ = SQLQ & " AND HE_TDATE = " & Date_SQL(dlpTo.Text)
'
'    gdbAdoIhr001.Execute SQLQ
'
'    fTablHREMP.AddNew
'
'    fTablHREMP("HE_EMPNBR") = SnapAddEntitle("ED_EMPNBR")
'    fTablHREMP("HE_COMPNO") = "001"
'    fTablHREMP("HE_TYPE_TABL") = "ADRE"
'    fTablHREMP("HE_TYPE") = clpCode(2).Text
'    fTablHREMP("HE_FDATE") = dlpFrom.Text
'    fTablHREMP("HE_TDATE") = dlpTo.Text
'    fTablHREMP("HE_ENTITLE") = dblEntitleUpd
'    fTablHREMP("HE_COE") = True
'    fTablHREMP("HE_DHRS") = SnapAddEntitle("ED_DHRS")
'    fTablHREMP("HE_LDATE") = Now
'    fTablHREMP("HE_LTIME") = Time$
'    fTablHREMP("HE_LUSER") = glbUserID
'    fTablHREMP.Update
'    '    xID = fTablHREMP("HE_ID")
'    If NumRec > 0 Then  'if accumulate and found duplicate record
'        'Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpFrom.Text, dblEntitleUpd - oldEntitleUpd, "U", "Mass modified the Hourly Entitlement")
'        Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpAsOf.Text, dblEntitleUpd - oldEntitleUpd, "U", "Mass modified the Hourly Entitlement")
'    Else
'        'Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpFrom.Text, dblEntitleUpd, "A", "Mass added the Hourly Entitlement")
'        Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpAsOf.Text, dblEntitleUpd, "A", "Mass added the Hourly Entitlement")
'    End If
'
'End If
'    DoEvents
'    xKey = SnapAddEntitle("ED_EMPNBR")
'    xKey = xKey & "|" & Format(dlpFrom.Text, "dd-mmm-yyyy")
'    xKey = xKey & "|" & Format(dlpTo.Text, "dd-mmm-yyyy")
'    xKey = xKey & "|" & clpCode(2).Text
'    xKey = xKey & "|" & dblEntitleUpd
'    xKey = xKey & "|" & Format(Date, "dd-mmm-yyyy") 'Transaction Date
'    Call Entitlements_Master_Integration(xKey, 0)
'    DoEvents
'lblNext2Rec:
'    SnapAddEntitle.MoveNext
'Wend
'
'modInsSelection = True
'
'MDIMain.panHelp(0).FloodPercent = 100
'MDIMain.panHelp(0).FloodType = 0
'
'CommitTrans
'
'fTablHREMP.Close
'
'SnapAddEntitle.Close
'
'Screen.MousePointer = DEFAULT
'
'Exit Function
'
'modInsSelection_Err:
'
'If Err = 13 Or Err = 94 Or Err = 3018 Then
'    Err = 0
'    Resume Next
'   'MsgBox "Conflicting Dates"
'    Screen.MousePointer = DEFAULT
'    Exit Function
'End If
'
'Screen.MousePointer = DEFAULT
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "InsertEntitle", "HR_EMP", "edit/Add")
'Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
'    'Rollback
'    Resume Next
'Else
'    Unload Me
'End If
'
'
'End Function
'
'Private Function modUpdateSelection()
'Dim HEID&
'Dim strDivision$
'Dim strJob$, dblServiceYears#
'Dim spt As Variant, varStartDate As Variant, lngRecs&
'Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
'Dim dblFTEHours#
'Dim dblEntitleUpd#, DtTm As Variant
'Dim Msg$, Title$, DgDef As Variant
'Dim Response%, pct%
'Dim prec%
'Dim rsHE As New ADODB.Recordset
'Dim oldEntitleUpd
'Dim xKey
'On Error GoTo modUpdateSelection_Err
'modUpdateSelection = False
'
'If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
'
'Screen.MousePointer = DEFAULT
'If snapEntitle.BOF And snapEntitle.EOF Then
'    MsgBox "Employees for this selection do not exist!"
'    Exit Function
'Else
'    lngRecs& = snapEntitle.RecordCount
''    Msg$ = lngRecs& & " Records to process" & Chr(10) & "Would You Like To Proceed?"
''    Title$ = "Update Entitlements"
''    DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
''    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
''    If Response% = IDNO Then    ' Evaluate response
''        Exit Function
'End If
'Screen.MousePointer = HOURGLASS
''End If
'MDIMain.panHelp(0).FloodType = 1
'MDIMain.panHelp(0).FloodPercent = 5
'For x% = 0 To 19
'    If IsNumeric(medGTServ(x%)) Then
'        If glbFrench Then
'            If medGTServ(x%) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
'        Else
'            If Val(medGTServ(x%)) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
'        End If
'    End If
'    If Len(medLTServ(x%)) > 0 And Len(medGTServ(x%)) = 0 Then medGTServ(x%) = 9999999
'Next
'BeginTrans
'
'While Not snapEntitle.EOF
'    prec% = prec% + 1
'    pct% = Int(100 * (prec% / lngRecs&))
'    MDIMain.panHelp(0).FloodPercent = pct%
'
'    HEID& = snapEntitle("HE_ID")
'    oldEntitleUpd = snapEntitle("HE_ENTITLE")
'    If Accum = True Then
'        dblEntitleUpd = snapEntitle("HE_ENTITLE")
'    Else
'        dblEntitleUpd = 0
'    End If
'    spt = snapEntitle("ED_PT")
'    strDivision$ = snapEntitle("ED_DIV")
'
'    If IsNull(snapEntitle(fglbWDate$)) Then
'            GoTo lblNextRec
'    End If
'
'    varStartDate = snapEntitle(fglbWDate$)
'
'    If Not IsNumeric(snapEntitle("JH_DHRS")) Then
'        dblDHours# = 0
'    Else
'        dblDHours# = snapEntitle("JH_DHRS")
'    End If
'    If Not IsNumeric(snapEntitle("JH_FTENUM")) Then
'        dblFTEHours# = 0
'    Else
'        dblFTEHours# = snapEntitle("JH_FTENUM")
'    End If
'
'    'dblServiceYears# = (DateDiff("d", varStartDate, Now) / 365) * 12
'    dblServiceYears# = MonthDiff(CVDate(varStartDate), Date)
'
'    intWhereFit& = -1   ' first record can be just less than
'    For x% = 0 To 19
'        If medLTServ(x%) = "" And Not medGTServ(x%) = "" Then Exit Function
'
'        If IsNumeric(medLTServ(x%)) And medGTServ(x%) = "" Then
'            If dblServiceYears# >= CDbl(medLTServ(x%)) Then
'                intWhereFit& = x%
'                Exit For
'            End If
'        End If
'
'        If IsNumeric(medLTServ(x%)) And IsNumeric(medGTServ(x%)) Then
'            If dblServiceYears# >= CDbl(medLTServ(x%)) And dblServiceYears# <= CDbl(medGTServ(x%)) Then
'                intWhereFit& = x%
'                Exit For
'            End If
'        End If
'
'    Next x%
'
'    If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges
'
'    dblNewEntitle# = medEntitle(intWhereFit&)
'    If optD(intWhereFit&) = True Then           ' Entitlements entered in days
'        dblNewEntitle# = dblNewEntitle# * dblDHours#
'    End If
'    If optH(intWhereFit&) = True Then           ' Entitlements entered in Hours
'        dblNewEntitle# = dblNewEntitle#
'    End If
'    If optF(intWhereFit&) = True Then           ' Entitlements entered in FTE
'        ' (Entitlement * Hrs/Day) * FTE Factor
'        dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
'    End If
'    If Accum = True Then
'        dblEntitleUpd = dblEntitleUpd + dblNewEntitle
'    Else
'        dblEntitleUpd = dblNewEntitle
'    End If
'
'    DtTm = Now
'
'
'    rsHE.Open "SELECT * FROM HRENTHRS WHERE HE_ID= " & HEID&, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    rsHE("HE_ENTITLE") = dblEntitleUpd
'    rsHE("HE_LDATE") = Now
'    rsHE("HE_LTIME") = Time$
'    rsHE("HE_LUSER") = glbUserID
'    rsHE.Update
'    rsHE.Close
'
'    If Accum = True Then
'        Call Append_Accrual(snapEntitle("ED_EMPNBR"), clpCode(2).Text, Date, dblEntitleUpd - oldEntitleUpd, "U", "Mass modified the Hourly Entitlement")
'    Else
'        Call Append_Accrual(snapEntitle("ED_EMPNBR"), clpCode(2).Text, dlpFrom.Text, dblEntitleUpd - oldEntitleUpd, "U", "Mass modified the Hourly Entitlement")
'    End If
'    DoEvents
'    xKey = snapEntitle("HE_EMPNBR")
'    xKey = xKey & "|" & Format(dlpFrom.Text, "dd-mmm-yyyy")
'    xKey = xKey & "|" & Format(dlpTo.Text, "dd-mmm-yyyy")
'    xKey = xKey & "|" & clpCode(2).Text
'    xKey = xKey & "|" & dblEntitleUpd
'    xKey = xKey & "|" & Format(Date, "dd-mmm-yyyy") 'Transaction Date
'    Call Entitlements_Master_Integration(xKey, HEID&)
'
'    DoEvents
'lblNextRec:
'    snapEntitle.MoveNext
'
'Wend
'modUpdateSelection = True
'MDIMain.panHelp(0).FloodType = 0
'CommitTrans
'
''fTablHREMP.Close
'
'snapEntitle.Close
'
'Screen.MousePointer = DEFAULT
'
'Exit Function
'
'modUpdateSelection_Err:
'
'If Err = 13 Or Err = 94 Or Err = 3018 Then
'    Err = 0
'    Resume Next
'   'MsgBox "Conflicting Dates"
'    Screen.MousePointer = DEFAULT
'    Exit Function
'End If
'
'Screen.MousePointer = DEFAULT
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateEntitle", "HR_EMP", "edit/Add")
'Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
'    'Rollback
'    Resume Next
'Else
'    Unload Me
'End If
'End Function

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

If Not glbCBrant Then
    fglbWSQLQ = fglbWSQLQ & " AND HE_FDATE >= " & Date_SQL(dlpFrom.Text)
    fglbWSQLQ = fglbWSQLQ & " AND HE_TDATE <= " & Date_SQL(dlpTo.Text)
End If

If xType = "" Then Exit Function

'If xType = "O" Then
'    xDiv = ODIV
'    xDept = ODept
'    xORG = oOrg
'    xFDate = oFDate
'    xTDate = OTDate
'    xEMP = oEMP
'    xEmpMode = oEmpMode
'    xHETYPE = oHETYPE
'    xLoc = OLoc
'    xSection = OSection
'Else
If xType = "D" Then
    xDiv = data1.Recordset("EH_DIV")
    xDept = data1.Recordset("EH_DEPT")
    xORG = data1.Recordset("EH_ORG")
    xFDate = data1.Recordset("EH_FDATE")
    xTDate = data1.Recordset("EH_TDATE")
    xEMP = data1.Recordset("EH_EMP")
    xEmpMode = data1.Recordset("EH_PT")
    xHETYPE = data1.Recordset("EH_HETYPE")
    xLoc = data1.Recordset("EH_LOC")
    xSection = data1.Recordset("EH_SECTION")
End If
'Else
'    xDiv = clpDiv.Text
'    xDept = clpDept.Text
'    xORG = clpCode(0).Text
'    xFDate = dlpFrom.Text
'    xTDate = dlpTo.Text
'    xEMP = clpCode(1).Text
'    xEmpMode = clpPT.Text
'    xHETYPE = clpCode(2).Text
'    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
'        xLoc = clpProv.Text
'    Else
'        xLoc = clpCode(4).Text
'    End If
'    xSection = clpCode(3).Text
'End If
'
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
    
TF = True

UpdateState = OPENING

Call set_Buttons(UpdateState)

If Not UpdateRight Then TF = False

End Sub

Public Property Get ChangeAction() As UpdateStateEnum
    ChangeAction = OPENING
End Property

Public Property Let ChangeAction(vData As UpdateStateEnum)

End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = nothingrelate
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Hrly_Entitlements
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = False
End Property

Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = False
End Property

Public Sub cmdZeroOutHr_Click()
    cmdZeroOut.Visible = True
    cmdZeroOut.Enabled = True
    cmdZeroOutAll.Visible = True
    cmdZeroOutAll.Enabled = True
    cmdRollover.Enabled = False
    cmdRollover.Visible = False
    cmdRolloverAll.Enabled = False
    cmdRolloverAll.Visible = False
    fFormCalled = "ZeroOut"
    frmZero.Visible = True
    lblMaxRollover.Visible = False
    medMaxRollover.Visible = False
    optDH(0).Visible = False
    optDH(1).Visible = False
End Sub

Public Sub cmdRolloverHr_Click()
    cmdZeroOut.Enabled = False
    cmdZeroOut.Visible = False
    cmdZeroOutAll.Enabled = False
    cmdZeroOutAll.Visible = False
    cmdRollover.Visible = True
    cmdRollover.Enabled = True
    cmdRolloverAll.Visible = True
    cmdRolloverAll.Enabled = True
    fFormCalled = "Rollover"
    frmZero.Visible = False
    lblMaxRollover.Visible = True
    medMaxRollover.Visible = True
    medMaxRollover.Text = ""
    optDH(1).Visible = True
    optDH(0).Visible = True
    optDH(0).Value = True
End Sub

Private Function getHourlyEntMax(xEmpNo, xType)
    Dim rsEnt As New ADODB.Recordset
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ As String
    Dim xDiv, xDept, xLoc, xEMP, xFT, xORG, xSection
    Dim retval
    retval = 0
    
    SQLQ = "SELECT * FROM HR_HOURLYENT WHERE EH_HETYPE = '" & xType & "' "
    SQLQ = SQLQ & "ORDER BY EH_TDATE DESC "
    rsEnt.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEnt.EOF Then
        If Not IsNull(rsEnt("EH_MAX")) Then
            retval = rsEnt("EH_MAX")
        End If
        xDiv = "": xDept = "": xLoc = "": xEMP = "": xFT = "": xORG = "": xSection = ""
        If Not IsNull(rsEnt("EH_DIV")) Then xDiv = Trim(rsEnt("EH_DIV"))
        If Not IsNull(rsEnt("EH_DEPT")) Then xDept = Trim(rsEnt("EH_DEPT"))
        If Not IsNull(rsEnt("EH_LOC")) Then xLoc = Trim(rsEnt("EH_LOC"))
        If Not IsNull(rsEnt("EH_EMP")) Then xEMP = Trim(rsEnt("EH_EMP"))
        If Not IsNull(rsEnt("EH_PT")) Then xFT = Trim(rsEnt("EH_PT"))
        If Not IsNull(rsEnt("EH_ORG")) Then xORG = Trim(rsEnt("EH_ORG"))
        If Not IsNull(rsEnt("EH_SECTION")) Then xSection = Trim(rsEnt("EH_SECTION"))
        
        SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmp.EOF Then
            SQLQ = "SELECT * FROM HR_HOURLYENT WHERE EH_HETYPE = '" & xType & "' "
            If Len(xDiv) > 0 Then
                If Not IsNull(rsEmp("ED_DIV")) Then
                    SQLQ = SQLQ & "AND EH_DIV = '" & rsEmp("ED_DIV") & "' "
                End If
            End If
            If Len(xDept) > 0 Then
                If Not IsNull(rsEmp("ED_DEPTNO")) Then
                    SQLQ = SQLQ & "AND EH_DEPT = '" & rsEmp("ED_DEPTNO") & "' "
                End If
            End If
            If Len(xLoc) > 0 Then
                If Not IsNull(rsEmp("ED_LOC")) Then
                    SQLQ = SQLQ & "AND EH_LOC = '" & rsEmp("ED_LOC") & "' "
                End If
            End If
            If Len(xEMP) > 0 Then
                If Not IsNull(rsEmp("ED_EMP")) Then
                    SQLQ = SQLQ & "AND EH_EMP = '" & rsEmp("ED_EMP") & "' "
                End If
            End If
            If Len(xFT) > 0 Then
                If Not IsNull(rsEmp("ED_PT")) Then
                    SQLQ = SQLQ & "AND EH_PT = '" & rsEmp("ED_PT") & "' "
                End If
            End If
            If Len(xORG) > 0 Then
                If Not IsNull(rsEmp("ED_ORG")) Then
                    SQLQ = SQLQ & "AND EH_ORG = '" & rsEmp("ED_ORG") & "' "
                End If
            End If
            If Len(xSection) > 0 Then
                If Not IsNull(rsEmp("ED_SECTION")) Then
                    SQLQ = SQLQ & "AND EH_SECTION = '" & rsEmp("ED_SECTION") & "' "
                End If
            End If
            
            SQLQ = SQLQ & "ORDER BY EH_TDATE DESC "
            
            If rsEnt.State <> 0 Then rsEnt.Close
            rsEnt.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsEnt.EOF Then
                If Not IsNull(rsEnt("EH_MAX")) Then
                    If rsEnt("EH_MAX") > 0 Then
                        retval = rsEnt("EH_MAX")
                    End If
                End If
            End If
        End If
    End If

    getHourlyEntMax = retval
End Function
