VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSDashboardRule 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Dashboard Rule"
   ClientHeight    =   5760
   ClientLeft      =   90
   ClientTop       =   1005
   ClientWidth     =   7800
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
   ScaleHeight     =   5760
   ScaleWidth      =   7800
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkTopDash 
      Caption         =   "Show on the Top Dashboard"
      DataField       =   "DB_TOP_DASH"
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
      Height          =   375
      Left            =   4200
      TabIndex        =   33
      Tag             =   "40-Select to show on the Top Dashboard"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.TextBox txtItemDesc 
      Appearance      =   0  'Flat
      DataField       =   "DB_ITEM_DESC"
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
      Left            =   6120
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtItemPriority 
      Appearance      =   0  'Flat
      DataField       =   "DB_ITEM_PRIORITY"
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
      Height          =   285
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   2
      Tag             =   "10-Enter the Sequence Number (for display purposes)"
      Top             =   4275
      Width           =   495
   End
   Begin VB.TextBox txtItemCode 
      Appearance      =   0  'Flat
      DataField       =   "DB_ITEM_CODE"
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
      Left            =   6120
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3735
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox comItem 
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
      ItemData        =   "frmSDashboardRule.frx":0000
      Left            =   1320
      List            =   "frmSDashboardRule.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "10-Dashboard Item to display"
      Top             =   3720
      Width           =   4695
   End
   Begin VB.TextBox txtSubCategory 
      Appearance      =   0  'Flat
      DataField       =   "DB_SUB_CATEGORY"
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
      Left            =   5280
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4635
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox comCategory 
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
      ItemData        =   "frmSDashboardRule.frx":0004
      Left            =   1320
      List            =   "frmSDashboardRule.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "10-Dashboard Category"
      Top             =   3180
      Width           =   3135
   End
   Begin VB.TextBox txtCategory 
      Appearance      =   0  'Flat
      DataField       =   "DB_CATEGORY"
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
      Left            =   4560
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3195
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DB_LUSER"
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DB_LDATE"
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
      Index           =   0
      Left            =   840
      MaxLength       =   12
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DB_LTIME"
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
      Left            =   1590
      MaxLength       =   8
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   645
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   1320
      Top             =   5400
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
   Begin VB.VScrollBar scrControl 
      Height          =   5325
      LargeChange     =   315
      Left            =   6960
      Max             =   100
      SmallChange     =   315
      TabIndex        =   23
      Top             =   2880
      Visible         =   0   'False
      Width           =   300
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   360
      Top             =   5280
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   8730
      TabIndex        =   5
      Tag             =   "00-Administered By"
      Top             =   9465
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   8730
      TabIndex        =   9
      Tag             =   "00-Union"
      Top             =   10065
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   2850
      TabIndex        =   4
      Tag             =   "00-Department"
      Top             =   10365
      Visible         =   0   'False
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   2850
      TabIndex        =   3
      Tag             =   "00-Division"
      Top             =   10065
      Visible         =   0   'False
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   2850
      TabIndex        =   7
      Tag             =   "00-Employment Status"
      Top             =   9465
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   2850
      TabIndex        =   8
      Tag             =   "EDPT-Category"
      Top             =   9765
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   8730
      TabIndex        =   10
      Tag             =   "00-Location"
      Top             =   10365
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   8730
      TabIndex        =   6
      Tag             =   "00-Section"
      Top             =   9765
      Visible         =   0   'False
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmSDashboardRule.frx":0008
      Height          =   2535
      Left            =   240
      OleObjectBlob   =   "frmSDashboardRule.frx":001C
      TabIndex        =   24
      Top             =   240
      Width           =   7335
   End
   Begin VB.Label Label30 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sequence #"
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
      Left            =   240
      TabIndex        =   31
      Top             =   4320
      Width           =   885
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
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
      Left            =   240
      TabIndex        =   27
      Top             =   3780
      Width           =   300
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Category"
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
      Left            =   4200
      TabIndex        =   26
      Top             =   4680
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label27 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Left            =   240
      TabIndex        =   25
      Top             =   3240
      Width           =   630
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      DataField       =   "DB_COMPNO"
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
      Left            =   4320
      TabIndex        =   13
      Top             =   5280
      Visible         =   0   'False
      Width           =   405
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
      Left            =   7440
      TabIndex        =   22
      Top             =   10410
      Visible         =   0   'False
      Width           =   615
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
      Left            =   7440
      TabIndex        =   21
      Top             =   9810
      Visible         =   0   'False
      Width           =   1260
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
      Left            =   1560
      TabIndex        =   20
      Top             =   9810
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   7440
      TabIndex        =   19
      Top             =   9510
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label lblEmpStatys 
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
      Left            =   1560
      TabIndex        =   18
      Top             =   9510
      Visible         =   0   'False
      Width           =   1350
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
      Left            =   7440
      TabIndex        =   17
      Top             =   10110
      Visible         =   0   'False
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
      Left            =   1560
      TabIndex        =   16
      Top             =   10410
      Visible         =   0   'False
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
      Left            =   1560
      TabIndex        =   15
      Top             =   10110
      Visible         =   0   'False
      Width           =   555
   End
End
Attribute VB_Name = "frmSDashboardRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean
Dim fglbSort As Boolean
Dim fUPMode As Integer
Dim rsDATA As New ADODB.Recordset
Dim UpdateState As UpdateStateEnum

Private Function chkDashboardRule()
Dim Msg As String
Dim flgExists As Boolean

chkDashboardRule = False

'Check if valid numeric values entered for Priorities
'Employee Functions - Entitlements
If Len(Trim(txtItemPriority.Text)) > 0 Then
    If Not IsNumeric(txtItemPriority.Text) Then
        MsgBox "Invalid Sequence number."
        txtItemPriority.SetFocus
        Exit Function
    End If
End If

'Check if the dashboard item already exists in the rule
flgExists = False
If fglbNew Then
    'Check by Category and Item
    flgExists = Rule_Exists
    If flgExists Then
        MsgBox "Dashboard Rule for this Category/Item already exists."
        comCategory.SetFocus
        Exit Function
    End If
Else
    'Check by Category and Item but exclude the current record, DB_ID
    flgExists = Rule_Exists(Data1.Recordset("DB_ID"))
    If flgExists Then
        MsgBox "Dashboard Rule for this Category/Item already exists."
        comCategory.SetFocus
        Exit Function
    End If
End If

chkDashboardRule = True

End Function

Sub cmdCancel_Click()

On Error GoTo Can_Err

Data1.Recordset.CancelBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    Data1.Refresh
End If

fglbNew = False

Call ST_UPD_MODE(False)

comCategory.ListIndex = -1
comItem.ListIndex = -1

'rsDATA.CancelUpdate
'Call Display_Value

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HR_DASHBOARD_RULE", "Cancel")
Call RollBack '09June99 js

End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    Call ST_UPD_MODE(False)
Else
    Call ST_UPD_MODE(True)
End If

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_DASHBOARD_RULE", "Modify")
Call RollBack '09June99 js

End Sub

Sub cmdDelete_Click()
Dim Msg As String, a%

On Error GoTo DelErr

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then
    Exit Sub
End If

'gdbAdoIhr001.BeginTrans
rsDATA.Delete
'gdbAdoIhr001.CommitTrans
Data1.Refresh

'If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
'End If

fglbNew = False

'Data1.Recordset.Delete

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

Data1.Refresh

Call SET_UP_MODE

Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_DASHBOARD_RULE", "Delete")
Call RollBack '09June99 js

End Sub

Sub cmdNew_Click()

On Error GoTo AddN_Err

'Data1.Recordset.AddNew
Call Set_Control("B", Me)

rsDATA.AddNew

lblCNum.Caption = "001"
comCategory.ListIndex = -1
comItem.ListIndex = -1

fglbNew = True

Call SET_UP_MODE

'Call ST_UPD_MODE(True)
'clpCode(3).SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_DASHBOARD_RULE", "Add")
Call RollBack '09June99 js

End Sub

Sub cmdOK_Click()
Dim x%
Dim bmk As Variant

On Error GoTo cmdOK_Err

If Not chkDashboardRule() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)

If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    bmk = 0
Else
    bmk = Data1.Recordset.Bookmark
End If

Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans
Data1.Refresh

If Not bmk = 0 Then
    Data1.Recordset.Bookmark = bmk
End If

fglbNew = False

Call Display_Value

'Call EERetrieve

''Turnover Ratio prioritised then disable Current Year New Hire and Termination
'If Len(Trim(txtSupALTurnoverRatio.Text)) > 0 Then
'    If IsNumeric(txtSupALTurnoverRatio.Text) Then
'        'txtSupALCurrNewHires.Text = ""
'        'txtSupALCurrTerms.Text = ""
'        txtSupALCurrNewHires.Enabled = False
'        txtSupALCurrTerms.Enabled = False
'    Else
'        txtSupALCurrNewHires.Enabled = True
'        txtSupALCurrTerms.Enabled = True
'    End If
'Else
'    txtSupALCurrNewHires.Enabled = True
'    txtSupALCurrTerms.Enabled = True
'
'    If IsNumeric(txtSupALCurrNewHires.Text) Or IsNumeric(txtSupALCurrTerms.Text) Then
'        txtSupALTurnoverRatio.Enabled = False
'    Else
'        txtSupALTurnoverRatio.Enabled = True
'    End If
'End If

Screen.MousePointer = DEFAULT

Exit Sub

cmdOK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_DASHBOARD_RULE", "Update")
Call RollBack '09June99 js

End Sub

Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError

    Screen.MousePointer = HOURGLASS

    SQLQ = "SELECT * FROM HR_DASHBOARD_RULE"
    SQLQ = SQLQ & " WHERE DB_ID = " & Data1.Recordset!DB_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

Data1.RecordSource = SQLQ
Data1.Refresh

Call Display_Value

EERetrieve = True

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HR_DASHBOARD_RULE", "SELECT")
Call RollBack

Exit Function

End Function

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = "Dashboard Rule"
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

RHeading = "Dashboard Rule"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub chkTopDash_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub comCategory_Change()
    'Populate Items based on Category selected
    Call Populate_Items
End Sub

Private Sub comCategory_Click()
    'Populate Items based on Category selected
    Call Populate_Items
    
    txtCategory.Text = comCategory.Text
End Sub

Private Sub comCategory_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub comItem_Click()
    'If Not fglbSort Then
        'Get the Item Code and update the txtItemCode.text
        Dim xDashItems
        xDashItems = Split(Get_ItemCode(comItem.Text), "|")
        If UBound(xDashItems) > 0 Then
            'txtItemCode.Text = Get_ItemCode(comItem.Text)
            txtItemCode.Text = xDashItems(0)
            txtItemDesc.Text = comItem.Text
            chkTopDash.Value = IIf(xDashItems(1), vbChecked, vbUnchecked)
        End If
    'End If
End Sub

Private Sub comItem_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtCategory_Change()
    Dim x As Integer
    
    If txtCategory.Text <> "" Then
        For x = 0 To comCategory.ListCount - 1
            If UCase(txtCategory.Text) = UCase(comCategory.List(x)) Then
                comCategory.ListIndex = x
                
                If Len(txtItemDesc.Text) > 0 Then
                    Call txtItemDesc_Change
                End If
                Exit For
            End If
        Next
    End If
End Sub

Private Sub txtItemCode_Change()
'    'Get corresponding Item Description and select the item in the Item combo
'    txtItemDesc.Text = Get_ItemDesc(txtItemCode.Text)
End Sub

Private Sub txtItemDesc_Change()
'    Dim X As Integer
'
'    For X = 0 To comItem.ListCount - 1
'        If UCase(txtItemDesc.Text) = UCase(comItem.List(X)) Then
'            comItem.ListIndex = X
'            Exit For
'        End If
'    Next
End Sub

Private Sub txtItemPriority_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()

Call SET_UP_MODE

Me.cmdModify_Click

End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim SQLQ

'Me.Show

glbOnTop = "FRMSDASHBOARDRULE"

Screen.MousePointer = HOURGLASS

'If HR_DASHBOARD_RULE is empty then create a default record
'Call Add_Default_Dashboard_Rule

Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "SELECT * FROM HR_DASHBOARD_RULE ORDER BY DB_USERID"
Data1.Refresh

Screen.MousePointer = DEFAULT

'Call setRptCaption(Me)
'Call setCaption(lblPT)
'Call setCaption(lblDiv)
'Call setCaption(lblDept)
'Call setCaption(lblLocation)
'Call setCaption(lblRegion)
'Call setCaption(lblAdmin)
'Call setCaption(lblSection)
'Call setCaption(lblUnion)
'Call setCaption(lblPT)

vbxTrueGrid.Columns(0).Visible = False
vbxTrueGrid.Columns(2).Visible = False
vbxTrueGrid.Columns(5).Visible = False
vbxTrueGrid.Columns(6).Visible = False


'Combo box items
Call Populate_Category


'Call EERetrieve


Call ST_UPD_MODE(False)

'vbxTrueGrid.Columns(0).Caption = lStr("Section")
'lblTitle(0).Caption = lStr(lblTitle(0).Caption)
'lblTitle(1).Caption = lStr(lblTitle(1).Caption)
'lblTitle(2).Caption = lStr(lblTitle(2).Caption)

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT                           '

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

'Private Sub Form_Resize()
'If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
'    'Vertical scroll bar
'    If Me.Height >= 9600 Then
'        scrControl.Value = 0
'        scrFrame.Top = 120
'        scrControl.Visible = False
'    Else
'        scrControl.Left = Me.Width - 600    '400
'        scrControl.Visible = True
'        scrControl.Height = Me.Height - 800
'        If Me.Height < 8000 Then
'            scrControl.Max = 4700
'        Else
'            scrControl.Max = 3000
'        End If
'        'scrControl.Left = Me.Width - scrControl.Width - 120
'        If Me.Height - scrControl.Top - 780 > 0 Then
'            scrControl.Height = Me.Height - scrControl.Top - 780
'        End If
'    End If
'End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
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

fUPMode = TF

'clpDiv.Enabled = TF
'clpDept.Enabled = TF
'clpPT.Enabled = TF
'clpCode(0).Enabled = TF
'clpCode(1).Enabled = TF
'clpCode(2).Enabled = TF
'clpCode(3).Enabled = TF
'clpCode(4).Enabled = TF
comCategory.Enabled = TF
comItem.Enabled = TF
txtItemPriority.Enabled = TF
End Sub

'Private Sub scrControl_Change()
'    scrFrame.Top = 120 - scrControl.Value
'End Sub

'Private Sub txtSupALCurrNewHires_LostFocus()
'    'If Turnover Ratio prioritize then clear it
'    If Len(Trim(txtSupALCurrNewHires.Text)) > 0 Then
'        txtSupALTurnoverRatio.Text = ""
'        txtSupALTurnoverRatio.Enabled = False
'    ElseIf IsNumeric(txtSupALCurrNewHires.Text) Or IsNumeric(txtSupALCurrTerms.Text) Then
'        txtSupALTurnoverRatio.Enabled = False
'    Else
'        txtSupALTurnoverRatio.Enabled = True
'    End If
'End Sub
'
'Private Sub txtSupALCurrTerms_LostFocus()
'    'If Turnover Ratio prioritize then clear it
'    If Len(Trim(txtSupALCurrTerms.Text)) > 0 Then
'        txtSupALTurnoverRatio.Text = ""
'        txtSupALTurnoverRatio.Enabled = False
'    ElseIf IsNumeric(txtSupALCurrNewHires.Text) Or IsNumeric(txtSupALCurrTerms.Text) Then
'        txtSupALTurnoverRatio.Enabled = False
'    Else
'        txtSupALTurnoverRatio.Enabled = True
'    End If
'End Sub
'
'Private Sub txtSupALTurnoverRatio_LostFocus()
'    'Turnover Ratio prioritised then disable Current Year New Hire and Termination
'    If Len(Trim(txtSupALTurnoverRatio.Text)) > 0 Then
'        If IsNumeric(txtSupALTurnoverRatio.Text) Then
'            txtSupALCurrNewHires.Text = ""
'            txtSupALCurrTerms.Text = ""
'            txtSupALCurrNewHires.Enabled = False
'            txtSupALCurrTerms.Enabled = False
'        Else
'            txtSupALCurrNewHires.Enabled = True
'            txtSupALCurrTerms.Enabled = True
'        End If
'    Else
'        txtSupALCurrNewHires.Enabled = True
'        txtSupALCurrTerms.Enabled = True
'    End If
'
'End Sub

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

Private Sub Display_Value()
    Dim SQLQ
    Dim x As Integer
        
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
        Call SET_UP_MODE
        
        comCategory.ListIndex = -1
        comItem.ListIndex = -1
        
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HR_DASHBOARD_RULE WHERE DB_ID = " & Data1.Recordset!DB_ID 'DB_CATEGORY = '" & Data1.Recordset!DB_CATEGORY & "' AND DB_ITEM_CODE = '" & Data1.Recordset!DB_ITEM_CODE & "'"
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    
    'Combo box items
    'Call Populate_Category
    
    Call Set_Control("R", Me, rsDATA)
    
    Call SET_UP_MODE
    
    
    'Select the corresponding item in the Item combobox based on the Item Desc. in the table
    For x = 0 To comItem.ListCount - 1
        If UCase(txtItemDesc.Text) = UCase(comItem.List(x)) Then
            comItem.ListIndex = x
            Exit For
        End If
    Next
    

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
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_DashboardRule
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
ElseIf Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If

Call ST_UPD_MODE(TF)

Call set_Buttons(UpdateState)

If Not UpdateRight Then TF = False

End Sub

Private Sub Add_Default_Dashboard_Rule()
Dim rsDashboard As New ADODB.Recordset
Dim SQLQ As String

    'Add a Default Dashboard setup record if table is empty
    SQLQ = "SELECT * FROM HR_DASHBOARD_RULE"
    rsDashboard.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsDashboard.EOF Then
        'Table is empty, add default record
        rsDashboard.AddNew
        rsDashboard("DB_COMPNO") = "001"
        rsDashboard("DB_EMPLOYEE") = 0
        rsDashboard("DB_EMP_ENTITLE") = 0
        rsDashboard("DB_EMP_VACTIME_REQ") = 0
        rsDashboard("DB_SUPERVISOR") = 0
        rsDashboard("DB_SUP_VACTIME_REQ") = 0
        rsDashboard("DB_SUP_ANALYTICAL") = 0
        rsDashboard("DB_LDATE") = Date
        rsDashboard("DB_LTIME") = Time$
        rsDashboard("DB_LUSER") = glbUserID
        rsDashboard.Update
    End If
    rsDashboard.Close
    Set rsDashboard = Nothing
End Sub

Private Sub Populate_Category()
Dim rsDashCat As New ADODB.Recordset
Dim SQLQ As String
    
    'Clear the Categorise list
    comCategory.Clear
    
    'Retrieve Categories
    SQLQ = "SELECT DISTINCT(DB_CATEGORY) FROM HR_DASHBOARD_ITEMS GROUP BY DB_CATEGORY ORDER BY DB_CATEGORY"
    rsDashCat.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    Do While Not rsDashCat.EOF
        comCategory.AddItem rsDashCat("DB_CATEGORY")
        rsDashCat.MoveNext
    Loop
    rsDashCat.Close
    Set rsDashCat = Nothing
    
End Sub

Private Sub Populate_Items()
Dim rsDashItems As New ADODB.Recordset
Dim SQLQ As String
    
    'Clear the Items list
    comItem.Clear
    
    'Retrieve Items
    SQLQ = "SELECT DB_ITEMNAME FROM HR_DASHBOARD_ITEMS WHERE UPPER(DB_CATEGORY) = '" & UCase(comCategory.Text) & "' ORDER BY DB_ITEMNAME"
    rsDashItems.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    Do While Not rsDashItems.EOF
        comItem.AddItem rsDashItems("DB_ITEMNAME")
        rsDashItems.MoveNext
    Loop
    rsDashItems.Close
    Set rsDashItems = Nothing
    
End Sub

Private Function Get_ItemCode(xItemDesc) As String
Dim rsDashItems As New ADODB.Recordset
Dim SQLQ As String
        
    Get_ItemCode = ""
    
    'Retrieve Item Code
    SQLQ = "SELECT DB_ITEMCODE, DB_TOP_DASH FROM HR_DASHBOARD_ITEMS WHERE UPPER(DB_ITEMNAME) = '" & UCase(xItemDesc) & "'"
    SQLQ = SQLQ & " AND UPPER(DB_CATEGORY) = '" & UCase(comCategory.Text) & "'"
    rsDashItems.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsDashItems.EOF Then
        'txtItemCode.Text = rsDashItems("DB_ITEMCODE")
        Get_ItemCode = rsDashItems("DB_ITEMCODE") & "|" & rsDashItems("DB_TOP_DASH")
    Else
        'txtItemCode.Text = ""
        Get_ItemCode = ""
    End If
    rsDashItems.Close
    Set rsDashItems = Nothing

End Function

Private Function Get_ItemDesc(xItemCode) As String
Dim rsDashItems As New ADODB.Recordset
Dim SQLQ As String
    
    Get_ItemDesc = ""
    
    'Retrieve Item Code
    SQLQ = "SELECT DB_ITEMNAME FROM HR_DASHBOARD_ITEMS WHERE UPPER(DB_ITEMCODE) = '" & UCase(xItemCode) & "'"
    rsDashItems.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsDashItems.EOF Then
        Get_ItemDesc = rsDashItems("DB_ITEMNAME")
    Else
        Get_ItemDesc = ""
    End If
    rsDashItems.Close
    Set rsDashItems = Nothing

End Function

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
    'fglbSort = True
    
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
                   
    SQLQ = "SELECT * FROM HR_DASHBOARD_RULE "
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

    Data1.RecordSource = SQLQ
    Data1.Refresh

    'fglbSort = False
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Display_Value
End Sub

Private Function Rule_Exists(Optional xID) As Boolean
    Dim rsDashboard As New ADODB.Recordset
    Dim SQLQ As String
    
    Rule_Exists = False
    
    SQLQ = "SELECT * FROM HR_DASHBOARD_RULE WHERE UPPER(DB_CATEGORY) = '" & UCase(comCategory.Text) & "'"
    SQLQ = SQLQ & " AND UPPER(DB_ITEM_CODE) = '" & UCase(txtItemCode.Text) & "'"
    If Not IsMissing(xID) Then
        SQLQ = SQLQ & " AND DB_ID <> " & xID
    End If
    rsDashboard.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsDashboard.EOF Then
        Rule_Exists = True
    Else
        Rule_Exists = False
    End If
    rsDashboard.Close
    Set rsDashboard = Nothing
End Function
