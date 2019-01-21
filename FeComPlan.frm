VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEComPlan 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Compensation Plan"
   ClientHeight    =   8595
   ClientLeft      =   195
   ClientTop       =   1200
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
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Height          =   1665
      Left            =   240
      OleObjectBlob   =   "FeComPlan.frx":0000
      TabIndex        =   0
      Tag             =   "Listing of Salary Records"
      Top             =   630
      Width           =   10575
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "CP_TDATE"
      Height          =   285
      Index           =   1
      Left            =   1650
      TabIndex        =   2
      Tag             =   "41-To Date of Compensation Plan"
      Top             =   2790
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "CP_FDATE"
      Height          =   285
      Index           =   0
      Left            =   1650
      TabIndex        =   1
      Tag             =   "41-From Date of Compensation Plan"
      Top             =   2460
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin VB.TextBox txtCommPlan 
      Appearance      =   0  'Flat
      DataField       =   "CP_COMMPLAN"
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
      Left            =   1950
      MaxLength       =   30
      TabIndex        =   7
      Tag             =   "00-Commission Plan"
      Top             =   4440
      Width           =   5055
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CP_LUSER"
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   2
      Left            =   3075
      MaxLength       =   25
      TabIndex        =   30
      TabStop         =   0   'False
      Text            =   "LUser"
      Top             =   6630
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CP_LTIME"
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   1
      Left            =   450
      MaxLength       =   25
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "LTime"
      Top             =   6630
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CP_LDATE"
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "Ldate"
      Top             =   6630
      Visible         =   0   'False
      Width           =   615
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   16
      Top             =   7935
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
      Begin MSAdodcLib.Adodc Data1 
         Height          =   375
         Left            =   7140
         Top             =   90
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Tag             =   "Close and exit this screen"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1080
         TabIndex        =   18
         Tag             =   "Edit the information on this screen"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1950
         TabIndex        =   19
         Tag             =   "Save the changes made"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2775
         TabIndex        =   20
         Tag             =   "Cancel the changes made"
         Top             =   30
         Width           =   855
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3720
         TabIndex        =   21
         Tag             =   "Add a new Record"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4560
         TabIndex        =   22
         Tag             =   "Delete the Record Selected"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5430
         TabIndex        =   23
         Tag             =   "Print a Salary History Report"
         Top             =   30
         Width           =   855
      End
      Begin VB.CommandButton cmdCalculate 
         Appearance      =   0  'Flat
         Caption         =   "&Recalculate"
         Height          =   405
         Left            =   9540
         TabIndex        =   24
         Top             =   60
         Width           =   1485
      End
   End
   Begin VB.TextBox txtComment 
      Appearance      =   0  'Flat
      DataField       =   "CP_COMMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   1980
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Tag             =   "00-Comments"
      Top             =   4770
      Width           =   5025
   End
   Begin Threed.SSPanel panEEDesc 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11880
      _Version        =   65536
      _ExtentX        =   20955
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
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1380
         TabIndex        =   11
         Top             =   120
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3030
         TabIndex        =   10
         Top             =   120
         Width           =   1740
      End
   End
   Begin MSMask.MaskEdBox medTargBonus 
      DataField       =   "CP_TARGBONUS"
      Height          =   285
      Left            =   1950
      TabIndex        =   5
      Tag             =   "21-Target Bonus"
      Top             =   3780
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medTargIncome 
      DataField       =   "CP_TARGINCOME"
      Height          =   285
      Left            =   1950
      TabIndex        =   6
      Tag             =   "20-Target Income"
      Top             =   4110
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   9465
      Top             =   8580
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
   Begin MSMask.MaskEdBox medSalary 
      DataField       =   "CP_SALARY"
      Height          =   285
      Left            =   1950
      TabIndex        =   3
      Tag             =   "20-Salary"
      Top             =   3120
      Width           =   1530
      _ExtentX        =   2699
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
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medAdjSalary 
      DataField       =   "CP_ADJSALARY"
      Height          =   285
      Left            =   1950
      TabIndex        =   4
      Tag             =   "20-Adjustment salary"
      Top             =   3450
      Width           =   1530
      _ExtentX        =   2699
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
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Label lblTitle 
      Caption         =   "Commission Plan"
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
      Index           =   7
      Left            =   450
      TabIndex        =   34
      Top             =   4440
      Width           =   1245
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   450
      TabIndex        =   33
      Top             =   2790
      Width           =   1095
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      DataField       =   "CP_COMPNO"
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
      Left            =   1320
      TabIndex        =   32
      Top             =   6510
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EEId"
      DataField       =   "CP_EMPNBR"
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
      Left            =   1350
      TabIndex        =   31
      Top             =   6750
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Target Income"
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
      Left            =   450
      TabIndex        =   27
      Top             =   4140
      Width           =   1350
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustment Salary"
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
      Left            =   450
      TabIndex        =   26
      Top             =   3450
      Width           =   1260
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
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
      Left            =   450
      TabIndex        =   25
      Top             =   3120
      Width           =   540
   End
   Begin VB.Label lblTitle 
      Caption         =   "Comments"
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
      Index           =   8
      Left            =   450
      TabIndex        =   15
      Top             =   4740
      Width           =   855
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   450
      TabIndex        =   14
      Top             =   2490
      Width           =   1245
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Target Bonus"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   450
      TabIndex        =   13
      Tag             =   "21-Target Bonus"
      Top             =   3810
      Width           =   1350
   End
End
Attribute VB_Name = "frmEComPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim flagFrmLoad
Dim fglbNew As Integer, Actn
Dim cnCP As New ADODB.Connection
Dim oFDate
Dim dynaSals As New ADODB.Recordset
Dim xSalary(2) As Double
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control


Private Function chkComPlan()
Dim x%
Dim SQLQ As String, Msg$, dd&
Dim DgDef As Variant, Title$, Response%, DCurSHDate  As Variant

chkComPlan = False

On Error GoTo chkComPlan_Err

If Len(dlpDate(0).Text) < 1 Then
    Msg$ = "From Date is required"
    dlpDate(0).SetFocus
    MsgBox Msg$
    Exit Function
Else
    If Not IsDate(dlpDate(0).Text) Then
        Msg$ = "Not a Valid From Date"
        dlpDate(0).SetFocus
        MsgBox Msg$
        Exit Function
    End If
End If

If Not IsNumeric(medTargBonus) Then
    Msg$ = "Target Bonus is invalid"
    MsgBox Msg$
    medTargBonus.SetFocus
    Exit Function
End If

chkComPlan = True

Exit Function

chkComPlan_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkComPlan", "Compensation Plan", "edit/Add")
Resume Next

End Function




Private Sub cmdCalculate_Click()
Dim xEDate
Dim rsCP As New ADODB.Recordset
On Error GoTo JS_Err

If Data1.Recordset.RecordCount < 1 Then Exit Sub

Set rsCP = Data1.Recordset.Clone
dynaSals.Requery
rsCP.MoveFirst
Do Until rsCP.EOF
    If rsCP("CP_FDATE") <= Date And rsCP("CP_TDATE") > Date Then
        Call getSalary(rsCP("CP_FDATE"), rsCP("CP_TDATE"))
        rsCP("CP_SALARY") = xSalary(0)
        rsCP("CP_ADJSALARY") = xSalary(1) - xSalary(0)
        rsCP("CP_TARGINCOME") = xSalary(1) + rsCP("CP_TARGBONUS")
        rsCP.Update
    End If
    rsCP.MoveNext
Loop
If Not glbSQL And Not glbOracle Then Pause (0.5)
Data1.Refresh
rsCP.Close
Exit Sub

JS_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Set SALARY History", "HR_SALARY_HISTORY", "SELECT")
Resume Next

End Sub

Private Sub cmdCancel_Click()

On Error GoTo Can_Err


rsDATA.CancelUpdate
Call Display_Value

Call ST_UPD_MODE(False)  ' reset screen's attributes

Me.vbxTrueGrid.SetFocus
fglbNew = False

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_SALARY_HISTORY", "Cancel")
Resume Next

End Sub

Private Sub cmdCancel_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMECOMPLAN" Then glbOnTop = ""

End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdDelete_Click()
Dim a As Integer, Msg As String
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, rc%, DtTm As Variant

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

DtTm = Now

gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh
If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

Call ST_UPD_MODE(False)

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_SALARY_HISTORY", "Delete")
Call RollBack '28July99 js
End Sub

Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()
Dim SQLQ As String, x%
Dim Response%, Msg$, Title$, DgDef As Double

On Error GoTo Mod_Err

Call ST_UPD_MODE(True)
Actn = "M"
dlpDate(0).SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_SALARY_HISTORY", "Modify")
Call RollBack '28July99 js

End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()
Dim SQLQ As String, Msg$
Dim x%
Dim orgMarketLine, orgSalCD
On Error GoTo AddN_Err


'Data1.Recordset.AddNew

''' Sam add July 2002 * Remove Binding Control
Call Set_Control("B", Me)
rsDATA.AddNew


Actn = "A"
fglbNew = True

lblCNum.Caption = "001"
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID

medSalary = 0

Call ST_UPD_MODE(True)
dlpDate(0).SetFocus



Exit Sub

AddN_Err:
If Err = 3021 Then
    Err = 0
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_SALARY_HISTORY", "Add")
Resume Next

End Sub

Private Sub CmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub cmdOK_Click()
Dim rsCP As New ADODB.Recordset
Dim x, xID
Dim SHMark
On Error GoTo Add_Err


If Not chkComPlan() Then Exit Sub
Screen.MousePointer = HOURGLASS
Call UpdUStats(Me) ' update user's stats (who did it and when)

If glbtermopen Then Data1.Recordset("TERM_SEQ") = glbTERM_Seq
Data1.Recordset("CP_TARGINCOME") = Val(medTargIncome)

gdbAdoIhr001.BeginTrans
Call Set_Control("U", Me, rsDATA)
rsDATA.Update
gdbAdoIhr001.CommitTrans
rsDATA.Resync
xID = rsDATA!CP_ID
Data1.Refresh


Data1.Recordset.Find "CP_ID=" & xID


Call ST_UPD_MODE(False)
Screen.MousePointer = DEFAULT
fglbNew = False
Exit Sub

Add_Err:
If Err = 3021 Then
    Err = 0
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_SALARY_HISTORY", "Update")
Resume Next
Unload Me

End Sub

Private Sub cmdOK_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub cmdCalculate_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub cmdPrint_Click()
Dim RHeading As String, xReport, x%

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

cmdPrint.Enabled = False

RHeading = lblEEName & "'s Compensation Plan"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Action = 1

cmdPrint.Enabled = True
End Sub

Private Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Function CR_Sal_Snap()
Dim SQLQ As String

On Error GoTo JS_Err

SQLQ = "Select SH_EDATE,SH_SALARY,SH_SALCD,SH_WHRS "
SQLQ = SQLQ & " from HR_SALARY_HISTORY "
SQLQ = SQLQ & " WHERE SH_EMPNBR = " & glbLEE_ID & " "
SQLQ = SQLQ & " ORDER BY SH_EDATE DESC"
dynaSals.Open SQLQ, gdbAdoIhr001, adOpenStatic
Exit Function

JS_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SALARY History Snap", "HR_SALARY_HISTORY", "SELECT")
Resume Next

End Function

Private Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError
SQLQ = SQLQ & "SELECT * FROM COMPENSATION_PLAN "
SQLQ = SQLQ & "WHERE CP_EMPNBR = " & glbLEE_ID
SQLQ = SQLQ & " ORDER BY CP_FDATE DESC"

Data1.RecordSource = SQLQ
Data1.Refresh

EERetrieve = True

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Salary", "HR_SALARY_HISTORY", "SELECT")

Resume Next

Exit Function

End Function

Private Sub Form_Activate()
    glbOnTop = "FRMECOMPLAN"
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMECOMPLAN"
End Sub

Private Sub Form_Load()
flagFrmLoad = True 'carmen may 00
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim x%

glbOnTop = "FRMECOMPLAN"
If glbCompSerial = "S/N - 2291W" Then
    Data1.ConnectionString = Replace(UCase(glbAdoIHRDB), "IHR001.MDB", "SN2291.MDB")
End If
If glbCompSerial = "S/N - 2325W" Then
    Data1.ConnectionString = Replace(UCase(glbAdoIHRDB), "IHR001.MDB", "SN2325.MDB")
End If

cnCP.Open Data1.ConnectionString
Call DecSetup
oFDate = 0
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

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = ShowEmpnbr(glbLEE_ID)
Call CR_Sal_Snap
Call ST_UPD_MODE(False)
If Not gSec_Upd_Salary Then
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
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
    Set frmEComPlan = Nothing
End Sub

Private Sub medAdjSalary_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medTargBonus_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub
Private Sub medSalary_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Function Round2DEC(tmpNUM) 'laura nov 10, 1997
Dim strNUM As String, x%

If glbCompDecHR <> 2 And glbCompDecHR <> 3 And glbCompDecHR <> 4 Then
    glbCompDecHR = 2  'THIS SHOULD NOT HAPPEN BUT IS A VALID DEFAULT
End If
Round2DEC = Round(tmpNUM, glbCompDecHR)

End Function


Private Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

cmdOK.Enabled = TF
cmdCancel.Enabled = TF

cmdClose.Enabled = FT
cmdModify.Enabled = FT
cmdNew.Enabled = FT
cmdDelete.Enabled = FT
cmdPrint.Enabled = FT

cmdCalculate.Enabled = FT
vbxTrueGrid.Enabled = FT        'Jaddy 8/10/99

medTargBonus.Enabled = TF
medTargIncome.Enabled = False
medSalary.Enabled = False
medAdjSalary.Enabled = False
dlpDate(0).Enabled = TF
dlpDate(1).Enabled = False
txtCommPlan.Enabled = TF
txtComment.Enabled = TF

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End If
If glbtermopen Then
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
End If

End Sub




Private Sub medTargBonus_LostFocus()
    If Trim(medTargBonus) = "" Then medTargBonus = 0
    Call setIncome
End Sub

Private Sub medTargIncome_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtComment_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtCommPlan_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub
'Private Sub txtDate_Change(Index As Integer)
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtDate_DblClick(Index As Integer)
'    Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub txtDate_GotFocus(Index As Integer)
'    Call SetPanHelp(ActiveControl)
'    If Index = 0 Then oFDate = txtDate(0)
'End Sub
'Private Sub txtDate_KeyPress(Index As Integer, KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub
'Private Sub txtDate_LostFocus(Index As Integer)
'If Index = 0 Then
'    If IsDate(txtDate(0)) Then
'    'If Not IsNumeric(OFDate) And IsDate(txtDate(0)) Then
'    '    If OFDate <> txtDate(0) Then
'            txtDate(1) = DateAdd("d", -1, DateAdd("yyyy", 1, CVDate(txtDate(0))))
'            Call setSalary
'    '    End If
'    End If
'    oFDate = 0
'End If
'End Sub

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
        
        SQLQ = SQLQ & "SELECT * FROM COMPENSATION_PLAN "
        SQLQ = SQLQ & "WHERE CP_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
    If cmdOK.Enabled Then
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


Private Sub setIncome()
If Trim(medTargBonus) = "" Then medTargBonus = 0
medTargIncome = Val(medTargBonus) + Val(medSalary) + Val(medAdjSalary)
End Sub
Private Sub setSalary()
dynaSals.Requery
Call getSalary(dlpDate(0).Text, dlpDate(1).Text)
medSalary = xSalary(0)
medAdjSalary = xSalary(1) - xSalary(0)
Call setIncome
End Sub


Private Sub getSalary(xDate0, xDate1)
Dim xEDate
On Error GoTo JS_Err
dynaSals.Find "SH_EDATE<=#" & xDate1 & "#"
If dynaSals.EOF Then
    xSalary(0) = 0
    xSalary(1) = 0
    Exit Sub
End If
'M & D added by Bryan 28/Sep/05 Ticket#9354
If dynaSals("SH_SALCD") = "A" Then
    xSalary(1) = dynaSals("SH_SALARY")
ElseIf dynaSals("SH_SALCD") = "H" Then
    xSalary(1) = dynaSals("SH_SALARY") * 52 * dynaSals("SH_WHRS")
ElseIf dynaSals("SH_SALCD") = "M" Then
    xSalary(1) = dynaSals("SH_SALARY") * 12
ElseIf dynaSals("SH_SALCD") = "D" Then
    If GetLeapYear(Year(Date)) Then
        xSalary(1) = dynaSals("SH_SALARY") * 366
    Else
        xSalary(1) = dynaSals("SH_SALARY") * 365
    End If
    
End If

If dynaSals("SH_EDATE") <= CVDate(dlpDate(0).Text) Then
    xSalary(0) = xSalary(1)
    Exit Sub
End If

dynaSals.Find "SH_EDATE<=#" & dlpDate(0).Text & "#"
If dynaSals.EOF Then
    xSalary(0) = 0
Else
    'M and D added by Bryan 28/Sep/05 Ticket#9354
    If dynaSals("SH_SALCD") = "A" Then
        xSalary(0) = dynaSals("SH_SALARY")
    ElseIf dynaSals("SH_SALCD") = "H" Then
        xSalary(0) = dynaSals("SH_SALARY") * 52 * dynaSals("SH_WHRS")
    ElseIf dynaSals("SH_SALCD") = "M" Then
        xSalary(0) = dynaSals("SH_SALARY") * 12
    ElseIf dynaSals("SH_SALCD") = "D" Then
        If GetLeapYear(Year(Date)) Then
            xSalary(0) = dynaSals("SH_SALARY") * 366
        Else
            xSalary(0) = dynaSals("SH_SALARY") * 365
        End If
    End If
'    If dynaSals("SH_EDATE") = CVDate(txtDate(0)) Then
'        dynaSals.MoveNext
'        If Not dynaSals.EOF Then
'            xSalary(1) = xSalary(1) + (xSalary(0) - dynaSals("SH_EDATE"))
'        End If
'    End If
End If
Exit Sub

JS_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Set SALARY History", "HR_SALARY_HISTORY", "SELECT")
Resume Next

End Sub
Private Sub DecSetup()
Select Case glbCompDecHR
Case 3
    medSalary.Format = "#,##0.000;(#,##0.000)"
    medAdjSalary.Format = "#,##0.000;(#,##0.000)"
    medTargBonus.Format = "#,##0.000;(#,##0.000)"
    medTargIncome.Format = "#,##0.000;(#,##0.000)"
    vbxTrueGrid.Columns(2).NumberFormat = "#,##0.000;(#,##0.000)"
    vbxTrueGrid.Columns(3).NumberFormat = "#,##0.000;(#,##0.000)"
    vbxTrueGrid.Columns(4).NumberFormat = "#,##0.000;(#,##0.000)"
    vbxTrueGrid.Columns(5).NumberFormat = "#,##0.000;(#,##0.000)"
Case 4
    medSalary.Format = "#,##0.0000;(#,##0.0000)"
    medAdjSalary.Format = "#,##0.0000;(#,##0.0000)"
    medTargBonus.Format = "#,##0.0000;(#,##0.0000)"
    medTargIncome.Format = "#,##0.0000;(#,##0.0000)"
    vbxTrueGrid.Columns(2).NumberFormat = "#,##0.0000;(#,##0.0000)"
    vbxTrueGrid.Columns(3).NumberFormat = "#,##0.0000;(#,##0.0000)"
    vbxTrueGrid.Columns(4).NumberFormat = "#,##0.0000;(#,##0.0000)"
    vbxTrueGrid.Columns(5).NumberFormat = "#,##0.0000;(#,##0.0000)"
Case Else
    medSalary.Format = "#,##0.00;(#,##0.00)"
    medAdjSalary.Format = "#,##0.00;(#,##0.00)"
    medTargBonus.Format = "#,##0.00;(#,##0.00)"
    medTargIncome.Format = "#,##0.00;(#,##0.00)"
    vbxTrueGrid.Columns(2).NumberFormat = "#,##0.00;(#,##0.00)"
    vbxTrueGrid.Columns(3).NumberFormat = "#,##0.00;(#,##0.00)"
    vbxTrueGrid.Columns(4).NumberFormat = "#,##0.00;(#,##0.00)"
    vbxTrueGrid.Columns(5).NumberFormat = "#,##0.00;(#,##0.00)"
End Select
End Sub

''' Sam add July 2002 * Remove Binding Control
Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
      
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

        Exit Sub
    End If
    
   SQLQ = SQLQ & "SELECT * FROM COMPENSATION_PLAN "
   SQLQ = SQLQ & "WHERE CP_ID = " & Data1.Recordset!CP_ID
   SQLQ = SQLQ & " ORDER BY CP_FDATE DESC"
   If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
   rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

   If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
   Call Set_Control("R", Me, rsDATA)

End Sub






