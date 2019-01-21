VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmPosBudget 
   Caption         =   "Budgeted Positions"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   4185
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   10995
   WindowState     =   2  'Maximized
   Begin VB.TextBox medFTEHrs 
      Appearance      =   0  'Flat
      DataField       =   "JG_FTEHRS"
      Height          =   285
      Left            =   2115
      TabIndex        =   6
      Tag             =   "10-FTE Hours/Year"
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox medFTENum 
      Appearance      =   0  'Flat
      DataField       =   "JG_FTENUM"
      Height          =   285
      Left            =   2115
      TabIndex        =   5
      Tag             =   "10-Number of FTE "
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtNoPos 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      DataField       =   "JG_BUDGNBR"
      Height          =   285
      Left            =   2115
      MaxLength       =   3
      TabIndex        =   4
      Tag             =   "01-Number of positions that exist for this job"
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JG_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   5280
      MaxLength       =   25
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5640
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JG_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   4230
      MaxLength       =   25
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5640
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JG_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   2550
      MaxLength       =   25
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5640
      Visible         =   0   'False
      Width           =   1065
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10995
      _Version        =   65536
      _ExtentX        =   19394
      _ExtentY        =   873
      _StockProps     =   15
      ForeColor       =   0
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
      Begin VB.Label lblPosDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Descr"
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
         TabIndex        =   19
         Top             =   135
         Width           =   630
      End
      Begin VB.Label lblPosition 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ABCD"
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
         TabIndex        =   18
         Top             =   120
         Width           =   630
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   165
         Width           =   690
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   6195
      Width           =   10995
      _Version        =   65536
      _ExtentX        =   19394
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
      Begin VB.CommandButton cmdCountPos 
         Appearance      =   0  'Flat
         Caption         =   "&Count Budgeted Positions"
         Height          =   495
         Left            =   6240
         TabIndex        =   15
         Tag             =   "Count positions filled; total the points - for all pos'ns"
         Top             =   120
         Width           =   1905
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5265
         TabIndex        =   14
         Tag             =   "Print Listing "
         Top             =   195
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4425
         TabIndex        =   13
         Tag             =   "Delete the Record Selected"
         Top             =   195
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3585
         TabIndex        =   12
         Tag             =   "Add a new Record"
         Top             =   195
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2685
         TabIndex        =   11
         Tag             =   "Cancel the changes made"
         Top             =   195
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1845
         TabIndex        =   10
         Tag             =   "Save the changes made"
         Top             =   195
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1005
         TabIndex        =   9
         Tag             =   "Edit the information on this screen"
         Top             =   195
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   165
         TabIndex        =   8
         Tag             =   "Close and exit this screen"
         Top             =   195
         Visible         =   0   'False
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   9360
         Top             =   120
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
         PrintFileUseRptDateFmt=   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   8400
         Top             =   240
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Caption         =   "Adodc2"
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
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmPosBudget.frx":0000
      Height          =   2115
      Left            =   120
      OleObjectBlob   =   "frmPosBudget.frx":0014
      TabIndex        =   7
      Tag             =   "Skills Lookup"
      Top             =   600
      Width           =   10635
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      DataField       =   "JG_DIV"
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "00-Specific Division Desired"
      Top             =   3120
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      DataField       =   "JG_DEPTNO"
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   3480
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpGLNum 
      DataField       =   "JG_GLNO"
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Tag             =   "00-General Ledger - Code"
      Top             =   3840
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   3
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   39
      Top             =   3165
      Width           =   1410
   End
   Begin VB.Label lblFTEHrs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FTE Hours/Year"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   38
      Top             =   4965
      Width           =   1275
   End
   Begin VB.Label lblTotHrs 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total FTE Hours/Year"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3360
      TabIndex        =   37
      Top             =   4965
      Width           =   1575
   End
   Begin VB.Label lblFTETotHrs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      DataField       =   "JG_FTETOTHR"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5070
      TabIndex        =   36
      Top             =   4965
      Width           =   690
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      DataField       =   "JG_FTENUMVACN"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6480
      TabIndex        =   35
      Top             =   4605
      Width           =   570
   End
   Begin VB.Label lblPosFiled 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      DataField       =   "JG_NBRFIL"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4440
      TabIndex        =   34
      Top             =   4245
      Width           =   570
   End
   Begin VB.Label lblFTETotNum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      DataField       =   "JG_FTENUMFILL"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4440
      TabIndex        =   33
      Top             =   4605
      Width           =   570
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vacancy # FTE"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   5400
      TabIndex        =   32
      Top             =   4605
      Width           =   1440
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " FTE # Filled"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   3360
      TabIndex        =   31
      Top             =   4605
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "FTE #"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   360
      TabIndex        =   30
      Top             =   4605
      Width           =   1200
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Positions Filled"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   3360
      TabIndex        =   29
      Top             =   4245
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Budgeted #Pos'ns"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   28
      Top             =   4245
      Width           =   1440
   End
   Begin VB.Label lblID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   27
      Top             =   5640
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblPositions 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "POST"
      DataField       =   "JG_CODE"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   26
      Top             =   5760
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CompNo"
      DataField       =   "JG_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1080
      TabIndex        =   25
      Top             =   5640
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      Height          =   195
      Left            =   360
      TabIndex        =   21
      Top             =   3525
      Width           =   1560
   End
   Begin VB.Label lblGLNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "G/L Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   20
      Top             =   3885
      Width           =   870
   End
End
Attribute VB_Name = "frmPosBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRecords%, fglbEditMode%
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim LGR_snap As New ADODB.Recordset
Dim snapDiv As New ADODB.Recordset
Dim RDept, RGLNum
Dim rsDATA As New ADODB.Recordset
Dim fglbNew As Boolean

Private Sub clpDept_Change()
    Call Dept_GL
End Sub

Public Sub cmdCancel_Click()

On Error GoTo Can_Err
fglbNew = False
rsDATA.CancelUpdate
Call Display_Value

'Call ST_UPD_MODE(False)  ' reset screen's attributes
Call SET_UP_MODE
'Data1.Recordset.CancelUpdate
'If Not glbSQL Then Call Pause(0.5)
'Data1.Refresh


Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRJOBBUD", "Cancel")
Call RollBack   '15June99 js

End Sub

Public Sub cmdClose_Click()
Unload Me
End Sub


Private Sub cmdCountPos_Click()
On Error GoTo CountErr

If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
    If mod_Upd_Pos_Budget(True) Then
        Beep
        MsgBox "Budgeted Positions Counted"
    End If
    Data1.Refresh
End If

Exit Sub

CountErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Count Pos Error", "HRJOBBUD Refresh", "Refresh")
Resume Next
Call RollBack

End Sub

Public Sub cmdDelete_Click()
Dim a As Integer, Msg As String, INo&

If Not gSec_Upd_BudgetedPos Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    fglbRecords% = False
    Exit Sub
Else
    fglbRecords% = True
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Record?  "
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub
fglbNew = False
gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

'Call ST_UPD_MODE(False)
Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRJOBBUD", "Delete")
Call RollBack   '15June99 js

End Sub

Public Sub cmdModify_Click()
Dim SQLQ As String

If Not gSec_Upd_BudgetedPos Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

'Call ST_UPD_MODE(True)
Call SET_UP_MODE
On Error GoTo Edit_Err


fglbEditMode% = True

RDept = clpDept

clpDiv.SetFocus

Exit Sub

Edit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdModify", "HRJOBSKL", "Edit")
Call RollBack   '15June99 js
End Sub

Public Sub cmdNew_Click()
Dim SQLQ As String

If Not gSec_Upd_BudgetedPos Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

'Call ST_UPD_MODE(True)
fglbNew = True
Call SET_UP_MODE
On Error GoTo AddN_Err

Call Set_Control("B", Me, rsDATA)
rsDATA.AddNew

'Data1.Recordset.AddNew
fglbEditMode% = True
lblCNum.Caption = "001"
lblPositions.Caption = glbPos$

clpDept.Enabled = True
clpDiv.SetFocus
RDept = ""
Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRJOBBUD", "Add")
Call RollBack
End Sub

Public Sub cmdOK_Click()
On Error GoTo OK_Err

If Not chkBudgetPos() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)
Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans
Data1.Refresh

fglbNew = False
'Call ST_UPD_MODE(False)
Call SET_UP_MODE
fglbEditMode% = False

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRJOBBUD", "Update")
Call RollBack   '15June99 js
Unload Me

End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = Me.Caption
RHeading = Mid(RHeading, 1, InStr(RHeading, "-"))
RHeading = RHeading & " " & lblPosDesc.Caption

Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Public Sub cmdView_Click()
Dim RHeading As String

RHeading = Me.Caption
RHeading = Mid(RHeading, 1, InStr(RHeading, "-"))
RHeading = RHeading & " " & lblPosDesc.Caption

Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMPOSBUDGET"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim X%

On Error GoTo FLErr

glbOnTop = "FRMPOSBUDGET"

Screen.MousePointer = HOURGLASS
If glbPos = "" Then frmJOBS.Show 1
If glbPos = "" Then glbUserUploadMode = UploadFormWithoutCheck: Unload Me: Exit Sub

'Ticket #20583 Franks 07/07/2011, make Div optional for all customers except WorkSafe NB
If Not glbCompSerial = "S/N - 2336W" Then 'WorkSafe NB
    lblDiv.FontBold = False
End If

lblPosition.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$
Me.Caption = "Budgeted Positions - " & lblPosition

Data1.ConnectionString = glbAdoIHRDB
'Call CR_Lgr_Snap

If Not EERetrieve() Then
    Exit Sub        '  modGet it sets fglbRecords
End If
lblDiv.Caption = lStr(lblDiv)
lblDept.Caption = lStr(lblDept)
lblGLNo.Caption = lStr(lblGLNo)
Call INI_Controls(Me)
Call Display_Value


'Call SET_UP_MODE
Call SET_UP_MODE
'If glbWHSCC And Not gSec_Upd_WHSCC_BUDPOS% Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
'    cmdCountPos.Enabled = False
'End If

Exit Sub

FLErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form load Error", "Budgeted Positions", "Select")
Call RollBack   '15June99 js


End Sub

Public Function EERetrieve() 'StrPos$)
Dim SQLQ$

EERetrieve = False
Screen.MousePointer = HOURGLASS

On Error GoTo EERetrieveErr


' out or left join query not updateable - so do straight.
SQLQ$ = "SELECT * FROM HRJOBBUD "
SQLQ$ = SQLQ$ & "WHERE JG_CODE = '" & glbPos$ & "' "
SQLQ$ = SQLQ$ & "ORDER BY JG_CODE"

Data1.RecordSource = SQLQ$
Data1.Refresh

lblPosition.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    fglbRecords% = False
    cmdModify.Enabled = False       'Laura jan 06, 1998
Else
    fglbRecords% = True
End If
EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERetrieveErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Budgeted Positions", "HRJOBBUD", "SELECT")
Call RollBack   '15June99 js

End Function



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub

Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)

End Sub

Private Sub lblPositions_Change()
lblPosition.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$
End Sub

Private Sub medFTEHrs_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medFTENum_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Dept_GL()
Dim snapDepts As New ADODB.Recordset
Dim Response%, Msg$, Title$, DgDef As Double
Dim SQLQ
If Not cmdOK.Enabled Then RDept = clpDept
If Len(clpDept.Text) > 0 Then

    SQLQ = "Select DISTINCT HRDEPT.DF_NBR, HRDEPT.DF_NAME, HRDEPT.DF_GLNO from HRDEPT"
    SQLQ = SQLQ & " Where " & glbSeleDept
    SQLQ = SQLQ & " AND DF_NBR = '" & clpDept.Text & "'"
    If glbOracle Then
        SQLQ = SQLQ & " ORDER BY DF_NAME "
    Else
        SQLQ = SQLQ & " ORDER BY [DF_NAME] "
    End If
    If snapDepts.State <> 0 Then snapDepts.Close
    snapDepts.Open SQLQ, gdbAdoIhr001, adOpenStatic

    If Not snapDepts.EOF Then
        RGLNum = snapDepts("DF_GLNO")
        If RDept <> clpDept Then
            If IsNull(RGLNum) Then
                RGLNum = ""
                'txtGLNum = ""
            Else
                Msg$ = "Do you want the associated G/L #?"
                Title$ = "info:HR"
                DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
                Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                If Response% = IDYES Then clpGLNum = RGLNum
            End If
            RDept = clpDept

        End If
    End If
End If
End Sub

Private Sub CR_Lgr_Snap()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String

On Error GoTo Job_Err

Screen.MousePointer = HOURGLASS
SQLQ = "SELECT * FROM HRGL "

If LGR_snap.State <> 0 Then LGR_snap.Close
LGR_snap.Open SQLQ, gdbAdoIhr001, adOpenStatic

Screen.MousePointer = DEFAULT

Exit Sub

Job_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Descriptions", "HRGL", "SELECT")
Call RollBack '21June99 js

End Sub

Private Function chkBudgetPos()
Dim SQLQ As String, Msg As String, dd#, PID&, xDiv, xDeptno$, xGLNO$, xPosCtrl

chkBudgetPos = False

On Error GoTo chkBudgetPos_Err

If Len(clpDiv) < 1 Then
    ''Added 2363 by Bryan 28/Sep/05 Ticket#9415
    'If glbCompSerial <> "S/N - 2366W" And glbCompSerial <> "S/N - 2363W" Then
    'Ticket #20583 Franks 07/07/2011, make Div optional for all customers except WorkSafe NB
    If glbCompSerial = "S/N - 2336W" Then 'WorkSafe NB
        MsgBox lStr("Division is a required field")
        clpDiv.SetFocus
        Exit Function
    End If
Else
    If clpDiv.Caption = "Unassigned" Then
        MsgBox lStr("Division must be valid")
        clpDiv.SetFocus
        Exit Function
    End If
End If

If Len(clpDept) < 1 Then
    MsgBox lStr("Department is a required field")
    clpDept.SetFocus
    Exit Function
Else
    If clpDept.Caption = "Unassigned" Then
        MsgBox lStr("Department must be valid")
        clpDept.SetFocus
        Exit Function
    End If
End If

If Len(clpGLNum) < 1 Then
    If glbWHSCC Then
        MsgBox lStr("G/L Code is a required field")
        clpGLNum.SetFocus
        Exit Function
    End If
Else
    If clpGLNum.Caption = "Unassigned" Then
        MsgBox lStr("G/L Code must be valid")
        clpGLNum.SetFocus
        Exit Function
    End If
End If
If IsNull(rsDATA("JG_ID")) Then
    PID& = 0
Else
    PID& = rsDATA("JG_ID") ' CLng(Val(lblID))
End If
xDiv = clpDiv
xDeptno$ = clpDept
xGLNO$ = clpGLNum

'If Not glbWHSCC Then
'    If modISDupBudgetPosCtrl(glbPos$, xPosCtrl, PID&) Then
'        MsgBox "Position Control # must be unique"
'        clpDiv.SetFocus
'        Exit Function
'    End If
'Else
    If modISDupBudget(glbPos$, xDiv, xDeptno$, xGLNO$, PID&) Then
        If Len(xGLNO$) > 0 Then
            If Len(xDiv) > 0 Then
                MsgBox lStr("[Division]") & " + " & lStr("[Department]") & " + " & lStr("[G/L Code]") & " must be unique"
            Else
                MsgBox lStr("[Department]") & " + " & lStr("[G/L Code]") & " must be unique"
            End If
        Else
            If Len(xDiv) > 0 Then
                MsgBox lStr("[Division]") & " + " & lStr("[Department]") & " must be unique"
            Else
                MsgBox lStr("[Department]") & " must be unique"
            End If
        End If
        clpDiv.SetFocus
        Exit Function
    End If
'End If
chkBudgetPos = True

Exit Function

chkBudgetPos_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HRJOBSKL", "edit/Add")
Call RollBack   '15June99 js

End Function
Private Function modISDupBudgetPosCtrl(Pos$, xPosCtrl, ID&)
Dim SQLQ$
Dim snapBudget As New ADODB.Recordset

modISDupBudgetPosCtrl = True

On Error GoTo modISDupBudget_Err

Screen.MousePointer = HOURGLASS

SQLQ$ = "SELECT * FROM HRJOBBUD "
SQLQ$ = SQLQ$ & "Where "
SQLQ$ = SQLQ$ & " (JG_CODE = '" & Pos$ & "' "
SQLQ$ = SQLQ$ & "AND JG_POSCTRLNO = '" & xPosCtrl & "' "
SQLQ$ = SQLQ$ & "AND JG_ID <> " & ID& & ") "
If snapBudget.State <> 0 Then snapBudget.Close
snapBudget.Open SQLQ$, gdbAdoIhr001, adOpenStatic

If snapBudget.BOF And snapBudget.EOF Then
    modISDupBudgetPosCtrl = False
End If

Screen.MousePointer = DEFAULT
snapBudget.Close

Exit Function

modISDupBudget_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Code Snap", "TABL", "SELECT")
Call RollBack   '15June99 js
End Function
Private Function modISDupBudget(Pos$, xDiv, DeptNo$, GLNO$, ID&)
Dim SQLQ$
Dim snapBudget As New ADODB.Recordset

modISDupBudget = True

On Error GoTo modISDupBudget_Err

Screen.MousePointer = HOURGLASS

SQLQ$ = "SELECT * FROM HRJOBBUD "
SQLQ$ = SQLQ$ & "Where "
SQLQ$ = SQLQ$ & " (JG_CODE = '" & Pos$ & "' "
'2363 Added by Bryan 04/Oct/05 Ticket# 9415
If glbCompSerial = "S/N - 2366W" Or glbCompSerial = "S/N - 2363W" Then
    If Len(xDiv) > 0 Then
        SQLQ$ = SQLQ$ & "AND JG_DIV = '" & xDiv & "' "
    End If
Else
    SQLQ$ = SQLQ$ & "AND JG_DIV = '" & xDiv & "' "
End If

SQLQ$ = SQLQ$ & "AND JG_DEPTNO = '" & DeptNo$ & "' "

If Len(GLNO$) > 0 Then
    SQLQ$ = SQLQ$ & "AND JG_GLNO = '" & GLNO$ & "' "
End If

SQLQ$ = SQLQ$ & "AND JG_ID <> " & ID& & ") "

If snapBudget.State <> 0 Then snapBudget.Close
snapBudget.Open SQLQ$, gdbAdoIhr001, adOpenStatic

If snapBudget.BOF And snapBudget.EOF Then
    modISDupBudget = False
End If

Screen.MousePointer = DEFAULT
snapBudget.Close

Exit Function

modISDupBudget_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Code Snap", "TABL", "SELECT")
Call RollBack   '15June99 js

End Function

Function mod_Upd_Pos_Budget(updPCtComp%)
Dim snapJobCount As New ADODB.Recordset
Dim rsHrJob As New ADODB.Recordset
Dim Comp$, Job$, JobCount&, SQLQ As String, pct#, ipct#, rcount&, spct%
Dim JobPoints#
Dim snapEvalPoints As New ADODB.Recordset
Dim FTENum#, FTEHrs#
Dim snapFTENum As New ADODB.Recordset
Dim snapFTEHrs As New ADODB.Recordset
Dim snapBudget As New ADODB.Recordset
Dim xJOB, xDiv, xDeptno, xGLNO, xPosCtrl

mod_Upd_Pos_Budget = False
On Error GoTo mod_Upd_Pos_Budget_Err
MDIMain.panHelp(0).FloodShowPct = True
MDIMain.panHelp(0).ForeColor = &HFFFFFF
pct# = 1

MDIMain.panHelp(0).FloodType = 1

SQLQ = "SELECT * FROM HRJOBBUD"
If snapBudget.State <> 0 Then snapBudget.Close
snapBudget.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not (snapBudget.EOF And snapBudget.BOF) Then
    snapBudget.MoveLast
    rcount& = snapBudget.RecordCount
    snapBudget.MoveFirst
End If
pct# = 0
Do While Not snapBudget.EOF
    MDIMain.panHelp(0).FloodPercent = (pct# / rcount&) * 100
    pct# = pct# + 1
    xDiv = "": xDeptno = "": xGLNO = "": xPosCtrl = ""
    xJOB = snapBudget("JG_CODE")
    If Not IsNull(snapBudget("JG_DIV")) Then
        xDiv = snapBudget("JG_DIV")
    End If
    If Not IsNull(snapBudget("JG_DEPTNO")) Then
        xDeptno = snapBudget("JG_DEPTNO")
    End If
    If Not IsNull(snapBudget("JG_GLNO")) Then
        xGLNO = snapBudget("JG_GLNO")
    End If
    'If Not IsNull(snapBudget("JG_POSCTRLNO")) Then
    '    xPosCtrl = snapBudget("JG_POSCTRLNO")
    'End If
    
    'Position filled
    If glbMulti Then
        SQLQ = "SELECT HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB,HR_JOB_HISTORY.JH_DEPTNO, "
        'Ticket #21537 - The JH_DIV grouping is causing an issue with the FTE count. Div is not mandatory,
        'therefore it should not be part of the grouping. The WHERE clause picks the right group and gives
        'the count.
        'Added by Bryan 17/Oct/05 Ticket# 9415
        If (glbCompSerial = "S/N - 2363W" And xDiv <> "") Then 'Or glbCompSerial <> "S/N - 2363W" Then
            SQLQ = SQLQ & "HR_JOB_HISTORY.JH_DIV, "
        Else
            'Ticket #21537
            If Len(xDiv) > 0 Then
                SQLQ = SQLQ & "HR_JOB_HISTORY.JH_DIV, "
            End If
        End If
        
        If Len(xGLNO) > 0 Then
            SQLQ = SQLQ & "HR_JOB_HISTORY.JH_GLNO, "
        End If
        
        'If Len(xPosCtrl) > 0 Then
        '    SQLQ = SQLQ & "HR_JOB_HISTORY.JH_POSITION_CONTROL, "
        'End If
        
        SQLQ = SQLQ & "COUNT(HR_JOB_HISTORY.JH_EMPNBR) AS NoPosFilled  "
        SQLQ = SQLQ & "FROM HR_JOB_HISTORY "
        SQLQ = SQLQ & "WHERE (JH_CURRENT <> 0) AND JH_JOB = '" & xJOB & "' "
        
        'Added by Bryan 04/Oct/05 Ticket# 9415
        If glbCompSerial = "S/N - 2363W" And xDiv = "" Then
            SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_DIV is null "
        Else
            'Ticket #21537
            If Len(xDiv) > 0 Then
                SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_DIV = '" & xDiv & "' "
            End If
        End If
        
        SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_DEPTNO = '" & xDeptno & "' "
        
        If Len(xGLNO) > 0 Then
            SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_GLNO = '" & xGLNO & "' "
        End If
        'If Len(xPosCtrl) > 0 Then
        '    SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_POSITION_CONTROL = '" & xPosCtrl & "' "
        'End If
        'SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_EMP <> 'CAS' "
        'SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_EMP <> 'CONT' "
        'SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_EMP <> 'TA' "
        If glbCompSerial = "S/N - 2343W" Then 'OCCAC
            SQLQ = SQLQ & "AND (HR_JOB_HISTORY.JH_EMP = 'PERM' "
            SQLQ = SQLQ & "OR HR_JOB_HISTORY.JH_EMP = 'EIS' "
            SQLQ = SQLQ & "OR HR_JOB_HISTORY.JH_EMP = 'LTD' "
            SQLQ = SQLQ & "OR HR_JOB_HISTORY.JH_EMP = 'MAT' "
            SQLQ = SQLQ & "OR HR_JOB_HISTORY.JH_EMP = 'PAR' "
            SQLQ = SQLQ & "OR HR_JOB_HISTORY.JH_EMP = 'WCB') "
        End If
        
        SQLQ = SQLQ & "GROUP BY HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB,HR_JOB_HISTORY.JH_DEPTNO "
        'Ticket #21537 - The JH_DIV grouping is causing an issue with the FTE count. Div is not mandatory,
        'therefore it should not be part of the grouping. The WHERE clause picks the right group and gives
        'the count.
        'Added by Bryan 17/Oct/05 Ticket# 9415
        If (glbCompSerial = "S/N - 2363W" And xDiv <> "") Then 'Or glbCompSerial <> "S/N - 2363W" Then
            SQLQ = SQLQ & ", HR_JOB_HISTORY.JH_DIV "
        Else
            If Len(xDiv) > 0 Then
                SQLQ = SQLQ & ", HR_JOB_HISTORY.JH_DIV "
            End If
        End If
        If Len(xGLNO) > 0 Then
            SQLQ = SQLQ & ",HR_JOB_HISTORY.JH_GLNO "
        End If
        'If Len(xPosCtrl) > 0 Then
        '    SQLQ = SQLQ & ",HR_JOB_HISTORY.JH_POSITION_CONTROL "
        'End If
    Else
        If glbOracle Then
            SQLQ = "SELECT HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB,HR_JOB_HISTORY.JH_DEPTNO, "
            'Ticket #21537 - The JH_DIV grouping is causing an issue with the FTE count. Div is not mandatory,
            'therefore it should not be part of the grouping. The WHERE clause picks the right group and gives
            'the count.
            'Added by Bryan 17/Oct/05 Ticket# 9415
            If (glbCompSerial = "S/N - 2363W" And xDiv <> "") Then 'Or glbCompSerial <> "S/N - 2363W" Then
                SQLQ = SQLQ & "HR_JOB_HISTORY.JH_DIV, "
            Else
                'Ticket #21537
                If Len(xDiv) > 0 Then
                    SQLQ = SQLQ & "HR_JOB_HISTORY.JH_DIV, "
                End If
            End If

            If Len(xGLNO) > 0 Then
                SQLQ = SQLQ & "HREMP.ED_GLNO, "
            End If
            'If Len(xPosCtrl) > 0 Then
            '    SQLQ = SQLQ & "HR_JOB_HISTORY.JH_POSITION_CONTROL, "
            'End If
            SQLQ = SQLQ & "COUNT(HR_JOB_HISTORY.JH_EMPNBR) AS NoPosFilled  "
            SQLQ = SQLQ & "FROM HR_JOB_HISTORY,HREMP "
            SQLQ = SQLQ & "WHERE (JH_CURRENT <> 0) AND JH_JOB = '" & xJOB & "' "
            'Added by Bryan 04/Oct/05 Ticket# 9415
            If glbCompSerial <> "S/N - 2363W" Then
                'Ticket #21537
                If Len(xDiv) > 0 Then
                    SQLQ = SQLQ & "AND HREMP.ED_DIV = '" & xDiv & "' "
                End If
            End If
            SQLQ = SQLQ & "AND HREMP.ED_DEPTNO = '" & xDeptno & "' "
            If Len(xGLNO) > 0 Then
                SQLQ = SQLQ & "AND HREMP.ED_GLNO = '" & xGLNO & "' "
            End If
            'If Len(xPosCtrl) > 0 Then
            '    SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_POSITION_CONTROL = '" & xPosCtrl & "' "
            'End If
            SQLQ = SQLQ & "AND HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
            SQLQ = SQLQ & "GROUP BY HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB,HR_JOB_HISTORY.JH_DEPTNO "
            'Ticket #21537 - The JH_DIV grouping is causing an issue with the FTE count. Div is not mandatory,
            'therefore it should not be part of the grouping. The WHERE clause picks the right group and gives
            'the count.
            'Added by Bryan 17/Oct/05 Ticket# 9415
            If (glbCompSerial = "S/N - 2363W" And xDiv <> "") Then 'Or glbCompSerial <> "S/N - 2363W" Then
                SQLQ = SQLQ & ", HR_JOB_HISTORY.JH_DIV "
            Else
                'Ticket #21537
                If Len(xDiv) > 0 Then
                    SQLQ = SQLQ & ", HR_JOB_HISTORY.JH_DIV "
                End If
            End If

            If Len(xGLNO) > 0 Then
                SQLQ = SQLQ & ",HREMP.ED_GLNO "
            End If
            'If Len(xPosCtrl) > 0 Then
            '    SQLQ = SQLQ & ",HR_JOB_HISTORY.JH_POSITION_CONTROL "
            'End If
        Else
            SQLQ = "SELECT HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB,HR_JOB_HISTORY.JH_DEPTNO, "
            'Ticket #21537 - The JH_DIV grouping is causing an issue with the FTE count. Div is not mandatory,
            'therefore it should not be part of the grouping. The WHERE clause picks the right group and gives
            'the count.
            'Added by Bryan 17/Oct/05 Ticket# 9415
            If (glbCompSerial = "S/N - 2363W" And xDiv <> "") Then 'Or glbCompSerial <> "S/N - 2363W" Then
                SQLQ = SQLQ & "HR_JOB_HISTORY.JH_DIV, "
            Else
                'Ticket #21537
                If Len(xDiv) > 0 Then
                    SQLQ = SQLQ & "HR_JOB_HISTORY.JH_DIV, "
                End If
            End If

            If Len(xGLNO) > 0 Then
                SQLQ = SQLQ & "HREMP.ED_GLNO, "
            End If
            'If Len(xPosCtrl) > 0 Then
            '    SQLQ = SQLQ & "HR_JOB_HISTORY.JH_POSITION_CONTROL, "
            'End If
            
            SQLQ = SQLQ & "COUNT(HR_JOB_HISTORY.JH_EMPNBR) AS NoPosFilled  "
            SQLQ = SQLQ & "FROM HR_JOB_HISTORY "
            SQLQ = SQLQ & "INNER JOIN HREMP ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
            SQLQ = SQLQ & "WHERE (JH_CURRENT <> 0) AND JH_JOB = '" & xJOB & "' "
            If glbCompSerial <> "S/N - 2363W" Then
                'Ticket #21537
                If Len(xDiv) > 0 Then
                    SQLQ = SQLQ & "AND HREMP.ED_DIV = '" & xDiv & "' "
                End If
            End If
            SQLQ = SQLQ & "AND HREMP.ED_DEPTNO = '" & xDeptno & "' "
            If Len(xGLNO) > 0 Then
                SQLQ = SQLQ & "AND HREMP.ED_GLNO = '" & xGLNO & "' "
            End If
            'If Len(xPosCtrl) > 0 Then
            '    SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_POSITION_CONTROL = '" & xPosCtrl & "' "
            'End If
            
            SQLQ = SQLQ & "GROUP BY HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB,HR_JOB_HISTORY.JH_DEPTNO "
            'Ticket #21537 - The JH_DIV grouping is causing an issue with the FTE count. Div is not mandatory,
            'therefore it should not be part of the grouping. The WHERE clause picks the right group and gives
            'the count.
            'Added by Bryan 17/Oct/05 Ticket# 9415
            If (glbCompSerial = "S/N - 2363W" And xDiv <> "") Then 'Or glbCompSerial <> "S/N - 2363W" Then
                SQLQ = SQLQ & ", HR_JOB_HISTORY.JH_DIV "
            Else
                'Ticket #21537
                If Len(xDiv) > 0 Then
                    SQLQ = SQLQ & ", HR_JOB_HISTORY.JH_DIV "
                End If
            End If
            If Len(xGLNO) > 0 Then
                SQLQ = SQLQ & ",HREMP.ED_GLNO "
            End If
            'If Len(xPosCtrl) > 0 Then
            '    SQLQ = SQLQ & ",HR_JOB_HISTORY.JH_POSITION_CONTROL "
            'End If
        End If
    End If


    snapJobCount.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not snapJobCount.EOF Then
        If Not IsNull(snapJobCount("NoPosFilled")) Then
            JobCount& = snapJobCount("NoPosFilled")
        Else
            JobCount& = 0
        End If
        snapBudget("JG_NBRFIL") = JobCount&
        snapBudget.Update
    Else
        snapBudget("JG_NBRFIL") = 0
    End If
    snapJobCount.Close

    'FTE # filled & FTE Hours/Year
    If glbMulti Then
        SQLQ = "SELECT HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB,HR_JOB_HISTORY.JH_DEPTNO, "
        'Added by Bryan 17/Oct/05 Ticket# 9415
        If (glbCompSerial = "S/N - 2363W" And xDiv <> "") Or glbCompSerial <> "S/N - 2363W" Then
            SQLQ = SQLQ & "HR_JOB_HISTORY.JH_DIV, "
        End If
        
        If Len(xGLNO) > 0 Then
            SQLQ = SQLQ & "HR_JOB_HISTORY.JH_GLNO, "
        End If
        'If Len(xPosCtrl) > 0 Then
        '    SQLQ = SQLQ & "HR_JOB_HISTORY.JH_POSITION_CONTROL, "
        'End If
        SQLQ = SQLQ & "SUM(JH_FTENUM) AS FTENumTot, SUM(JH_FTEHRS) AS FTEHrsTot "
        SQLQ = SQLQ & "FROM HR_JOB_HISTORY  "
        SQLQ = SQLQ & "WHERE (JH_CURRENT <> 0) AND JH_JOB = '" & xJOB & "' "
         'Added by Bryan 04/Oct/05 Ticket# 9415
        If (glbCompSerial <> "S/N - 2363W") Or (glbCompSerial = "S/N - 2363W" And xDiv <> "") Then
            'Ticket #21537
            If Len(xDiv) > 0 Then
                SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_DIV = '" & xDiv & "' "
            End If
        End If
        SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_DEPTNO = '" & xDeptno & "' "
        If Len(xGLNO) > 0 Then
            SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_GLNO = '" & xGLNO & "' "
        End If
        'If Len(xPosCtrl) > 0 Then
        '    SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_POSITION_CONTROL = '" & xPosCtrl & "' "
        'End If
        'SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_EMP <> 'CAS' "
        'SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_EMP <> 'CONT' "
        'SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_EMP <> 'TA' "
        If glbCompSerial = "S/N - 2343W" Then 'OCCAC
            SQLQ = SQLQ & "AND (HR_JOB_HISTORY.JH_EMP = 'PERM' "
            SQLQ = SQLQ & "OR HR_JOB_HISTORY.JH_EMP = 'EIS' "
            SQLQ = SQLQ & "OR HR_JOB_HISTORY.JH_EMP = 'LTD' "
            SQLQ = SQLQ & "OR HR_JOB_HISTORY.JH_EMP = 'MAT' "
            SQLQ = SQLQ & "OR HR_JOB_HISTORY.JH_EMP = 'PAR' "
            SQLQ = SQLQ & "OR HR_JOB_HISTORY.JH_EMP = 'WCB') "
        End If
        SQLQ = SQLQ & "GROUP BY HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB,HR_JOB_HISTORY.JH_DEPTNO "
        'Added by Bryan 17/Oct/05 Ticket# 9415
        If (glbCompSerial = "S/N - 2363W" And xDiv <> "") Or glbCompSerial <> "S/N - 2363W" Then
            SQLQ = SQLQ & ", HR_JOB_HISTORY.JH_DIV "
        End If
        If Len(xGLNO) > 0 Then
            SQLQ = SQLQ & ",HR_JOB_HISTORY.JH_GLNO "
        End If
        'If Len(xPosCtrl) > 0 Then
        '    SQLQ = SQLQ & ",HR_JOB_HISTORY.JH_POSITION_CONTROL "
        'End If
    Else
        If glbOracle Then
            SQLQ = "SELECT HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB,HR_JOB_HISTORY.JH_DEPTNO, "
            'Added by Bryan 17/Oct/05 Ticket# 9415
            If (glbCompSerial = "S/N - 2363W" And xDiv <> "") Or glbCompSerial <> "S/N - 2363W" Then
                SQLQ = SQLQ & "HR_JOB_HISTORY.JH_DIV, "
            End If

            If Len(xGLNO) > 0 Then
                SQLQ = SQLQ & "HREMP.ED_GLNO, "
            End If
            'If Len(xPosCtrl) > 0 Then
            '    SQLQ = SQLQ & "HR_JOB_HISTORY.JH_POSITION_CONTROL, "
            'End If
            SQLQ = SQLQ & "SUM(JH_FTENUM) AS FTENumTot, SUM(JH_FTEHRS) AS FTEHrsTot "
            SQLQ = SQLQ & "FROM HR_JOB_HISTORY,HREMP  "
            SQLQ = SQLQ & "WHERE (JH_CURRENT <> 0) AND JH_JOB = '" & xJOB & "' "
            'Added by Bryan 04/Oct/05 Ticket# 9415
            If (glbCompSerial <> "S/N - 2363W") Or (glbCompSerial = "S/N - 2363W" And xDiv <> "") Then
                'Ticket #21537
                If Len(xDiv) > 0 Then
                    SQLQ = SQLQ & "AND HREMP.ED_DIV = '" & xDiv & "' "
                End If
            End If
            SQLQ = SQLQ & "AND HREMP.ED_DEPTNO = '" & xDeptno & "' "
            If Len(xGLNO) > 0 Then
                SQLQ = SQLQ & "AND HREMP.ED_GLNO = '" & xGLNO & "' "
            End If
            'If Len(xPosCtrl) > 0 Then
            '    SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_POSITION_CONTROL = '" & xPosCtrl & "' "
            'End If
            SQLQ = SQLQ & "AND HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
            SQLQ = SQLQ & "GROUP BY HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB,HR_JOB_HISTORY.JH_DEPTNO "
            'Added by Bryan 17/Oct/05 Ticket# 9415
            If (glbCompSerial = "S/N - 2363W" And xDiv <> "") Or glbCompSerial <> "S/N - 2363W" Then
                SQLQ = SQLQ & ", HR_JOB_HISTORY.JH_DIV "
            End If
            If Len(xGLNO) > 0 Then
                SQLQ = SQLQ & ",HREMP.ED_GLNO "
            End If
            'If Len(xPosCtrl) > 0 Then
            '    SQLQ = SQLQ & ",HR_JOB_HISTORY.JH_POSITION_CONTROL "
            'End If
        Else
            SQLQ = "SELECT HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB,HR_JOB_HISTORY.JH_DEPTNO, "
            'Added by Bryan 17/Oct/05 Ticket# 9415
            If (glbCompSerial = "S/N - 2363W" And xDiv <> "") Or glbCompSerial <> "S/N - 2363W" Then
                SQLQ = SQLQ & "HR_JOB_HISTORY.JH_DIV, "
            End If

            If Len(xGLNO) > 0 Then
                SQLQ = SQLQ & "HREMP.ED_GLNO, "
            End If
            'If Len(xPosCtrl) > 0 Then
            '    SQLQ = SQLQ & "HR_JOB_HISTORY.JH_POSITION_CONTROL, "
            'End If
            SQLQ = SQLQ & "SUM(JH_FTENUM) AS FTENumTot, SUM(JH_FTEHRS) AS FTEHrsTot "
            SQLQ = SQLQ & "FROM HR_JOB_HISTORY  "
            SQLQ = SQLQ & "INNER JOIN HREMP ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
            SQLQ = SQLQ & "WHERE (JH_CURRENT <> 0) AND JH_JOB = '" & xJOB & "' "
            'Added by Bryan 04/Oct/05 Ticket# 9415
            If (glbCompSerial <> "S/N - 2363W") Or (glbCompSerial = "S/N - 2363W" And xDiv <> "") Then
                'Ticket #21537
                If Len(xDiv) > 0 Then
                    SQLQ = SQLQ & "AND HREMP.ED_DIV = '" & xDiv & "' "
                End If
            End If
            SQLQ = SQLQ & "AND HREMP.ED_DEPTNO = '" & xDeptno & "' "
            If Len(xGLNO) > 0 Then
                SQLQ = SQLQ & "AND HREMP.ED_GLNO = '" & xGLNO & "' "
            End If
            'If Len(xPosCtrl) > 0 Then
            '    SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_POSITION_CONTROL = '" & xPosCtrl & "' "
            'End If
            SQLQ = SQLQ & "GROUP BY HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB,HR_JOB_HISTORY.JH_DEPTNO "
            'Added by Bryan 17/Oct/05 Ticket# 9415
            If (glbCompSerial = "S/N - 2363W" And xDiv <> "") Or glbCompSerial <> "S/N - 2363W" Then
                SQLQ = SQLQ & ", HR_JOB_HISTORY.JH_DIV "
            End If

            If Len(xGLNO) > 0 Then
                SQLQ = SQLQ & ",HREMP.ED_GLNO "
            End If
            'If Len(xPosCtrl) > 0 Then
            '    SQLQ = SQLQ & ",HR_JOB_HISTORY.JH_POSITION_CONTROL "
            'End If
        End If
    End If
    

    '--------
    snapFTENum.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not snapFTENum.EOF Then
        If Not IsNull(snapFTENum("FTENumTot")) Then
            FTENum# = snapFTENum("FTENumTot")
        Else
            FTENum# = 0
        End If
        If Not IsNull(snapFTENum("FTEHrsTot")) Then
            FTEHrs# = snapFTENum("FTEHrsTot")
        Else
            FTEHrs# = 0
        End If
        snapBudget("JG_FTENUMFILL") = FTENum#
        snapBudget("JG_FTENUMVACN") = snapBudget("JG_FTENUM") - FTENum#
        snapBudget("JG_FTETOTHR") = FTEHrs#
        snapBudget.Update
    Else
        snapBudget("JG_FTENUMFILL") = 0
        snapBudget("JG_FTENUMVACN") = snapBudget("JG_FTENUM")
        snapBudget.Update
    End If
    snapFTENum.Close
    
    
    snapBudget.MoveNext
Loop
snapBudget.Close

MDIMain.panHelp(0).FloodPercent = 0
MDIMain.panHelp(0).ForeColor = &H0&
MDIMain.panHelp(0).FloodType = 0
mod_Upd_Pos_Budget = True

Exit Function


mod_Upd_Pos_Budget_Err:
If Err = 94 Then
Err = 0
Resume Next
End If
glbFrmCaption$ = "Module - Count Budgeted Positions"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update HRJOBBUD Count", "HRJOBBUD", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Private Sub txtNoPos_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Public Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "SELECT * FROM HRJOBBUD "
    SQLQ = SQLQ & "WHERE JG_ID = " & Data1.Recordset!JG_ID
    SQLQ = SQLQ & " ORDER BY JG_CODE"

    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    lblID = rsDATA!JG_ID
    Call Set_Control("R", Me, rsDATA)
End If
    Call SET_UP_MODE
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        ' out or left join query not updateable - so do straight.
        SQLQ$ = "SELECT * FROM HRJOBBUD "
        SQLQ$ = SQLQ$ & "WHERE JG_CODE = '" & glbPos$ & "' "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$, X%
Dim SQLQ As String

On Error GoTo Tab1_Err
Call Display_Value

Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HRJOBBUD", "Add")
Call RollBack   '15June99 js

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
UpdateRight = gSec_Upd_BudgetedPos
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

clpDiv.Enabled = TF
clpDept.Enabled = TF
clpGLNum.Enabled = TF
txtNoPos.Enabled = TF

medFTENum.Enabled = TF
medFTEHrs.Enabled = TF
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
