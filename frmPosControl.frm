VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmPosControl 
   Caption         =   "CCAC Position  Number"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   4185
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   10995
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      DataField       =   "PC_EMPNBR"
      Height          =   285
      Left            =   1740
      TabIndex        =   14
      Top             =   3630
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
      Enabled         =   0   'False
   End
   Begin VB.TextBox txtPosCtrl 
      Appearance      =   0  'Flat
      DataField       =   "PC_POSITION_CONTROL"
      Height          =   285
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   13
      Tag             =   "10-CCAC Position #"
      Top             =   3240
      Width           =   870
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PC_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   5280
      MaxLength       =   25
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5640
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PC_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   4230
      MaxLength       =   25
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5640
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PC_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   2550
      MaxLength       =   25
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5640
      Visible         =   0   'False
      Width           =   1065
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   2
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
      Begin VB.CommandButton cmdPrintAll 
         Caption         =   "&Print All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1740
         TabIndex        =   17
         Top             =   120
         Width           =   1275
      End
      Begin VB.CommandButton cmdRecal 
         Caption         =   "&Recalculate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   16
         Top             =   120
         Width           =   1275
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
      Bindings        =   "frmPosControl.frx":0000
      Height          =   2115
      Left            =   120
      OleObjectBlob   =   "frmPosControl.frx":0014
      TabIndex        =   1
      Tag             =   "Skills Lookup"
      Top             =   600
      Width           =   9675
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee #"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Top             =   3720
      Width           =   1560
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CCAC Position #"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Width           =   1560
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
      TabIndex        =   11
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
      DataField       =   "PC_JOB"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   10
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
      DataField       =   "PC_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1080
      TabIndex        =   9
      Top             =   5640
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "frmPosControl"
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



Public Sub cmdDelete_Click()
Dim a As Integer, Msg As String, INo&

If Not gSec_Upd_Job_Master Then
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

If Not gSec_Upd_Job_Master Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

'Call ST_UPD_MODE(True)
Call SET_UP_MODE
On Error GoTo Edit_Err


fglbEditMode% = True


txtPosCtrl.SetFocus

Exit Sub

Edit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdModify", "HRJOBSKL", "Edit")
Call RollBack   '15June99 js
End Sub

Public Sub cmdNew_Click()
Dim SQLQ As String

If Not gSec_Upd_Job_Master Then
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

txtPosCtrl.SetFocus
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

If Not chkPosCtrl() Then Exit Sub

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

    Me.vbxCrystal.SelectionFormula = "{HR_JOB_CONTROL.PC_JOB}='" & glbPos & "'"
    Me.vbxCrystal.Destination = crptToPrinter
    
    Me.vbxCrystal.WindowTitle = "CCAC Positions"
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgPosNbr.rpt"
    Me.vbxCrystal.Connect = RptODBC_SQL
    Me.vbxCrystal.Destination = 1
    Me.vbxCrystal.Action = 1

End Sub
Public Sub cmdView_Click()

    Me.vbxCrystal.SelectionFormula = "{HR_JOB_CONTROL.PC_JOB}='" & glbPos & "'"
    Me.vbxCrystal.Destination = crptToWindow
    
    Me.vbxCrystal.WindowTitle = "CCAC Positions"
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgPosNbr.rpt"
    Me.vbxCrystal.Connect = RptODBC_SQL
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrintAll_Click()
 '   Me.vbxCrystal.SelectionFormula = glbstrSelCri
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    Me.vbxCrystal.SelectionFormula = ""
    Me.vbxCrystal.WindowTitle = "CCAC Positions"
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgPosNbr.rpt"
    Me.vbxCrystal.Connect = RptODBC_SQL
    Me.vbxCrystal.Action = 1
End Sub


Private Sub cmdRecal_Click()
Dim SQLQ
Dim rsPOS As New ADODB.Recordset
Do Until Data1.Recordset.EOF
    SQLQ = "SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 "
    SQLQ = SQLQ & " AND JH_POSITION_CONTROL='" & Data1.Recordset("PC_POSITION_CONTROL") & "'"
    rsPOS.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsPOS.EOF Then
        Data1.Recordset("PC_EMPNBR") = rsPOS("JH_EMPNBR")
    Else
        Data1.Recordset("PC_EMPNBR") = Null
    End If
    Data1.Recordset.Update
    Data1.Recordset.MoveNext
    rsPOS.Close
Loop
Data1.Refresh
MsgBox "     Recalculate is Finished.     "
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMPOSCONTROL"

End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim x%

On Error GoTo FLErr
glbOnTop = "FRMPOSCONTROL"

Screen.MousePointer = HOURGLASS
If glbPos = "" Then frmJOBS.Show 1
If glbPos = "" Then glbUserUploadMode = UploadFormWithoutCheck: Unload Me: Exit Sub

lblPosition.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$
Me.Caption = "CCAC Position - " & lblPosition

Data1.ConnectionString = glbAdoIHRDB

If Not EERetrieve() Then
    Exit Sub        '  modGet it sets fglbRecords
End If

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
SQLQ$ = "SELECT * FROM HR_JOB_CONTROL "
SQLQ$ = SQLQ$ & "WHERE PC_JOB = '" & glbPos$ & "' "
SQLQ$ = SQLQ$ & "ORDER BY PC_POSITION_CONTROL"

Data1.RecordSource = SQLQ$
Data1.Refresh

lblPosition.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    fglbRecords% = False
Else
    fglbRecords% = True
End If
EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERetrieveErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CCAC Position", "HR_JOB_CONTROL", "SELECT")
Call RollBack   '15June99 js

End Function



Private Sub lblPositions_Change()
lblPosition.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$
End Sub


Private Function chkPosCtrl()
Dim SQLQ As String, Msg As String, dd#, PID&, xPosCtrl, xDeptno$, xGLNO$

chkPosCtrl = False

On Error GoTo chkPosCtrl_Err

If Len(txtPosCtrl) < 1 Then
    MsgBox "CCAC Position # is a required field"
    txtPosCtrl.SetFocus
    Exit Function
End If

If IsNull(rsDATA("PC_ID")) Then
    PID& = 0
Else
    PID& = rsDATA("PC_ID") ' CLng(Val(lblID))
End If
xPosCtrl = txtPosCtrl

If modISDupPosCtrl(glbPos$, xPosCtrl, PID&) Then
    MsgBox "[Position Code] + [CCAC Position #] must be unique"
    txtPosCtrl.SetFocus
    Exit Function
End If

chkPosCtrl = True

Exit Function

chkPosCtrl_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HRJOBSKL", "edit/Add")
Call RollBack   '15June99 js

End Function
Private Function modISDupPosCtrl(Pos$, xPosCtrl, ID&)
Dim SQLQ$
Dim snapPosCtrl As New ADODB.Recordset

modISDupPosCtrl = True

On Error GoTo modISDupPosCtrl_Err

Screen.MousePointer = HOURGLASS

SQLQ$ = "SELECT * FROM HR_JOB_CONTROL "
SQLQ$ = SQLQ$ & "WHERE "
SQLQ$ = SQLQ$ & " (PC_JOB = '" & Pos$ & "' "
SQLQ$ = SQLQ$ & "AND PC_POSITION_CONTROL = '" & xPosCtrl & "' "
SQLQ$ = SQLQ$ & "AND PC_ID <> " & ID& & ") "
If snapPosCtrl.State <> 0 Then snapPosCtrl.Close
snapPosCtrl.Open SQLQ$, gdbAdoIhr001, adOpenStatic

If snapPosCtrl.BOF And snapPosCtrl.EOF Then
    modISDupPosCtrl = False
End If

Screen.MousePointer = DEFAULT
snapPosCtrl.Close

Exit Function

modISDupPosCtrl_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Code Snap", "TABL", "SELECT")
Call RollBack   '15June99 js

End Function

Public Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HR_JOB_CONTROL "
    SQLQ = SQLQ & "WHERE PC_ID = " & Data1.Recordset!PC_ID
    SQLQ = SQLQ & " ORDER BY PC_JOB"

    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    lblID = rsDATA!PC_ID
    Call Set_Control("R", Me, rsDATA)
    
    Call SET_UP_MODE
End Sub

Private Sub txtPosCtrl_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        ' out or left join query not updateable - so do straight.
        SQLQ$ = "SELECT * FROM HR_JOB_CONTROL "
        SQLQ$ = SQLQ$ & "WHERE PC_JOB = '" & glbPos$ & "' "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$, x%
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
UpdateRight = gSec_Upd_Job_Master
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

txtPosCtrl.Enabled = TF

End Sub

