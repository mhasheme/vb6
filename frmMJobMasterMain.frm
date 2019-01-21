VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmMJobMasterMain 
   Caption         =   "Job Master"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   10500
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkHideInactive 
      Caption         =   "Hide Inactive Positions"
      Height          =   315
      Left            =   7800
      TabIndex        =   26
      Top             =   2520
      Value           =   1  'Checked
      Width           =   2355
   End
   Begin VB.TextBox medUserDef2 
      Appearance      =   0  'Flat
      DataField       =   "JB_USERDEF2"
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Tag             =   "00-Job User Defined 2"
      Top             =   3960
      Width           =   1545
   End
   Begin VB.TextBox txtUserDef1 
      Appearance      =   0  'Flat
      DataField       =   "JB_USERDEF1"
      Height          =   285
      Left            =   2280
      MaxLength       =   25
      TabIndex        =   4
      Tag             =   "00-Job User Defined 1"
      Top             =   3600
      Width           =   1545
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JB_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   4080
      MaxLength       =   25
      TabIndex        =   14
      Text            =   "Ldate"
      Top             =   5445
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JB_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   5760
      MaxLength       =   25
      TabIndex        =   13
      Text            =   "LTime"
      Top             =   5445
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JB_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   7440
      MaxLength       =   25
      TabIndex        =   12
      Text            =   "LUser"
      Top             =   5445
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox txtJob 
      Appearance      =   0  'Flat
      DataField       =   "JB_JOBCODE"
      Height          =   285
      Left            =   2280
      MaxLength       =   25
      TabIndex        =   0
      Tag             =   "01-Job Code"
      Top             =   2520
      Width           =   1545
   End
   Begin VB.TextBox txtJobDescr 
      Appearance      =   0  'Flat
      DataField       =   "JB_JOBDESCR"
      Height          =   285
      Left            =   3960
      MaxLength       =   100
      TabIndex        =   1
      Tag             =   "01-Job Description"
      Top             =   2520
      Width           =   3495
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmMJobMasterMain.frx":0000
      Height          =   2295
      Left            =   0
      OleObjectBlob   =   "frmMJobMasterMain.frx":0014
      TabIndex        =   7
      Top             =   120
      Width           =   10275
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "JB_GRPCD"
      Height          =   285
      Index           =   2
      Left            =   1965
      TabIndex        =   2
      Tag             =   "01-Job Group Code "
      Top             =   2880
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "JB_STATUS"
      Height          =   285
      Index           =   1
      Left            =   1965
      TabIndex        =   3
      Tag             =   "01-Job Status - Code "
      Top             =   3240
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBST"
      MaxLength       =   6
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   8805
      TabIndex        =   6
      Tag             =   "01-Job Level Code "
      Top             =   6480
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBLE"
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7320
      Top             =   5805
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
      TabIndex        =   15
      Top             =   7155
      Visible         =   0   'False
      Width           =   10500
      _Version        =   65536
      _ExtentX        =   18521
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
         Left            =   15
         TabIndex        =   23
         Tag             =   "Select this Division"
         Top             =   105
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   855
         TabIndex        =   22
         Tag             =   "Close and exit this screen"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1680
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
         Tag             =   "Cancel changes made"
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   4260
         TabIndex        =   18
         Tag             =   "Create a new Division"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5070
         TabIndex        =   17
         Tag             =   "Delete Division listed"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5895
         TabIndex        =   16
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
         ReportSource    =   1
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.Label lblUserDef2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Job User Defined 2"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   3990
      Width           =   1725
   End
   Begin VB.Label lblUserDef1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Job User Defined 1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   3630
      Width           =   1725
   End
   Begin VB.Label lblJob 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Job Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   2550
      Width           =   675
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Job Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblGroup 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Job Group"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblLevel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Job Level"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7920
      TabIndex        =   8
      Top             =   6480
      Visible         =   0   'False
      Width           =   1170
   End
End
Attribute VB_Name = "frmMJobMasterMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRSOld As String, glbEmptyNew  As Integer
'Dim fglbNewRec% ' new record
Dim RSDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim Ctrl As Control 'Sam add July 2002 * Remove ADO
Dim FRS As ADODB.Recordset
Dim fglbNew As Boolean

Private Sub chkHideInactive_Click()
Dim SQLQ As String
SQLQ = "SELECT * FROM HRJOBMASTER WHERE (1=1) "
If chkHideInactive Then
    SQLQ = SQLQ & " AND JB_STATUS<>'INAC'"
    SQLQ = SQLQ & " AND UPPER(LEFT(JB_JOBDESCR, 2)) <> 'Z ' "
End If
SQLQ = SQLQ & "ORDER BY JB_JOBDESCR"
Data1.RecordSource = SQLQ
Data1.LockType = adLockReadOnly
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

If glbJobMaster <> "" Then
    Data1.Recordset.Find "JB_JOBCODE='" & glbJobMaster & "' "
    If Data1.Recordset.RecordCount > 7 Then
        vbxTrueGrid.ScrollBars = vbVertical
    End If
End If
End Sub

Public Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err

RSDATA.CancelUpdate
Call Set_Control("R", Me, RSDATA)


Call modSTUPD(False)  ' reset screen's attributes

fglbNew = False
'cmdClose.SetFocus


Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRJOBMASTER", "Cancel")
Resume Next
End Sub

Public Sub cmdClose_Click()
    'glbDiv = ""
    'glbDivDesc = ""
    fglbNew = False
    
    Unload Me
End Sub

Public Sub cmdDelete_Click()
Dim Div As String, SQLQ As String, Msg$, a%
Dim snapEEDivs As New ADODB.Recordset

On Error GoTo DelErr

If Len(txtJob) < 1 Then Exit Sub

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

gdbAdoIhr001.BeginTrans
RSDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

fglbNew = False
Call SET_UP_MODE

Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRJOBMASTER", "Delete")
Resume Next

End Sub


Public Sub cmdModify_Click()
On Error GoTo Mod_Err

Call modSTUPD(True)
txtJob.Enabled = False
txtJobDescr.Enabled = True
txtJobDescr.SetFocus
fglbNew = False

Exit Sub
Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HRJOBMASTER", "Modify")
Call RollBack '08June99 js

End Sub

Public Sub cmdNew_Click()
Dim SQLQ As String
glbCodeRef = True

On Error GoTo NewErr

'Call modSTUPD(True)
fglbNew = True
Call SET_UP_MODE

Call Set_Control("B", Me)

If RSDATA.State = 0 Then
    If Data1.Recordset.EOF Then
        SQLQ = "SELECT * FROM HRJOBMASTER "
    Else
        SQLQ = "SELECT * FROM HRJOBMASTER WHERE JB_JOBCODE = '" & Data1.Recordset!JB_JOBCODE & "' " '& " order by Division_Name"
    End If
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

RSDATA.AddNew

txtJob.Enabled = True
txtJob.SetFocus


Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HRJOBMASTER", "AddNew")
Resume Next
End Sub

Public Sub cmdOK_Click()
Dim JobCode, ctylist
Dim xMsg As String
Dim SQLQ As String

On Error GoTo OK_Err

If Not chkJobMaster() Then Exit Sub

Call UpdUStats(Me)
JobCode = txtJob.Text

Call Set_Control("U", Me, RSDATA)

gdbAdoIhr001.BeginTrans
RSDATA.Update
gdbAdoIhr001.CommitTrans

If Not (glbWFC And glbPlantCode = "GREN") Then   'Greensboro
    Call Codes_Master_Integration("POSITION", JobCode)
End If

SQLQ = "SELECT * FROM HRJOBMASTER WHERE (1=1) "
If chkHideInactive Then
    SQLQ = SQLQ & " AND JB_STATUS<>'INAC'"
    SQLQ = SQLQ & " AND UPPER(LEFT(JB_JOBDESCR, 2)) <> 'Z ' "
End If
SQLQ = SQLQ & "ORDER BY JB_JOBDESCR"

Data1.RecordSource = SQLQ '"SELECT * FROM HRJOBMASTER ORDER BY JB_JOBDESCR "
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Data1.Recordset.Find "JB_JOBCODE = '" & JobCode & "' "

fglbNew = False
Call SET_UP_MODE
'Call modSTUPD(False)

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRJOBMASTER", "Update")
Resume Next
Unload Me

End Sub

Public Sub cmdView_Click()
    Call cmdPrint_Click
End Sub

Public Sub cmdPrint_Click()
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

Me.vbxCrystal.ReportTitle = "All Table Codes"
Me.vbxCrystal.BoundReportHeading = Me.Caption
Me.vbxCrystal.WindowTitle = Me.Caption & " Report"
Me.vbxCrystal.Action = 1

End Sub

Public Sub cmdSelect_Click()
glbJobMaster = Data1.Recordset("JB_JOBCODE")
glbJobMasterDesc = Data1.Recordset("JB_JOBDESCR")
Unload Me

End Sub

Private Sub Form_Activate()
'Data1.RecordSource = "SELECT * FROM HRJOBMASTER ORDER BY JB_JOBDESCR "
'Data1.Refresh
'
'Set FRS = Data1.Recordset.Clone
'vbxTrueGrid.FetchRowStyle = True
glbOnTop = "frmMJobMasterMain"
End Sub

Private Sub Form_GotFocus()
glbOnTop = "frmMJobMasterMain"
End Sub

Private Sub Form_Load()
Dim SQLQ, I, ctylist, x

glbOnTop = "frmMJobMasterMain"


If Not EERetrieve() Then Exit Sub

If glbJobMaster <> "" Then
    Data1.Recordset.Find "JB_JOBCODE='" & glbJobMaster & "' "
    If Data1.Recordset.RecordCount > 7 Then
        vbxTrueGrid.ScrollBars = vbVertical
    End If
End If

Screen.MousePointer = HOURGLASS

'Me.vbxTrueGrid.Refresh

Screen.MousePointer = DEFAULT

Call modSTUPD(False)

If Not gSec_Upd_Job_Master Then     '
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
End If                          '

Call WFCScreenSetup
    
If Not gSec_Upd_Job_Master Then
    Call set_Buttons
End If

Call INI_Controls(Me)

End Sub

Private Sub WFCScreenSetup() 'Ticket #25911 Franks 09/30/2014
    lblGroup.Caption = lStr("Job Group")
    'lblLevel.Caption = lStr("Job Level")
    lblStatus.Caption = lStr("Job Status")
    lblUserDef1.Caption = lStr("Job User Defined 1")
    lblUserDef2.Caption = lStr("Job User Defined 2")
    vbxTrueGrid.Columns(2).Caption = lStr("Job Group")
    'vbxTrueGrid.Columns(3).Caption = lStr("Job Level")
    vbxTrueGrid.Columns(3).Caption = lStr("Job Status")
    vbxTrueGrid.Columns(4).Caption = lStr("Job User Defined 1")
    vbxTrueGrid.Columns(5).Caption = lStr("Job User Defined 2")
    
    clpCode(2).TABLTitle = UCase(lStr("Job Group")) & " CODES"
    'clpCode(3).TABLTitle = UCase(lStr("Job Level")) & " CODES"
    clpCode(1).TABLTitle = UCase(lStr("Job Status")) & " CODES"
End Sub
Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
If fglbNew Then
    UpdateState = NewRecord
    TF = True
ElseIf RSDATA.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)

If Not UpdateRight Then TF = False
Call modSTUPD(TF)

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

cmdModify.Enabled = FT          '
cmdDelete.Enabled = FT          '
cmdNew.Enabled = FT             '
cmdCancel.Enabled = TF          '
cmdOK.Enabled = TF              '
txtJob.Enabled = TF
txtJobDescr.Enabled = TF
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
txtUserDef1.Enabled = TF
medUserDef2.Enabled = TF
cmdClose.Enabled = FT           '
cmdSelect.Enabled = FT          '
cmdPrint.Enabled = FT           '
        
If glbJobMasterInhSel% Then
    cmdSelect.Enabled = False
End If
End Sub

Public Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        Exit Sub
    End If
  
    SQLQ = "select * from HRJOBMASTER WHERE JB_JOBCODE = '" & Data1.Recordset!JB_JOBCODE & "' " '& " order by Division_Name"
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If RSDATA.EOF Or RSDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, RSDATA)
    
    Call SET_UP_MODE
    
End Sub



Private Sub medUserDef2_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtJob_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtJob_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtJobDescr_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtUserDef1_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_DblClick()
If cmdSelect.Enabled Then
    If Not Me.vbxTrueGrid.EditActive Then
        glbJobMaster = Data1.Recordset("JB_JOBCODE")
        glbJobMasterDesc = Data1.Recordset("JB_JOBDESCR")
        'Unload Me
    Else
        MsgBox "Save/cancel changes first"
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

    SQLQ = "SELECT * FROM HRJOBMASTER WHERE (1=1) "
    If chkHideInactive Then
        SQLQ = SQLQ & " AND JB_STATUS<>'INAC'"
        SQLQ = SQLQ & " AND UPPER(LEFT(JB_JOBDESCR, 2)) <> 'Z ' "
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh

    Set FRS = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Data1.Recordset.EOF Then
        glbJobMaster = ""
        glbJobMasterDesc = ""
    Else
        glbJobMaster = Data1.Recordset("JB_JOBCODE")
        glbJobMasterDesc = Data1.Recordset("JB_JOBDESCR")
    End If
        
    Call Display_Value
End Sub

Private Function chkJobMaster()
Dim Div As String, SQLQ As String, Msg$
Dim snapDivs As New ADODB.Recordset
Dim x
chkJobMaster = False
On Error GoTo chkJobMaster_Err

If Len(txtJob) < 1 Then
    MsgBox ("Job Code is a required field")
    txtJob.SetFocus
    Exit Function
End If

If Len(txtJobDescr) < 1 Then
    MsgBox lStr("Job Description is a required field")
    txtJobDescr.SetFocus
    Exit Function
End If

If fglbNew Then
    SQLQ = "SELECT * FROM HRJOBMASTER "
    SQLQ = SQLQ & "WHERE JB_JOBCODE = '" & txtJob.Text & "'"
    
    If snapDivs.State <> 0 Then snapDivs.Close
    snapDivs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If snapDivs.BOF And snapDivs.EOF Then
        snapDivs.Close
    Else
        Msg$ = ("This Job Code already exists")
        MsgBox Msg$
        snapDivs.Close
        Exit Function
    End If
End If

If Len(clpCode(2).Text) = 0 Then
    MsgBox ("Group is a required field")
    clpCode(2).SetFocus
    Exit Function
End If
'If Len(clpCode(3).Text) = 0 Then
'    MsgBox ("Level is a required field")
'    clpCode(3).SetFocus
'    Exit Function
'End If
If Len(clpCode(1).Text) = 0 Then
    MsgBox ("Status is a required field")
    clpCode(1).SetFocus
    Exit Function
End If

For x = 1 To 3
    If Len(clpCode(x).Text) > 0 And clpCode(x).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(x).SetFocus
        Exit Function
    End If
Next x

chkJobMaster = True

Exit Function

chkJobMaster_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkJobMaster", "HRJOBMASTER", "Cancel")
Resume Next

End Function


Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateJobMaster  'RelatePos
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


Public Function EERetrieve()

Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError

Data1.ConnectionString = glbAdoIHRDB

SQLQ = "SELECT * FROM HRJOBMASTER WHERE (1=1) "
If chkHideInactive Then
    SQLQ = SQLQ & " AND JB_STATUS<>'INAC'"
    SQLQ = SQLQ & " AND UPPER(LEFT(JB_JOBDESCR, 2)) <> 'Z ' "
End If
SQLQ = SQLQ & "ORDER BY JB_JOBDESCR"
Data1.RecordSource = SQLQ
Data1.LockType = adLockReadOnly
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True


If glbJobMaster <> "" Then
    Data1.Recordset.Find "JB_JOBCODE='" & glbJobMaster & "' "
    If Data1.Recordset.RecordCount > 7 Then
        vbxTrueGrid.ScrollBars = vbVertical
    End If
End If


'''Data1.RecordSource = "SELECT * FROM HRJOB ORDER BY JB_DESCR"
''SQLQ = SQLQ & "SELECT * FROM HRJOB WHERE 1 = 1"
''If Len(glbWFCUserSecList) > 0 Then 'Ticket #27609 Franks 10/13/2015
''    SQLQ = SQLQ & " AND JB_SECTION IN " & glbWFCUserSecList & " "
''End If
''If chkHideInactive Then
''    SQLQ = SQLQ & " AND JB_STATUS<>'INAC'"
''    If glbOracle Then 'Ticket #16416
''        SQLQ = SQLQ & " AND UPPER(SUBSTR(JB_DESCR,1,2)) <> 'Z '"
''    ElseIf glbSQL Then
''        SQLQ = SQLQ & " AND UPPER(LEFT(JB_DESCR, 2)) <> 'Z '"
''    Else
''        SQLQ = SQLQ & " AND UCASE(LEFT(JB_DESCR, 2)) <> 'Z '"
''    End If
''End If
''SQLQ = SQLQ & "ORDER BY JB_DESCR"
''
''Data1.RecordSource = SQLQ '"SELECT * FROM HRJOB ORDER BY JB_DESCR"
''Data1.Refresh
''
''If glbPos <> "" Then
''    Data1.Recordset.Find "JB_CODE='" & glbPos & "' "
''    If Data1.Recordset.RecordCount > 7 Then
''        vbxTrueGrid.ScrollBars = vbVertical
''    End If
''End If

EERetrieve = True

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HRJOBMASTER", "SELECT")
Resume Next

Exit Function

End Function

