VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmWorkFlowMaster 
   Caption         =   "Work Flow Master"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtByWhom 
      Appearance      =   0  'Flat
      DataField       =   "WK_BYWHOM"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   5640
      TabIndex        =   25
      Tag             =   "00-Employee Number of individual's supervisor"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      DataField       =   "WK_COMPNO"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7200
      MaxLength       =   3
      TabIndex        =   23
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "WK_LUSER"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   8160
      MaxLength       =   25
      TabIndex        =   22
      Text            =   "LUser"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "WK_LTIME"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   6480
      MaxLength       =   25
      TabIndex        =   21
      Text            =   "LTime"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "WK_LDATE"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   4800
      MaxLength       =   25
      TabIndex        =   20
      Text            =   "Ldate"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox txtTask 
      Appearance      =   0  'Flat
      DataField       =   "WK_TASK"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      MaxLength       =   100
      TabIndex        =   3
      Tag             =   "01-Task"
      Top             =   4080
      Width           =   3135
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   5505
      Width           =   9165
      _Version        =   65536
      _ExtentX        =   16166
      _ExtentY        =   952
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
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5895
         TabIndex        =   13
         Tag             =   "Print Division Listing"
         Top             =   105
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4350
         TabIndex        =   12
         Tag             =   "Delete Division listed"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3540
         TabIndex        =   11
         Tag             =   "Create a new Division"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Tag             =   "Cancel changes made"
         Top             =   105
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Tag             =   "Save changes made"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Tag             =   "Edit the information "
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   135
         TabIndex        =   7
         Tag             =   "Close and exit this screen"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
         Height          =   375
         Left            =   6720
         TabIndex        =   6
         Tag             =   "Select this Division"
         Top             =   105
         Visible         =   0   'False
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
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmWorkFlowMaster.frx":0000
      Height          =   3075
      Left            =   120
      OleObjectBlob   =   "frmWorkFlowMaster.frx":0014
      TabIndex        =   14
      Tag             =   "Division Listings"
      Top             =   120
      Width           =   8865
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "WK_WORKFLOW"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   1850
      TabIndex        =   1
      Tag             =   "01-Work Flow Type Code"
      Top             =   3360
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WKFL"
   End
   Begin MSMask.MaskEdBox medStep 
      DataField       =   "WK_STEP"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Tag             =   "00-Step"
      Top             =   3720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###0;(###)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medTarget 
      DataField       =   "WK_TARGET"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Tag             =   "00-Target"
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###0;(###)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.EmployeeLookup elpByWhom 
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1850
      TabIndex        =   5
      Tag             =   "10-Employee Number of individual's supervisor"
      Top             =   4800
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   9120
      Top             =   5280
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
   Begin VB.Label lblDays 
      Caption         =   "Days"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   4460
      Width           =   1335
   End
   Begin VB.Label lblTarget 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Target"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   4440
      Width           =   465
   End
   Begin VB.Label lblByWhom 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "By Whom"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   4800
      Width           =   690
   End
   Begin VB.Label lblStep 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   330
   End
   Begin VB.Label lbType 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Work Flow Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   1170
   End
   Begin VB.Label lblTask 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Task"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   360
   End
End
Attribute VB_Name = "frmWorkFlowMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRSOld As String, glbEmptyNew  As Integer
Dim fglbNewRec% ' new record
Dim rsDATA As New ADODB.Recordset '
Dim Ctrl As Control

Private Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err


Data1.Recordset.CancelBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
'Data1.Refresh
If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    Data1.Refresh
End If

Call modSTUPD(False)  ' reset screen's attributes
cmdClose.SetFocus


Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRPROv", "Cancel")
Resume Next

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim Div As String, SQLQ As String, Msg$, a%
Dim snapEEDivs As New ADODB.Recordset

On Error GoTo DelErr

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub


gdbAdoIhr001.BeginTrans
Data1.Recordset.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End If

Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRWORKFLOW_MASTER", "Delete")
Resume Next
End Sub

Private Sub cmdModify_Click()
On Error GoTo Mod_Err

Call modSTUPD(True)

clpCode(0).SetFocus
fglbNewRec% = False

Exit Sub
Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '08June99 js
End Sub

Private Sub cmdNew_Click()

glbCodeRef = True

On Error GoTo NewErr

Call modSTUPD(True)

fglbNewRec% = True

'rsDATA.AddNew
Data1.Recordset.AddNew
txtComp.Text = glbCompNo
'Call Set_Control("B", Me)

clpCode(0).Enabled = True
'clpCode(0).SetFocus


Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HRWORKFLOW_MASTER", "AddNew")
End Sub

Private Sub cmdOK_Click()
Dim TypeCode, ctylist
Dim fglbID As Integer
On Error GoTo OK_Err

If Not chkWorkFlow() Then Exit Sub

'Call UpdUStats(Me)
'TypeCode = clpCode(0).Text
'
'Call Set_Control("U", Me, rsDATA)
'
'gdbAdoIhr001.BeginTrans
'rsDATA.Update
'gdbAdoIhr001.CommitTrans
'
'Data1.Refresh
'Data1.Recordset.Find "WK_WORKFLOW='" & TypeCode & " '"

Data1.Recordset("WK_COMPNO") = "001"
Data1.Recordset("WK_LDATE") = Date
Data1.Recordset("WK_LTIME") = Time$
Data1.Recordset("WK_LUSER") = glbUserID
Data1.Recordset.UpdateBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    fglbID = Data1.Recordset("WK_ID")
    Data1.Refresh
    Data1.Recordset.Find "WK_ID=" & fglbID & " "
End If


fglbNewRec% = False
Call modSTUPD(False)

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRWORKFLOW_MASTER", "Update")
Resume Next
Unload Me


End Sub

Private Function chkWorkFlow()
Dim Div As String, SQLQ As String, Msg$
Dim snapWL As New ADODB.Recordset
Dim X
chkWorkFlow = False
On Error GoTo chkWorkFlow_Err

If Len(clpCode(0).Text) < 1 Then
    MsgBox ("Type of Work Flow is a required field")
    clpCode(0).SetFocus
    Exit Function
End If

If Len(medStep) < 1 Then
    MsgBox ("Step is a required field")
    medStep.SetFocus
    Exit Function
End If
If Not IsNumeric(medStep) Then
    MsgBox ("Step must be numeric.")
    medStep.SetFocus
    Exit Function
End If
If Len(txtTask.Text) < 1 Then
    MsgBox ("Task is a required field")
    txtTask.SetFocus
    Exit Function
End If

If fglbNewRec% Then
    SQLQ = "SELECT * from HRWORKFLOW_MASTER "
    SQLQ = SQLQ & "WHERE WK_WORKFLOW = '" & clpCode(0).Text & "' "
    SQLQ = SQLQ & "AND WK_STEP = " & medStep.Text & " "
    
    If snapWL.State <> 0 Then snapWL.Close
    snapWL.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If snapWL.BOF And snapWL.EOF Then
        snapWL.Close
    Else
        Msg$ = lStr("This [Type + Step] already exists")
        MsgBox Msg$
        snapWL.Close
        Exit Function
    End If
End If

For X = 0 To 0
    If Len(clpCode(X).Text) > 0 And clpCode(X).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(X).SetFocus
        Exit Function
    End If
Next X

chkWorkFlow = True

Exit Function

chkWorkFlow_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkWorkFlow", "HR_Div", "Cancel")
Resume Next

End Function

Private Sub cmdPrint_Click()
'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

End Sub

Private Sub elpByWhom_Change()
txtByWhom.Text = elpByWhom.Text
End Sub

Private Sub Form_Load()
Dim SQLQ, I, ctylist, X
'glbOnTop = "frmWorkFlowMaster"
'Data1.DatabaseName = glbIHRDB
Data1.ConnectionString = glbAdoIHRDB
SQLQ = "select * from HRWORKFLOW_MASTER  order by WK_WORKFLOW,WK_STEP"
Data1.RecordSource = SQLQ
Data1.Refresh

Screen.MousePointer = HOURGLASS
Me.vbxTrueGrid.Refresh
Screen.MousePointer = DEFAULT
Call modSTUPD(False)

'Call setCaption(lblDiv)

Call INI_Controls(Me)

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

cmdModify.Enabled = FT
cmdDelete.Enabled = FT          '
cmdNew.Enabled = FT             '
cmdCancel.Enabled = TF          '
cmdOK.Enabled = TF              '
clpCode(0).Enabled = TF
medStep.Enabled = TF '
txtTask.Enabled = TF
medTarget.Enabled = TF
elpByWhom.Enabled = TF
vbxTrueGrid.Enabled = FT
cmdClose.Enabled = FT           '
cmdSelect.Enabled = FT          '
cmdPrint.Enabled = FT           '

End Sub


Private Sub medStep_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medTarget_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtByWhom_Change()
elpByWhom.Text = txtByWhom.Text
End Sub

Private Sub txtTask_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If

        SQLQ = "SELECT * FROM HRWORKFLOW_MASTER"
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

        Data1.RecordSource = SQLQ
        Data1.Refresh

End Sub
