VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEPositionBK 
   Caption         =   "Backup Positions"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JH_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   8520
      MaxLength       =   25
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JH_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   6840
      MaxLength       =   25
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JH_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   5040
      MaxLength       =   25
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feposbk.frx":0000
      Height          =   1425
      Left            =   120
      OleObjectBlob   =   "feposbk.frx":0014
      TabIndex        =   0
      Top             =   480
      Width           =   9015
   End
   Begin INFOHR_Controls.DateLookup dlpEDate 
      DataField       =   "JH_SDATE"
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Tag             =   "40-Date when course is to be renewed"
      Top             =   2640
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
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
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2850
         TabIndex        =   12
         Top             =   120
         Width           =   585
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1320
         TabIndex        =   11
         Top             =   135
         Width           =   1005
      End
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
         TabIndex        =   10
         Top             =   160
         Width           =   1005
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   13
      Top             =   7365
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
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
      Begin VB.CommandButton cmdPosition 
         Appearance      =   0  'Flat
         Caption         =   "&Position"
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
         Left            =   120
         TabIndex        =   14
         Tag             =   "Load Beneficiary screen"
         Top             =   120
         Width           =   1485
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   7140
         Top             =   165
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   "c:\newihr\rgedsem.rpt"
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   9360
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         ConnectMode     =   3
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
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      DataField       =   "JH_JOB"
      Height          =   285
      Left            =   1560
      TabIndex        =   15
      Tag             =   "01-Position code"
      Top             =   2160
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   6
      LookupType      =   5
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "JH_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3480
      TabIndex        =   8
      Top             =   6840
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "JH_EMPNBR"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4290
      TabIndex        =   7
      Top             =   6840
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Effective Date"
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
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1365
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Code"
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
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1560
   End
End
Attribute VB_Name = "frmEPositionBK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim fglbNew  As Integer
Dim fglHredsem As String       '
Dim fglCursName As String      '
Dim fglExtName As String       '
Dim rsDATA As New ADODB.Recordset

Private Sub cmdPosition_Click()
Unload Me
Load frmEPOSITION
End Sub


Private Sub dlpEDate_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
glbOnTop = "frmEPositionBK"
Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
glbOnTop = "frmEPositionBK"
End Sub

Private Sub Form_Load()
Dim x%
    
    Screen.MousePointer = DEFAULT
     
    glbOnTop = "frmEPositionBK"
    
    If glbtermopen Then
        Data1.ConnectionString = glbAdoIHRAUDIT
    Else
        Data1.ConnectionString = glbAdoIHRDB
    End If
    
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
    
    
    Screen.MousePointer = HOURGLASS
    
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
        frmEPositionBK.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    lblEENum.Caption = ShowEmpnbr(lblEEID)
    Call Display_Value
    Call ST_UPD_MODE(True)

    Call INI_Controls(Me)
    For x% = 1 To 15
        Call setCaption(lblTitle(x%))
    Next

    Screen.MousePointer = DEFAULT
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub


Private Sub lblEEID_Change()
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    frmEPositionBK.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblEENum = ShowEmpnbr(lblEEID)
End Sub


Private Sub medPPE_Change()

End Sub

Private Sub medPPE_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Function EERetrieve()
Dim SQLQ As String
EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS

If glbtermopen Then
    SQLQ = "Select * from Term_JOB_BACKUP"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY JH_SDATE DESC"
Else
    SQLQ = "Select * from HR_JOB_BACKUP"
    SQLQ = SQLQ & " where JH_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY JH_SDATE DESC"
End If


Data1.RecordSource = SQLQ
Data1.Refresh

EERetrieve = True

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DEPRetrieve", "HR_JOB_BACKUP", "SELECT")
Call RollBack '23July99 js

Exit Function

End Function

Sub Display_Value()
Dim SQLQ
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    If glbtermopen Then
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
Else
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    If glbtermopen Then
        SQLQ = "Select * from Term_JOB_BACKUP"
        SQLQ = SQLQ & " WHERE JH_ID = " & Data1.Recordset!JH_ID
        SQLQ = SQLQ & " ORDER BY JH_SDATE DESC"
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "Select * "
        SQLQ = SQLQ & " from HR_JOB_BACKUP "
        SQLQ = SQLQ & " where JH_ID = " & Data1.Recordset!JH_ID
        SQLQ = SQLQ & " ORDER BY JH_SDATE DESC"
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
End If
Call SET_UP_MODE
'Me.cmdModify_Click
End Sub

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
Call ST_UPD_MODE(TF)
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

fUPMode = TF    ' update mode

dlpEDate.Enabled = TF
clpJob.Enabled = TF

End Sub

Sub cmdCancel_Click()
Dim I%
Dim x
On Error GoTo Can_Err

fglbNew = False
rsDATA.CancelUpdate
Call Display_Value

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_JOB_BACKUP", "Cancel")
Call RollBack '23July99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "frmEPositionBK" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, x

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub


fglHredsem = dlpEDate.Text

If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001X.CommitTrans
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001.CommitTrans
    Data1.Refresh
End If
If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(True)


Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HREDSEM", "Delete")
Call RollBack '23July99 js

End Sub

Sub cmdNew_Click()
Dim x%
fglbNew = True

'Call ST_UPD_MODE(True)
On Error GoTo AddN_Err

Call Set_Control("B", Me)
rsDATA.AddNew
Call SET_UP_MODE

If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_JOB_BACKUP", "Add")
Call RollBack '23July99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
'Dim xChange1, xChange2
Dim x
Dim xID As Long

On Error GoTo Add_Err


If Not chkPosition() Then Exit Sub

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    Call UpdUStats(Me) ' update user's stats (who did it and when)
    Call Set_Control("U", Me, rsDATA)
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    xID = rsDATA("JH_ID")
    Data1.Refresh
  Else
    fglHredsem = dlpEDate.Text
    Call UpdUStats(Me) ' update user's stats (who did it and when)
    Call Set_Control("U", Me, rsDATA)
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    xID = rsDATA("JH_ID")
    Data1.Refresh
End If


fglbNew = False
Call SET_UP_MODE

Exit Sub
Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_JOB_BACKUP", "Update")
Call RollBack '23July99 js

End Sub

Private Function chkPosition()
Dim oCode As String, OCodeD As String

chkPosition = False
If Len(clpJob.Text) <= 0 Then
    MsgBox "Position Code is required"
    clpJob.SetFocus
    Exit Function
Else
    If clpJob.Caption = "Unassigned" Then
        MsgBox "Position Code is required"
        clpJob.SetFocus
        Exit Function
    End If
End If

If Len(dlpEDate.Text) = 0 Then
    MsgBox "Effective Date is required"
    dlpEDate.SetFocus
    Exit Function
Else
    If Not IsDate(dlpEDate.Text) Then
        MsgBox ("Effective Date is invalid")
        dlpEDate.SetFocus
        Exit Function
    End If
End If

If ChkDup() Then
        Exit Function
End If

If ChkCurrentPos() Then
        Exit Function
End If

chkPosition = True

End Function
Private Function ChkCurrentPos()
Dim SQLQ, Logx, Msg$, SavReviewDate
Dim Response%
Dim rsTB As New ADODB.Recordset
Dim Title$, Msg1$, DgDef
Dim xESID As Integer

ChkCurrentPos = False

Logx = False
If glbtermopen Then
    SQLQ = "SELECT JH_EMPNBR FROM Term_JOB_HISTORY WHERE JH_CURRENT<>0 AND TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR = " & glbLEE_ID
End If

If clpJob.Text <> "" Then
    SQLQ = SQLQ & " AND JH_JOB = '" & clpJob.Text & "'"
End If
If glbtermopen Then
    rsTB.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockReadOnly
Else
    rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockReadOnly
End If
If Not rsTB.EOF Then Logx = True
rsTB.Close
If Logx = True Then
    Msg$ = "   This Code has been assigned to current position!  "
    MsgBox Msg$
    ChkCurrentPos = True
End If
End Function
Private Function ChkDup()
Dim SQLQ, Logx, Msg$, SavReviewDate
Dim Response%
Dim rsTB As New ADODB.Recordset
Dim Title$, Msg1$, DgDef
Dim xESID As Integer

ChkDup = False

Logx = False
If glbtermopen Then
    SQLQ = "SELECT JH_EMPNBR FROM Term_JOB_BACKUP WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "SELECT JH_EMPNBR FROM HR_JOB_BACKUP WHERE JH_EMPNBR = " & glbLEE_ID
End If

If clpJob.Text <> "" Then
    SQLQ = SQLQ & " AND JH_JOB = '" & clpJob.Text & "'"
End If
If Len(dlpEDate.Text) > 0 Then
    SQLQ = SQLQ & " AND JH_SDATE = " & Date_SQL(dlpEDate.Text) & " "
End If
If rsDATA.EditMode <> adEditAdd Then SQLQ = SQLQ & " AND JH_ID<>" & Data1.Recordset("JH_ID")
If glbtermopen Then
    rsTB.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockReadOnly
Else
    rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockReadOnly
End If
If Not rsTB.EOF Then Logx = True
rsTB.Close
If Logx = True Then
    Msg$ = "     Duplicate Record!     "
    MsgBox Msg$
    ChkDup = True
End If
End Function


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
RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Position
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
Printable = False
End Property

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
End Sub
