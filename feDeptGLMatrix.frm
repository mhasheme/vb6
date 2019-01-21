VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmDeptGLMatrix 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Department / GL # Matrix"
   ClientHeight    =   6885
   ClientLeft      =   90
   ClientTop       =   1005
   ClientWidth     =   10560
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
   ScaleHeight     =   6885
   ScaleWidth      =   10560
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkInactive 
      Alignment       =   1  'Right Justify
      Caption         =   "Inactive Department / GL # Matrix"
      DataField       =   "DG_INACTIVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   8
      Tag             =   "Company Paid "
      Top             =   3840
      Value           =   1  'Checked
      Width           =   2955
   End
   Begin VB.CheckBox chkHideInactive 
      Caption         =   "Hide Inactive"
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
      Left            =   9120
      TabIndex        =   10
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1395
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feDeptGLMatrix.frx":0000
      Height          =   2535
      Left            =   240
      OleObjectBlob   =   "feDeptGLMatrix.frx":0014
      TabIndex        =   0
      Top             =   240
      Width           =   10215
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      DataField       =   "DG_DEPTNO"
      Height          =   285
      Left            =   2775
      TabIndex        =   4
      Tag             =   "00-Specific Department Desired"
      Top             =   3150
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6840
      Top             =   6120
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DG_LUSER"
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
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DG_LDATE"
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
      Left            =   6840
      MaxLength       =   12
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3465
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DG_LTIME"
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
      Left            =   7590
      MaxLength       =   8
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   645
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6240
      Top             =   6120
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
   Begin INFOHR_Controls.CodeLookup clpGLNum 
      DataField       =   "DG_GLNO"
      Height          =   285
      Left            =   2775
      TabIndex        =   7
      Tag             =   "00-General Ledger - Code"
      Top             =   3480
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   3
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "G/L #"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   12
      Left            =   330
      TabIndex        =   9
      Top             =   3525
      Width           =   525
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   330
      TabIndex        =   6
      Top             =   3150
      Width           =   990
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      DataField       =   "DG_COMPNO"
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
      Left            =   9600
      TabIndex        =   3
      Top             =   6240
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "frmDeptGLMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fGLBNew As Boolean
Dim fglbSDate As Variant
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim DefType(0 To 3)
Dim SystType(0 To 3)
Dim RSDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim UpdateState As UpdateStateEnum
Dim FRS As ADODB.Recordset

Private Function chkEPayroll()
Dim Msg As String
Dim X%, xchk
Dim rsDGMat As New ADODB.Recordset
Dim SQLQ As String

chkEPayroll = False

If Len(clpDept.Text) = 0 Then
    MsgBox "Department must be entered"
    clpDept.SetFocus
    Exit Function
End If
If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
    clpDept.SetFocus
    Exit Function
End If
If clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
    clpDept.SetFocus
    Exit Function
End If

If Len(clpGLNum.Text) = 0 Then
    MsgBox lStr("G/L # is a required field")
    clpGLNum.SetFocus
    Exit Function
End If

If Len(clpGLNum.Text) > 0 And clpGLNum.Caption = "Unassigned" Then
    MsgBox lStr("G/L # must be valid")
    clpGLNum.SetFocus
    Exit Function
End If

'Check if the duplicate Department/GL Code already exists
SQLQ = "SELECT * FROM HR_DEPTGL_MATRIX"
SQLQ = SQLQ & " WHERE DG_DEPTNO = '" & clpDept.Text & "'"
SQLQ = SQLQ & " AND DG_GLNO = '" & clpGLNum.Text & "'"
If Not fGLBNew Then
    SQLQ = SQLQ & " AND DG_ID <> " & Data1.Recordset!DG_ID
End If
SQLQ = SQLQ & " ORDER BY DG_DEPTNO,DG_GLNO "
rsDGMat.Open SQLQ$, gdbAdoIhr001, adOpenStatic
If rsDGMat.EOF Then
    rsDGMat.Close
    Set rsDGMat = Nothing
Else
    MsgBox "The " & lStr("Department") & " / " & lStr("G/L") & " Matrix" & " already exists."
    rsDGMat.Close
    Set rsDGMat = Nothing
    clpDept.SetFocus
    Exit Function
End If

xchk = False


chkEPayroll = True

End Function

Sub cmdCancel_Click()

On Error GoTo Can_Err

fGLBNew = False

If fglbEmptyNew Then
    Me.vbxTrueGrid.Enabled = True
    Me.vbxTrueGrid.Refresh
End If

'Data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh
''' Sam add July 2002 * Remove Binding Control
RSDATA.CancelUpdate
Call Display_Value


'Call ST_UPD_MODE(True) ' reset screen's attributes

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HR_DEPTGL_MATRIX", "Cancel")
Call RollBack '09June99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String
Dim X As Integer
Dim xID As Integer

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
If vbxTrueGrid.SelBookmarks.count > 1 Then
    Msg = Msg & "These Records?"
Else
    Msg = Msg & "This Record?"
End If
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

'7.9 Enhancement to delete multiple records.
If vbxTrueGrid.SelBookmarks.count = 0 Then vbxTrueGrid.SelBookmarks.Add Data1.Recordset.Bookmark
For X = 0 To vbxTrueGrid.SelBookmarks.count - 1
    Data1.Recordset.Bookmark = vbxTrueGrid.SelBookmarks(X)
    xID = Data1.Recordset("DG_ID")
    
    gdbAdoIhr001.BeginTrans
    'rsDATA.Delete
    gdbAdoIhr001.Execute "DELETE FROM HR_DEPTGL_MATRIX WHERE DG_ID=" & xID
    gdbAdoIhr001.CommitTrans
    DoEvents
Next

Data1.Refresh

Call SET_UP_MODE
'Call ST_UPD_MODE(False)

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_DEPTGL_MATRIX", "Delete")
Call RollBack '09June99 js

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Call ST_UPD_MODE(True)
clpDept.SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_DEPTGL_MATRIX", "Modify")
Call RollBack '09June99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()

On Error GoTo AddN_Err

Call Set_Control("B", Me)

RSDATA.AddNew

lblCNum.Caption = "001"

fGLBNew = True

Call SET_UP_MODE

'Call ST_UPD_MODE(True)
clpDept.SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_DEPTGL_MATRIX", "Add")
Call RollBack '09June99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim X%
Dim bmk As Variant

On Error GoTo cmdOK_Err
If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    bmk = 0
Else
    bmk = Data1.Recordset.Bookmark
End If

If Not chkEPayroll() Then Exit Sub


Call UpdUStats(Me) ' update user's stats (who did it and when)
Call Set_Control("U", Me, RSDATA)

gdbAdoIhr001.BeginTrans
RSDATA.Update
gdbAdoIhr001.CommitTrans

Data1.Refresh
If Not bmk = 0 Then
    Data1.Recordset.Bookmark = bmk
End If

'Release 8.1
Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

fGLBNew = False

Call Display_Value

Me.vbxTrueGrid.Enabled = True
Me.vbxTrueGrid.SetFocus
Screen.MousePointer = DEFAULT

Exit Sub

cmdOK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_DEPTGL_MATRIX", "Update")
Call RollBack '09June99 js

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lStr("Department") & "/" & lStr("G/L") & " Matrix"
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

RHeading = lStr("Department") & "/" & lStr("G/L") & " Matrix"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub chkHideInactive_Click()
'Release 8.1
If chkHideInactive.Value = 1 Then
    Data1.RecordSource = "SELECT * FROM HR_DEPTGL_MATRIX WHERE DG_INACTIVE = 0 ORDER BY DG_INACTIVE, DG_DEPTNO,DG_GLNO "
    Data1.Refresh
Else
    Data1.RecordSource = "SELECT * FROM HR_DEPTGL_MATRIX ORDER BY DG_INACTIVE, DG_DEPTNO,DG_GLNO "
    Data1.Refresh
End If

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HR_DEPTGL_MATRIX", "SELECT")

End Sub

Private Sub Form_Activate()
Call SET_UP_MODE

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Me.cmdModify_Click
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim I%, SQLQ


Me.Show

glbOnTop = "FRMDEPTGLMATRIX"

Screen.MousePointer = HOURGLASS

Data1.ConnectionString = glbAdoIHRDB

'Release 8.1
Data1.RecordSource = "SELECT * FROM HR_DEPTGL_MATRIX WHERE DG_INACTIVE = 0 ORDER BY DG_INACTIVE, DG_DEPTNO,DG_GLNO "
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Call setCaption(lblDept)
Call setCaption(lblTitle(12))
chkInactive.Caption = "Inactive " & lStr("Department") & "/" & lStr("G/L")

Screen.MousePointer = DEFAULT

'Call Display_Value

Call ST_UPD_MODE(False)

If Not gSec_Upd_DeptGL_Matrix Then                                      'May99 js
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
End If

vbxTrueGrid.Columns(0).Caption = lStr("Department")
vbxTrueGrid.Columns(1).Caption = lStr(lblTitle(12).Caption)

Call INI_Controls(Me)

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

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
Dim I As Integer
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
'vbxTrueGrid.Enabled = FT
'cmdOK.Enabled = TF              '
'cmdCancel.Enabled = TF          '
'cmdClose.Enabled = FT           '
'cmdModify.Enabled = FT          '
'cmdNew.Enabled = FT             '
'cmdDelete.Enabled = FT          '
'cmdPrint.Enabled = FT           '


clpDept.Enabled = TF                    '
clpGLNum.Enabled = TF
chkInactive.Enabled = TF    'Release 8.1
If Data1.Recordset.BOF Or Data1.Recordset.EOF Then
'    cmdModify.Enabled = False
 '   cmdDelete.Enabled = False
End If

End Sub

Private Sub txtConvert_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    'Release 8.1
    If Not fGLBNew Then
        FRS.Requery
        FRS.Bookmark = Bookmark
        If FRS("DG_INACTIVE") Then
            RowStyle.ForeColor = vbRed
        End If
    End If
End Sub

'Private Sub  clpDept_Change()
'End Sub
'Private Sub txtDept_DblClick()
'    Call Get_Dept(False)
'    txtDept.Text = glbDept
'    lblDeptDesc.Caption = glbDeptDesc
'End Sub
'Private Sub txtDept_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub txtDept_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
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
    
    SQLQ = "SELECT * FROM HR_DEPTGL_MATRIX "
    'Release 8.1
    If chkHideInactive.Value = 1 Then
        SQLQ = SQLQ & " WHERE DG_INACTIVE = 0 "
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

    Data1.RecordSource = SQLQ
    Data1.Refresh
    
    'Release 8.1
    Set FRS = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True
        
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I%

On Error GoTo vbxTrueGrid_Err

Call Display_Value

If Data1.Recordset.EOF Or Data1.Recordset.BOF = 0 Then
    Exit Sub
End If


Exit Sub

vbxTrueGrid_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_DEPTGL_MATRIX", "Select")
Call RollBack '09June99 js

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

   
''' Sam add July 2002 * Remove Binding Control
Private Sub Display_Value()
    Dim SQLQ
    
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
        RSDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Call SET_UP_MODE
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HR_DEPTGL_MATRIX WHERE DG_ID= " & Data1.Recordset!DG_ID
    
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If RSDATA.EOF Or RSDATA.BOF Then Exit Sub
    
    Call Set_Control("R", Me, RSDATA)
    Call SET_UP_MODE
    
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fGLBNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fGLBNew = True
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_DeptGL_Matrix
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
If fGLBNew Then
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

