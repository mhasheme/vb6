VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmTABLATTGroupMASTER 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendance Group Matrix"
   ClientHeight    =   6720
   ClientLeft      =   1080
   ClientTop       =   1050
   ClientWidth     =   9870
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
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6720
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin INFOHR_Controls.CodeLookup clpGroups 
      DataSource      =   "Data1"
      Height          =   252
      Left            =   2040
      TabIndex        =   18
      Top             =   4200
      Width           =   4332
      _ExtentX        =   7646
      _ExtentY        =   503
      TABLName        =   "ADGP"
      TABLTitle       =   "Attendance Group Code"
      Object.Height          =   255
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmTABLATTGroupMaster.frx":0000
      Height          =   3972
      Left            =   0
      OleObjectBlob   =   "frmTABLATTGroupMaster.frx":0014
      TabIndex        =   0
      Top             =   0
      Width           =   9132
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   336
      Left            =   7560
      Top             =   6360
      Visible         =   0   'False
      Width           =   1932
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
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   6060
      Width           =   9876
      _Version        =   65536
      _ExtentX        =   17420
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
      Begin VB.CommandButton BtnClose 
         Caption         =   "Close"
         Height          =   372
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   612
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5340
         TabIndex        =   15
         Tag             =   "Print Code Listing Report"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4485
         TabIndex        =   14
         Tag             =   "Delete code listed above"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3615
         TabIndex        =   13
         Tag             =   "Add a new Code"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2580
         TabIndex        =   12
         Tag             =   "Cancel the changes made"
         Top             =   135
         Width           =   915
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1740
         TabIndex        =   11
         Tag             =   "Save the changes made"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   900
         TabIndex        =   10
         Tag             =   "Edit the Information"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   60
         TabIndex        =   9
         Tag             =   "Close and exit this screen"
         Top             =   -480
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6210
         Top             =   15
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   1
         DiscardSavedData=   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      DataField       =   "CompanyNo"
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
      Left            =   5040
      MaxLength       =   3
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   372
      Left            =   2040
      TabIndex        =   6
      Tag             =   "Find specific record"
      Top             =   5280
      Width           =   612
   End
   Begin VB.TextBox txtFindKey 
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
      Height          =   285
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   5
      Tag             =   "00-Search Code"
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox txtFindTabl 
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
      Height          =   285
      Left            =   240
      MaxLength       =   4
      TabIndex        =   4
      Tag             =   "00-Search Table"
      Top             =   5280
      Width           =   750
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LUSER"
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
      Height          =   315
      Index           =   2
      Left            =   3360
      MaxLength       =   25
      TabIndex        =   3
      Text            =   "LUser"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LTIME"
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
      Height          =   315
      Index           =   1
      Left            =   1680
      MaxLength       =   25
      TabIndex        =   2
      Text            =   "LTime"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LDATE"
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
      Height          =   315
      Index           =   0
      Left            =   0
      MaxLength       =   25
      TabIndex        =   1
      Text            =   "Ldate"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1590
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataSource      =   "Data1"
      Height          =   252
      Index           =   1
      Left            =   2040
      TabIndex        =   16
      Tag             =   "01-Attendance Code"
      Top             =   4560
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ADRE"
      MaxLength       =   10
      Object.Height          =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Group Code"
      Height          =   252
      Left            =   240
      TabIndex        =   19
      Top             =   4200
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "Attendance Code"
      Height          =   252
      Left            =   240
      TabIndex        =   17
      Top             =   4560
      Width           =   1572
   End
End
Attribute VB_Name = "frmTABLATTGroupMASTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNewRec%
Dim fglbUDMode As Integer ', glbEmptyNew As Integer
Dim fglbRSOld As String
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim fRS As ADODB.Recordset


'validatation function
Private Function ValidateForm() As Boolean

    If Len(clpCode(1).Text) > 0 And clpCode(1).Caption = "Unassigned" Then
        MsgBox "Invalid Attendance Code"
        clpCode(1).SetFocus
        ValidateForm = False
        Exit Function
    End If
    
     If Len(clpGroups.Text) > 0 And clpGroups.Caption = "Unassigned" Then
        MsgBox "Invalid Group Code"
        clpGroups.SetFocus
        ValidateForm = False
        Exit Function
    End If

Dim SQLQ As String, Msg$, Tabl As String, Ky As String
Dim snapTabs As New ADODB.Recordset


On Error GoTo chkMastTable_Err

If fglbNewRec Then
    Tabl = clpGroups.Text
    Ky = clpCode(1).Text
    SQLQ = "SELECT * FROM  HRATTGRP "
    SQLQ = SQLQ & "WHERE AD_ATTENDANCE_CODE_TABL = 'ADRE'"
    SQLQ = SQLQ & " AND AD_ATTENDANCE_TABLE = 'ADGP' and AD_ATTENDANCE_GROUP = '" & Tabl & "'"
    SQLQ = SQLQ & " AND AD_ATTENDANCE_CODE = '" & Ky & "'"
    
    
    snapTabs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If snapTabs.BOF And snapTabs.EOF Then
        snapTabs.Close
    Else
        Msg$ = "This Group Code is already associated with the selected attendance code"
        MsgBox Msg$
        snapTabs.Close
        Exit Function
    End If
End If

ValidateForm = True

Exit Function

chkMastTable_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HRTABLE", "HRTABL", "Cancel")
Call RollBack '10June99 js

End Function


'not in use here
Private Sub cmbDWM_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


'close button
Private Sub BtnClose_Click()

    Unload Me

End Sub

'change event function
Private Sub clpCode_Change(Index As Integer)
  Call SetPanHelp(ActiveControl)
End Sub


'cancel button
Private Sub cmdCancel_Click()

On Error GoTo Can_Err
Dim bk
'Data1.UpdateControls
Data1.Recordset.CancelUpdate
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
Call ST_UPD_MODE(False)  ' reset screen's attributes

fglbNewRec% = False
txtFindTabl.Visible = True
txtFindKey.Visible = True
'txtFindDesc.Visible = True
cmdFind.Visible = True

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRTABL", "Cancel")
Call RollBack '10June99 js

End Sub


Private Sub cmdCancel_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub cmdDelete_Click()
Dim a%, Msg
On Error GoTo DelErr

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

If Data1.Recordset.RecordCount < 2 Then
    MsgBox "You can not delete the last reference for this code"
Else
    'Call Codes_Master_Integration(txtTable, txtKey, , True)
    
    Data1.Recordset.Delete
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    Data1.Refresh
End If


Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRATTGRP", "Delete")
Call RollBack '10June99 js

End Sub

Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim SQLQ As String

    'Hemu - 05/29/2003 Begin - Ticket # 4204
    If glbCompSerial = "S/N - 2161W" Then
        'Since the txtKey maxlength is changing on Find to 4
        'txtKey.MaxLength = 8
    End If
    'Hemu - 05/29/2003 End

If Len(txtFindTabl) > 0 Then
    SQLQ = " AD_ATTENDANCE_GROUP >= '" & txtFindTabl.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
    Else
        txtFindTabl = ""
    End If
    Exit Sub
End If

If Len(txtFindKey) > 0 Then
    SQLQ = "AD_ATTENDANCE_CODE >= '" & txtFindKey.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
    Else
        txtFindKey = ""
    End If
    Exit Sub
End If




End Sub

Private Sub cmdFind_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()

On Error GoTo Mod_Err

Call ST_UPD_MODE(True)
'txtTable.Enabled = False
'txtKey.Enabled = False
'txtDesc.Enabled = True
'txtDesc.SetFocus
clpCode(1).Enabled = True
clpGroups.Enabled = True

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '10June99 js

End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()

On Error GoTo NewErr

glbCodeRef = True   'global call for code refresh
fglbNewRec% = True
Call ST_UPD_MODE(True)
Data1.Recordset.AddNew
'ChkEMELEA.Value = False
txtComp.Text = glbCompNo
'chkInactiveCode.Value = 0
'fglbNewRec% = True
'txtTable.SetFocus
txtFindTabl.Visible = False
txtFindKey.Visible = False
'txtFindDesc.Visible = False
cmdFind.Visible = False
If glbBurlTech Then
 '   chkLE.Value = False
End If
clpCode(1).Enabled = True
clpGroups.Enabled = True

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HRATTENDANCE", "HTATTGRP", "add new")
Call RollBack '10June99 js

End Sub

Private Sub clpCode_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub CmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim strT, strK As String
Dim bk
On Error GoTo OK_Err

'do validation
If Not ValidateForm() Then Exit Sub

glbCodeRef = True   'table entrie modified/added - forces refresh
                    ' at form level of codes/descriptions.
Call UpdUStats(Me)

strT = clpCode(1).Text
strK = clpGroups.Text
'If cmbDWM.Visible Then
 '   txtDWM = Left(cmbDWM, 1)
  '  If cmbDWM.ListIndex <> -1 And Len(txtWaitPeriod) = 0 Then
   '     txtWaitPeriod = 0
   ' End If
    'If txtWaitPeriod = "" And txtWaitPeriod.DataChanged Then txtWaitPeriod.DataChanged = False: Data1.Recordset("TB_USR2") = Null
'End If
Data1.Recordset("AD_ATTENDANCE_GROUP") = strK & ""
Data1.Recordset("AD_ATTENDANCE_CODE") = strT & ""

Data1.Recordset.UpdateBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Data1.Recordset.Filter = " AD_ATTENDANCE_GROUP = '" & strK & "' AND AD_ATTENDANCE_CODE = '" & strT & "'"
If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
    bk = Data1.Recordset.Bookmark
    Data1.Recordset.Filter = ""
    Data1.Recordset.Bookmark = bk
Else
    Data1.Recordset.Filter = ""
End If

'Call Codes_Master_Integration(txtTable, txtKey)

fglbNewRec% = False


Call ST_UPD_MODE(False)

cmdClose.SetFocus
txtFindTabl.Visible = True
txtFindKey.Visible = True
'txtFindDesc.Visible = True
cmdFind.Visible = True

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRATTGRP", "Update")
Call RollBack '10June99 js

End Sub

Private Sub cmdOK_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPrint_Click()

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

Me.vbxCrystal.ReportTitle = "All Table Codes"
Me.vbxCrystal.BoundReportHeading = Me.Caption
Me.vbxCrystal.WindowTitle = Me.Caption & " Report"
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "data1 error", "HRATTGRP", "SELECT")
Call RollBack '10June99 js

End Sub


'form load event
Private Sub Form_Load()
Dim SQLQ As String
Dim rsSR As New ADODB.Recordset

'SQLQ = "UPDATE HRTABL SET TB_INACTIVE = 0 WHERE TB_INACTIVE IS NULL"
'gdbAdoIhr001.Execute SQLQ


Data1.ConnectionString = glbAdoIHRDB
If glbOracle Then
    Data1.RecordSource = "SELECT * FROM HRATTGRP WHERE AD_TB_NAME IN(SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) ORDER BY 1,2"
Else
    'Data1.RecordSource = "SELECT * FROM HRTABL WHERE TB_NAME IN (SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & glbUserID & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) ORDER BY TB_NAME,TB_KEY"
    Data1.RecordSource = "SELECT * FROM HRATTGRP"
    
End If
Data1.Refresh
Set fRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

'If glbCompSerial = "S/N - 2347W" Then 'Surrey Place
'    ChkInc.Caption = "LTD"
'    ChkInc.Tag = "LTD"
'End If
clpGroups.DataField = "AD_ATTENDANCE_GROUP"
clpCode(1).DataField = "AD_ATTENDANCE_CODE"

clpCode(1).Enabled = False
clpGroups.Enabled = False


If glbBurlTech Then
    'ChkInc.Caption = "Unexcused"
    'chkSen.Caption = "Excused"
    'lblPoint.Caption = "Absence Point"
    'chkLE.DataField = "TB_LEFLAG"
    'txtLEPoint.DataField = "TB_LEPOINT"
    
End If

glbCode = ""    'set to null - implies none found/cancel
glbCodeDesc = ""

Screen.MousePointer = HOURGLASS
Me.vbxTrueGrid.Refresh
Call ST_UPD_MODE(False)
'SQLQ = "SELECT * FROM HRATTGRP "
''' WHERE Attendance_Code_TABL NOT IN (SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & glbUserID & "' AND CODENAME IS NOT NULL AND Maintainable<>0) "
'rsSR.Open SQLQ, gdbAdoIhr001, adOpenKeyset
'If Not rsSR.EOF Then
'    cmdNew.Enabled = False
'   cmdModify.Enabled = False
'    cmdDelete.Enabled = False
'End If
Call INI_Controls(Me)

If Not gSec_Upd_Attendance_Group_Code_Matrix Then
    cmdModify.Enabled = False
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
End If
                                
Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

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

fglbUDMode = TF     'in update/new mode

'txtWaitPeriod.Enabled = TF
'cmbDWM.Enabled = TF

cmdOK.Enabled = TF
cmdCancel.Enabled = TF
cmdModify.Enabled = FT
cmdClose.Enabled = FT
cmdFind.Enabled = FT
cmdNew.Enabled = FT
cmdDelete.Enabled = FT
cmdPrint.Enabled = FT

'txtGrpDesc.Enabled = TF
'txtGrpCode.Enabled = TF

'txtFindDesc.Enabled = FT
txtFindKey.Enabled = FT
txtFindTabl.Enabled = FT
vbxTrueGrid.Enabled = FT

End Sub

Private Sub txtDesc_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


'Hemu Commented
'Private Sub txtDWM_Change()
'cmbDWM.ListIndex = -1
'Select Case txtDWM
'Case "D"
'    cmbDWM.ListIndex = 0
'Case "W"
'    cmbDWM.ListIndex = 1
'Case "M"
'    cmbDWM.ListIndex = 2
'End Select
'End Sub

Private Sub txtFindDesc_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub txtFindKey_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtFindTabl_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub txtFindTabl_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

'Hemu Commented
'Private Sub txtKey_Change()
'    'Hemu - 05/29/2003 Begin - Ticket # 4204
'    If glbCompSerial = "S/N - 2161W" Then
'        If txtTable = "ESCT" Or txtTable = "ESCD" Then
'            txtKey.MaxLength = 8
'        End If
'    End If
'    'Hemu - 05/29/2003 End
'    'Bryan 15/Sep/05 Ticket #9327
'    If txtTable = "BNCD" Then
'            txtKey.MaxLength = 10
'    End If
'End Sub

Private Sub txtKey_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub txtKey_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtTable_Change()

'If txtTable = "ADRE" Then
'  ChkInc.Visible = True
'  chkSen.Visible = True
'  chkAbsence.Visible = True
'  ChkEMELEA.Visible = True
'  lblPoint.Visible = True
'  txtPoint.Visible = True
'  If glbBurlTech Then
'    chkLE.Visible = True
'    lblLEPoint.Visible = True
'    txtLEPoint.Visible = True
'  End If
'ElseIf txtTable = "EDEM" And Not glbLinamar Then
'  ChkInc.Visible = False
'  chkSen.Visible = False
'  chkAbsence.Visible = False
'  ChkEMELEA.Visible = True
'  ChkEMELEA.Caption = "Leave of Absence"
'  lblPoint.Visible = False
'  txtPoint.Visible = False
'  If glbBurlTech Then
'    chkLE.Visible = False
'    lblLEPoint.Visible = False
'    txtLEPoint.Visible = False
'  End If
'Else
'  ChkInc.Visible = False
'  chkSen.Visible = False
'  chkAbsence.Visible = False
'  ChkEMELEA.Visible = False
'  lblPoint.Visible = False
'  txtPoint.Visible = False
'  If glbBurlTech Then
'    chkLE.Visible = False
'    lblLEPoint.Visible = False
'    txtLEPoint.Visible = False
'  End If
'End If
'
'If txtTable = "SDLB" Then txtKey.MaxLength = 1 Else txtKey.MaxLength = 4
'
'    'Hemu - 05/29/2003 Begin - Ticket # 4204
'    If glbCompSerial = "S/N - 2161W" Then
'        If txtTable = "ESCT" Or txtTable = "ESCD" Then
'            txtKey.MaxLength = 8
'        End If
'    End If
'    'Hemu - 05/29/2003 End
'
''Bryan 15/Sep/05 Ticket #9327
'    If txtTable = "BNCD" Then
'        txtKey.MaxLength = 10
'    End If
''Bryan
'
'If txtTable = "BNCD" And glbLinamar Then
'    lblTitle(0).Visible = True
'    txtWaitPeriod.Visible = True
'    cmbDWM.Visible = True
'Else
'    lblTitle(0).Visible = False
'    txtWaitPeriod.Visible = False
'    cmbDWM.Visible = False
'End If
End Sub

Private Sub txtTable_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub txtTable_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtWaitPeriod_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_DblClick()

If Not Me.vbxTrueGrid.EditActive Then
    glbCode = Data1.Recordset("AD_ATTENDANCE_GROUP")
    glbCodeDesc = Data1.Recordset("AD_ATTENDANCE_TABLE")
    Unload frmTABLATTGroupMASTER
Else
    MsgBox "Save/cancel changes first"
End If

End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    'added by Frank 16/Apr/07 Ticket #12913
    If Not fglbNewRec% Then
        fRS.Requery
        fRS.Bookmark = Bookmark
        'If fRS("TB_INACTIVE") Then
           ' RowStyle.ForeColor = vbRed
        'End If
    End If
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Function RollBack()
On Error GoTo RR
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
RR:
End Function

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        If glbOracle Then
            SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME IN(SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) "
        Else
            SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME IN (SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME IS NOT NULL AND ACCESSABLE<>0) "
        End If
        SQLQ = "SELECT * FROM HRATTGRP "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh


End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rsSR As New ADODB.Recordset

If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
    'If Data1.Recordset("TB_NAME") = "SDLB" Then txtKey.MaxLength = 1 Else txtKey.MaxLength = 4
    
    'Hemu - 05/29/2003 Begin - Ticket # 4204
    If glbCompSerial = "S/N - 2161W" Then
        'If txtTable = "ESCT" Or txtTable = "ESCD" Then
         '   txtKey.MaxLength = 8
       ' End If
    End If
    'Hemu - 05/29/2003 End

    If glbVadim Then
        'If txtTable = "EDOR" Then txtKey.MaxLength = 1
       ' If txtTable = Vadim_PayType_TABLName Then txtKey.MaxLength = 1
       ' If txtTable = Vadim_EmpType_TABLName Then txtKey.MaxLength = 2
    End If

    rsSR.Open "SELECT CODENAME FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbUserID, "'", "''") & "' AND CODENAME='" & Data1.Recordset("AD_ATTENDANCE_TABLE") & "' AND Maintainable<>0 ", gdbAdoIhr001, adOpenKeyset
    If rsSR.EOF Then
        cmdNew.Enabled = False
        cmdModify.Enabled = False
        cmdDelete.Enabled = False
    Else
        cmdNew.Enabled = True
        cmdModify.Enabled = True
        cmdDelete.Enabled = True
    End If
Else
    cmdModify.Enabled = False
End If


End Sub




