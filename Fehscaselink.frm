VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEHSCauseLinks 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Root Cause Link - Type of Event and Immediate / Direct Causes"
   ClientHeight    =   4575
   ClientLeft      =   1485
   ClientTop       =   885
   ClientWidth     =   7875
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4575
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   7080
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   5
      Top             =   3915
      Width           =   7875
      _Version        =   65536
      _ExtentX        =   13891
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
         Left            =   6720
         TabIndex        =   6
         Tag             =   "Select Province listed above"
         Top             =   165
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Tag             =   "Close and exit screen"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1035
         TabIndex        =   8
         Tag             =   "Edit the information above"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1845
         TabIndex        =   9
         Tag             =   "Save the changes made"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Tag             =   "Cancel the changes made"
         Top             =   165
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3525
         TabIndex        =   11
         Tag             =   "Add a new Province to the list"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4335
         TabIndex        =   12
         Tag             =   "Delete the Province listed above"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5745
         TabIndex        =   13
         Tag             =   "Print the Province listing report"
         Top             =   165
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   5235
      TabIndex        =   2
      Tag             =   "Find specific record"
      Top             =   4890
      Width           =   735
   End
   Begin VB.TextBox txtFindDesc 
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
      TabIndex        =   1
      Tag             =   "00-Search Description"
      Top             =   4920
      Width           =   3975
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
      Left            =   360
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "00-Search Code"
      Top             =   4920
      Width           =   540
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      DataField       =   "RL_COMPNO"
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
      Left            =   6720
      MaxLength       =   3
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "Fehscaselink.frx":0000
      Height          =   2745
      Left            =   120
      OleObjectBlob   =   "Fehscaselink.frx":0014
      TabIndex        =   3
      Tag             =   "Province Listings"
      Top             =   120
      Width           =   7695
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7080
      Top             =   4080
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
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "RL_FIRSTCODE"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   2460
      TabIndex        =   14
      Tag             =   "00-Type of Event"
      Top             =   3120
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECRE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "RL_SECONDCODE"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   2460
      TabIndex        =   15
      Tag             =   "00-Immediate / Direct Causes"
      Top             =   3480
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECRI"
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Event"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   3165
      Width           =   2355
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Immediate / Direct Causes"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   3525
      Width           =   2355
   End
End
Attribute VB_Name = "frmEHSCauseLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRSOld As String, glbEmptyNew As Integer
Dim fglbNewRec%, xOldCode As String
Dim fglbID As Integer
Dim xLinkItem As String

Private Function chkRootCause()
Dim Prov$, SQLQ$, Msg$
Dim snapProv As New ADODB.Recordset

chkRootCause = False

On Error GoTo chkRootCause_Err

If Len(clpCode(0)) < 1 Then
    MsgBox lblTitle(0).Caption & " is a required field"
    clpCode(0).SetFocus
    Exit Function
End If
If Len(clpCode(1)) < 1 Then
    MsgBox lblTitle(1).Caption & " is a required field"
    clpCode(1).SetFocus
    Exit Function
End If


'Prov$ = txtCode
'SQLQ$ = "SELECT * from HR_OHS_RLINK_EVENT "
'SQLQ$ = SQLQ$ & " WHERE HR_OHS_RLINK_EVENT.CODE = '" & Prov$ & "'"
'If Not fglbNewRec Then
'    SQLQ$ = SQLQ$ & " AND HR_OHS_RLINK_EVENT.CODE <> '" & xOldCode & "'"
'End If
'SQLQ$ = SQLQ$ & " ORDER BY HR_OHS_RLINK_EVENT.CODE "
'
'snapProv.Open SQLQ$, gdbAdoIhr001, adOpenStatic
'If snapProv.BOF And snapProv.EOF Then
'    snapProv.Close
'Else
'    Msg$ = "This Province/State already exists"
'    MsgBox Msg$
'    snapProv.Close
'    Exit Function
'End If

chkRootCause = True

Exit Function

chkRootCause_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Select", "HR_OHS_RLINK_EVENT", "Cancel")
Resume Next

End Function



Private Sub cmdCancel_Click()
Dim bk
'On Error GoTo Can_Err

'Data1.UpdateControls    ' returns without saving
Data1.Recordset.CancelBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
'Data1.Refresh
If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    'fglbID = Data1.Recordset("RL_ID")
    Data1.Refresh
    'Data1.Recordset.Find "RL_ID=" & fglbID & " "
End If

Call modSTUPD(False)  ' reset screen's attributes


cmdClose.SetFocus

fglbNewRec% = False

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_OHS_RLINK_EVENT", "Cancel")
Resume Next

End Sub

Private Sub cmdCancel_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdDelete_Click()
Dim Msg As String, a%

On Error GoTo DelErr

'If Data1.Recordset.RecordCount <= 1 Then
'    MsgBox "You Can Not Delete The Last Province."
'    Exit Sub
'End If

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then
    Exit Sub
End If

Data1.Recordset.Delete
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End If

Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "Single", "Delete")
Call RollBack '09June99 js

End Sub

Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim SQLQ$

txtFindKey.SetFocus 'added by Marlon Cowan 9/16/97

If Len(txtFindKey) > 0 Then
    SQLQ$ = "CODE >= '" & txtFindKey.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ$
    If Data1.Recordset.EOF Then
        Data1.Refresh
    Else
        txtFindKey = ""
    End If
    Exit Sub
End If

If Len(txtFindDesc) > 0 Then
    SQLQ$ = "DESCR >= '" & txtFindDesc.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ$
    If Data1.Recordset.EOF Then
        Data1.Refresh
    Else
        txtFindDesc = ""
    End If
    Exit Sub
End If

End Sub

Private Sub cmdFind_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()

On Error GoTo Mod_Err

Call modSTUPD(True)
'txtCode.Enabled = True
'txtCode.SetFocus
'xOldCode = txtCode

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '09June99 js

End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()

On Error GoTo NewErr


glbCodeRef = True


Data1.Recordset.AddNew
txtComp.Text = glbCompNo
fglbNewRec% = True
Call modSTUPD(True)
'txtCode.SetFocus
clpCode(0).SetFocus

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HR_OHS_RLINK_EVENT", "AddNew")
Resume Next

End Sub

Private Sub CmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
'Dim SQLQ As String
Dim Desc As String
Dim ProvCode
On Error GoTo OK_Err

If Not chkRootCause() Then Exit Sub

Data1.Recordset("RL_COMPNO") = txtComp 'txtCode & ""
Data1.Recordset.UpdateBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    fglbID = Data1.Recordset("RL_ID")
    Data1.Refresh
    Data1.Recordset.Find "RL_ID=" & fglbID & " "
End If

fglbNewRec% = False


Call modSTUPD(False)

cmdClose.SetFocus

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption

Resume Next
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OHS_RLINK_EVENT", "Update")
Resume Next
Unload Me

End Sub

Private Sub cmdOK_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPrint_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = "Province/State Codes"
Me.vbxCrystal.WindowTitle = "Province/State Code Report"
Me.vbxCrystal.BoundReportHeading = "Province/State Codes"
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdSelect_Click()

If Me.vbxTrueGrid.EditActive = True Then
    MsgBox "Save/Cancel changes first"
Else
    glbProv = Data1.Recordset("CODE")
    glbProvDesc = Data1.Recordset("DESCR")
    Unload frmEHSCauseLinks
End If

End Sub

Private Sub cmdSelect_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
'
'glbFrmCaption$ = Me.Caption
'glbErrNum& = ErrorNumber
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HR_OHS_RLINK_EVENT", "SELECT")
'
'End Sub

Private Sub Form_Load()
Dim SQLQ As String

Screen.MousePointer = HOURGLASS
If Me.LinkItem = "EVENT" Then
    Data1.ConnectionString = glbAdoIHRDB
    Data1.RecordSource = "SELECT * FROM HR_OHS_RLINK_EVENT ORDER BY RL_FIRSTCODE,RL_SECONDCODE "
    Data1.Refresh
    If glbLinamar Then  'Ticket #14573
        Me.Caption = "Root Cause Link - Substandard Act and Substandard Condition"
        lblTitle(0).Caption = "Substandard Act"
        vbxTrueGrid.Columns(0).Caption = "Substandard Act"
        lblTitle(1).Caption = "Substandard Condition"
        vbxTrueGrid.Columns(1).Caption = "Substandard Condition"
    End If
End If
If Me.LinkItem = "IMMEDIATE" Then
    Data1.ConnectionString = glbAdoIHRDB
    Data1.RecordSource = "SELECT * FROM HR_OHS_RLINK_IMMEDIATE ORDER BY RL_FIRSTCODE,RL_SECONDCODE "
    Data1.Refresh
    Me.Caption = "Root Cause Link - Immediate / Direct Causes and Basic Categories"
    lblTitle(0).Caption = "Immediate / Direct Causes"
    lblTitle(1).Caption = "Basic Categories"
    vbxTrueGrid.Columns(0).Caption = "Immediate / Direct Causes"
    vbxTrueGrid.Columns(1).Caption = "Basic Categories"
    clpCode(0).TablName = "ECRI"
    clpCode(1).TablName = "ECRA"
    If glbLinamar Then  'Ticket #14573
        Me.Caption = "Root Cause Link - Substandard Condition and Personal Factor"
        lblTitle(0).Caption = "Substandard Condition"
        vbxTrueGrid.Columns(0).Caption = "Substandard Condition"
        lblTitle(1).Caption = "Personal Factor"
        vbxTrueGrid.Columns(1).Caption = "Personal Factor"
    End If
End If
If Me.LinkItem = "BASIC" Then
    Data1.ConnectionString = glbAdoIHRDB
    Data1.RecordSource = "SELECT * FROM HR_OHS_RLINK_BASIC ORDER BY RL_FIRSTCODE,RL_SECONDCODE "
    Data1.Refresh
    Me.Caption = "Root Cause Link - Basic Categories and Basic / Underlying Causes"
    lblTitle(0).Caption = "Basic Categories"
    lblTitle(1).Caption = "Basic / Underlying Causes"
    vbxTrueGrid.Columns(0).Caption = "Basic Categories"
    vbxTrueGrid.Columns(1).Caption = "Basic / Underlying Causes"
    clpCode(0).TablName = "ECRA"
    clpCode(1).TablName = "ECRB"
    If glbLinamar Then 'Ticket #12758
        'Me.Caption = "Root Cause Link - Sub Basic/Root Causes and Basic / Underlying Causes"
        'lblTitle(0).Caption = "Sub Basic/Root Causes"
        'vbxTrueGrid.Columns(0).Caption = "Sub Basic/Root Causes"
        'Ticket #14573
        Me.Caption = "Root Cause Link - Personal Factor and Job/System Factor"
        lblTitle(0).Caption = "Personal Factor"
        vbxTrueGrid.Columns(0).Caption = "Personal Factor"
        lblTitle(1).Caption = "Job/System Factor"
        vbxTrueGrid.Columns(1).Caption = "Job/System Factor"
    End If
End If

If Me.LinkItem = "TERMCAUSE" Then
    Data1.ConnectionString = glbAdoIHRDB
    Data1.RecordSource = "SELECT * FROM HR_TERMCAUSE_LINK ORDER BY RL_FIRSTCODE,RL_SECONDCODE "
    Data1.Refresh
    Me.Caption = "Termination Link - Reason / Cause"
    'lblTitle(0).Caption = "Termination Cause"
    'lblTitle(1).Caption = "Termination Reason"
    'vbxTrueGrid.Columns(0).Caption = "Termination Cause"
    'vbxTrueGrid.Columns(1).Caption = "Termination Reason"
    'clpCode(0).TablName = "TECA"
    'clpCode(1).TablName = "TERM"
    lblTitle(0).Caption = "Termination Reason"
    lblTitle(1).Caption = "Termination Cause"
    vbxTrueGrid.Columns(0).Caption = "Termination Reason"
    vbxTrueGrid.Columns(1).Caption = "Termination Cause"
    clpCode(0).TablName = "TERM"
    clpCode(1).TablName = "TECA"
End If

Call modSTUPD(False)            'Jaddy 10/18/99

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT   '
                                
End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

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

cmdOK.Enabled = TF          'May99 js
cmdCancel.Enabled = TF      '
cmdClose.Enabled = FT
cmdPrint.Enabled = FT       '
cmdFind.Enabled = FT        '
cmdSelect.Enabled = FT
cmdDelete.Enabled = FT
'Ticket #18335 Frank 04/12/2010, Jerry asked to use OHS Cause Code security
If gSec_Upd_Master_Table("ECCA") Then
'If gSec_Upd_Company Then '
    cmdModify.Enabled = FT      '
    cmdNew.Enabled = FT         '
    cmdDelete.Enabled = FT      '
Else
    cmdModify.Enabled = False      '
    cmdNew.Enabled = False   '
    cmdDelete.Enabled = False
End If

If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End If
clpCode(0).Enabled = TF
clpCode(1).Enabled = TF
txtFindDesc.Enabled = FT
txtFindKey.Enabled = FT
vbxTrueGrid.Enabled = FT

If glbDivInhSel Then
    cmdSelect.Enabled = False
End If

End Sub



Private Sub txtFindDesc_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub txtFindKey_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub


Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        If Me.LinkItem = "EVENT" Then
            SQLQ = "SELECT * FROM HR_OHS_RLINK_EVENT"
            SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        End If
        
        If Me.LinkItem = "BASIC" Then
            SQLQ = "SELECT * FROM HR_OHS_RLINK_BASIC"
            SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        End If
        
        If Me.LinkItem = "IMMEDIATE" Then
            SQLQ = "SELECT * FROM HR_OHS_RLINK_IMMEDIATE "
            SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        End If
        
        If Me.LinkItem = "TERMCAUSE" Then
            SQLQ = "SELECT * FROM HR_TERMCAUSE_LINK "
            SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        End If

        Data1.RecordSource = SQLQ
        Data1.Refresh

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

Public Property Let LinkItem(vData As String)
    xLinkItem = vData
End Property
Public Property Get LinkItem() As String
    LinkItem = xLinkItem
End Property

