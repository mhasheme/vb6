VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSPersCompltgF7 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form 7 Filled By"
   ClientHeight    =   5955
   ClientLeft      =   1485
   ClientTop       =   885
   ClientWidth     =   8445
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
   ScaleHeight     =   5955
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkInactive 
      Alignment       =   1  'Right Justify
      Caption         =   "Inactive Code"
      DataField       =   "P7_INACTIVE"
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
      Height          =   315
      Left            =   120
      TabIndex        =   20
      Top             =   4800
      Width           =   1395
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      DataField       =   "P7_TITLE"
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
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   2
      Tag             =   "00-Official Title of the Person Completing the Form 7"
      Top             =   3900
      Width           =   5175
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      DataField       =   "P7_NAME"
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
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "00-Name of Person Completing the Form 7"
      Top             =   3480
      Width           =   5175
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   7080
      Top             =   3840
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
      TabIndex        =   14
      Top             =   5295
      Width           =   8445
      _Version        =   65536
      _ExtentX        =   14896
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
         TabIndex        =   15
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
         TabIndex        =   6
         Tag             =   "Close and exit screen"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1035
         TabIndex        =   7
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
         TabIndex        =   4
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
         TabIndex        =   5
         Tag             =   "Cancel the changes made"
         Top             =   165
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3525
         TabIndex        =   8
         Tag             =   "Add a new Province to the list"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4335
         TabIndex        =   9
         Tag             =   "Delete the Province listed above"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5745
         TabIndex        =   16
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
      TabIndex        =   12
      Tag             =   "Find specific record"
      Top             =   5610
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
      TabIndex        =   11
      Tag             =   "00-Search Description"
      Top             =   5640
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
      TabIndex        =   10
      Tag             =   "00-Search Code"
      Top             =   5640
      Width           =   540
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      DataField       =   "P7_COMPNO"
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
      Left            =   6360
      MaxLength       =   3
      TabIndex        =   13
      Text            =   "001"
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "FsPersCompltgF7.frx":0000
      Height          =   3105
      Left            =   120
      OleObjectBlob   =   "FsPersCompltgF7.frx":0014
      TabIndex        =   0
      Tag             =   "Province Listings"
      Top             =   120
      Width           =   8175
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7080
      Top             =   4440
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
   Begin MSMask.MaskEdBox medTelephone 
      DataField       =   "P7_PHONE"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Tag             =   "11-Telephone Number of the Person Completing Form 7"
      Top             =   4320
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   27
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "(###) ###-####  Ext(######)"
      PromptChar      =   "_"
   End
   Begin VB.Label lblPhone 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
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
      Left            =   120
      TabIndex        =   19
      Top             =   4365
      Width           =   465
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   3525
      Width           =   795
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Official Title"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   3945
      Width           =   1050
   End
End
Attribute VB_Name = "frmSPersCompltgF7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRSOld As String, glbEmptyNew As Integer
Dim fglbNewRec%, xOldCode As String
Dim fglbID As Integer
Dim xLinkItem As String
Dim xName, xTitle As String
Dim FRS As ADODB.Recordset
'Dim rsDATA As New ADODB.Recordset

Private Function chkPers_CompletingF7()
Dim prov$, Msg$
Dim rsPersF7 As New ADODB.Recordset
Dim SQLQ As String

chkPers_CompletingF7 = False

On Error GoTo chkPers_CompletingF7_Err

If Len(txtName.Text) < 1 Then
    MsgBox lblTitle(0).Caption & " is a required field"
    txtName.SetFocus
    Exit Function
End If

If Len(txtTitle.Text) < 1 Then
    MsgBox lblTitle(1).Caption & " is a required field"
    txtTitle.SetFocus
    Exit Function
End If

'Check if this combination of Employment Status and Category already exists
Set rsPersF7 = Nothing
SQLQ = "SELECT * FROM HR_OHS_PERSON_COMPLTG_F7"
SQLQ = SQLQ & " WHERE P7_NAME = '" & txtName.Text & "'"
SQLQ = SQLQ & " AND P7_TITLE = '" & txtTitle.Text & "'"
rsPersF7.Open SQLQ, gdbAdoIhr001, adOpenStatic
If fglbNewRec% = True Then
    If Not rsPersF7.EOF Then
        'Combination already exist
        MsgBox "This Person already exists."
        txtName.SetFocus
        Exit Function
    End If
    rsPersF7.Close
Else
    If (xName <> txtName.Text) Or (xTitle <> txtTitle.Text) Then
        If Not rsPersF7.EOF Then
            'Combination already exist
            MsgBox "This Person already exists."
            txtName.SetFocus
            Exit Function
        End If
        rsPersF7.Close
    End If
End If


chkPers_CompletingF7 = True

Exit Function

chkPers_CompletingF7_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Select", "HR_OHS_PERSON_COMPLTG_F7", "chkPers_CompletingF7")
Resume Next

End Function

Private Sub clpCode_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdCancel_Click()
Dim bk
'On Error GoTo Can_Err

Data1.Recordset.CancelBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    Data1.Refresh
End If

'rsDATA.CancelUpdate
'Call Set_Control("R", Me, rsDATA)

Call modSTUPD(False)  ' reset screen's attributes

cmdClose.SetFocus

fglbNewRec% = False
xName = ""
xTitle = ""
medTelephone.Text = ""
medTelephone.Mask = "(###) ###-####  Ext(######)"

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_OHS_PERSON_COMPLTG_F7", "Cancel")
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

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then
    Exit Sub
End If

'gdbAdoIhr001.BeginTrans
'rsDATA.Delete
'gdbAdoIhr001.CommitTrans

Data1.Recordset.Delete

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End If

Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_OHS_PERSON_COMPLTG_F7", "Delete")
Call RollBack '09June99 js

End Sub

Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub cmdFind_Click()
'Dim SQLQ$
'
'txtFindKey.SetFocus 'added by Marlon Cowan 9/16/97
'
'If Len(txtFindKey) > 0 Then
'    SQLQ$ = "CODE >= '" & txtFindKey.Text & "'"
'    Data1.Recordset.Requery
'    Data1.Recordset.Find SQLQ$
'    If Data1.Recordset.EOF Then
'        Data1.Refresh

'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
'    Else
'        txtFindKey = ""
'    End If
'    Exit Sub
'End If
'
'If Len(txtFindDesc) > 0 Then
'    SQLQ$ = "DESCR >= '" & txtFindDesc.Text & "'"
'    Data1.Recordset.Requery
'    Data1.Recordset.Find SQLQ$
'    If Data1.Recordset.EOF Then
'        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
'    Else
'        txtFindDesc = ""
'    End If
'    Exit Sub
'End If
'
'End Sub

Private Sub cmdFind_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()

On Error GoTo Mod_Err

Call modSTUPD(True)
xName = txtName.Text
xTitle = txtTitle.Text

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_OHS_PERSON_COMPLTG_F7", "Modify")
Call RollBack '09June99 js

End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()

On Error GoTo NewErr

fglbNewRec% = True

Data1.Recordset.AddNew

chkInactive.Value = 0
txtComp.Text = glbCompNo
xName = ""
xTitle = ""
medTelephone.Text = ""
medTelephone.Mask = "(###) ###-####  Ext(######)"


Call modSTUPD(True)

'Call Set_Control("B", Me)
'rsDATA.AddNew

txtName.SetFocus

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HR_OHS_PERSON_COMPLTG_F7", "AddNew")
Resume Next

End Sub

Private Sub cmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim Desc As String
Dim ProvCode

On Error GoTo OK_Err

If Not chkPers_CompletingF7() Then Exit Sub

'Call UpdUStats(Me)

fglbID = Data1.Recordset("P7_ID")

'Call Set_Control("U", Me, rsDATA)
'gdbAdoIhr001.BeginTrans
'rsDATA.Update
'gdbAdoIhr001.CommitTrans
'Data1.RecordSource = "SELECT * FROM HR_OHS_PERSON_COMPLTG_F7 ORDER BY P7_INACTIVE, P7_NAME"
'Data1.Refresh
'Set FRS = Data1.Recordset.Clone
'vbxTrueGrid.FetchRowStyle = True
'Data1.Recordset.Find "P7_ID='" & fglbID & "'"

Data1.Recordset("P7_COMPNO") = txtComp
Data1.Recordset("P7_PHONE") = medTelephone.Text
Data1.Recordset("P7_LDATE") = Format(Now, "SHORT DATE")
Data1.Recordset("P7_LTIME") = Time$
Data1.Recordset("P7_LUSER") = glbUserID
Data1.Recordset.UpdateBatch

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    fglbID = Data1.Recordset("P7_ID")
    Data1.Refresh
    
    Set FRS = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True
    
    Data1.Recordset.Find "P7_ID=" & fglbID & " "
End If

fglbNewRec% = False

Call modSTUPD(False)

cmdClose.SetFocus

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OHS_PERSON_COMPLTG_F7", "Update")

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

RHeading = "Form 7 Filled By"
Me.vbxCrystal.WindowTitle = "Form 7 Filled By Report"
Me.vbxCrystal.BoundReportHeading = "Form 7 Filled By"
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdSelect_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
Data1.RecordSource = "SELECT * FROM HR_OHS_PERSON_COMPLTG_F7 ORDER BY P7_INACTIVE,P7_NAME"
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

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
    
    Data1.ConnectionString = glbAdoIHRDB
    
    SQLQ = "UPDATE HR_OHS_PERSON_COMPLTG_F7 SET P7_INACTIVE = 0 WHERE P7_INACTIVE IS NULL"
    gdbAdoIhr001.Execute SQLQ

    Data1.RecordSource = "SELECT * FROM HR_OHS_PERSON_COMPLTG_F7 ORDER BY P7_INACTIVE,P7_NAME"
    'Data1.LockType = adLockReadOnly
    Data1.Refresh
    
    Set FRS = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True

    Call modSTUPD(False)            'Jaddy 10/18/99
    
    'If Not glbWSIBModule Then
    '    cmdModify.Enabled = False
    '    cmdNew.Enabled = False
    '    cmdDelete.Enabled = False
    'End If                          '
    
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
'If gSec_Upd_Basic Then  '
    cmdModify.Enabled = FT      '
    cmdNew.Enabled = FT         '
    cmdDelete.Enabled = FT      '
'Else
'    cmdModify.Enabled = False      '
'    cmdNew.Enabled = False   '
'    cmdDelete.Enabled = False
'End If

If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End If
txtName.Enabled = TF
txtTitle.Enabled = TF
medTelephone.Enabled = TF
chkInactive.Enabled = TF

txtFindDesc.Enabled = FT
txtFindKey.Enabled = FT
vbxTrueGrid.Enabled = FT

'If glbDivInhSel Then
'    cmdSelect.Enabled = False
'End If

End Sub

Private Sub medTelephone_Change()
medTelephone.Mask = "(###) ###-####  Ext(######)"
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

Private Sub txtName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If Not fglbNewRec% And Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
        FRS.Requery
        FRS.Bookmark = Bookmark
        If FRS("P7_INACTIVE") Then
            RowStyle.ForeColor = vbRed
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
                   
    SQLQ = "SELECT * FROM HR_OHS_PERSON_COMPLTG_F7 "
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

    Data1.RecordSource = SQLQ
    Data1.Refresh

    Set FRS = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True
End Sub

'Private Sub Display_Value()
'    Dim SQLQ
'    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
'        Call Set_Control("B", Me)
'        Exit Sub
'    End If
'
'    SQLQ = "SELECT * FROM HR_OHS_PERSON_COMPLTG_F7 WHERE P7_ID='" & Data1.Recordset!P7_ID & "'" & " order by P7_NAME"
'    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
'    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
'    Call Set_Control("R", Me, rsDATA)
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

'Public Property Let LinkItem(vData As String)
'    xLinkItem = vData
'End Property
'
'Public Property Get LinkItem() As String
'    LinkItem = xLinkItem
'End Property

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'Call Display_Value
End Sub
