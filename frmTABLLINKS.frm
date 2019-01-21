VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmTABLLINKS 
   Caption         =   "Table Master Edit Links"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLinkDesc 
      Appearance      =   0  'Flat
      DataField       =   "TB_SECOND_TABL_DESC"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7440
      MaxLength       =   35
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtMainDesc 
      Appearance      =   0  'Flat
      DataField       =   "TB_FIRST_TABL_DESC"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7440
      MaxLength       =   35
      TabIndex        =   21
      Top             =   2880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtLink 
      Appearance      =   0  'Flat
      DataField       =   "TB_SECOND_TABL"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      MaxLength       =   20
      TabIndex        =   20
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtMain 
      Appearance      =   0  'Flat
      DataField       =   "TB_FIRST_TABL"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      MaxLength       =   20
      TabIndex        =   19
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      DataField       =   "TB_COMPNO"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   8520
      MaxLength       =   3
      TabIndex        =   18
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox comLink 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmTABLLINKS.frx":0000
      Left            =   1560
      List            =   "frmTABLLINKS.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "10-Type of Employee "
      Top             =   3840
      Width           =   5055
   End
   Begin VB.ComboBox comMain 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmTABLLINKS.frx":0004
      Left            =   1560
      List            =   "frmTABLLINKS.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "10-Type of Employee "
      Top             =   2880
      Width           =   5055
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   5880
      Top             =   4560
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
      TabIndex        =   11
      Top             =   4980
      Width           =   9510
      _Version        =   65536
      _ExtentX        =   16775
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
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4335
         TabIndex        =   10
         Tag             =   "Delete the Province listed above"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3525
         TabIndex        =   9
         Tag             =   "Add a new Province to the list"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Tag             =   "Cancel the changes made"
         Top             =   165
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1845
         TabIndex        =   7
         Tag             =   "Save the changes made"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1035
         TabIndex        =   6
         Tag             =   "Edit the information above"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Tag             =   "Close and exit screen"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
         Height          =   375
         Left            =   6720
         TabIndex        =   12
         Tag             =   "Select Province listed above"
         Top             =   165
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmTABLLINKS.frx":0008
      Height          =   2625
      Left            =   120
      OleObjectBlob   =   "frmTABLLINKS.frx":001C
      TabIndex        =   0
      Tag             =   "Province Listings"
      Top             =   120
      Width           =   9255
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "TB_FIRSTCODE"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   1260
      TabIndex        =   2
      Tag             =   "00-Type of Event"
      Top             =   3360
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECRE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "TB_SECONDCODE"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   1260
      TabIndex        =   4
      Tag             =   "00-Immediate / Direct Causes"
      Top             =   4320
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECRE"
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Child Code"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   4320
      Width           =   915
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Child Mnemonic"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   1155
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Parent Code"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   1035
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Parent Mnemonic"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   1395
   End
End
Attribute VB_Name = "frmTABLLINKS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRSOld As String, glbEmptyNew As Integer
Dim fglbNewRec%, xOldCode As String
Dim fglbID As Integer
Dim xLinkItem As String
Dim ArrTabCode(150, 2)
Dim ArrMax As Integer
Dim SQLQ As String

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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRTABL_LINKS", "Cancel")
Resume Next

End Sub

Private Sub cmdClose_Click()
    Unload Me
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
Call RollBack

End Sub

Private Sub cmdModify_Click()
On Error GoTo Mod_Err

Call modSTUPD(True)

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack
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
'comMain
txtMain.Text = Left(comMain.Text, 4)
txtMainDesc.Text = Trim(Mid(comMain.Text, 8, 35))
txtLink.Text = Left(comLink.Text, 4)
txtLinkDesc.Text = Trim(Mid(comLink.Text, 8, 35))

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HRTABL_LINKS", "AddNew")
Resume Next

End Sub

Private Sub cmdOK_Click()
Dim Desc As String
Dim ProvCode
On Error GoTo OK_Err

If Not chkRootCause() Then Exit Sub

Data1.Recordset("TB_COMPNO") = txtComp 'txtCode & ""
Data1.Recordset.UpdateBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    fglbID = Data1.Recordset("TB_ID]")
    Data1.Refresh
    Data1.Recordset.Find "TB_ID]=" & fglbID & " "
End If

fglbNewRec% = False


Call modSTUPD(False)

cmdClose.SetFocus

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption

Resume Next
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRTABL_LINKS", "Update")
Resume Next
Unload Me

End Sub

Private Sub comLink_Click()
txtLink.Text = Left(comLink.Text, 4)
txtLinkDesc.Text = Trim(Mid(comLink.Text, 8, 35))
End Sub

Private Sub comMain_Click()
txtMain.Text = Left(comMain.Text, 4)
txtMainDesc.Text = Trim(Mid(comMain.Text, 8, 35))
End Sub

Private Sub Form_Load()
Dim SQLQ As String

Screen.MousePointer = HOURGLASS

Call Pupulate_Combo

Data1.ConnectionString = glbAdoIHRDB

SQLQ = "SELECT * FROM HRTABL_LINKS ORDER BY TB_FIRST_TABL_DESC,TB_FIRSTCODE,TB_SECOND_TABL_DESC,TB_SECONDCODE "
Data1.RecordSource = SQLQ
Data1.Refresh

'Call modSTUPD(False)

Call INI_Controls(Me)

If Not gSec_Upd_SAMTableMasterLinks Then
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
    comMain.Enabled = False
    clpCode(0).Enabled = False
    comLink.Enabled = False
    clpCode(1).Enabled = False
End If
    
Screen.MousePointer = DEFAULT
                                
End Sub

Private Sub Pupulate_Combo()
Dim rsTBL As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim I As Integer 'HRSALWRK
    comMain.Clear
    comLink.Clear
    SQLQ = "DELETE FROM HRSALWRK WHERE WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ
    
    'SQLQ = "SELECT * FROM HRTABDES WHERE TD_NAME IN (SELECT CODENAME FROM HR_CODERELATE) ORDER BY TD_NAME" ' TD_DESC " ' TD_NAME"
    SQLQ = "SELECT * FROM HRTABDES " 'WHERE TD_NAME IN (SELECT CODENAME FROM HR_CODERELATE) ORDER BY TD_NAME" ' TD_DESC " ' TD_NAME"
    rsTBL.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    SQLQ = "SELECT * FROM HRSALWRK WHERE WRKEMP = '" & glbUserID & "' "
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    Do While Not rsTBL.EOF
        rsWRK.AddNew
        rsWRK("SH_COMMENT") = rsTBL("TD_NAME")
        rsWRK("SH_COMMENT2") = UCase(lStr(rsTBL("TD_DESC")))
        rsWRK("WRKEMP") = glbUserID
        rsWRK("SH_CURRENT") = 1
        rsWRK.Update
        rsTBL.MoveNext
     Loop
    
    rsTBL.Close
    rsWRK.Close
    SQLQ = "SELECT * FROM HRSALWRK WHERE WRKEMP = '" & glbUserID & "' "
    SQLQ = SQLQ & "ORDER BY SH_COMMENT2 "
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    I = 0
    Do While Not rsWRK.EOF
        comMain.AddItem rsWRK("SH_COMMENT") & " - " & rsWRK("SH_COMMENT2")
        comLink.AddItem rsWRK("SH_COMMENT") & " - " & rsWRK("SH_COMMENT2")
        ArrTabCode(I, 0) = rsWRK("SH_COMMENT")
        ArrTabCode(I, 1) = rsWRK("SH_COMMENT2")
        ArrTabCode(I, 2) = I 'combo index
        rsWRK.MoveNext
        I = I + 1
    Loop
    ArrMax = I - 1
    rsWRK.Close
    
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

cmdOK.Enabled = TF
cmdCancel.Enabled = TF      '
cmdClose.Enabled = FT
cmdPrint.Enabled = FT       '
cmdSelect.Enabled = FT
cmdDelete.Enabled = FT

cmdModify.Enabled = FT      '
cmdNew.Enabled = FT         '
cmdDelete.Enabled = FT      '

If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End If
comMain.Enabled = TF
comLink.Enabled = TF
clpCode(0).Enabled = TF
clpCode(1).Enabled = TF
vbxTrueGrid.Enabled = FT

End Sub

Private Function chkRootCause()
Dim prov$, SQLQ$, Msg$
Dim snapProv As New ADODB.Recordset

chkRootCause = False

On Error GoTo chkRootCause_Err

If Len(comMain.Text) = 0 Then
    MsgBox "Main Table is a required field"
    comMain.SetFocus
    Exit Function
End If
If Len(clpCode(0)) < 1 Then
    MsgBox lblTitle(0).Caption & " is a required field"
    clpCode(0).SetFocus
    Exit Function
End If
If Len(comLink.Text) = 0 Then
    MsgBox "Link Table is a required field"
    comLink.SetFocus
    Exit Function
End If
If Len(clpCode(1)) < 1 Then
    MsgBox lblTitle(1).Caption & " is a required field"
    clpCode(1).SetFocus
    Exit Function
End If


chkRootCause = True

Exit Function

chkRootCause_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Select", "HRTABL_LINKS", "Cancel")
Resume Next

End Function


Private Sub txtLink_Change()
    clpCode(1).TablName = txtLink.Text
    comLink.ListIndex = getDropdownIndex(txtLink.Text)
    Call INI_Controls(Me) 'Ticket #24055 Franks 07/11/2013 - fixed the error for samuel
End Sub

Private Sub txtMain_Change()
    clpCode(0).TablName = txtMain.Text
    comMain.ListIndex = getDropdownIndex(txtMain.Text)
End Sub

Private Function getDropdownIndex(xVal)
Dim I As Integer
Dim retVal As Integer
    retVal = -1
    For I = 0 To ArrMax
        If ArrTabCode(I, 0) = xVal Then
            retVal = ArrTabCode(I, 2)
        End If
    Next
    getDropdownIndex = retVal
End Function
