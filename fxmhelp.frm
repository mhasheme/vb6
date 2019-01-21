VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmMHelp 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help Descriptions Setup - Corrective Actiion Code"
   ClientHeight    =   5520
   ClientLeft      =   1485
   ClientTop       =   885
   ClientWidth     =   8040
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
   ScaleHeight     =   5520
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
      DataField       =   "HP_MEMO"
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
      Height          =   1530
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Tag             =   "00-Description"
      Top             =   3000
      Width           =   7755
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   6240
      Top             =   4440
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
      TabIndex        =   4
      Top             =   4860
      Width           =   8040
      _Version        =   65536
      _ExtentX        =   14182
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
         TabIndex        =   5
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
         TabIndex        =   8
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
         TabIndex        =   9
         Tag             =   "Cancel the changes made"
         Top             =   165
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3525
         TabIndex        =   10
         Tag             =   "Add a new Province to the list"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4335
         TabIndex        =   11
         Tag             =   "Delete the Province listed above"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5745
         TabIndex        =   12
         Tag             =   "Print the Province listing report"
         Top             =   165
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      DataField       =   "HP_COMPNO"
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
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxmhelp.frx":0000
      Height          =   2265
      Left            =   120
      OleObjectBlob   =   "fxmhelp.frx":0014
      TabIndex        =   0
      Tag             =   "Province Listings"
      Top             =   120
      Width           =   7455
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   5880
      Top             =   4920
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
      DataField       =   "HP_TABL_KEY"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   2460
      TabIndex        =   1
      Tag             =   "00-Corrective Action Code"
      Top             =   2520
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECCR"
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Corrective Action Code"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   2355
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Help Description"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   2355
   End
End
Attribute VB_Name = "frmMHelp"
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


chkRootCause = True

Exit Function

chkRootCause_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Select", "HRHELP", "Cancel")
Resume Next

End Function



Private Sub clpCode_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmdCancel_Click()
Dim bk
'On Error GoTo Can_Err


Data1.Recordset.CancelBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
'Data1.Refresh
If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRHELP", "Cancel")
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
Dim SQLQ As String
On Error GoTo DelErr


Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then
    Exit Sub
End If

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    SQLQ = "DELETE FROM HRHELP WHERE HP_TABL_NAME = '" & Data1.Recordset("HP_TABL_NAME") & "' "
    SQLQ = SQLQ & "AND HP_TABL_KEY = '" & Data1.Recordset("HP_TABL_KEY") & "' "
    gdbAdoIhr001.Execute SQLQ
Else
    Exit Sub
End If
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
vbxTrueGrid.SetFocus

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

Private Sub cmdModify_Click()

On Error GoTo Mod_Err

Call modSTUPD(True)

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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HRHELP", "AddNew")
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

Data1.Recordset("HP_COMPNO") = txtComp 'txtCode & ""
Data1.Recordset("HP_TABL_NAME") = "ECCR"
Data1.Recordset.UpdateBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    fglbID = Data1.Recordset("HP_ID")
    Data1.Refresh
    Data1.Recordset.Find "HP_ID=" & fglbID & " "
End If

fglbNewRec% = False


Call modSTUPD(False)

cmdClose.SetFocus

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption

Resume Next
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRHELP", "Update")
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
    Unload frmMHelp
End If

End Sub

Private Sub cmdSelect_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub Form_Load()
Dim SQLQ As String

Screen.MousePointer = HOURGLASS
If glbOracle Then
    SQLQ = "SELECT HRHELP.*, HRTABL.TB_NAME, HRTABL.TB_KEY, HRTABL.TB_DESC FROM HRHELP,HRTABL "
    SQLQ = SQLQ & "WHERE HRHELP.HP_TABL_NAME = HRTABL.TB_NAME  AND HRHELP.HP_TABL_KEY = HRTABL.TB_KEY ORDER BY HRTABL.TB_NAME, HRTABL.TB_DESC"
Else
    SQLQ = "SELECT HRHELP.*, HRTABL.TB_NAME, HRTABL.TB_KEY, HRTABL.TB_DESC FROM HRHELP "
    SQLQ = SQLQ & "LEFT JOIN HRTABL ON HRHELP.HP_TABL_NAME = HRTABL.TB_NAME  AND HRHELP.HP_TABL_KEY = HRTABL.TB_KEY ORDER BY HRTABL.TB_NAME, HRTABL.TB_DESC "
End If
Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = SQLQ '"SELECT * FROM HRHELP ORDER BY HP_TABL_KEY "
Data1.Refresh

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
cmdSelect.Enabled = FT
If gSec_Upd_Company Then '
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
End If
clpCode(0).Enabled = TF
memComments.Enabled = TF

vbxTrueGrid.Enabled = FT

If glbDivInhSel Then
    cmdSelect.Enabled = False
End If

End Sub

Private Sub memComments_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If

        SQLQ = "SELECT * FROM HRHELP"
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag


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

