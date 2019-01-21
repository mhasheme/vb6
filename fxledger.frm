VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLEDGER 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Ledger"
   ClientHeight    =   5940
   ClientLeft      =   1320
   ClientTop       =   720
   ClientWidth     =   6720
   ControlBox      =   0   'False
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
   ScaleHeight     =   5940
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkInactiveCode 
      Alignment       =   1  'Right Justify
      Caption         =   "Inactive Code"
      DataField       =   "GL_INACTIVE"
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
      TabIndex        =   19
      Top             =   4800
      Width           =   1395
   End
   Begin VB.TextBox txtShortCode 
      Appearance      =   0  'Flat
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      MaxLength       =   25
      TabIndex        =   6
      Tag             =   "01-General Ledger Number"
      Top             =   3600
      Visible         =   0   'False
      Width           =   915
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4320
      Top             =   4920
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      TabIndex        =   10
      Top             =   5280
      Width           =   6720
      _Version        =   65536
      _ExtentX        =   11853
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
         Left            =   75
         TabIndex        =   11
         Tag             =   "Select Ledger Description"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   900
         TabIndex        =   12
         Tag             =   "Close and exit this screen"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1710
         TabIndex        =   13
         Tag             =   "Edit the information "
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Tag             =   "Save changes made"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3300
         TabIndex        =   15
         Tag             =   "Cancel changes made"
         Top             =   150
         Width           =   835
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   4200
         TabIndex        =   16
         Tag             =   "Create a new Ledger"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5010
         TabIndex        =   17
         Tag             =   "Delete Ledger listed"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5805
         TabIndex        =   18
         Tag             =   "Print Ledger Listing"
         Top             =   150
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6225
         Top             =   30
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowTitle     =   "General Leager Codes and Descriptions"
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
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   5835
      TabIndex        =   9
      Tag             =   "Find specific record"
      Top             =   4150
      Width           =   735
   End
   Begin VB.TextBox txtFndLgDsc 
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
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   8
      Tag             =   "00-Search Description"
      Top             =   4200
      Width           =   3645
   End
   Begin VB.TextBox txtFndLgNum 
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
      Left            =   120
      MaxLength       =   25
      TabIndex        =   7
      Tag             =   "00-Search for General Ledger Number"
      Top             =   4200
      Width           =   1275
   End
   Begin VB.TextBox txtDescptn 
      Appearance      =   0  'Flat
      DataField       =   "GL_DESCR"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "01-Description of General Ledger"
      Top             =   3600
      Width           =   3645
   End
   Begin VB.TextBox txtLeager 
      Appearance      =   0  'Flat
      DataField       =   "GL_NO"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   100
      MaxLength       =   25
      TabIndex        =   4
      Tag             =   "01-General Ledger Number"
      Top             =   3600
      Width           =   1275
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "GL_DLAST"
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
      Left            =   600
      MaxLength       =   25
      TabIndex        =   0
      Text            =   "Ldate"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "GL_TLAST"
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
      Left            =   2280
      MaxLength       =   25
      TabIndex        =   1
      Text            =   "LTime"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "GL_USER"
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
      Left            =   3960
      MaxLength       =   25
      TabIndex        =   2
      Text            =   "LUser"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1590
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxledger.frx":0000
      Height          =   3360
      Left            =   0
      OleObjectBlob   =   "fxledger.frx":0014
      TabIndex        =   3
      Tag             =   "General Ledger List"
      Top             =   90
      Width           =   6720
   End
End
Attribute VB_Name = "frmLEDGER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRSOld As String, glbEmptyNew  As Integer
Dim fglbNewRec% ' new record
Dim fglOldShortCode As String '14796
Dim FRS As ADODB.Recordset

Private Function chkLgr()
Dim Leager As String, SQLQ As String, Msg$
Dim snapLeagers As New ADODB.Recordset
chkLgr = False
On Error GoTo chkLgr_Err

If Len(txtLeager) < 1 Then
    MsgBox lStr("G/L Number is a required field")
    txtLeager.SetFocus
    Exit Function
End If

If Len(txtDescptn) < 1 Then
    MsgBox lStr("G/L Description is a required field")
    txtDescptn.SetFocus
    Exit Function
End If

If fglbNewRec Then
    Leager = CStr(txtLeager)
    SQLQ = "SELECT GL_NO from HRGL "
    SQLQ = SQLQ & "WHERE GL_NO = '" & Leager & "'"
    SQLQ = SQLQ & "ORDER BY GL_DESCR"
    
    snapLeagers.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If snapLeagers.BOF And snapLeagers.EOF Then
        snapLeagers.Close
    Else
        Msg$ = lStr("This G/L Number already exists")
        MsgBox Msg$
        snapLeagers.Close
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2394W" Then  'St. John #14796
    Dim rsTemp As New ADODB.Recordset
    If Len(txtShortCode.Text) > 0 Then
        If glbCompSerial = "S/N - 2394W" Then  'St. John #14796
            If Not Len(txtShortCode.Text) = 6 Then
                Msg$ = ("Short Code must be 6 characters")
                MsgBox Msg$
                Exit Function
            End If
        End If
        SQLQ = "SELECT * from HRGL "
        SQLQ = SQLQ & "WHERE GL_SHORTNO = '" & txtShortCode & "'"
        If Not fglbNewRec Then
            SQLQ = SQLQ & " AND NOT GL_SHORTNO = '" & fglOldShortCode & "'"
        End If
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            Msg$ = ("This Short Code already exists")
            MsgBox Msg$
            'snapLeagers.Close
            rsTemp.Close
            Exit Function
        End If
        rsTemp.Close
    End If
End If
    
chkLgr = True

Exit Function

chkLgr_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkLeager", "HRGL", "Cancel")
Resume Next

End Function

Private Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err

Data1.Recordset.CancelUpdate
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Call modSTUPD(False)    ' reset screen's attributes
cmdClose.Enabled = True
cmdClose.SetFocus

fglbNewRec% = False
txtLeager.Enabled = False

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HRGL", "Cancel")
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
Dim Ledger As String, Msg$, a%

On Error GoTo DelErr

If Len(txtLeager) < 1 Then Exit Sub

Ledger = CStr(txtLeager)

If chkDelete(Ledger) = False Then Exit Sub

Lok:    'looks ok to me

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> vbYes Then Exit Sub

If glbCompSerial = "S/N - 2394W" Then  'St. John #14796
    If Len(txtShortCode.Text) > 0 Then
        Call Codes_Master_Integration("DEPT", txtShortCode.Text, , True)
    End If
End If

Data1.Recordset.Delete
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh      'laura 03/17/98

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRDEPT", "Delete")
Resume Next

End Sub

Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim Sch As String, SN As String, SQLQ As String, Squ As Double
Dim Pres_Sur As String, SearchField As String, Comp As Integer
Dim KeyAscii As Integer
Dim bkmark

On Error GoTo Srch_Err

If Len(txtFndLgNum) > 0 Then
    SearchField = "GL_NO"
    Pres_Sur = txtFndLgNum
Else
    SearchField = "GL_DESCR"
    Pres_Sur = txtFndLgDsc
End If
If Trim(Pres_Sur) = "" Then
    txtFndLgNum.Enabled = True
    txtFndLgNum.SetFocus
    Exit Sub
End If
    
Sch = Pres_Sur
Squ = InStr(Sch, "'")
If Squ = 1 Then Sch = Right(Sch, Len(Sch) - 1)
If Squ > 1 Then Sch = Left(Sch, Squ - 1)

Screen.MousePointer = HOURGLASS

Data1.Recordset.Requery
If SearchField = "GL_NO" Then
    Data1.Recordset.Find SearchField & " = '" & Sch & "'"
    If Data1.Recordset.EOF Then
        Data1.Refresh
        
        Set FRS = Data1.Recordset.Clone
        vbxTrueGrid.FetchRowStyle = True
        
    Else
        txtFndLgNum = ""
    End If
Else
    Data1.Recordset.Find SearchField & " >= '" & Sch & "'"
    If Data1.Recordset.EOF Then
        Data1.Refresh
    
        Set FRS = Data1.Recordset.Clone
        vbxTrueGrid.FetchRowStyle = True
    
    Else
        txtFndLgDsc = ""
    End If
End If
Screen.MousePointer = DEFAULT

If SearchField = "GL_NO" Then
    txtFndLgNum.SetFocus
Else
    txtFndLgDsc.SetFocus
End If

Exit Sub

Srch_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdFind_Click", "HRGL", "Find Next")
Call RollBack   '08June99 js

End Sub

Private Sub cmdFind_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()

On Error GoTo Mod_Err

Call modSTUPD(True)
'Data1.Recordset.Edit
txtLeager.Enabled = False
If Len(txtShortCode.Text) = 0 Then
    txtShortCode.Enabled = True
Else
    txtShortCode.Enabled = False
End If
txtDescptn.Enabled = True
txtDescptn.SetFocus
If glbCompSerial = "S/N - 2394W" Then  'St. John #14796
    fglOldShortCode = txtShortCode.Text
End If

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack  '08June99 js

End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()

glbCodeRef = True

On Error GoTo NewErr

Call modSTUPD(True)

fglbNewRec% = True

Data1.Recordset.AddNew

chkInactiveCode.Value = 0
txtLeager.Enabled = True
txtLeager.SetFocus

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HRGL", "AddNew")
Resume Next

End Sub

Private Sub CmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim LedgerNbr
Dim LedgerStNbr '#14796

On Error GoTo OK_Err

If Not chkLgr() Then Exit Sub

Call UpdUStats(Me)

LedgerNbr = txtLeager
LedgerStNbr = txtShortCode
Data1.Recordset("GL_NO") = txtLeager & ""
Data1.Recordset.UpdateBatch

If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Data1.Recordset.Find "GL_NO='" & LedgerNbr & "'"

fglbNewRec% = False

Call modSTUPD(False)

txtLeager.Enabled = False

If glbCompSerial = "S/N - 2394W" Then  'St. John #14796
    If Len(LedgerStNbr) > 0 Then
        Call Codes_Master_Integration("DEPT", LedgerStNbr)
    End If
End If

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRGL", "Update")
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

RHeading = "Ledgers"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdSelect_Click()

glbLgr = Data1.Recordset("GL_NO")
glbLgrDesc = Data1.Recordset("GL_DESCR")
Unload Me

End Sub

Private Sub cmdSelect_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRLgr", "SELECT")

End Sub

Private Sub Form_Activate()
Data1.RecordSource = "SELECT * FROM HRGL order by GL_INACTIVE, GL_DESCR"
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

End Sub

Private Sub Form_Load()
Dim SQLQ As String

glbOnTop = "FRMLEDGER"

If glbCompSerial = "S/N - 2394W" Then 'St. John's Rehab Hospital Ticket #14796
    txtShortCode.DataField = "GL_SHORTNO"
    txtShortCode.Visible = True
    vbxTrueGrid.Columns(2).DataField = "GL_SHORTNO"
Else
    vbxTrueGrid.Columns(2).Visible = False
End If

SQLQ = "UPDATE HRGL SET GL_INACTIVE = 0 WHERE GL_INACTIVE IS NULL"
gdbAdoIhr001.Execute SQLQ

Screen.MousePointer = HOURGLASS
Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "SELECT * FROM HRGL order by GL_INACTIVE, GL_DESCR"

Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Call setCaption(Me)
Call setCaption(Me.vbxTrueGrid.Columns(0))
Call setCaption(Me.vbxTrueGrid.Columns(1))

Call modSTUPD(False)

If Not gSec_Upd_Ledgers Then
    cmdModify.Enabled = False
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
cmdNew.Enabled = FT         '
cmdClose.Enabled = FT       '
cmdModify.Enabled = FT      '
cmdDelete.Enabled = FT      '
cmdPrint.Enabled = FT       '
cmdFind.Enabled = FT        '
txtFndLgNum.Enabled = FT
txtFndLgDsc.Enabled = FT
vbxTrueGrid.Enabled = FT    'Jaddy 11/12/99
txtDescptn.Enabled = TF     '
txtShortCode.Enabled = TF
txtLeager.Enabled = TF      '
chkInactiveCode.Enabled = TF

If Data1.Recordset.EOF Or Data1.Recordset.BOF Or glbLgrInhSel Then
    cmdSelect.Enabled = False
Else
    cmdSelect.Enabled = True
End If
End Sub

Private Sub txtDescptn_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFndLgDsc_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFndLgNum_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFndLgNum_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtLeager_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtLeager_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub vbxTrueGrid_DblClick()

glbLgr = Data1.Recordset("GL_NO")
glbLgrDesc = Data1.Recordset("GL_DESCR")
Unload Me

End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If Not fglbNewRec% Then
        FRS.Requery
        FRS.Bookmark = Bookmark
        If FRS("GL_INACTIVE") Then
            RowStyle.ForeColor = vbRed
        End If
    End If
End Sub

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
        
        SQLQ = "select * from HRGL " 'order by GL_DESCR"
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
        
    Set FRS = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True

End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the enter key was struck
    KeyAscii = 0
    If Not Me.vbxTrueGrid.EditActive Then   '08June99 js
        cmdClose.SetFocus                   '
    End If                                  '
End If

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

Private Function chkDelete(xLedger) As Boolean
On Error GoTo chkLgr_Err
Dim SQLQ As String, snapEELgrs As New ADODB.Recordset
Dim Msg As String

chkDelete = False

SQLQ = "SELECT DF_NBR, DF_NAME FROM HRDEPT "
SQLQ = SQLQ & "WHERE DF_GLNO = '" & xLedger & "'"

snapEELgrs.Open SQLQ, gdbAdoIhr001, adOpenStatic
If snapEELgrs.BOF = False And snapEELgrs.EOF = False Then
    Msg$ = lStr("Department presently assigned to this G/L#")
    Msg$ = Msg$ & Chr(10) & CStr(snapEELgrs("DF_NBR"))
    Msg$ = Msg$ & Chr(10) & snapEELgrs("DF_NAME")
    Msg$ = Msg$ & Chr(10) & "Delete aborted."
    MsgBox Msg$
    snapEELgrs.Close
    Exit Function
End If
snapEELgrs.Close

SQLQ = "SELECT DISTINCT AD_EMPNBR FROM HR_ATTENDANCE "
If glbCompSerial = "S/N - 2192W" Then
    Msg = "AD_CHRGCODE"
Else
    Msg = "AD_GLNO"
End If
SQLQ = SQLQ & "WHERE " & Msg & "='" & xLedger & "'"
snapEELgrs.Open SQLQ, gdbAdoIhr001, adOpenStatic
If snapEELgrs.BOF = False And snapEELgrs.EOF = False Then
    Msg$ = lStr("Attendance records are presently assigned to this G/L#")
'    Msg$ = Msg$ & Chr(10) & CStr(snapEELgrs("AD_EMPNBR"))
    Msg$ = Msg$ & Chr(10) & "Delete aborted."
    MsgBox Msg$
    snapEELgrs.Close
    Exit Function
End If
snapEELgrs.Close


SQLQ = "SELECT DISTINCT AH_EMPNBR FROM HR_ATTENDANCE_HISTORY "
If glbCompSerial = "S/N - 2192W" Then
    Msg = "AH_CHRGCODE"
Else
    Msg = "AH_GLNO"
End If
SQLQ = SQLQ & "WHERE " & Msg & "='" & xLedger & "'"
snapEELgrs.Open SQLQ, gdbAdoIhr001, adOpenStatic
If snapEELgrs.BOF = False And snapEELgrs.EOF = False Then
    Msg$ = lStr("Attendance History records are presently assigned to this G/L#")
'    Msg$ = Msg$ & Chr(10) & CStr(snapEELgrs("AD_EMPNBR"))
    Msg$ = Msg$ & Chr(10) & "Delete aborted."
    MsgBox Msg$
    snapEELgrs.Close
    Exit Function
End If
snapEELgrs.Close


SQLQ = "SELECT DISTINCT AD_EMPNBR FROM TERM_ATTENDANCE "
If glbCompSerial = "S/N - 2192W" Then
    Msg = "AD_CHRGCODE"
Else
    Msg = "AD_GLNO"
End If
SQLQ = SQLQ & "WHERE " & Msg & "='" & xLedger & "'"
snapEELgrs.Open SQLQ, gdbAdoIhr001, adOpenStatic
If snapEELgrs.BOF = False And snapEELgrs.EOF = False Then
    Msg$ = lStr("Terminated Attendance records are presently assigned to this G/L#")
'    Msg$ = Msg$ & Chr(10) & CStr(snapEELgrs("AD_EMPNBR"))
    Msg$ = Msg$ & Chr(10) & "Delete aborted."
    MsgBox Msg$
    snapEELgrs.Close
    Exit Function
End If
snapEELgrs.Close


SQLQ = "SELECT DISTINCT JH_EMPNBR FROM HR_JOB_HISTORY "
SQLQ = SQLQ & "WHERE JH_GLNO='" & xLedger & "'"
snapEELgrs.Open SQLQ, gdbAdoIhr001, adOpenStatic
If snapEELgrs.BOF = False And snapEELgrs.EOF = False Then
    Msg$ = lStr("Position records are presently assigned to this G/L#")
'    Msg$ = Msg$ & Chr(10) & CStr(snapEELgrs("AD_EMPNBR"))
    Msg$ = Msg$ & Chr(10) & "Delete aborted."
    MsgBox Msg$
    snapEELgrs.Close
    Exit Function
End If
snapEELgrs.Close


SQLQ = "SELECT DISTINCT JH_EMPNBR FROM TERM_JOB_HISTORY "
SQLQ = SQLQ & "WHERE JH_GLNO='" & xLedger & "'"
snapEELgrs.Open SQLQ, gdbAdoIhr001, adOpenStatic
If snapEELgrs.BOF = False And snapEELgrs.EOF = False Then
    Msg$ = lStr("Position records are presently assigned to this G/L#")
'    Msg$ = Msg$ & Chr(10) & CStr(snapEELgrs("AD_EMPNBR"))
    Msg$ = Msg$ & Chr(10) & "Delete aborted."
    MsgBox Msg$
    snapEELgrs.Close
    Exit Function
End If
snapEELgrs.Close


SQLQ = "SELECT MACHINE_NUM, DESCRIPTION FROM HR_MACHINE_NUM "
SQLQ = SQLQ & "WHERE GL_NUMBER='" & xLedger & "'"
snapEELgrs.Open SQLQ, gdbAdoIhr001, adOpenStatic
If snapEELgrs.BOF = False And snapEELgrs.EOF = False Then
    Msg$ = lStr("Position records are presently assigned to this G/L#")
    Msg$ = Msg$ & Chr(10) & CStr(snapEELgrs("MACHINE_NUM"))
    Msg$ = Msg$ & Chr(10) & snapEELgrs("DESCRIPTION")
    Msg$ = Msg$ & Chr(10) & "Delete aborted."
    MsgBox Msg$
    snapEELgrs.Close
    Exit Function
End If
snapEELgrs.Close

chkDelete = True

Exit Function

chkLgr_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkLeager", "HRGL", "Cancel")
Resume Next
End Function


