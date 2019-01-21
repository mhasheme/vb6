VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmHomeMaster 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Home Master"
   ClientHeight    =   6240
   ClientLeft      =   1080
   ClientTop       =   1050
   ClientWidth     =   8175
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
   ScaleHeight     =   6240
   ScaleWidth      =   8175
   Begin VB.TextBox txtShowKey 
      Appearance      =   0  'Flat
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1260
      MaxLength       =   12
      TabIndex        =   2
      Tag             =   "01-Code"
      Top             =   4380
      Width           =   1965
   End
   Begin VB.TextBox txtTable 
      Appearance      =   0  'Flat
      DataField       =   "TB_NAME"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4560
      Top             =   5400
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
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   21
      Top             =   5580
      Width           =   8175
      _Version        =   65536
      _ExtentX        =   14420
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
         Left            =   120
         TabIndex        =   8
         Tag             =   "Select the Code listed above"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   6240
         TabIndex        =   15
         Tag             =   "Print Code Listing Report"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5400
         TabIndex        =   14
         Tag             =   "Delete code listed above"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   4530
         TabIndex        =   13
         Tag             =   "Add a new Code"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   12
         Tag             =   "Cancel the changes made"
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Tag             =   "Save the changes made"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Tag             =   "Edit the Information"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Tag             =   "Close and exit this screen"
         Top             =   120
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
      DataField       =   "TB_COMPNO"
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
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5250
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      DataField       =   "TB_DESC"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3330
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "01-Description of the Code"
      Top             =   4380
      Width           =   3795
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   7290
      TabIndex        =   7
      Tag             =   "Find specific record"
      Top             =   4830
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
      Left            =   3330
      MaxLength       =   30
      TabIndex        =   6
      Tag             =   "00-Search Description"
      Top             =   4860
      Width           =   3795
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
      Left            =   1260
      MaxLength       =   12
      TabIndex        =   5
      Tag             =   "00-Search Code"
      Top             =   4860
      Width           =   1965
   End
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      DataField       =   "TB_KEY"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1290
      MaxLength       =   16
      TabIndex        =   19
      Tag             =   "01-Code"
      Top             =   4380
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TB_LUSER"
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
      Left            =   3240
      MaxLength       =   25
      TabIndex        =   18
      Text            =   "LUser"
      Top             =   5865
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TB_LTIME"
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
      TabIndex        =   17
      Text            =   "LTime"
      Top             =   5865
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TB_LDATE"
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
      Left            =   120
      MaxLength       =   25
      TabIndex        =   16
      Text            =   "Ldate"
      Top             =   5865
      Visible         =   0   'False
      Width           =   1590
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxhmmst.frx":0000
      Height          =   4185
      Left            =   90
      OleObjectBlob   =   "fxhmmst.frx":0014
      TabIndex        =   0
      Tag             =   "Codes Listings"
      Top             =   0
      Width           =   7875
   End
   Begin INFOHR_Controls.CodeLookup clpDIV 
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Top             =   4380
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpFindDIV 
      Height          =   285
      Left            =   30
      TabIndex        =   4
      Top             =   4860
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin VB.Label lblDIVDesc 
      AutoSize        =   -1  'True
      Caption         =   "Unassigned"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   7290
      TabIndex        =   23
      Top             =   5730
      Visible         =   0   'False
      Width           =   840
   End
End
Attribute VB_Name = "frmHomeMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim fglbNewRec%
Dim fglbUDMode As Integer
Dim fglbRSOld As String
Dim ODIV, ODivD, xGlbDiv, xGlbDivDesc
Private Function chkMastTable()
Dim SQLQ As String, Msg$, Tabl As String, Ky As String
Dim snapTabs As New ADODB.Recordset

On Error GoTo chkMastTable_Err

chkMastTable = False

If txtTable = "HMOP" Or txtTable = "HMLN" Then
    If Len(clpDiv) < 1 Then
        MsgBox lStr("Division is a required field")
        clpDiv.SetFocus
        Exit Function
    Else
        If clpDiv.Caption = "Unassigned" Then
            MsgBox lStr("If Division Entered - it must be known")
            clpDiv.SetFocus
            Exit Function
        End If
    End If
End If

If Len(txtShowKey) < 1 Then
    MsgBox "Key (or Code) is a required field"
    txtShowKey.SetFocus
    Exit Function
End If

If Len(txtDesc) < 1 Then
    MsgBox "Description is a required field"
    txtDesc.SetFocus
    Exit Function
End If

If fglbNewRec Then
    Tabl = txtTable
    If glbTabNam = "HMOP" Or glbTabNam = "HMLN" Then
        Ky = clpDiv & txtShowKey
    Else
        Ky = txtShowKey
    End If
    SQLQ = "SELECT TB_NAME, TB_KEY from LN_HOMES "
    SQLQ = SQLQ & "WHERE TB_NAME = '" & Tabl & "' "
    SQLQ = SQLQ & " AND TB_KEY = '" & Ky & "'"
    snapTabs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If snapTabs.BOF And snapTabs.EOF Then
        snapTabs.Close
    Else
        Msg$ = "Code already exists in database"
        MsgBox Msg$
        snapTabs.Close
        Exit Function
    End If
End If

chkMastTable = True

Exit Function

chkMastTable_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HRTABLE", "HRTABL", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function





Private Sub cmdCancel_Click()
On Error GoTo Can_Err

Data1.Recordset.CancelBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
Call ST_UPD_MODE(False)  ' reset screen's attributes

fglbNewRec% = False

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRTABL", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If




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
Dim rsEMP As New ADODB.Recordset
Dim SQLQ, Msg$

On Error GoTo DelErr

If Data1.Recordset.RecordCount < 2 Then
    MsgBox "You can not delete the last reference for this code"
    Exit Sub
End If

If txtTable = "EDSK" Then 'Ticket# 7459
    SQLQ = "SELECT SE_EMPNBR, SE_EMPNBR FROM HREMPSKL "
Else
    SQLQ = "SELECT ED_EMPNBR, ED_SURNAME, ED_DEPTNO FROM HREMP "
End If
Select Case txtTable
Case "HMOP"
    SQLQ = SQLQ & "WHERE ED_HOMEOPRTNBR = '" & txtKey & "'"
Case "HMSF"
    SQLQ = SQLQ & "WHERE ED_HOMESHIFT = '" & txtKey & "'"
Case "HMLN"
    SQLQ = SQLQ & "WHERE ED_HOMELINE = '" & txtKey & "'"
Case "HMWC"
    SQLQ = SQLQ & "WHERE ED_HOMEWRKCNT = '" & txtKey & "'"
Case "EDSK"
    SQLQ = SQLQ & "WHERE SE_SKILL = '" & txtKey & "'"
End Select

rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic

If rsEMP.BOF And rsEMP.EOF Then
    GoTo Lok
Else
    Msg$ = lStr("Employee presently assigned to this Code")
    If txtTable <> "EDSK" Then
    Msg$ = Msg$ & Chr(10) & ShowEmpnbr(rsEMP("ED_EMPNBR"))
    Msg$ = Msg$ & Chr(10) & rsEMP("ED_SURNAME")
    End If
    Msg$ = Msg$ & Chr(10) & "Delete aborted."
    MsgBox Msg$
    rsEMP.Close
    Exit Sub
End If

Lok:        'looks ok to me
rsEMP.Close

Data1.Recordset.Delete
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Exit Sub
DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRTable", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub cmdDelete_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim SQLQ As String
If Len(txtFindKey) > 0 Then
    If clpFindDIV = "" Then clpFindDIV = clpDiv
    SQLQ = "TB_KEY >= '" & clpFindDIV & txtFindKey & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
    Else
        txtFindKey = ""
    End If
    clpFindDIV = ""
    Exit Sub
End If

If Len(txtFindDesc) > 0 Then
    SQLQ = "TB_DESC >= '" & txtFindDesc & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
    Else
        txtFindDesc = ""
    End If
    clpFindDIV = ""
    Exit Sub
End If


End Sub

Private Sub cmdFind_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()


On Error GoTo Mod_Err
Call ST_UPD_MODE(True)
clpDiv.Enabled = False
txtShowKey.Enabled = False
txtDesc.SetFocus

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub cmdModify_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()

On Error GoTo NewErr


Call ST_UPD_MODE(True)
fglbNewRec% = True

'glbhomeRef = True


Data1.Recordset.AddNew
If glbTabNam = "HMOP" Or glbTabNam = "HMLN" Then
    If glbDIVCount = 1 Then clpDiv = glbSDIV
    If glbTransDiv <> "" Then clpDiv = glbTransDiv
End If
txtShowKey = ""
txtTable.Text = glbTabNam
txtComp.Text = glbCompNo

If clpDiv.Visible Then
    clpDiv.SetFocus
Else
    txtShowKey.SetFocus
End If

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HRTABLE", "HRTABL", "add new")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub CmdNew_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim SQLQ As String
Dim strK
On Error GoTo OK_Err

'glbhomeRef = True   'table entrie modified/added - forces refresh
                    ' at form level of codes/descriptions.
txtKey = clpDiv & txtShowKey
If Not chkMastTable() Then Exit Sub

Call UpdUStats(Me)
strK = txtKey

Data1.Recordset("TB_KEY") = clpDiv & txtShowKey
Data1.Recordset.UpdateBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Data1.Recordset.Find "TB_KEY = '" & strK & "'"
Call ST_UPD_MODE(False)
fglbNewRec% = False


Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRTABL", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Private Sub cmdOK_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPrint_Click()

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

Me.vbxCrystal.ReportTitle = "Table Codes for - " & glbTabNam
Me.vbxCrystal.BoundReportHeading = frmHomeMaster.Caption
Me.vbxCrystal.WindowTitle = frmHomeMaster.Caption & " Report"
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdSelect_Click()


If Data1.Recordset.EOF And Data1.Recordset.BOF Then
  Exit Sub
End If

If Len(Data1.Recordset("TB_KEY")) > 0 Then
    glbHome = Data1.Recordset("TB_SHOWKEY")
    glbTransDiv = Data1.Recordset("TB_DIV")
    glbHomeDesc = Data1.Recordset("TB_DESC")
    Unload frmHomeMaster
Else
    Exit Sub
End If


End Sub

Private Sub cmdSelect_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "data1 error", "HRTABL", "SELECT")


End Sub

Private Sub Form_Activate()
If glbTabNam = "HMSF" Then
    txtKey.MaxLength = 5
    txtFindKey.MaxLength = 5
End If
End Sub

Private Sub Form_Load()
Dim SQLQ As String
Dim VS
Screen.MousePointer = HOURGLASS
glbOnTop = "FRMHOMEMASTER"
'glbhomeRef = False  'table entrie modified/added false
VS = True
SQLQ = " SELECT * FROM QRY_LN_HOMES WHERE TB_NAME='" & glbTabNam & "'"

If glbTabNam = "HMOP" Or glbTabNam = "HMLN" Then
    SQLQ = SQLQ & " AND TB_DIV in " & glbDIVList
    If glbTransDiv <> "" Then
        SQLQ = SQLQ & " AND (TB_DIV = '" & glbTransDiv & "' or TB_DIV='ALL')"
    End If
Else
    VS = False
End If
If glbDIVCount = 1 Then VS = False
If glbTransDiv <> "" Then VS = False
clpDiv.Visible = VS
clpFindDIV.Visible = VS
vbxTrueGrid.Columns(0).Visible = VS
If Not VS Then
    txtShowKey.Left = clpDiv.Left
    txtFindKey.Left = clpDiv.Left
End If
Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = SQLQ
Data1.Refresh

glbHome = ""    'set to null - implies none found/cancel
glbHomeDesc = ""

    Call INI_Controls(Me)

Call ST_UPD_MODE(False)

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

cmdOK.Enabled = TF
cmdCancel.Enabled = TF

cmdModify.Enabled = FT
cmdClose.Enabled = FT
cmdNew.Enabled = FT
cmdDelete.Enabled = FT
cmdPrint.Enabled = FT
cmdFind.Enabled = FT
cmdSelect.Enabled = FT


clpDiv.Enabled = TF
txtShowKey.Enabled = TF
txtKey.Enabled = TF
txtDesc.Enabled = TF

clpFindDIV.Enabled = FT
txtFindKey.Enabled = FT
txtFindDesc.Enabled = FT
vbxTrueGrid.Enabled = FT
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    cmdSelect.Enabled = False
    cmdModify.Enabled = False
    cmdFind.Enabled = False
    cmdDelete.Enabled = False
End If
If glbHOMEInhSel Then
    cmdSelect.Enabled = False
End If
On Error GoTo ERR_EXIT
If Not gSec_Upd_Master_Table(glbTabNam) Then
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
End If
ERR_EXIT:
If Err.Number = 5 Then
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
End If
End Sub

Private Sub txtDesc_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub





Private Sub txtFindDesc_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub txtFindKey_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtKey_Change()
If txtTable = "HMOP" Or txtTable = "HMLN" Then
    clpDiv = Left(txtKey, 3)
    txtShowKey = Mid(txtKey, 4)
Else
    clpDiv = ""
    txtShowKey = txtKey
End If
End Sub

Private Sub txtKey_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Private Sub txtKey_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtShowKey_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtShowKey_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub vbxTrueGrid_DblClick()

If Not Me.vbxTrueGrid.EditActive Then
    If Not (Data1.Recordset.BOF Or Data1.Recordset.EOF) Then
        glbHome = Data1.Recordset("TB_SHOWKEY")
        glbTransDiv = Data1.Recordset("TB_DIV")
        glbHomeDesc = Data1.Recordset("TB_DESC")
    End If
    Unload frmHomeMaster
Else
    MsgBox "Save/cancel changes first"
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
        
        SQLQ = " SELECT * FROM QRY_LN_HOMES WHERE TB_NAME='" & glbTabNam & "'"

        If glbTabNam = "HMOP" Or glbTabNam = "HMLN" Then
            SQLQ = SQLQ & " AND TB_DIV in " & glbDIVList
            If glbTransDiv <> "" Then
                SQLQ = SQLQ & " AND (TB_DIV = '" & glbTransDiv & "' or TB_DIV='ALL')"
            End If
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub
