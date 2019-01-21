VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmMTABLin 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Codes"
   ClientHeight    =   6015
   ClientLeft      =   1350
   ClientTop       =   1650
   ClientWidth     =   8550
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
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6015
   ScaleWidth      =   8550
   Begin VB.TextBox txtWaitPeriod 
      Appearance      =   0  'Flat
      DataField       =   "TB_USR2"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1380
      MaxLength       =   7
      TabIndex        =   4
      Tag             =   "00-Waiting Period"
      Top             =   4470
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.ComboBox cmbDWM 
      Height          =   315
      ItemData        =   "fxmtablin.frx":0000
      Left            =   2250
      List            =   "fxmtablin.frx":000D
      TabIndex        =   5
      Tag             =   "40-Select Day, Week or Month"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtDWM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "TB_USR1"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3630
      MaxLength       =   7
      TabIndex        =   29
      Tag             =   "01-Department - Code"
      Top             =   4500
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtShowKey 
      Appearance      =   0  'Flat
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1380
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "01-Code"
      Top             =   4080
      Width           =   1125
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   18
      Top             =   5355
      Width           =   8550
      _Version        =   65536
      _ExtentX        =   15081
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
         TabIndex        =   19
         Tag             =   "Select the Code listed above"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   855
         TabIndex        =   20
         Tag             =   "Close and exit this screen"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1635
         TabIndex        =   21
         Tag             =   "Edit the Information"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2415
         TabIndex        =   22
         Tag             =   "Save the changes made"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3225
         TabIndex        =   23
         Tag             =   "Cancel the changes made"
         Top             =   150
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   4080
         TabIndex        =   24
         Tag             =   "Add a new Code"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4890
         TabIndex        =   25
         Tag             =   "Delete code listed above"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5700
         TabIndex        =   26
         Tag             =   "Print Code Listing Report"
         Top             =   150
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   5505
         Top             =   150
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
      Begin MSAdodcLib.Adodc Data1 
         Height          =   480
         Left            =   4770
         Top             =   30
         Visible         =   0   'False
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   847
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
   End
   Begin VB.CheckBox chkSen 
      Alignment       =   1  'Right Justify
      Caption         =   "Seniority"
      DataField       =   "TB_SEN"
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
      Left            =   8160
      TabIndex        =   11
      Top             =   4110
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CheckBox ChkInc 
      Alignment       =   1  'Right Justify
      Caption         =   "Incentive"
      DataField       =   "TB_INDICATOR"
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
      Left            =   7110
      TabIndex        =   6
      Top             =   4110
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   285
      Left            =   7110
      TabIndex        =   10
      Tag             =   "Find specific record"
      Top             =   4920
      Width           =   840
   End
   Begin VB.TextBox txtFindDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2550
      TabIndex        =   9
      Tag             =   "00-Search Description"
      Top             =   4920
      Width           =   4410
   End
   Begin VB.TextBox txtFindKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1380
      MaxLength       =   4
      TabIndex        =   8
      Tag             =   "00-Search Code"
      Top             =   4920
      Width           =   1125
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      DataField       =   "TB_COMPNO"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5130
      MaxLength       =   3
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5550
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtTable 
      Appearance      =   0  'Flat
      DataField       =   "TB_NAME"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5550
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      DataField       =   "TB_DESC"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2550
      MaxLength       =   30
      TabIndex        =   3
      Top             =   4080
      Width           =   4410
   End
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      DataField       =   "TB_KEY"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1410
      TabIndex        =   12
      Tag             =   "01-Code"
      Top             =   4080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TB_LUSER"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   3300
      MaxLength       =   25
      TabIndex        =   15
      Text            =   "LUser"
      Top             =   5550
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TB_LTIME"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   1740
      MaxLength       =   25
      TabIndex        =   14
      Text            =   "LTime"
      Top             =   5550
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TB_LDATE"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   180
      MaxLength       =   25
      TabIndex        =   13
      Text            =   "Ldate"
      Top             =   5550
      Visible         =   0   'False
      Width           =   1590
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxmtablin.frx":002C
      Height          =   3885
      Left            =   180
      OleObjectBlob   =   "fxmtablin.frx":0040
      TabIndex        =   0
      Tag             =   "Codes Listings"
      Top             =   0
      Width           =   8235
   End
   Begin INFOHR_Controls.CodeLookup clpDIV 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpFindDIV 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting Period"
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
      Index           =   0
      Left            =   180
      TabIndex        =   30
      Top             =   4500
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblDeptDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
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
      Left            =   3390
      TabIndex        =   28
      Top             =   5310
      Visible         =   0   'False
      Width           =   840
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
      Left            =   7110
      TabIndex        =   27
      Top             =   5640
      Visible         =   0   'False
      Width           =   900
   End
End
Attribute VB_Name = "frmMTABLin"
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

If Len(clpDIV) < 1 Then
    MsgBox lStr("Division is a required field")
    clpDIV.SetFocus
    Exit Function
Else
    If clpDIV.Caption = "Unassigned" Then
        MsgBox lStr("If Division Entered - it must be known")
        clpDIV.SetFocus
        Exit Function
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
If Len(txtWaitPeriod) > 0 And txtWaitPeriod.Visible Then
    If IsNumeric(txtWaitPeriod) Then
        If cmbDWM.ListIndex = -1 Then
            MsgBox "Please select Day/Week/Month"
            cmbDWM.SetFocus
            Exit Function
        End If
    Else
        MsgBox "Waiting Period must be numeric"
        txtWaitPeriod.SetFocus
        Exit Function
    End If
    
End If
If fglbNewRec Then
    Tabl = txtTable
    Ky = clpDIV & txtShowKey
    SQLQ = "SELECT TB_NAME, TB_KEY from HRTABL "
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





Private Sub cmbDWM_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub
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
SQLQ = "SELECT ED_EMPNBR, ED_SURNAME, ED_DEPTNO FROM HREMP "
Select Case txtTable
Case "EDSE"
    SQLQ = SQLQ & "WHERE ED_SECTION = '" & txtKey & "'"
Case "EDRG"
    SQLQ = SQLQ & "WHERE ED_REGION = '" & txtKey & "'"
Case "BNCD"
    SQLQ = SQLQ & "WHERE ED_EMPNBR IN (SELECT BF_EMPNBR FROM HRBENFT WHERE BF_BCODE = '" & txtKey & "')"
End Select

rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic

If rsEMP.BOF And rsEMP.EOF Then
    GoTo Lok
Else
    Msg$ = lStr("Employee presently assigned to this Code")
    Msg$ = Msg$ & Chr(10) & ShowEmpnbr(rsEMP("ED_EMPNBR"))
    Msg$ = Msg$ & Chr(10) & rsEMP("ED_SURNAME")
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
    If clpFindDIV = "" Then clpFindDIV = clpDIV
    SQLQ = "TB_KEY >= '" & clpDIV & txtFindKey & "'"
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
clpDIV.Enabled = False
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

glbCodeRef = True


Data1.Recordset.AddNew
If glbDIVCount = 1 Then clpDIV = glbSDIV
If glbTransDiv <> "" Then clpDIV = glbTransDiv
txtTable.Text = glbTabNam
txtComp.Text = glbCompNo

If clpDIV.Visible Then
    clpDIV.SetFocus
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

glbCodeRef = True   'table entrie modified/added - forces refresh
                    ' at form level of codes/descriptions.

If Not chkMastTable() Then Exit Sub

Call UpdUStats(Me)
strK = clpDIV & txtShowKey
txtKey = strK
If cmbDWM.Visible Then
    txtDWM = Left(cmbDWM, 1)
    If cmbDWM.ListIndex <> -1 And Len(txtWaitPeriod) = 0 Then
        txtWaitPeriod = 0
    End If
    If txtWaitPeriod = "" And txtWaitPeriod.DataChanged Then txtWaitPeriod.DataChanged = False: Data1.Recordset("TB_USR2") = Null
End If
Data1.Recordset("TB_KEY") = strK
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
Me.vbxCrystal.BoundReportHeading = frmMTABLin.Caption
Me.vbxCrystal.WindowTitle = frmMTABLin.Caption & " Report"
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
    glbCode = Data1.Recordset("TB_SHOWKEY")
    glbTransDiv = Data1.Recordset("TB_DIV")
    glbCodeDesc = Data1.Recordset("TB_DESC")
    Unload frmMTABLin
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
If glbCompSerial = "S/N - 2191W" And glbTabNam = "JBGC" Then
    txtKey.MaxLength = 6
    txtFindKey.MaxLength = 6
End If
End Sub

Private Sub Form_Load()

Dim SQLQ As String
Dim VS
Screen.MousePointer = HOURGLASS
glbOnTop = "FRMmtaBLIN"
glbCodeRef = False  'table entrie modified/added false
     

SQLQ = "SELECT * FROM QRY_LN_TABL WHERE TB_NAME = '" & glbTabNam & "'"
SQLQ = SQLQ & " AND TB_DIV in " & glbDIVList
If glbTransDiv <> "" Then
    SQLQ = SQLQ & " AND (TB_DIV = '" & glbTransDiv & "' or TB_DIV = 'ALL')"
End If
If glbTabNam = "EDSE" Then
    SQLQ = SQLQ & " AND " & glbSeleSection
Else
    txtShowKey.MaxLength = 8
    txtFindKey.MaxLength = 8
End If

SQLQ = SQLQ & " ORDER BY TB_DIV,TB_NAME, TB_DESC, TB_KEY"

Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = SQLQ
Data1.Refresh
VS = True
If glbDIVCount = 1 Then VS = False
If glbTransDiv <> "" Then VS = False
clpDIV.Visible = VS
clpFindDIV.Visible = VS
vbxTrueGrid.Columns(0).Visible = VS
If Not VS Then
    txtShowKey.Left = clpDIV.Left
    txtFindKey.Left = clpDIV.Left
End If
glbCode = ""    'set to null - implies none found/cancel
glbCodeDesc = ""
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

ChkInc.Enabled = TF
chkSen.Enabled = TF

txtWaitPeriod.Enabled = TF
cmbDWM.Enabled = TF

clpDIV.Enabled = TF
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

Private Sub txtDWM_Change()
cmbDWM.ListIndex = -1
Select Case txtDWM
Case "D"
    cmbDWM.ListIndex = 0
Case "W"
    cmbDWM.ListIndex = 1
Case "M"
    cmbDWM.ListIndex = 2
End Select
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

Private Sub txtKey_Change()

txtShowKey = Mid(txtKey, 4)
clpDIV = Left(txtKey, 3)

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

Private Sub txtTable_Change()
Dim xTabName

If Len(Trim(txtTable)) = 0 Then
    xTabName = glbTabNam
Else
    xTabName = txtTable
End If
If xTabName = "BNCD" Then
    lblTitle(0).Visible = True
    txtWaitPeriod.Visible = True
    cmbDWM.Visible = True
Else
    lblTitle(0).Visible = False
    txtWaitPeriod.Visible = False
    cmbDWM.Visible = False
End If
End Sub

Private Sub txtWaitPeriod_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub
Private Sub vbxTrueGrid_DblClick()

If Not Me.vbxTrueGrid.EditActive Then
    If Not (Data1.Recordset.BOF Or Data1.Recordset.EOF) Then
        glbCode = Data1.Recordset("TB_SHOWKEY")
        glbTransDiv = Data1.Recordset("TB_DIV")
        glbCodeDesc = Data1.Recordset("TB_DESC")
    End If
    Unload frmMTABLin
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
        
        SQLQ = "SELECT * FROM QRY_LN_TABL WHERE TB_NAME = '" & glbTabNam & "'"
        SQLQ = SQLQ & " AND TB_DIV in " & glbDIVList
        If glbTransDiv <> "" Then
            SQLQ = SQLQ & " AND (TB_DIV = '" & glbTransDiv & "' or TB_DIV = 'ALL')"
        End If
        If glbTabNam = "EDSE" Then
            SQLQ = SQLQ & " AND " & glbSeleSection
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub
