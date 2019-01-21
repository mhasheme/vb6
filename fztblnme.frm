VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTblName 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Table Master Report"
   ClientHeight    =   7365
   ClientLeft      =   705
   ClientTop       =   1605
   ClientWidth     =   9810
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7365
   ScaleWidth      =   9810
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7560
      Top             =   6000
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin Threed.SSPanel panTablname 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   600
      Width           =   915
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   15
      Caption         =   "All Tables"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   21.71
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   1
      Font3D          =   1
      Alignment       =   6
      Enabled         =   0   'False
      Begin VB.TextBox txtTblName 
         Appearance      =   0  'Flat
         DataField       =   "TD_NAME"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   910
      End
   End
   Begin Threed.SSCheck chkTable 
      Height          =   255
      Left            =   630
      TabIndex        =   2
      Top             =   1350
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "All Tables                         "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Font3D          =   3
   End
   Begin Threed.SSCheck ChkPage 
      Height          =   255
      Left            =   630
      TabIndex        =   3
      Top             =   1830
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Page Break on Table Name"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Alignment       =   1
      Font3D          =   3
   End
   Begin TrueOleDBGrid60.TDBGrid tblTables 
      Bindings        =   "fztblnme.frx":0000
      Height          =   4875
      Left            =   4350
      OleObjectBlob   =   "fztblnme.frx":0014
      TabIndex        =   5
      Tag             =   "Tables Names Lookup"
      Top             =   600
      Width           =   4335
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6720
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowWidth     =   480
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   2
      WindowControls  =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "List of Table Names:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   4330
      TabIndex        =   4
      Top             =   330
      Width           =   1770
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Table Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   630
      TabIndex        =   0
      Top             =   690
      Width           =   1035
   End
End
Attribute VB_Name = "frmTblName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Tabl_Snap As New ADODB.Recordset


Private Sub chkTable_Click(Value As Integer)
If chkTable.Value = True Then
    ChkPage.Enabled = True
    ChkPage.Value = True
    txtTblName.Visible = False
Else
    ChkPage.Enabled = False
    ChkPage.Value = False
    txtTblName.Visible = True
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub



Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Table Master Report Criteria", Me) Then Exit Sub
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Call set_PrintState(False)
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    'Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztable.rpt"
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

PrntErr:
MsgBox "Error Printing - check your Windows Printer setup"
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub



Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Call set_PrintState(False)
    Screen.MousePointer = HOURGLASS
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
    Call set_PrintState(True)
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ATTEND", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub



Private Sub CR_Table_Snap()
Dim SQLQ As String
Dim countr   As Integer  ' CR_EENames_Snap is definded at form level
                                                 
On Error GoTo CR_EENames_Err
         
Screen.MousePointer = HOURGLASS

Data1.RecordSource = "SELECT TD_NAME, TD_DESC,TD_SHOWDESC FROM HRTABDES WHERE TD_NAME IN (SELECT DISTINCT TB_NAME FROM HRTABL)"
Data1.Refresh
Do Until Data1.Recordset.EOF
    Data1.Recordset!TD_SHOWDESC = UCase(lStr(Data1.Recordset!TD_DESC))
    
    If glbCompSerial = "S/N - 2355W" Then
        If UCase(Data1.Recordset!TD_DESC) = "HIRE CODES" Then
            Data1.Recordset!TD_SHOWDESC = "VADIM STATUS CODE"
        End If
    End If
    
    Data1.Recordset.Update
    Data1.Recordset.MoveNext
Loop
SQLQ = "Select TB_NAME from HRTABL "

If Tabl_Snap.State <> 0 Then Tabl_Snap.Close
Tabl_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic

Screen.MousePointer = DEFAULT

Exit Sub

CR_EENames_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_Tables", "HRTABLE", "Select")
Resume Next

End Sub

Private Function Cri_SetAll()
Dim x%, strPage$, strSFormat$

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""
If chkTable.Value = False Then
' call cri models set both glbiONeWhere and strSelCri
    Call Cri_Table    ' sets fglbCriteria and fglbiOneWhere
Else
    txtTblName.Visible = False
End If
' report name
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztable.rpt"

'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For x% = 0 To 2
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next x%
    
    ' set security for database
    'Me.vbxCrystal.Password = gstrAccPWord$
    'Me.vbxCrystal.UserName = gstrAccUID$
End If
' window title if appropriate
Me.vbxCrystal.WindowTitle = "Table Master Report"

Cri_SetAll = True
If ChkPage Then
    strPage$ = "T;"
Else
    strPage$ = "F;"
End If
strSFormat$ = "GH" & CStr(1) & ";X;X;" & strPage$ & "X;X;X;X"
Me.vbxCrystal.SectionFormat(1) = strSFormat$

Screen.MousePointer = DEFAULT
Exit Function


modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "Comp Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Sub Cri_Table()
Dim CodeCri As String
        
If Len(txtTblName.Text) > 0 Then
    CodeCri = "({tblTable.TB_NAME} = '" & txtTblName.Text & "')"
End If

If Len(CodeCri) >= 1 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & CodeCri
    Else
        glbstrSelCri = CodeCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Function CriCheck()
CriCheck = True
End Function


Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
glbOnTop = Me.name
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

''frmTblName.Left = 1000
''frmTblName.Top = 1000

'Data1.DatabaseName = glbIHRDB 'ADDED BY RAUBREY 6/4/97
Data1.ConnectionString = glbAdoIHRDB



Call CR_Table_Snap
'tblTables.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub

Private Sub panCriteria_Click()

End Sub

Private Sub tblTables_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub tblTables_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If tblTables.Tag = "ASC" Then
            tblTables.Tag = "DESC"
        Else
            tblTables.Tag = "ASC"
        End If
        
        SQLQ = "SELECT TD_NAME, TD_DESC,TD_SHOWDESC FROM HRTABDES WHERE TD_NAME IN (SELECT DISTINCT TB_NAME FROM HRTABL)"
        SQLQ = SQLQ & " ORDER BY " & tblTables.Columns(ColIndex).DataField & " " & tblTables.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub txtTblName_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub


Private Sub txtTblName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = Reports
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = False
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property
Public Property Get Updateble() As Boolean
Updateble = False
End Property
Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Call set_Buttons
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub


