VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMDataSource 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Multiple Data Sources Setup"
   ClientHeight    =   6675
   ClientLeft      =   90
   ClientTop       =   1005
   ClientWidth     =   9480
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
   ScaleHeight     =   6675
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtUserPswEncrypt 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      DataField       =   "DS_PASSWORD"
      ForeColor       =   &H80000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   5280
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   20
      Tag             =   "01- Password"
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cmbDriver 
      Height          =   315
      ItemData        =   "fxmdatasourcel.frx":0000
      Left            =   2280
      List            =   "fxmdatasourcel.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "11-Driver"
      Top             =   3270
      Width           =   2745
   End
   Begin VB.TextBox txtUserPsw 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   7
      Tag             =   "01- Password"
      Top             =   5160
      Width           =   2775
   End
   Begin VB.TextBox txtUserID 
      Appearance      =   0  'Flat
      DataField       =   "DS_USERID"
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
      Left            =   2280
      MaxLength       =   25
      TabIndex        =   6
      Tag             =   "01-User Name"
      Top             =   4800
      Width           =   2775
   End
   Begin VB.TextBox txtDatabase 
      Appearance      =   0  'Flat
      DataField       =   "DS_DATABASE"
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
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "01-Database"
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox txtServer 
      Appearance      =   0  'Flat
      DataField       =   "DS_SERVER"
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
      Left            =   2280
      MaxLength       =   100
      TabIndex        =   4
      Tag             =   "01-Server"
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      DataField       =   "DS_NAME"
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "01-Data Source Name"
      Top             =   2880
      Width           =   2775
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxmdatasourcel.frx":0004
      Height          =   2535
      Left            =   240
      OleObjectBlob   =   "fxmdatasourcel.frx":0018
      TabIndex        =   0
      Top             =   240
      Width           =   8775
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
      DataField       =   "DS_LUSER"
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
      Left            =   8400
      MaxLength       =   10
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txtDSN 
      Appearance      =   0  'Flat
      DataField       =   "DS_DSN"
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "01-DSN Name"
      Top             =   3720
      Width           =   2775
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DS_LDATE"
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
      Left            =   6960
      MaxLength       =   12
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5145
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DS_LTIME"
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
      Left            =   7710
      MaxLength       =   8
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5160
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
   Begin VB.TextBox txtDriver 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "DS_DRIVER"
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
      Left            =   5880
      MaxLength       =   50
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   360
      TabIndex        =   19
      Top             =   5160
      Width           =   1275
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   360
      TabIndex        =   18
      Top             =   4800
      Width           =   1275
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Database"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   17
      Top             =   4440
      Width           =   1275
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   16
      Top             =   4080
      Width           =   1275
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Driver"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   3240
      Width           =   1035
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DSN Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   12
      Top             =   3720
      Width           =   1275
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Data Source Name"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Top             =   2880
      Width           =   1635
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      DataField       =   "DS_COMPNO"
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
      Left            =   6960
      TabIndex        =   10
      Top             =   5520
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "frmMDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Const KEY_READ = &H20019
Const REG_SZ = 1
Const ERROR_MORE_DATA = 234
Dim gdbAdoIHRDS As New ADODB.Connection
Dim fglbNew As Boolean
Dim fglbSDate As Variant
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim DefType(0 To 3)
Dim SystType(0 To 3)
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim UpdateState As UpdateStateEnum



Private Function chkDataSource()
Dim rsTemp As New ADODB.Recordset
Dim msg As String, SQLQ
Dim x%, xchk

chkDataSource = False

If Len(txtName.Text) < 1 Then
    MsgBox "Data Source Name must be entered"
    txtName.SetFocus
    Exit Function
End If
If Len(cmbDriver.Text) < 1 Then
    MsgBox "Driver must be entered"
    txtName.SetFocus
    Exit Function
End If
If Len(txtDSN.Text) < 1 Then
    MsgBox "DSN Name must be entered"
    txtName.SetFocus
    Exit Function
End If
If Len(txtServer.Text) < 1 Then
    MsgBox "Server must be entered"
    txtName.SetFocus
    Exit Function
End If
If glbSQL Then
    If Len(txtDatabase.Text) < 1 Then
        MsgBox "Database must be entered"
        txtName.SetFocus
        Exit Function
    End If
End If
If Len(txtUserID.Text) < 1 Then
    MsgBox "User Name must be entered"
    txtName.SetFocus
    Exit Function
End If
If Len(txtUserPsw.Text) < 1 Then
    MsgBox "Password must be entered"
    txtName.SetFocus
    Exit Function
End If

'SQLQ = "SELECT * FROM HR_DATA_SOURCE WHERE DS_NAME = '" & txtName & "' "
''SQLQ = SQLQ & "AND DS_ID <> " & rsDATA("DS_ID")
'rsTemp.Open SQLQ, xAdoIHRDB, adOpenStatic
'If Not rsTemp.EOF Then
'    MsgBox "Duplicate Data Source Name"
'    txtName.SetFocus
'    Exit Function
'End If
'rsTemp.Close
chkDataSource = True

End Function



Private Sub cmbDriver_Change()
txtDriver.Text = cmbDriver.Text
End Sub

Sub cmdCancel_Click()

On Error GoTo Can_Err
fglbNew = False
If fglbEmptyNew Then
    Me.vbxTrueGrid.Enabled = True
    Me.vbxTrueGrid.Refresh
End If


'Data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh
''' Sam add July 2002 * Remove Binding Control
rsDATA.CancelUpdate
Call Display_Value


'Call ST_UPD_MODE(True) ' reset screen's attributes

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HR_DATA_SOURCE", "Cancel")
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
Dim a As Integer, msg As String

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

msg = "Are You Sure You Want To Delete "
msg = msg & "This Record?"
a% = MsgBox(msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

'Data1.Recordset.Delete
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh
''' Sam add July 2002 * Remove Binding Control
gdbAdoIHRDS.BeginTrans
rsDATA.Delete
gdbAdoIHRDS.CommitTrans
Data1.Refresh


Call SET_UP_MODE
'Call ST_UPD_MODE(False)


Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_DATA_SOURCE", "Delete")
Call RollBack '09June99 js

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Call ST_UPD_MODE(True)
'cmbDriver.SetFocus
Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_DATA_SOURCE", "Modify")
Call RollBack '09June99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()

On Error GoTo AddN_Err

Call Set_Control("B", Me)
rsDATA.AddNew

txtDSN.Text = "INFOHR"

lblCNum.Caption = "001"

fglbNew = True
Call SET_UP_MODE

'Call ST_UPD_MODE(True)
'cmbDriver.SetFocus
Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_DATA_SOURCE", "Add")
Call RollBack '09June99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim x%

On Error GoTo cmdOK_Err

txtDriver.Text = cmbDriver.Text
If Not chkDataSource() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)
Call Set_Control("U", Me, rsDATA)

gdbAdoIHRDS.BeginTrans
rsDATA.Update
gdbAdoIHRDS.CommitTrans
Data1.Refresh

fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(False)

Me.vbxTrueGrid.Enabled = True
Me.vbxTrueGrid.SetFocus
Screen.MousePointer = DEFAULT

Exit Sub

cmdOK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_DATA_SOURCE", "Update")
Call RollBack '09June99 js

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = "Multiple Data Source"
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

RHeading = "Multiple Data Source"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
Private Sub combType()
End Sub

Private Sub cmbDriver_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "PAYROLL", "SELECT")

End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
Me.cmdModify_Click
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim Answer, DefVal, msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim i%, SQLQ
Dim xAdoIHRDB

If glbSQL Then
    cmbDriver.AddItem "SQL Server"
    cmbDriver.ListIndex = 0
ElseIf glbOracle Then
    Call GetODBCDrivers
    If cmbDriver.ListCount > 0 Then
        cmbDriver.Text = SQLDriver
    Else
        'cmbDriver = "Oracle ODBC Driver"
    End If
End If

xAdoIHRDB = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=petman;Data Source=" & glbIHRREPORTS & "IHRDS.mdb"
If gdbAdoIHRDS.State = adStateOpen Then gdbAdoIHRDS.Close
gdbAdoIHRDS.CommandTimeout = 600
gdbAdoIHRDS.Mode = adModeReadWrite
gdbAdoIHRDS.Open xAdoIHRDB

Me.Show
glbOnTop = "frmMDataSource"
'Call combType
Screen.MousePointer = HOURGLASS

Data1.ConnectionString = gdbAdoIHRDS
Data1.RecordSource = "SELECT * FROM HR_DATA_SOURCE ORDER BY DS_NAME"
Data1.Refresh
'Frank May 17,2002
'Call setRptCaption(Me)
Screen.MousePointer = DEFAULT
'Call Display_Value
Call ST_UPD_MODE(False)

Call INI_Controls(Me)
If glbOracle Then
    lblTitle(4).Visible = False
    txtDatabase.Visible = False
    lblTitle(6).Top = lblTitle(5).Top
    lblTitle(5).Top = lblTitle(4).Top
    txtUserPsw.Top = txtUserID.Top
    txtUserID.Top = txtDatabase.Top
End If
Screen.MousePointer = DEFAULT                           '
End Sub

Private Sub GetODBCDrivers()
Dim res As Collection
Dim values As Variant
For Each values In EnumRegistryValues(HKEY_LOCAL_MACHINE, "Software\ODBC\ODBCINST.INI\ODBC Drivers")
    If InStr(values(0), "Oracle") <> 0 And Left(values(0), 1) = "O" Then
        cmbDriver.AddItem values(0)
    End If
Next

End Sub

Function EnumRegistryValues(ByVal hKey As Long, ByVal keyname As String) As Collection
    Dim handle As Long
    Dim Index As Long
    Dim valueType As Long
    Dim name As String
    Dim nameLen As Long
    Dim resLong As Long
    Dim resString As String
    Dim dataLen As Long
    Dim valueInfo(0 To 1) As Variant
    Dim retVal As Long
    
    ' initialize the result
    Set EnumRegistryValues = New Collection
    
    ' Open the key, exit if not found.
    If Len(keyname) Then
        If RegOpenKeyEx(hKey, keyname, 0, KEY_READ, handle) Then Exit Function
        ' in all cases, subsequent functions use hKey
        hKey = handle
    End If
    
    Do
        ' this is the max length for a key name
        nameLen = 260
        name = Space$(nameLen)
        ' prepare the receiving buffer for the value
        dataLen = 4096
        ReDim resBinary(0 To dataLen - 1) As Byte
        
        ' read the value's name and data
        ' exit the loop if not found
        retVal = RegEnumValue(hKey, Index, name, nameLen, ByVal 0&, valueType, _
            resBinary(0), dataLen)
        
        ' enlarge the buffer if you need more space
        If retVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To dataLen - 1) As Byte
            retVal = RegEnumValue(hKey, Index, name, nameLen, ByVal 0&, _
                valueType, resBinary(0), dataLen)
        End If
        ' exit the loop if any other error (typically, no more values)
        If retVal Then Exit Do
        
        ' retrieve the value's name
        valueInfo(0) = Left$(name, nameLen)
        
        ' return a value corresponding to the value type
        If valueType = REG_SZ Then
            EnumRegistryValues.Add valueInfo, valueInfo(0)
        End If
        
        Index = Index + 1
    Loop
   
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
        
End Function

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
Dim i As Integer
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

If Data1.Recordset.BOF Or Data1.Recordset.EOF Then
    If Not fglbNew Then
        TF = False
    End If
End If

txtName.Enabled = TF
cmbDriver.Enabled = TF
'txtDSN.Enabled = TF
txtServer.Enabled = TF
txtDatabase.Enabled = TF
txtUserID.Enabled = TF
txtUserPsw.Enabled = TF


End Sub

Private Sub txtConvert_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtDatabase_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtDriver_Change()
cmbDriver.Text = txtDriver.Text
End Sub

Private Sub txtDSN_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtName_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtServer_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtUserID_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtUserPsw_Change()
If gsMultiLang = "YES" Then 'whscc
    txtUserPswEncrypt.Text = EncryptPasswordMultiLang(txtUserPsw.Text)
Else
    txtUserPswEncrypt.Text = EncryptPassword(txtUserPsw.Text)
End If

End Sub

Private Sub txtUserPsw_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtUserPswEncrypt_Change()
If gsMultiLang = "YES" Then 'whscc
    txtUserPsw.Text = DecryptPasswordMultiLang(txtUserPswEncrypt.Text)
Else
    txtUserPsw.Text = DecryptPassword(txtUserPswEncrypt.Text)
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
        
        SQLQ = "SELECT * FROM HR_DATA_SOURCE "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i%

On Error GoTo vbxTrueGrid_Err
Call Display_Value

If Data1.Recordset.EOF Or Data1.Recordset.BOF = 0 Then
    Exit Sub
End If

'cmbDriver.ListIndex = i%

Exit Sub

vbxTrueGrid_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "Multiple Data Source", "Select")
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
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIHRDS, adOpenKeyset, adLockOptimistic
        Call SET_UP_MODE
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HR_DATA_SOURCE where DS_ID= " & Data1.Recordset!DS_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIHRDS, adOpenKeyset, adLockOptimistic
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
    Call SET_UP_MODE
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property
Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fglbNew = True
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_MultiDataSourceSetup
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
If fglbNew Then
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
