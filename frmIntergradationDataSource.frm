VERSION 5.00
Begin VB.Form frmIntergradationDataSource 
   Caption         =   "Vadim Data Source"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraConnection 
      Caption         =   "Connection Values"
      Height          =   3225
      Left            =   210
      TabIndex        =   2
      Top             =   150
      Width           =   4935
      Begin VB.TextBox txtServer 
         Height          =   330
         Left            =   1440
         TabIndex        =   8
         Top             =   1215
         Width           =   2745
      End
      Begin VB.ComboBox cboDSNList 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmIntergradationDataSource.frx":0000
         Left            =   1440
         List            =   "frmIntergradationDataSource.frx":0002
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   7
         Text            =   "cboDSNList"
         Top             =   810
         Width           =   2730
      End
      Begin VB.TextBox txtDatabase 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1455
         TabIndex        =   6
         Top             =   1650
         Width           =   2745
      End
      Begin VB.TextBox txtUserName 
         Height          =   330
         Left            =   1440
         TabIndex        =   5
         Top             =   2055
         Width           =   2745
      End
      Begin VB.TextBox txtUserPsw 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1455
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2490
         Width           =   2745
      End
      Begin VB.ComboBox cboDrivers 
         Height          =   315
         ItemData        =   "frmIntergradationDataSource.frx":0004
         Left            =   1440
         List            =   "frmIntergradationDataSource.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   390
         Width           =   2745
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Server:"
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   14
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Database:"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   13
         Top             =   1680
         Width           =   750
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "DSN Name:"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   12
         Top             =   840
         Width           =   810
      End
      Begin VB.Label lblLabels 
         Caption         =   "User Name:"
         Height          =   255
         Index           =   6
         Left            =   210
         TabIndex        =   11
         Top             =   2100
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Password:"
         Height          =   255
         Index           =   7
         Left            =   210
         TabIndex        =   10
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Driver:"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   9
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   450
      Left            =   2190
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   3630
      Width           =   1260
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "&OK"
      Height          =   450
      Left            =   510
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   3630
      Width           =   1440
   End
End
Attribute VB_Name = "frmIntergradationDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Const KEY_READ = &H20019
Const REG_SZ = 1
Const ERROR_MORE_DATA = 234
Dim fglbCGL As Boolean
Private Sub cmdCancel_Click()
  Unload Me
End Sub
Private Sub cmdRegister_Click()
Dim Response%, w%, x%, Y%, SECTION$, Key$, xPWD$, valtmp, I
On Error GoTo RegisterErr
If glbVadim Then
    SECTION$ = REG_NAME & "Network"
End If
x% = WriteRegistrySetting(lCurrentKey, SECTION$, "INTER_DATABASENAME", txtDatabase.Text)
InterDatabaseName = txtDatabase.Text

x% = WriteRegistrySetting(lCurrentKey, SECTION$, "INTER_SERVERNAME", txtServer.Text)
InterServerName = txtServer.Text

x% = WriteRegistrySetting(lCurrentKey, SECTION$, "INTER_USERNAME", txtUserName.Text)
InterUserName = txtUserName.Text

xPWD$ = EncryptPassword(txtUserPsw.Text)
x% = WriteRegistrySetting(lCurrentKey, SECTION$, "INTER_USERPSW", xPWD$)
InterUserPassword = txtUserPsw.Text
'Call glbAdo_Value


On Error GoTo DataOpenErr
'ODBCSetup:
'    DBEngine.RegisterDatabase cboDSNList.Text, cboDrivers, True, "Database=" & txtDatabase & vbCr & "Server=" & txtServer & vbCr
''    DBEngine.RegisterDatabase "IHREI", fglbDrivers, True, "Database=IHREI" & vbCr & "Server=" & txtServer & vbCr
    If gdbPayroll.State <> 0 Then gdbPayroll.Close
    glbPayrollDB = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & InterUserName & ";Password=" & InterUserPassword & ";Initial Catalog=" & InterDatabaseName & ";Data Source=" & InterServerName
    gdbPayroll.Open glbPayrollDB

    MsgBox "Datasource Registration Succeeded", vbInformation
    Unload Me
    Exit Sub
RegisterErr:
    MsgBox "Registration failed, Please check do you have right to change the system register."
    Unload Me
DataOpenErr:
    MsgBox "Could not open the payroll database."
    Unload Me

End Sub
Private Sub Form_Load()
Dim I As Integer

txtDatabase.Enabled = True

cboDSNList.Text = "INFOHR"

If fglbCGL Then
    Call GetODBCDrivers
    If cboDrivers.ListCount > 0 Then
        cboDrivers.ListIndex = 0
    End If
    lblLabels(0).Visible = False
    cboDSNList.Visible = False
    lblLabels(3).Visible = False
    txtDatabase.Visible = False
    Me.Caption = "Visual Quality System Data Source"
End If
If glbVadim Then
    cboDrivers.Clear
    cboDrivers.AddItem "SQL Server"
    cboDrivers.ListIndex = 0
    txtDatabase.Text = "Vadim"
    Me.Caption = "Vadim Payroll System Data Source"
End If
txtServer.Text = InterServerName
txtDatabase.Text = InterDatabaseName
txtUserName.Text = InterUserName
txtUserPsw.Text = InterUserPassword

End Sub

Private Sub GetODBCDrivers()
Dim res As Collection
Dim values As Variant
For Each values In EnumRegistryValues(HKEY_LOCAL_MACHINE, "Software\ODBC\ODBCINST.INI\ODBC Drivers")
    If InStr(values(0), "Oracle") <> 0 And Left(values(0), 1) = "O" Then
        cboDrivers.AddItem values(0)
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

Public Property Let CGLInterface(vData As Boolean)
fglbCGL = vData
End Property


Private Sub txtUID_Change()

End Sub
