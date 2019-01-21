VERSION 5.00
Begin VB.Form frmODBCLogon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Source"
   ClientHeight    =   5385
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   5220
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2016138
   Icon            =   "FrmODBCLogn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEncryptMethod 
      Caption         =   "Encryption Method"
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   3960
      Width           =   4935
      Begin VB.ComboBox cboEMList 
         Height          =   315
         ItemData        =   "FrmODBCLogn.frx":000C
         Left            =   1440
         List            =   "FrmODBCLogn.frx":000E
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   2745
      End
   End
   Begin VB.Frame fraMultiDS 
      Caption         =   "Change Data Source"
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   4935
      Begin VB.ComboBox cboDSList 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmODBCLogn.frx":0010
         Left            =   1440
         List            =   "FrmODBCLogn.frx":0012
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   2745
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   21
         Top             =   260
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   750
      MaskColor       =   &H00000000&
      TabIndex        =   7
      Top             =   4800
      Width           =   1440
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2940
      MaskColor       =   &H00000000&
      TabIndex        =   8
      Top             =   4800
      Width           =   1260
   End
   Begin VB.Frame fraConnection 
      Caption         =   "Connection Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox cboDrivers 
         Height          =   315
         ItemData        =   "FrmODBCLogn.frx":0014
         Left            =   1440
         List            =   "FrmODBCLogn.frx":0016
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   390
         Width           =   2745
      End
      Begin VB.TextBox txtUserPsw 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1455
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   2490
         Width           =   2745
      End
      Begin VB.TextBox txtUserName 
         Height          =   330
         Left            =   1440
         TabIndex        =   5
         Top             =   2055
         Width           =   2745
      End
      Begin VB.TextBox txtDatabase 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1455
         TabIndex        =   4
         Top             =   1650
         Width           =   2745
      End
      Begin VB.ComboBox cboDSNList 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmODBCLogn.frx":0018
         Left            =   1440
         List            =   "FrmODBCLogn.frx":001A
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   2
         Text            =   "cboDSNList"
         Top             =   810
         Width           =   2730
      End
      Begin VB.TextBox txtServer 
         Height          =   330
         Left            =   1440
         TabIndex        =   3
         Top             =   1215
         Width           =   2745
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Driver:"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   18
         Top             =   420
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Password:"
         Height          =   255
         Index           =   7
         Left            =   210
         TabIndex        =   17
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "User Name:"
         Height          =   255
         Index           =   6
         Left            =   210
         TabIndex        =   16
         Top             =   2100
         Width           =   975
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "DSN Name:"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   1
         Top             =   840
         Width           =   810
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Database:"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   9
         Top             =   1680
         Width           =   750
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Server:"
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   10
         Top             =   1260
         Width           =   540
      End
   End
   Begin VB.TextBox txtPWD 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1230
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   1050
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtUID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1230
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&UID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   750
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   1095
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmODBCLogon"
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
Dim gdbAdoIHRDS As New ADODB.Connection
Dim rsIHRDS As New ADODB.Recordset
Dim xAdoIHRDB As String, SQLQ

Private Sub cboDSList_Click()
    If Len(cboDSList.Text) > 0 Then
        SQLQ = "SELECT * FROM HR_DATA_SOURCE WHERE DS_NAME ='" & cboDSList.Text & "' "
        If rsIHRDS.State <> 0 Then rsIHRDS.Close
        rsIHRDS.Open SQLQ, gdbAdoIHRDS, adOpenKeyset, adLockOptimistic
        If Not rsIHRDS.EOF Then
            If Not IsNull(rsIHRDS("DS_DRIVER")) Then
            cboDrivers.Text = rsIHRDS("DS_DRIVER")
            End If
            cboDSNList.Text = rsIHRDS("DS_DSN")
            txtServer.Text = rsIHRDS("DS_SERVER")
            If glbSQL Then
                txtDatabase.Text = rsIHRDS("DS_DATABASE")
            End If
            txtUserName.Text = rsIHRDS("DS_USERID")
            If gsMultiLang = "YES" Then 'whscc
                txtUserPsw.Text = DecryptPasswordMultiLang(rsIHRDS("DS_PASSWORD"))
            Else
                txtUserPsw.Text = DecryptPassword(rsIHRDS("DS_PASSWORD"))
            End If
        End If
        rsIHRDS.Close
    End If
End Sub

Private Sub cboEMList_Click()
    If cboEMList.Text = "Multi Language" Then
        gsMultiLang = "YES"
    Else
        gsMultiLang = "NO"
    End If
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdRegister_Click()
Dim Response%, w%, X%, Y%, SECTION$, Key$, xPWD$, valtmp, I

'Ticket #24352 - This string is combined registry entry that needs encryption (PIPEDA)
Dim strHRSSLic As String
Dim strHRSSLicEncrypt As String

On Error GoTo RegisterErr


If glbHosted Then
    If FileLen(glbHostFile) <> 0 Then
        'Hosted Environment
        'Write the connection values from ihrhost.ini and Set it to the global connection variables
        SECTION$ = "ODBC Setup"
        
        X% = INIWrite(SECTION$, "DATABASENAME", txtDatabase.Text, glbHostFile)
        'X% = WriteRegistrySetting(lCurrentKey, SECTION$, "DATABASENAME", txtDatabase.Text)  'gbasINI_WritePrivateString
        SQLDatabaseName = txtDatabase.Text
        
        X% = INIWrite(SECTION$, "DRIVERNAME", cboDrivers.Text, glbHostFile)
        'X% = WriteRegistrySetting(lCurrentKey, SECTION$, "DRIVERNAME", cboDrivers.Text)  'gbasINI_WritePrivateString
        SQLDriver = cboDrivers.Text
        
        X% = INIWrite(SECTION$, "SERVERNAME", txtServer.Text, glbHostFile)
        'X% = WriteRegistrySetting(lCurrentKey, SECTION$, "SERVERNAME", txtServer.Text)  'gbasINI_WritePrivateString
        SQLServerName = txtServer.Text
        
        X% = INIWrite(SECTION$, "USERNAME", txtUserName.Text, glbHostFile)
        'X% = WriteRegistrySetting(lCurrentKey, SECTION$, "USERNAME", txtUserName.Text)  'gbasINI_WritePrivateString
        SQLUserName = txtUserName.Text
        
        If gsMultiLang = "Y" Then 'For Listowel only
            xPWD$ = EncryptPasswordMultiLang_First(txtUserPsw.Text)
        ElseIf UCase(gsMultiLang) = "YES" Then 'For general multi language clients
            xPWD$ = EncryptPasswordMultiLang(txtUserPsw.Text)
        'For version 7.6 only ticket# 9153
        Else
            xPWD$ = EncryptPassword(txtUserPsw.Text)
        End If
        X% = INIWrite(SECTION$, "USERPSW", xPWD$, glbHostFile)
        'X% = WriteRegistrySetting(lCurrentKey, SECTION$, "USERPSW", xPWD$)  'gbasINI_WritePrivateString
        SQLUserPassword = txtUserPsw.Text
                
        SECTION$ = "Options"
        X% = INIWrite(SECTION$, "MultiLang", gsMultiLang, glbHostFile)
        'X% = WriteRegistrySetting(lCurrentKey, SECTION$, "MultiLang", gsMultiLang)
        
        glbHosted = True
    End If
End If

strHRSSLic = ""

If Not glbHosted Then
    SECTION$ = REG_NAME & "ODBC Setup"
    
    'If the Encryption of Database Connection is turned-ON then do not update ODBC Setup registry, update License key instead
    If Not gsDB_CONNECT_ENCRYPT Then
        X% = WriteRegistrySetting(lCurrentKey, SECTION$, "DATABASENAME", txtDatabase.Text)  'gbasINI_WritePrivateString
    Else
        'Ticket #24352 - PIPEDA
        strHRSSLic = strHRSSLic & "DATABASENAME=" & txtDatabase.Text & "|"
    End If
    SQLDatabaseName = txtDatabase.Text
    
    'If the Encryption of Database Connection is turned-ON then do not update ODBC Setup registry, update License key instead
    If Not gsDB_CONNECT_ENCRYPT Then
        X% = WriteRegistrySetting(lCurrentKey, SECTION$, "DRIVERNAME", cboDrivers.Text)  'gbasINI_WritePrivateString
    Else
        'Ticket #24352 - PIPEDA
        strHRSSLic = strHRSSLic & "DRIVERNAME=" & cboDrivers.Text & "|"
    End If
    SQLDriver = cboDrivers.Text
    
    'If the Encryption of Database Connection is turned-ON then do not update ODBC Setup registry, update License key instead
    If Not gsDB_CONNECT_ENCRYPT Then
        X% = WriteRegistrySetting(lCurrentKey, SECTION$, "SERVERNAME", txtServer.Text)  'gbasINI_WritePrivateString
    Else
        'Ticket #24352 - PIPEDA
        strHRSSLic = strHRSSLic & "SERVERNAME=" & txtServer.Text & "|"
    End If
    SQLServerName = txtServer.Text
    
    'If the Encryption of Database Connection is turned-ON then do not update ODBC Setup registry, update License key instead
    If Not gsDB_CONNECT_ENCRYPT Then
        X% = WriteRegistrySetting(lCurrentKey, SECTION$, "USERNAME", txtUserName.Text)  'gbasINI_WritePrivateString
    Else
        'Ticket #24352 - PIPEDA
        strHRSSLic = strHRSSLic & "USERNAME=" & txtUserName.Text & "|"
    End If
    SQLUserName = txtUserName.Text
        
        
    If gsMultiLang = "Y" Then 'For Listowel only
        xPWD$ = EncryptPasswordMultiLang_First(txtUserPsw.Text)
    ElseIf UCase(gsMultiLang) = "YES" Then 'For general multi language clients
        xPWD$ = EncryptPasswordMultiLang(txtUserPsw.Text)
    'For version 7.6 only ticket# 9153
    Else
        xPWD$ = EncryptPassword(txtUserPsw.Text)
    End If
    
    'If the Encryption of Database Connection is turned-ON then do not update ODBC Setup registry, update License key instead
    If Not gsDB_CONNECT_ENCRYPT Then
        X% = WriteRegistrySetting(lCurrentKey, SECTION$, "USERPSW", xPWD$)  'gbasINI_WritePrivateString
    Else
        'Ticket #24352 - PIPEDA
        strHRSSLic = strHRSSLic & "USERPSW=" & xPWD$ & "|"
    End If
    SQLUserPassword = txtUserPsw.Text
    
    SECTION$ = REG_NAME & "Options"
    X% = WriteRegistrySetting(lCurrentKey, SECTION$, "MultiLang", gsMultiLang)
    
    'Ticket #24352 - PIPEDA
    'strHRSSLic = strHRSSLic & vbCrLf & vbCrLf & "[Options]" & vbCrLf & "MultiLang=" & gsMultiLang
    
    'The Encryption of Database Connection is turned-ON, save encrypted database connection information in License Key
    If gsDB_CONNECT_ENCRYPT Then
        'Ticket #24352 - PIPEDA
        'Encrypt the ODBC Setup and Options Registry settings
        If Len(strHRSSLic) > 0 Then
            strHRSSLicEncrypt = EncryptDatabaseSettings(strHRSSLic)
            If Len(strHRSSLicEncrypt) > 0 Then
                SECTION$ = REG_NAME & "Options"
                X% = WriteRegistrySetting(lCurrentKey, SECTION$, "License", strHRSSLicEncrypt)
            End If
        End If
    End If
End If

Call glbAdo_Value



On Error GoTo ODBCErr

ODBCSetup:
    DBEngine.RegisterDatabase cboDSNList.Text, cboDrivers, True, "Database=" & txtDatabase & vbCr & "Server=" & txtServer & vbCr
    'DBEngine.RegisterDatabase "IHREI", fglbDrivers, True, "Database=IHREI" & vbCr & "Server=" & txtServer & vbCr
    MsgBox "Datasource Registration Succeeded.", vbInformation
    Unload Me
    
    Exit Sub

RegisterErr:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Error Encrypting", "Registry Entry Fail", "OK")
    
    MsgBox "Registration failed. Please make sure that you have rights to change the System Registry."
    Unload Me

ODBCErr:
    MsgBox "ODBC setup failed. Please create ODBC DSN manually."
    
    Unload Me

End Sub

Private Sub Form_Load()
Dim I As Integer

'If Left(UCase(SQLServerName), 6) = "SERVER" Then txtDatabase.Enabled = True
txtDatabase.Enabled = True

cboDSNList.Text = "INFOHR"

If glbSQL Then
    cboDrivers.AddItem "SQL Server"
    cboDrivers.ListIndex = 0
    txtDatabase.Text = SQLDatabaseName
ElseIf glbOracle Then
    Call GetODBCDrivers
    If cboDrivers.ListCount > 0 Then
        cboDrivers.Text = SQLDriver
    Else
        'cboDrivers = "Oracle ODBC Driver"
    End If
    lblLabels(3).Visible = False
    txtDatabase.Visible = False
End If

txtServer.Text = SQLServerName
txtUserName.Text = SQLUserName
txtUserPsw.Text = SQLUserPassword

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

'Multi Data Source - Begin
If Not Dir(glbIHRREPORTS & "IHRDS.mdb") = "" Then
    xAdoIHRDB = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=petman;Data Source=" & glbIHRREPORTS & "IHRDS.mdb"
    
    If gdbAdoIHRDS.State = adStateOpen Then gdbAdoIHRDS.Close
    
    gdbAdoIHRDS.CommandTimeout = 600
    gdbAdoIHRDS.Mode = adModeReadWrite
    gdbAdoIHRDS.Open xAdoIHRDB
    
    SQLQ = "SELECT * FROM HR_DATA_SOURCE ORDER BY DS_NAME"
    If rsIHRDS.State <> 0 Then rsIHRDS.Close
    rsIHRDS.Open SQLQ, gdbAdoIHRDS, adOpenKeyset, adLockOptimistic
    Do While Not rsIHRDS.EOF
        cboDSList.Enabled = True
        cboDSList.AddItem rsIHRDS("DS_NAME")
        rsIHRDS.MoveNext
    Loop
    rsIHRDS.Close
    
    'Get the current Data source
    SQLQ = "SELECT * FROM HR_DATA_SOURCE WHERE DS_SERVER = '" & SQLServerName & "' "
    If Not glbOracle Then
        SQLQ = SQLQ & "AND DS_DATABASE = '" & SQLDatabaseName & "' "
    End If
    If rsIHRDS.State <> 0 Then rsIHRDS.Close
    rsIHRDS.Open SQLQ, gdbAdoIHRDS, adOpenKeyset, adLockOptimistic
    If Not rsIHRDS.EOF Then
        'xIHRDS = " [" & rsIHRDS("DS_NAME") & "] "
        cboDSList.Text = rsIHRDS("DS_NAME")
    End If
    rsIHRDS.Close
    
End If
'Multi Data Source - End

'Encryption method - Begin
If Not (gsMultiLang = "Y") Then
    cboEMList.AddItem "English"
    cboEMList.AddItem "Multi Language"
    If gsMultiLang = "YES" Then
        cboEMList.ListIndex = 1
    Else
         cboEMList.ListIndex = 0
    End If
End If
'Encryption method - End

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
