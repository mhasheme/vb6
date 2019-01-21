VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmOtherDatabaseSetup 
   Caption         =   "Database Setup"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6255
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox comVersion 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2745
   End
   Begin Threed.SSPanel panButtons 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   4785
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   1085
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin VB.CommandButton cmdTest 
         Appearance      =   0  'Flat
         Caption         =   "&Test Connection"
         Height          =   375
         Left            =   4050
         TabIndex        =   11
         Tag             =   "Update all matching records to the above"
         Top             =   120
         Width           =   1620
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2010
         TabIndex        =   10
         Tag             =   "Update all matching records to the above"
         Top             =   120
         Width           =   870
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Height          =   375
         Left            =   990
         TabIndex        =   9
         Tag             =   "Update all matching records to the above"
         Top             =   120
         Width           =   870
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Tag             =   "Close and exit this screen"
         Top             =   120
         Width           =   750
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   6600
         Top             =   150
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
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
   Begin VB.Frame frmSQL 
      Height          =   3585
      Left            =   360
      TabIndex        =   13
      Top             =   510
      Width           =   5535
      Begin VB.TextBox txtDatabaseName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2190
         TabIndex        =   1
         Top             =   210
         Width           =   2355
      End
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2190
         TabIndex        =   3
         Top             =   810
         Width           =   2355
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2190
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1110
         Width           =   2355
      End
      Begin VB.TextBox txtDatabaseServer 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2190
         TabIndex        =   2
         Top             =   510
         Width           =   2355
      End
      Begin VB.TextBox lblTest 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1515
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   1680
         Width           =   5115
      End
      Begin VB.Label lblHost 
         AutoSize        =   -1  'True
         Caption         =   "Database Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label lblUsername 
         AutoSize        =   -1  'True
         Caption         =   "Database Username:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   870
         Width           =   1785
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         Caption         =   "Database Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   1170
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Database Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   570
         Width           =   1440
      End
   End
   Begin VB.Frame frmAccess 
      Height          =   3705
      Left            =   330
      TabIndex        =   18
      Top             =   510
      Visible         =   0   'False
      Width           =   5535
      Begin VB.FileListBox File1 
         Height          =   3015
         Left            =   3660
         Pattern         =   "*.mdb"
         TabIndex        =   7
         Top             =   300
         Width           =   1695
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         TabIndex        =   6
         Tag             =   "Select drive"
         Top             =   3225
         Width           =   3225
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         Height          =   2790
         Left            =   360
         TabIndex        =   5
         Tag             =   "Select directory to import from"
         Top             =   300
         Width           =   3195
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Database Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   390
      TabIndex        =   17
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "frmOtherDatabaseSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbProduct As String
Dim rsDataSetup As New ADODB.Recordset


Private Sub cmdCancel_Click()
    If glbCompSerial <> "S/N - 2233W" Then
        Call RetrieveData
    Else
        Call RetrieveDataLG
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Set frmOtherDatabaseSetup = Nothing
End Sub

Function TestConnection() As Boolean
Dim rsTable As New ADODB.Recordset
Dim gdbAT As New ADODB.Connection
Dim glbAT
On Error Resume Next

TestConnection = False

If txtDatabaseServer <> "" Then
    gdbAT.CursorLocation = adUseClient
    If gdbAT.State = adStateOpen Then gdbAT.Close
    gdbAT.CommandTimeout = 600
    gdbAT.Mode = adModeReadWrite
    glbAT = "Provider=SQLOLEDB.1;Persist Security Info=False;" & _
                "User ID=" & txtUsername & _
                ";Password=" & txtPassword & _
                ";Initial Catalog=" & txtDatabaseName & _
                ";Data Source=" & txtDatabaseServer
    On Error GoTo conn_error
    gdbAT.Open glbAT
    lblTest = "Connect to " & fglbProduct & "  Database Successfully"
    
    If glbCompSerial = "S/N - 2233W" Then
        TestConnection = True 'Leeds Grenville ticket #14890
        Exit Function
    End If
    
check_details_table:
    On Error GoTo details_table_error
    If rsTable.State <> 0 Then rsTable.Close
    rsTable.Open "SELECT * FROM [infohr-etp.etp.employee]", gdbAT, adOpenForwardOnly
    lblTest = lblTest & vbNewLine & vbNewLine & "Found infohr-etp.etp.employee table"
    
TestConnection = True
End If
Exit Function
conn_error:
    lblTest = "Could not connect to " & fglbProduct & " database"
    lblTest = lblTest & vbNewLine & Space(3) & "Error:" & Err.Number & "-" & Err.Description
    Exit Function
details_table_error:
    lblTest = lblTest & vbNewLine & vbNewLine & "Could not find infohr-etp.etp.employee table"
    lblTest = lblTest & vbNewLine & Space(3) & "Error:" & Err.Number & "-" & Err.Description
End Function



Private Sub cmdOK_Click()
If glbCompSerial <> "S/N - 2233W" Then 'Simona - no combobox, SQL db by default - Leeds Grenville CAS ticket #14890
    If comVersion.ListIndex = 1 Then
        If Len(txtDatabaseName.Text) = 0 Then
            MsgBox "Database Name is a required field.", vbExclamation + vbOKOnly, "Missing Required Field"
            txtDatabaseName.SetFocus
            Exit Sub
        End If
        If Len(txtDatabaseServer.Text) = 0 Then
            MsgBox "Database Server is a required field.", vbExclamation + vbOKOnly, "Missing Required Field"
            txtDatabaseServer.SetFocus
            Exit Sub
        End If
        If Len(txtUsername.Text) = 0 Then
            MsgBox "User ID is a required field.", vbExclamation + vbOKOnly, "Missing Required Field"
            txtUsername.SetFocus
            Exit Sub
        End If
        If Len(txtPassword.Text) = 0 Then
            MsgBox "Password is a required field.", vbExclamation + vbOKOnly, "Missing Required Field"
            txtPassword.SetFocus
            Exit Sub
        End If
    
    End If
    Call SaveData

Else
        If Len(txtDatabaseName.Text) = 0 Then
            MsgBox "Database Name is a required field.", vbExclamation + vbOKOnly, "Missing Required Field"
            txtDatabaseName.SetFocus
            Exit Sub
        End If
        If Len(txtDatabaseServer.Text) = 0 Then
            MsgBox "Database Server is a required field.", vbExclamation + vbOKOnly, "Missing Required Field"
            txtDatabaseServer.SetFocus
            Exit Sub
        End If
        If Len(txtUsername.Text) = 0 Then
            MsgBox "User ID is a required field.", vbExclamation + vbOKOnly, "Missing Required Field"
            txtUsername.SetFocus
            Exit Sub
        End If
        If Len(txtPassword.Text) = 0 Then
            MsgBox "Password is a required field.", vbExclamation + vbOKOnly, "Missing Required Field"
            txtPassword.SetFocus
            Exit Sub
        End If
        Call SaveDataLG

End If 'Simona - end - Leeds Grenville CAS ticket #14890
'Call SaveData
On Error GoTo ODBCErr
ODBCSetup:
   
'    DBEngine.RegisterDatabase "IHRVADIM", "SQL Server", True, "Database=" & txtDatabaseName & vbCr & "Server=" & txtDatabaseServer & vbCr
    MsgBox "Datasource Registration Succeeded", vbInformation
    Unload Me
    Exit Sub
ODBCErr:
    MsgBox "ODBC setup failed, Please create ODBC DSN manually."
    Unload Me
End Sub

Private Sub cmdTest_Click()
Call TestConnection
End Sub

Private Sub comVersion_Click()
If comVersion.ListIndex = 0 Then
    frmAccess.Visible = True
    frmSQL.Visible = False
    cmdTest.Visible = False
Else
    frmAccess.Visible = False
    frmSQL.Visible = True
    cmdTest.Visible = True
End If
End Sub

Private Sub Dir1_Change()
Drive1.Drive = Dir1.Path
End Sub

Private Sub Drive1_Change()
File1.Path = Drive1
End Sub

Private Sub Form_Load()
    Dim SQLQ

    Me.Caption = fglbProduct & " Database Setup"
    If glbCompSerial <> "S/N - 2233W" Then
        
        comVersion.AddItem "MS Access"
        comVersion.AddItem "MS SQL Server"
        comVersion.ListIndex = 0
        Call RetrieveData
        
     Else   'Simona - Leeds Grenville CAS ticket #14890
        
        Label3.Visible = False
        comVersion.Visible = False
        frmSQL.Visible = True
        frmAccess.Visible = False
        cmdTest.Visible = True
        Call RetrieveDataLG
            
    End If
    
    'Call RetrieveData
End Sub
Private Sub RetrieveData()
    Dim SQLQ
    Dim x
    SQLQ = "SELECT * FROM APPLICATION_PARAMETER WHERE PARA_TYPE='Integration' AND PARA_CATEGORY='" & fglbProduct & "' "
    rsDataSetup.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

    comVersion.ListIndex = 0
    Do Until rsDataSetup.EOF
        If rsDataSetup("PARA_NAME") = "Version_Info" Then
            If rsDataSetup("PARA_VALUE") = "MS SQL Server" Then
                comVersion.ListIndex = 1
            Else
                comVersion.ListIndex = 0
            End If
    
        Else
            If comVersion.ListIndex = 1 Then
                If rsDataSetup("PARA_NAME") = "Database_Name" Then
                    txtDatabaseName.Text = rsDataSetup("PARA_VALUE")
                End If
                If rsDataSetup("PARA_NAME") = "Database_Server" Then
                    txtDatabaseServer.Text = rsDataSetup("PARA_VALUE")
                End If
                If rsDataSetup("PARA_NAME") = "User_Name" Then
                    txtUsername.Text = rsDataSetup("PARA_VALUE")
                End If
                If rsDataSetup("PARA_NAME") = "Password" Then
                    txtPassword.Text = rsDataSetup("PARA_VALUE")
                End If
            Else
                If rsDataSetup("PARA_NAME") = "Database_Path" Then
                    File1.Path = rsDataSetup("PARA_VALUE")
                    Dir1.Path = File1.Path
                    Drive1 = Dir1.Path
                    'File1.Path
                End If
                If rsDataSetup("PARA_NAME") = "Database_Name" Then
                    For x = 0 To File1.ListCount - 1
                        If rsDataSetup("PARA_VALUE") = File1.List(x) Then
                            File1.ListIndex = x
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
        rsDataSetup.MoveNext
    Loop
    rsDataSetup.Close

End Sub
Private Sub RetrieveDataLG()
Dim SQLQ
Dim x

    SQLQ = "SELECT * FROM APPLICATION_PARAMETER WHERE PARA_TYPE='Integration' AND PARA_CATEGORY='" & fglbProduct & "' AND PARA_CATEGORY2='Database Setup'"
    rsDataSetup.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    
    Do Until rsDataSetup.EOF
        
        If rsDataSetup("PARA_NAME") = "Database_Name" Then
            txtDatabaseName.Text = rsDataSetup("PARA_VALUE")
        End If
        If rsDataSetup("PARA_NAME") = "Database_Server" Then
            txtDatabaseServer.Text = rsDataSetup("PARA_VALUE")
        End If
        If rsDataSetup("PARA_NAME") = "User_Name" Then
            txtUsername.Text = rsDataSetup("PARA_VALUE")
        End If
        If rsDataSetup("PARA_NAME") = "Password" Then
            txtPassword.Text = rsDataSetup("PARA_VALUE")
        End If

        rsDataSetup.MoveNext
    Loop
    rsDataSetup.Close
End Sub


Private Sub SaveData()
    Dim SQLQ
    Dim x
    SQLQ = "DELETE FROM APPLICATION_PARAMETER WHERE PARA_TYPE='Integration' AND PARA_CATEGORY='" & fglbProduct & "' "
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    
    SQLQ = "SELECT * FROM APPLICATION_PARAMETER WHERE PARA_TYPE='Integration' AND PARA_CATEGORY='" & fglbProduct & "' "
    rsDataSetup.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            
    rsDataSetup.AddNew
    rsDataSetup("PARA_TYPE") = "Integration"
    rsDataSetup("PARA_CATEGORY") = fglbProduct
    rsDataSetup("PARA_NAME") = "Version_Info"
    
    
    If comVersion.ListIndex = 1 Then
        rsDataSetup("PARA_VALUE") = "MS SQL Server"
    Else
        rsDataSetup("PARA_VALUE") = "MS Access"
    End If
    
    rsDataSetup.Update
    
    If comVersion.ListIndex = 1 Then
        rsDataSetup.AddNew
        rsDataSetup("PARA_TYPE") = "Integration"
        rsDataSetup("PARA_CATEGORY") = fglbProduct
        rsDataSetup("PARA_NAME") = "Database_Name"
        rsDataSetup("PARA_VALUE") = txtDatabaseName
        rsDataSetup.Update
    
        rsDataSetup.AddNew
        rsDataSetup("PARA_TYPE") = "Integration"
        rsDataSetup("PARA_CATEGORY") = fglbProduct
        rsDataSetup("PARA_NAME") = "Database_Server"
        rsDataSetup("PARA_VALUE") = txtDatabaseServer
        rsDataSetup.Update
        
        rsDataSetup.AddNew
        rsDataSetup("PARA_TYPE") = "Integration"
        rsDataSetup("PARA_CATEGORY") = fglbProduct
        rsDataSetup("PARA_NAME") = "User_Name"
        rsDataSetup("PARA_VALUE") = txtUsername
        rsDataSetup.Update
        
        rsDataSetup.AddNew
        rsDataSetup("PARA_TYPE") = "Integration"
        rsDataSetup("PARA_CATEGORY") = fglbProduct
        rsDataSetup("PARA_NAME") = "Password"
        rsDataSetup("PARA_VALUE") = txtPassword
        rsDataSetup.Update
        
    Else
        rsDataSetup.AddNew
        rsDataSetup("PARA_TYPE") = "Integration"
        rsDataSetup("PARA_CATEGORY") = fglbProduct
        rsDataSetup("PARA_NAME") = "Database_Path"
        rsDataSetup("PARA_VALUE") = File1.Path
        rsDataSetup.Update
        
        rsDataSetup.AddNew
        rsDataSetup("PARA_TYPE") = "Integration"
        rsDataSetup("PARA_CATEGORY") = fglbProduct
        rsDataSetup("PARA_NAME") = "Database_Name"
        rsDataSetup("PARA_VALUE") = File1.ListIndex
        rsDataSetup.Update
    End If

rsDataSetup.Close
'End If
End Sub

Private Sub SaveDataLG() 'Simona - just SQL db, no combobox - Leeds Grenville CAS ticket #14890
Dim SQLQ
    Dim x
    SQLQ = "DELETE FROM APPLICATION_PARAMETER WHERE PARA_TYPE='Integration' AND PARA_CATEGORY='" & fglbProduct & "' AND PARA_CATEGORY2='Database Setup'"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    
    SQLQ = "SELECT * FROM APPLICATION_PARAMETER WHERE PARA_TYPE='Integration' AND PARA_CATEGORY='" & fglbProduct & "' AND PARA_CATEGORY2='Database Setup'"
    rsDataSetup.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            
    rsDataSetup.AddNew
    rsDataSetup("PARA_TYPE") = "Integration"
    rsDataSetup("PARA_CATEGORY") = fglbProduct
    rsDataSetup("PARA_CATEGORY2") = "Database Setup"
    rsDataSetup("PARA_NAME") = "Version_Info"
    
    If glbCompSerial = "S/N - 2233W" Then
        rsDataSetup("PARA_VALUE") = "MS SQL Server"
    End If
    
    rsDataSetup.Update
    
    If glbCompSerial = "S/N - 2233W" Then
  
            rsDataSetup.AddNew
            rsDataSetup("PARA_TYPE") = "Integration"
            rsDataSetup("PARA_CATEGORY") = fglbProduct
            rsDataSetup("PARA_CATEGORY2") = "Database Setup"
            rsDataSetup("PARA_NAME") = "Database_Name"
            rsDataSetup("PARA_VALUE") = txtDatabaseName
            rsDataSetup.Update
        
            rsDataSetup.AddNew
            rsDataSetup("PARA_TYPE") = "Integration"
            rsDataSetup("PARA_CATEGORY") = fglbProduct
            rsDataSetup("PARA_CATEGORY2") = "Database Setup"
            rsDataSetup("PARA_NAME") = "Database_Server"
            rsDataSetup("PARA_VALUE") = txtDatabaseServer
            rsDataSetup.Update
            
            rsDataSetup.AddNew
            rsDataSetup("PARA_TYPE") = "Integration"
            rsDataSetup("PARA_CATEGORY") = fglbProduct
            rsDataSetup("PARA_CATEGORY2") = "Database Setup"
            rsDataSetup("PARA_NAME") = "User_Name"
            rsDataSetup("PARA_VALUE") = txtUsername
            rsDataSetup.Update
            
            rsDataSetup.AddNew
            rsDataSetup("PARA_TYPE") = "Integration"
            rsDataSetup("PARA_CATEGORY") = fglbProduct
            rsDataSetup("PARA_CATEGORY2") = "Database Setup"
            rsDataSetup("PARA_NAME") = "Password"
            rsDataSetup("PARA_VALUE") = txtPassword
            rsDataSetup.Update
    End If 'Simona - end - Leeds Grenville CAS ticket #14890
    rsDataSetup.Close

End Sub
Public Property Let Product_Info(vData As String)
fglbProduct = vData
End Property
