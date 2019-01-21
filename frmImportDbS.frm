VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmImportDbSQL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Database(SQL to SQL)"
   ClientHeight    =   7140
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4218.547
   ScaleMode       =   0  'User
   ScaleWidth      =   10211.04
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSQLDiffConnection 
      Caption         =   "Source Database Connection Information"
      Height          =   3015
      Left            =   5640
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox txtSQLServer 
         Height          =   330
         Left            =   2160
         TabIndex        =   24
         Text            =   ".\SQLEXPRESS"
         Top             =   435
         Width           =   2385
      End
      Begin VB.TextBox txtSQLUserPsw 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   23
         Text            =   "sa"
         Top             =   1920
         Width           =   2385
      End
      Begin VB.TextBox txtSQLUserName 
         Height          =   330
         Left            =   2160
         TabIndex        =   22
         Text            =   "sa"
         Top             =   1440
         Width           =   2385
      End
      Begin VB.CheckBox chkIncSQLDoc 
         Caption         =   "Include DOC database"
         Height          =   375
         Left            =   2160
         TabIndex        =   21
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox txtSQLDBName 
         Height          =   330
         Left            =   2160
         TabIndex        =   20
         Text            =   "INFOHR"
         Top             =   960
         Width           =   2385
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&Server:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   435
         Width           =   540
      End
      Begin VB.Label lblLabels 
         Caption         =   "Password:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   27
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "User Name:"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Database Name:"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.Frame fraOraConnection 
      Caption         =   "Oracle Connection Information"
      Height          =   2415
      Left            =   5640
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox txtOraServer 
         Height          =   330
         Left            =   2160
         TabIndex        =   15
         Text            =   "IHRDEMO76"
         Top             =   435
         Width           =   2385
      End
      Begin VB.TextBox txtUserPsw 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   14
         Text            =   "infohr"
         Top             =   1440
         Width           =   2385
      End
      Begin VB.TextBox txtUserName 
         Height          =   330
         Left            =   2160
         TabIndex        =   13
         Text            =   "infohr"
         Top             =   960
         Width           =   2385
      End
      Begin VB.CheckBox chkIncDoc 
         Caption         =   "Include DOC database"
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&Server:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   435
         Width           =   540
      End
      Begin VB.Label lblLabels 
         Caption         =   "Password:"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "User Name:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.CheckBox chkScript 
      Caption         =   "Create Script"
      Height          =   315
      Left            =   3360
      TabIndex        =   10
      Top             =   3540
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame fraConnection 
      Caption         =   "Connection Information"
      Height          =   2655
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   5295
      Begin VB.TextBox txtDatabaseD 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   9
         Top             =   1800
         Width           =   2385
      End
      Begin VB.TextBox txtServer 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2160
         TabIndex        =   5
         Top             =   435
         Width           =   2385
      End
      Begin VB.TextBox txtDatabaseS 
         Height          =   300
         Left            =   2160
         TabIndex        =   4
         Top             =   1140
         Width           =   2385
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Destination  Database:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1620
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&Server:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   7
         Top             =   510
         Width           =   540
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Source Database:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1290
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   2
      Top             =   6585
      Width           =   10875
      _Version        =   65536
      _ExtentX        =   19182
      _ExtentY        =   979
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
      BevelOuter      =   1
      BevelInner      =   2
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Tag             =   "Accept values and proceed"
      Top             =   3540
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   1
      Tag             =   "Close and exit this screen"
      Top             =   3540
      Width           =   1425
   End
End
Attribute VB_Name = "frmImportDbSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xDatabaseType As String
Dim oraAdoIHRDB As String
Dim oraAdoIHRDB_DOC As String
Dim xDiffSQLServer As String 'Ticket #19762

Private Sub chkScript_Click()
If chkScript Then
  txtDatabaseS.Text = SQLDatabaseName
End If
End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)  '19Aug99 js
End Sub
Private Sub cmdOK1_Click()
Dim dbSource As New ADODB.Connection, dbDestination As New ADODB.Connection
Dim catSource As New ADOX.Catalog, catDestination As New ADOX.Catalog
Dim FldName, FldName1
Dim rstSource As New ADODB.Recordset, rstDestination As New ADODB.Recordset
Dim tblNameList
Dim iCounter As Integer, iTableCounter As Integer, iFieldCounter As Integer
Dim SQLQ, tblName, frtLetter
Dim lRecordCounter As Long
Dim asDatabaseName(1) As String        'Jaddy 6/4/99
Dim fldValue
Dim tblType
'''On Error GoTo LocalErrorHandler:

'get the directory that the use wants to import from...
Screen.MousePointer = vbHourglass
SSPanel1.Caption = "Please wait"

dbSource.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & SQLUserName & ";Password=" & SQLUserPassword & ";Initial Catalog=" & txtDatabaseS.Text & ";Data Source=" & SQLServerName
Set dbDestination = gdbAdoIhr001
Set catSource.ActiveConnection = dbSource
Set catDestination.ActiveConnection = dbDestination
tblNameList = ""
For iTableCounter = 0 To catDestination.Tables.count - 1
    If Mid(catDestination.Tables(iTableCounter).Type, 1, 6) <> "SYSTEM" Then
        If catSource.Tables(iTableCounter).Type = "TABLE" Then
            tblName = catDestination.Tables(iTableCounter).name
            'If InStr("Employee_Turnover,HRATTBAT,HRATTWRK,HRATTWRK1,INFO_HR_TABLES,HRSECWRK,HRCOEWRK,HREMPWRK,HRENTHRS,HRENTWRK,HRSENHRS,HRSENRPT,", tblName & ",") = 0 Then
            If InStr("Employee_Turnover,HRATTBAT,HRATTWRK,HRATTWRK1,INFO_HR_TABLES,HRSECWRK,HRCOEWRK,HREMPWRK,HRENTWRK,HRSENHRS,HRSENRPT,", tblName & ",") = 0 Then
                If InStr("HRJOBBUD,HRPOERPT,HRPOETMP,", tblName & ",") = 0 Then
                    If Left(tblName, 6) <> "PAYWEB" Then
                        If Left(tblName, 4) <> "Old_" Then
                            tblNameList = tblNameList & tblName & ","
                        End If
                    End If
                End If
            End If
        End If
    End If
Next
Dim OutPut
Dim xFieldList



Open "g:\oc\passdata.sql" For Output As #1

For iTableCounter = 0 To catSource.Tables.count - 1
    tblName = catSource.Tables(iTableCounter).name
    If catSource.Tables(iTableCounter).Type = "TABLE" Then
        If Left(tblName, 4) = "Old_" Then
            tblName = Mid(tblName, 5)
            SQLQ = tblName
            'If UCase(tblName) = "HRPROV" Then SQLQ = "SELECT * FROM HRPROV WHERE CODE IS NOT NULL"
            frtLetter = ""
            If InStr(UCase(tblNameList), UCase(tblName) & ",") > 0 Then
                rstSource.Open "Old_" & SQLQ, dbSource
                SSPanel1.Caption = tblName
                
                rstDestination.CursorType = adOpenKeyset
                rstDestination.LockType = adLockOptimistic
                rstDestination.Open tblName, dbDestination, , , adCmdTable
                    xFieldList = ""
                    For iFieldCounter = 0 To rstSource.Fields.count - 1
                        If InStr("IE_COMM1,AU_BSMOKER,BD_SIN,ED_USRC1,ED_USRC2,ED_USRC3,ED_USRC4,ED_USRC5,ED_USRN1,ED_USRN2,ED_USRN3,ED_USRD1,ED_USRD2,ED_USRD3,", UCase(rstSource.Fields(iFieldCounter).name) & ",") = 0 Then
                            If (Right(rstSource.Fields(iFieldCounter).name, 3) <> "_ID" And rstSource.Fields(iFieldCounter).name <> "ID") Or tblName = "HRNEWHIRE" Or Right(rstSource.Fields(iFieldCounter).name, 10) = "PAYROLL_ID" Or Right(rstSource.Fields(iFieldCounter).name, 6) = "JOB_ID" Then
                                FldName = rstSource.Fields(iFieldCounter).name
                                xFieldList = xFieldList & "[" & FldName & "],"
                            Else
                               Debug.Print tblName, rstSource.Fields(iFieldCounter).name
                            End If
                        End If
                    Next '
                    xFieldList = Left(xFieldList, Len(xFieldList) - 1)
                    
                    OutPut = "INSERT INTO " & tblName & " (" & xFieldList & ") SELECT " & xFieldList & " FROM Old_" & tblName
                    Print #1, "DELETE FROM " & tblName
                    Print #1, OutPut
                    Print #1, "GO"
                rstDestination.Close
                rstSource.Close
                Print #1, "DROP TABLE Old_" & tblName
                Print #1, "GO"
            Else
                'If InStr("ATTD_MATRIX,COURSE_MATRIX,HRSECWRK,POSITION_MATRIX,Employee_Turnover,HRATTBAT,HRATTWRK,HRATTWRK1,HRATTWRK2,HRCOEWRK,HREMPWRK,HRENTHRS,HRENTWRK,HRSENHRS,HRSENRPT,", tblName & ",") <> 0 Then
                If InStr("ATTD_MATRIX,COURSE_MATRIX,HRSECWRK,POSITION_MATRIX,Employee_Turnover,HRATTBAT,HRATTWRK,HRATTWRK1,HRATTWRK2,HRCOEWRK,HREMPWRK,HRENTWRK,HRSENHRS,HRSENRPT,", tblName & ",") <> 0 Then
                    Print #1, "DROP TABLE Old_" & tblName
                    Print #1, "GO"
                End If
                If Left(tblName, 3) = "LN_" Or tblName = "HR_WILL_TERM" Then
                    Print #1, "sp_rename 'Old_" & tblName & "','" & tblName & "'"
                    Print #1, "GO"
                End If
            End If
        End If

    End If
Next 'iTableCounter
Print #1, ""
Close #1
dbSource.Close
dbDestination.Close

MsgBox "Converted successfully, Please click OK to close the application."
End
LocalErrorHandlerExit:
Screen.MousePointer = vbDefault
'close data objects...
        
Unload Me
Exit Sub
    
LocalErrorHandler:
            
    Select Case Err.Number
        Case 91, 3420
            Resume Next
        Case 3022
            'duplicate key/primary/data...
            Resume Next
        Case Else
            MsgBox "Unhandled error " & Err.Number & " " & Err.Description
            Resume LocalErrorHandlerExit:
    End Select
    
End Sub
Private Sub cmdOK_GotFocus()
    Call SetPanHelp(ActiveControl)  '19Aug99 js
End Sub
Private Sub Form_Load()
 Dim I As Integer

  'Me.Caption = FORMCAPTION
    
    Me.Width = 5850
    glbOnTop = "FRMIMPORTDBSQL"
    Me.Height = 5250 'Ticket #19762
    If Me.DatabaseType = "SQL Server" Then
        If Me.DiffSQLServer = "Y" Then 'Ticket #19762
            frmImportDbSQL.Caption = "Import Database(SQL to SQL)"
            fraConnection.Visible = False
            fraSQLDiffConnection.Top = fraConnection.Top
            fraSQLDiffConnection.Left = fraConnection.Left
            fraSQLDiffConnection.Visible = True
        Else
            txtDatabaseS.Text = "INFOHR"
            txtDatabaseD.Text = SQLDatabaseName
            txtServer.Text = SQLServerName
        End If
    End If
    
    If Me.DatabaseType = "Oracle" Then
        frmImportDbSQL.Caption = "Import Database(Oracle to SQL)"
        fraConnection.Visible = False
        fraOraConnection.Top = fraConnection.Top
        fraOraConnection.Left = fraConnection.Left
        fraOraConnection.Visible = True
    End If
End Sub


Private Sub cmdOK_Click()
    If Me.DatabaseType = "SQL Server" Then
        If Me.DiffSQLServer = "Y" Then 'Ticket #19762
            Call ImpFromDiffSQL
        Else
            Call ImpFromSQL
        End If
    End If
    If Me.DatabaseType = "Oracle" Then
        Call ImpFromOracle
    End If
End Sub

Private Sub ImpFromDiffSQL() 'Ticket #19762
Dim dbSource As New ADODB.Connection, dbDestination As New ADODB.Connection
Dim catSource As New ADOX.Catalog, catDestination As New ADOX.Catalog
Dim FldName, FldName1
Dim rstSource As New ADODB.Recordset, rstDestination As New ADODB.Recordset
Dim tblNameList
Dim iCounter As Integer, iTableCounter As Integer, iFieldCounter As Integer
Dim SQLQ, tblName, frtLetter
Dim lRecordCounter As Long
Dim asDatabaseName(1) As String        '
Dim fldValue
Dim tblType
Dim K As Double
On Error GoTo LocalErrorHandler:
'get the directory that the use wants to import from...

'GoTo Ora_Doc

If Len(txtOraServer.Text) = 0 Then
    MsgBox "Please enter Server"
    Exit Sub
End If
If Len(txtUserName.Text) = 0 Then
    MsgBox "Please enter User Name"
    Exit Sub
End If
If Len(txtUserPsw.Text) = 0 Then
    MsgBox "Please enter Password"
    Exit Sub
End If

'Test the connection
'oraAdoIHRDB = "Provider=OraOLEDB.Oracle.1;Password=" & txtUserPsw.Text & ";Persist Security Info=True;User ID=" & txtUserName.Text & ";Data Source=" & txtOraServer.Text & ""
'oraAdoIHRDB_DOC = "Provider=OraOLEDB.Oracle.1;Password=" & txtUserPsw.Text & ";Persist Security Info=True;User ID=" & txtUserName.Text & ";Data Source=" & txtOraServer.Text & "_DOC"
oraAdoIHRDB = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & txtSQLUserName.Text & ";Password=" & txtSQLUserPsw.Text & ";Initial Catalog=" & txtSQLDBName.Text & ";Data Source=" & txtSQLServer.Text
oraAdoIHRDB_DOC = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & txtSQLUserName.Text & ";Password=" & txtSQLUserPsw.Text & ";Initial Catalog=" & txtSQLDBName.Text & "_DOC" & ";Data Source=" & txtSQLServer.Text

dbSource.Open oraAdoIHRDB

Screen.MousePointer = vbHourglass
SSPanel1.Caption = "Please wait"
'dbSource.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & SQLUserName & ";Password=" & SQLUserPassword & ";Initial Catalog=" & txtDatabaseS.Text & ";Data Source=" & SQLServerName
'Exit Sub

If chkIncSQLDoc.Value Then
GoTo SQL_Doc
End If

Set dbDestination = gdbAdoIhr001
Set catSource.ActiveConnection = dbSource
Set catDestination.ActiveConnection = dbDestination
tblNameList = ","
For iTableCounter = 0 To catDestination.Tables.count - 1
    If Mid(catDestination.Tables(iTableCounter).Type, 1, 6) <> "SYSTEM" Then
        tblName = catDestination.Tables(iTableCounter).name
        'If InStr("INFO_HR_TABLES,HRTABDES,HRNEWHIRE,Employee_Turnover,HRATTBAT,HRATTWRK,HRATTWRK1,HRCOEWRK,HREMPWRK,HRENTHRS,HRENTWRK,HRSENHRS,HRSENRPT,", tblName & ",") = 0 Then
        If InStr("INFO_HR_TABLES,HRTABDES,HRNEWHIRE,Employee_Turnover,HRATTBAT,HRATTWRK,HRATTWRK1,HRCOEWRK,HREMPWRK,HRENTWRK,HRSENHRS,HRSENRPT,", tblName & ",") = 0 Then
            tblNameList = tblNameList & tblName & ","
        End If
    End If
Next
For iTableCounter = 0 To catSource.Tables.count - 1
    tblName = catSource.Tables(iTableCounter).name
    'D.Muskoka only 'for District Muskoka  - Oracle Ticket #19111
    'If catSource.Tables(iTableCounter).Type = "TABLE" And Left(tblName, 2) = "DK" Then
    If catSource.Tables(iTableCounter).Type = "TABLE" Then
        DoEvents
        'SQLQ = tblName
        SQLQ = "SELECT * FROM " & tblName
        If UCase(tblName) = "HRPROV" Then SQLQ = "SELECT * FROM HRPROV WHERE CODE IS NOT NULL"
        frtLetter = ""
        If InStr(UCase(tblNameList), "," & UCase(tblName) & ",") > 0 Then
            'rstSource.Open SQLQ, dbSource
            rstSource.Open SQLQ, dbSource, adOpenStatic, adLockOptimistic
            SSPanel1.Caption = tblName
            dbDestination.BeginTrans
            dbDestination.Execute "DELETE FROM " & tblName
            dbDestination.CommitTrans
            rstDestination.CursorType = adOpenKeyset
            rstDestination.LockType = adLockOptimistic
            rstDestination.Open tblName, dbDestination, , , adCmdTable
            K = 0
            While Not rstSource.EOF
'If tblName = "HRAUDIT" Then
'Debug.Print rstSource.RecordCount
'End If
                DoEvents
                SSPanel1.Caption = tblName & " -  " & rstSource.RecordCount & ": " & Str(K): K = K + 1
                rstDestination.AddNew
                For iFieldCounter = 0 To rstSource.Fields.count - 1
                    If InStr("AU_BSMOKER,BD_SIN,ED_USRC1,ED_USRC2,ED_USRC3,ED_USRC4,ED_USRC5,ED_USRN1,ED_USRN2,ED_USRN3,ED_USRD1,ED_USRD2,ED_USRD3,", UCase(rstSource.Fields(iFieldCounter).name) & ",") = 0 Then
                        'If True Then 'D.Muskoka only Ticket #19111
                        If (Right(rstSource.Fields(iFieldCounter).name, 3) <> "_ID" And rstSource.Fields(iFieldCounter).name <> "ID") Or tblName = "HRNEWHIRE" Or tblName = "HREMP" Then
                            FldName = rstSource.Fields(iFieldCounter).name
                            If InStr(UCase(FldName), "LUSER") <> 0 Or InStr(UCase(FldName), "WRKEMP") <> 0 Then
                                'Debug.Print "Alter table " & tblName & " alter column " & FldName & " varchar(25)"
                                'Debug.Print "go"
                            End If
                            fldValue = rstSource(FldName)
                            If rstSource(FldName).Type = dbDate Then
                                If Year(fldValue) < 1900 Then
                                    fldValue = Format(fldValue, "mm/dd/") & "19" & Format(fldValue, "yy")
                                End If
                            End If
                            If UCase(FldName) = "LTIME" And UCase(tblName) = "HREARN" Then
                                fldValue = Left(rstSource(FldName) & "", 8)
                            End If
                            rstDestination(frtLetter & Replace(FldName, "-", "_")) = fldValue
                        End If
                    End If
                Next '
                rstDestination.Update
                rstSource.MoveNext
            Wend
            rstDestination.Close
            rstSource.Close
        End If
    End If
Next 'iTableCounter
dbSource.Close
dbDestination.Close

SQL_Doc:
If chkIncSQLDoc.Value Then
    Call ImpFromSQL_Doc
End If


MsgBox "Converted successfully, Please click OK to close the application."
End
LocalErrorHandlerExit:
Screen.MousePointer = vbDefault
'close data objects...
Unload Me
Exit Sub
LocalErrorHandler:
    If InStr(Err.Description, "TNS:") > 0 Then
        MsgBox "Incorrect connection information!"
        Exit Sub
    End If
    If InStr(Err.Description, "Cannot open database") > 0 Then
        MsgBox Err.Description
        Exit Sub
    End If
    Select Case Err.Number
        Case 91, 3420
            Resume Next
        Case 3022
            'duplicate key/primary/data...
            Resume Next
        Case Else
            'MsgBox "Unhandled error " & Err.Number & " " & Err.Description
            Resume Next
            'Resume LocalErrorHandlerExit:
    End Select

End Sub

Private Sub ImpFromOracle()
Dim dbSource As New ADODB.Connection, dbDestination As New ADODB.Connection
Dim catSource As New ADOX.Catalog, catDestination As New ADOX.Catalog
Dim FldName, FldName1
Dim rstSource As New ADODB.Recordset, rstDestination As New ADODB.Recordset
Dim tblNameList
Dim iCounter As Integer, iTableCounter As Integer, iFieldCounter As Integer
Dim SQLQ, tblName, frtLetter
Dim lRecordCounter As Long
Dim asDatabaseName(1) As String        'Jaddy 6/4/99
Dim fldValue
Dim tblType
Dim K As Double
On Error GoTo LocalErrorHandler:
'get the directory that the use wants to import from...

'GoTo Ora_Doc

If Len(txtOraServer.Text) = 0 Then
    MsgBox "Please enter Server"
    Exit Sub
End If
If Len(txtUserName.Text) = 0 Then
    MsgBox "Please enter User Name"
    Exit Sub
End If
If Len(txtUserPsw.Text) = 0 Then
    MsgBox "Please enter Password"
    Exit Sub
End If

'Test the connection
oraAdoIHRDB = "Provider=OraOLEDB.Oracle.1;Password=" & txtUserPsw.Text & ";Persist Security Info=True;User ID=" & txtUserName.Text & ";Data Source=" & txtOraServer.Text & ""
oraAdoIHRDB_DOC = "Provider=OraOLEDB.Oracle.1;Password=" & txtUserPsw.Text & ";Persist Security Info=True;User ID=" & txtUserName.Text & ";Data Source=" & txtOraServer.Text & "_DOC"
dbSource.Open oraAdoIHRDB

Screen.MousePointer = vbHourglass
SSPanel1.Caption = "Please wait"
'dbSource.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & SQLUserName & ";Password=" & SQLUserPassword & ";Initial Catalog=" & txtDatabaseS.Text & ";Data Source=" & SQLServerName
'Exit Sub

If chkIncDoc.Value Then
GoTo Ora_Doc
End If

Set dbDestination = gdbAdoIhr001
Set catSource.ActiveConnection = dbSource
Set catDestination.ActiveConnection = dbDestination
tblNameList = ","
For iTableCounter = 0 To catDestination.Tables.count - 1
    If Mid(catDestination.Tables(iTableCounter).Type, 1, 6) <> "SYSTEM" Then
        tblName = catDestination.Tables(iTableCounter).name
        'If InStr("INFO_HR_TABLES,HRTABDES,HRNEWHIRE,Employee_Turnover,HRATTBAT,HRATTWRK,HRATTWRK1,HRCOEWRK,HREMPWRK,HRENTHRS,HRENTWRK,HRSENHRS,HRSENRPT,", tblName & ",") = 0 Then
        If InStr("INFO_HR_TABLES,HRTABDES,HRNEWHIRE,Employee_Turnover,HRATTBAT,HRATTWRK,HRATTWRK1,HRCOEWRK,HREMPWRK,HRENTWRK,HRSENHRS,HRSENRPT,", tblName & ",") = 0 Then
            tblNameList = tblNameList & tblName & ","
        End If
    End If
Next
For iTableCounter = 0 To catSource.Tables.count - 1
    tblName = catSource.Tables(iTableCounter).name
    If catSource.Tables(iTableCounter).Type = "TABLE" Then
        DoEvents
        'SQLQ = tblName
        SQLQ = "SELECT * FROM " & tblName
        If UCase(tblName) = "HRPROV" Then SQLQ = "SELECT * FROM HRPROV WHERE CODE IS NOT NULL"
        frtLetter = ""
        If InStr(UCase(tblNameList), "," & UCase(tblName) & ",") > 0 Then
            'rstSource.Open SQLQ, dbSource
            rstSource.Open SQLQ, dbSource, adOpenStatic, adLockOptimistic
            SSPanel1.Caption = tblName
            dbDestination.BeginTrans
            dbDestination.Execute "DELETE FROM " & tblName
            dbDestination.CommitTrans
            rstDestination.CursorType = adOpenKeyset
            rstDestination.LockType = adLockOptimistic
            rstDestination.Open tblName, dbDestination, , , adCmdTable
            K = 0
            While Not rstSource.EOF
'If tblName = "HRAUDIT" Then
'Debug.Print rstSource.RecordCount
'End If
                DoEvents
                SSPanel1.Caption = tblName & " -  " & rstSource.RecordCount & ": " & Str(K): K = K + 1
                rstDestination.AddNew
                For iFieldCounter = 0 To rstSource.Fields.count - 1
                    If InStr("AU_BSMOKER,BD_SIN,ED_USRC1,ED_USRC2,ED_USRC3,ED_USRC4,ED_USRC5,ED_USRN1,ED_USRN2,ED_USRN3,ED_USRD1,ED_USRD2,ED_USRD3,", UCase(rstSource.Fields(iFieldCounter).name) & ",") = 0 Then
                        If (Right(rstSource.Fields(iFieldCounter).name, 3) <> "_ID" And rstSource.Fields(iFieldCounter).name <> "ID") Or tblName = "HRNEWHIRE" Or tblName = "HREMP" Then
                            FldName = rstSource.Fields(iFieldCounter).name
                            If InStr(UCase(FldName), "LUSER") <> 0 Or InStr(UCase(FldName), "WRKEMP") <> 0 Then
                                'Debug.Print "Alter table " & tblName & " alter column " & FldName & " varchar(25)"
                                'Debug.Print "go"
                            End If
                            fldValue = rstSource(FldName)
                            If rstSource(FldName).Type = dbDate Then
                                If Year(fldValue) < 1900 Then
                                    fldValue = Format(fldValue, "mm/dd/") & "19" & Format(fldValue, "yy")
                                End If
                            End If
                            If UCase(FldName) = "LTIME" And UCase(tblName) = "HREARN" Then
                                fldValue = Left(rstSource(FldName) & "", 8)
                            End If
                            rstDestination(frtLetter & Replace(FldName, "-", "_")) = fldValue
                        End If
                    End If
                Next '
                rstDestination.Update
                rstSource.MoveNext
            Wend
            rstDestination.Close
            rstSource.Close
        End If
    End If
Next 'iTableCounter
dbSource.Close
dbDestination.Close

Ora_Doc:
If chkIncDoc.Value Then
    Call ImpFromOra_Doc
End If


MsgBox "Converted successfully, Please click OK to close the application."
End
LocalErrorHandlerExit:
Screen.MousePointer = vbDefault
'close data objects...
Unload Me
Exit Sub
LocalErrorHandler:
    If InStr(Err.Description, "TNS:") > 0 Then
        MsgBox "Incorrect connection information!"
        Exit Sub
    End If
    Select Case Err.Number
        Case 91, 3420
            Resume Next
        Case 3022
            'duplicate key/primary/data...
            Resume Next
        Case Else
            'MsgBox "Unhandled error " & Err.Number & " " & Err.Description
            Resume Next
            'Resume LocalErrorHandlerExit:
    End Select
End Sub

Private Sub ImpFromSQL_Doc() 'Ticket #19762
Dim dbSource As New ADODB.Connection, dbDestination As New ADODB.Connection
Dim catSource As New ADOX.Catalog, catDestination As New ADOX.Catalog
Dim FldName, FldName1
Dim rstSource As New ADODB.Recordset, rstDestination As New ADODB.Recordset
Dim tblNameList
Dim iCounter As Integer, iTableCounter As Integer, iFieldCounter As Integer
Dim SQLQ, tblName, frtLetter
Dim lRecordCounter As Long
Dim asDatabaseName(1) As String        'Jaddy 6/4/99
Dim fldValue
Dim tblType
Dim K As Double
On Error GoTo LocalErrorHandler:
'get the directory that the use wants to import from...


'Test the connection
'oraAdoIHRDB = "Provider=OraOLEDB.Oracle.1;Password=" & txtUserPsw.Text & ";Persist Security Info=True;User ID=" & txtUserName.Text & ";Data Source=" & txtOraServer.Text & "_DOC"
oraAdoIHRDB = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & txtSQLUserName.Text & ";Password=" & txtSQLUserPsw.Text & ";Initial Catalog=" & txtSQLDBName.Text & "_DOC" & ";Data Source=" & txtSQLServer.Text
'oraAdoIHRDB = "Provider=OraOLEDB.Oracle.1;Password=" & txtUserPsw.Text & ";Persist Security Info=True;User ID=" & txtUserName.Text & ";Data Source=" & txtOraServer.Text
dbSource.Open oraAdoIHRDB

Screen.MousePointer = vbHourglass
SSPanel1.Caption = "Please wait"
'dbSource.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & SQLUserName & ";Password=" & SQLUserPassword & ";Initial Catalog=" & txtDatabaseS.Text & ";Data Source=" & SQLServerName
'Exit Sub

Set dbDestination = gdbAdoIhr001_DOC
Set catSource.ActiveConnection = dbSource
Set catDestination.ActiveConnection = dbDestination
tblNameList = ","
For iTableCounter = 0 To catDestination.Tables.count - 1
    If Mid(catDestination.Tables(iTableCounter).Type, 1, 6) <> "SYSTEM" Then
        tblName = catDestination.Tables(iTableCounter).name
        'If InStr("INFO_HR_TABLES,HRTABDES,HRNEWHIRE,Employee_Turnover,HRATTBAT,HRATTWRK,HRATTWRK1,HRCOEWRK,HREMPWRK,HRENTHRS,HRENTWRK,HRSENHRS,HRSENRPT,", tblName & ",") = 0 Then
            tblNameList = tblNameList & tblName & ","
        'End If
    End If
Next
For iTableCounter = 0 To catSource.Tables.count - 1
    tblName = catSource.Tables(iTableCounter).name
    If catSource.Tables(iTableCounter).Type = "TABLE" Then
        DoEvents
        'SQLQ = tblName
        SQLQ = "SELECT * FROM " & tblName
        If UCase(tblName) = "HRPROV" Then SQLQ = "SELECT * FROM HRPROV WHERE CODE IS NOT NULL"
        frtLetter = ""
        If InStr(UCase(tblNameList), "," & UCase(tblName) & ",") > 0 Then
            'rstSource.Open SQLQ, dbSource
            rstSource.Open SQLQ, dbSource, adOpenStatic, adLockOptimistic
            SSPanel1.Caption = tblName
            dbDestination.BeginTrans
            dbDestination.Execute "DELETE FROM " & tblName
            dbDestination.CommitTrans
            rstDestination.CursorType = adOpenKeyset
            rstDestination.LockType = adLockOptimistic
            rstDestination.Open tblName, dbDestination, , , adCmdTable
            K = 0
            While Not rstSource.EOF
'If tblName = "HRAUDIT" Then
'Debug.Print rstSource.RecordCount
'End If
                DoEvents
                SSPanel1.Caption = tblName & " -  " & rstSource.RecordCount & ": " & Str(K): K = K + 1
                rstDestination.AddNew
                For iFieldCounter = 0 To rstSource.Fields.count - 1
                    If InStr("AU_BSMOKER,BD_SIN,ED_USRC1,ED_USRC2,ED_USRC3,ED_USRC4,ED_USRC5,ED_USRN1,ED_USRN2,ED_USRN3,ED_USRD1,ED_USRD2,ED_USRD3,", UCase(rstSource.Fields(iFieldCounter).name) & ",") = 0 Then
                        If (Right(rstSource.Fields(iFieldCounter).name, 3) <> "_ID" And rstSource.Fields(iFieldCounter).name <> "ID") Or tblName = "HRNEWHIRE" Or tblName = "HREMP" Then
                            FldName = rstSource.Fields(iFieldCounter).name
                            If InStr(UCase(FldName), "LUSER") <> 0 Or InStr(UCase(FldName), "WRKEMP") <> 0 Then
                                'Debug.Print "Alter table " & tblName & " alter column " & FldName & " varchar(25)"
                                'Debug.Print "go"
                            End If
                            fldValue = rstSource(FldName)
                            If rstSource(FldName).Type = dbDate Then
                                If Year(fldValue) < 1900 Then
                                    fldValue = Format(fldValue, "mm/dd/") & "19" & Format(fldValue, "yy")
                                End If
                            End If
                            If UCase(FldName) = "LTIME" And UCase(tblName) = "HREARN" Then
                                fldValue = Left(rstSource(FldName) & "", 8)
                            End If
                            rstDestination(frtLetter & Replace(FldName, "-", "_")) = fldValue
                        End If
                    End If
                Next '
                rstDestination.Update
                rstSource.MoveNext
            Wend
            rstDestination.Close
            rstSource.Close
        End If
    End If
Next 'iTableCounter
dbSource.Close
dbDestination.Close

End

Exit Sub
LocalErrorHandler:
    If InStr(Err.Description, "TNS:") > 0 Then
        MsgBox "Incorrect connection information!"
        Exit Sub
    End If
    If InStr(Err.Description, "Cannot open database") > 0 Then
        MsgBox Err.Description
        Exit Sub
    End If
    Select Case Err.Number
        Case 91, 3420
            Resume Next
        Case 3022
            'duplicate key/primary/data...
            Resume Next
        Case Else
            'MsgBox "Unhandled error " & Err.Number & " " & Err.Description
            Resume Next
            'Resume LocalErrorHandlerExit:
    End Select

End Sub

Private Sub ImpFromOra_Doc()
Dim dbSource As New ADODB.Connection, dbDestination As New ADODB.Connection
Dim catSource As New ADOX.Catalog, catDestination As New ADOX.Catalog
Dim FldName, FldName1
Dim rstSource As New ADODB.Recordset, rstDestination As New ADODB.Recordset
Dim tblNameList
Dim iCounter As Integer, iTableCounter As Integer, iFieldCounter As Integer
Dim SQLQ, tblName, frtLetter
Dim lRecordCounter As Long
Dim asDatabaseName(1) As String        'Jaddy 6/4/99
Dim fldValue
Dim tblType
Dim K As Double
On Error GoTo LocalErrorHandler:
'get the directory that the use wants to import from...


'Test the connection
'oraAdoIHRDB = "Provider=OraOLEDB.Oracle.1;Password=" & txtUserPsw.Text & ";Persist Security Info=True;User ID=" & txtUserName.Text & ";Data Source=" & txtOraServer.Text & "_DOC"
oraAdoIHRDB = "Provider=OraOLEDB.Oracle.1;Password=" & txtUserPsw.Text & ";Persist Security Info=True;User ID=" & txtUserName.Text & ";Data Source=" & txtOraServer.Text
dbSource.Open oraAdoIHRDB

Screen.MousePointer = vbHourglass
SSPanel1.Caption = "Please wait"
'dbSource.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & SQLUserName & ";Password=" & SQLUserPassword & ";Initial Catalog=" & txtDatabaseS.Text & ";Data Source=" & SQLServerName
'Exit Sub

Set dbDestination = gdbAdoIhr001_DOC
Set catSource.ActiveConnection = dbSource
Set catDestination.ActiveConnection = dbDestination
tblNameList = ","
For iTableCounter = 0 To catDestination.Tables.count - 1
    If Mid(catDestination.Tables(iTableCounter).Type, 1, 6) <> "SYSTEM" Then
        tblName = catDestination.Tables(iTableCounter).name
        'If InStr("INFO_HR_TABLES,HRTABDES,HRNEWHIRE,Employee_Turnover,HRATTBAT,HRATTWRK,HRATTWRK1,HRCOEWRK,HREMPWRK,HRENTHRS,HRENTWRK,HRSENHRS,HRSENRPT,", tblName & ",") = 0 Then
            tblNameList = tblNameList & tblName & ","
        'End If
    End If
Next
For iTableCounter = 0 To catSource.Tables.count - 1
    tblName = catSource.Tables(iTableCounter).name
    If catSource.Tables(iTableCounter).Type = "TABLE" Then
        DoEvents
        'SQLQ = tblName
        SQLQ = "SELECT * FROM " & tblName
        If UCase(tblName) = "HRPROV" Then SQLQ = "SELECT * FROM HRPROV WHERE CODE IS NOT NULL"
        frtLetter = ""
        If InStr(UCase(tblNameList), "," & UCase(tblName) & ",") > 0 Then
            'rstSource.Open SQLQ, dbSource
            rstSource.Open SQLQ, dbSource, adOpenStatic, adLockOptimistic
            SSPanel1.Caption = tblName
            dbDestination.BeginTrans
            dbDestination.Execute "DELETE FROM " & tblName
            dbDestination.CommitTrans
            rstDestination.CursorType = adOpenKeyset
            rstDestination.LockType = adLockOptimistic
            rstDestination.Open tblName, dbDestination, , , adCmdTable
            K = 0
            While Not rstSource.EOF
'If tblName = "HRAUDIT" Then
'Debug.Print rstSource.RecordCount
'End If
                DoEvents
                SSPanel1.Caption = tblName & " -  " & rstSource.RecordCount & ": " & Str(K): K = K + 1
                rstDestination.AddNew
                For iFieldCounter = 0 To rstSource.Fields.count - 1
                    If InStr("AU_BSMOKER,BD_SIN,ED_USRC1,ED_USRC2,ED_USRC3,ED_USRC4,ED_USRC5,ED_USRN1,ED_USRN2,ED_USRN3,ED_USRD1,ED_USRD2,ED_USRD3,", UCase(rstSource.Fields(iFieldCounter).name) & ",") = 0 Then
                        If (Right(rstSource.Fields(iFieldCounter).name, 3) <> "_ID" And rstSource.Fields(iFieldCounter).name <> "ID") Or tblName = "HRNEWHIRE" Or tblName = "HREMP" Then
                            FldName = rstSource.Fields(iFieldCounter).name
                            If InStr(UCase(FldName), "LUSER") <> 0 Or InStr(UCase(FldName), "WRKEMP") <> 0 Then
                                'Debug.Print "Alter table " & tblName & " alter column " & FldName & " varchar(25)"
                                'Debug.Print "go"
                            End If
                            fldValue = rstSource(FldName)
                            If rstSource(FldName).Type = dbDate Then
                                If Year(fldValue) < 1900 Then
                                    fldValue = Format(fldValue, "mm/dd/") & "19" & Format(fldValue, "yy")
                                End If
                            End If
                            If UCase(FldName) = "LTIME" And UCase(tblName) = "HREARN" Then
                                fldValue = Left(rstSource(FldName) & "", 8)
                            End If
                            rstDestination(frtLetter & Replace(FldName, "-", "_")) = fldValue
                        End If
                    End If
                Next '
                rstDestination.Update
                rstSource.MoveNext
            Wend
            rstDestination.Close
            rstSource.Close
        End If
    End If
Next 'iTableCounter
dbSource.Close
dbDestination.Close

End

Exit Sub
LocalErrorHandler:
    If InStr(Err.Description, "TNS:") > 0 Then
        MsgBox "Incorrect connection information!"
        Exit Sub
    End If
    Select Case Err.Number
        Case 91, 3420
            Resume Next
        Case 3022
            'duplicate key/primary/data...
            Resume Next
        Case Else
            'MsgBox "Unhandled error " & Err.Number & " " & Err.Description
            Resume Next
            'Resume LocalErrorHandlerExit:
    End Select
End Sub


Private Sub ImpFromSQL()
If chkScript Then Call cmdOK1_Click: Exit Sub


Dim dbSource As New ADODB.Connection, dbDestination As New ADODB.Connection
Dim catSource As New ADOX.Catalog, catDestination As New ADOX.Catalog
Dim FldName, FldName1
Dim rstSource As New ADODB.Recordset, rstDestination As New ADODB.Recordset
Dim tblNameList
Dim iCounter As Integer, iTableCounter As Integer, iFieldCounter As Integer
Dim SQLQ, tblName, frtLetter
Dim lRecordCounter As Long
Dim asDatabaseName(1) As String        'Jaddy 6/4/99
Dim fldValue
Dim tblType
'''On Error GoTo LocalErrorHandler:
'get the directory that the use wants to import from...
Screen.MousePointer = vbHourglass
SSPanel1.Caption = "Please wait"
dbSource.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & SQLUserName & ";Password=" & SQLUserPassword & ";Initial Catalog=" & txtDatabaseS.Text & ";Data Source=" & SQLServerName
Set dbDestination = gdbAdoIhr001
Set catSource.ActiveConnection = dbSource
Set catDestination.ActiveConnection = dbDestination
tblNameList = ""
For iTableCounter = 0 To catDestination.Tables.count - 1
    If Mid(catDestination.Tables(iTableCounter).Type, 1, 6) <> "SYSTEM" Then
        tblName = catDestination.Tables(iTableCounter).name
        'If InStr("INFO_HR_TABLES,HRTABDES,HRNEWHIRE,Employee_Turnover,HRATTBAT,HRATTWRK,HRATTWRK1,HRCOEWRK,HREMPWRK,HRENTHRS,HRENTWRK,HRSENHRS,HRSENRPT,", tblName & ",") = 0 Then
        If InStr("INFO_HR_TABLES,HRTABDES,HRNEWHIRE,Employee_Turnover,HRATTBAT,HRATTWRK,HRATTWRK1,HRCOEWRK,HREMPWRK,HRENTWRK,HRSENHRS,HRSENRPT,", tblName & ",") = 0 Then
            tblNameList = tblNameList & tblName & ","
        End If
    End If
Next
For iTableCounter = 0 To catSource.Tables.count - 1
    tblName = catSource.Tables(iTableCounter).name
    If catSource.Tables(iTableCounter).Type = "TABLE" Then
        DoEvents
        SQLQ = tblName
        If UCase(tblName) = "HRPROV" Then SQLQ = "SELECT * FROM HRPROV WHERE CODE IS NOT NULL"
        frtLetter = ""
        If InStr(UCase(tblNameList), UCase(tblName) & ",") > 0 Then
            rstSource.Open SQLQ, dbSource
            SSPanel1.Caption = tblName
            dbDestination.Execute "DELETE FROM " & tblName
            rstDestination.CursorType = adOpenKeyset
            rstDestination.LockType = adLockOptimistic
            rstDestination.Open tblName, dbDestination, , , adCmdTable
            While Not rstSource.EOF
                rstDestination.AddNew
                For iFieldCounter = 0 To rstSource.Fields.count - 1
                    If InStr("AU_BSMOKER,BD_SIN,ED_USRC1,ED_USRC2,ED_USRC3,ED_USRC4,ED_USRC5,ED_USRN1,ED_USRN2,ED_USRN3,ED_USRD1,ED_USRD2,ED_USRD3,", UCase(rstSource.Fields(iFieldCounter).name) & ",") = 0 Then
                        If (Right(rstSource.Fields(iFieldCounter).name, 3) <> "_ID" And rstSource.Fields(iFieldCounter).name <> "ID") Or tblName = "HRNEWHIRE" Or tblName = "HREMP" Then
                            FldName = rstSource.Fields(iFieldCounter).name
                            If InStr(UCase(FldName), "LUSER") <> 0 Or InStr(UCase(FldName), "WRKEMP") <> 0 Then
                                Debug.Print "Alter table " & tblName & " alter column " & FldName & " varchar(25)"
                                Debug.Print "go"
                            End If
                            fldValue = rstSource(FldName)
                            If rstSource(FldName).Type = dbDate Then
                                If Year(fldValue) < 1900 Then
                                    fldValue = Format(fldValue, "mm/dd/") & "19" & Format(fldValue, "yy")
                                End If
                            End If
                            If UCase(FldName) = "LTIME" And UCase(tblName) = "HREARN" Then
                                fldValue = Left(rstSource(FldName) & "", 8)
                            End If
                            rstDestination(frtLetter & Replace(FldName, "-", "_")) = fldValue
                        End If
                    End If
                Next '
                rstDestination.Update
                rstSource.MoveNext
            Wend
            rstDestination.Close
            rstSource.Close
        End If
    End If
Next 'iTableCounter
dbSource.Close
dbDestination.Close
MsgBox "Converted successfully, Please click OK to close the application."
End
LocalErrorHandlerExit:
Screen.MousePointer = vbDefault
'close data objects...
Unload Me
Exit Sub
LocalErrorHandler:
    Select Case Err.Number
        Case 91, 3420
            Resume Next
        Case 3022
            'duplicate key/primary/data...
            Resume Next
        Case Else
            MsgBox "Unhandled error " & Err.Number & " " & Err.Description
            Resume LocalErrorHandlerExit:
    End Select
End Sub

Public Property Let DatabaseType(vData As String)
    xDatabaseType = vData
End Property

Public Property Get DatabaseType() As String
    DatabaseType = xDatabaseType
End Property

Public Property Let DiffSQLServer(vData As String) 'Ticket #19762
    xDiffSQLServer = vData
End Property
Public Property Get DiffSQLServer() As String 'Ticket #19762
    DiffSQLServer = xDiffSQLServer
End Property
