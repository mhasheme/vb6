VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmImportDb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Database"
   ClientHeight    =   5640
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3332.297
   ScaleMode       =   0  'User
   ScaleWidth      =   5422.412
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkImpFromSQL 
      Caption         =   "Export From SQL Servser Into Access"
      Height          =   285
      Left            =   240
      TabIndex        =   13
      Top             =   4200
      Visible         =   0   'False
      Width           =   4965
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   3
      Top             =   5085
      Width           =   5775
      _Version        =   65536
      _ExtentX        =   10186
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
      Left            =   1860
      TabIndex        =   0
      Tag             =   "Accept values and proceed"
      Top             =   4560
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
      Left            =   180
      TabIndex        =   1
      Tag             =   "Close and exit this screen"
      Top             =   4560
      Width           =   1425
   End
   Begin VB.CheckBox chkOracle 
      Caption         =   "Create Control file and Data file for Oracle user"
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   3930
      Width           =   4965
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Import From"
      Height          =   3705
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5535
      Begin VB.CheckBox chkDiffServer 
         Caption         =   "Different Server"
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.Frame frmImportDb 
         Caption         =   "Import From"
         Height          =   1515
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   5175
         Begin VB.OptionButton optOracle 
            Caption         =   "Oracle"
            Height          =   495
            Left            =   3600
            TabIndex        =   14
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton optAccess 
            Caption         =   "MS Access"
            Height          =   495
            Left            =   360
            TabIndex        =   11
            Top             =   600
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optSQL 
            Caption         =   "MS SQL Server"
            Height          =   495
            Left            =   1860
            TabIndex        =   10
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.CheckBox chkAPP 
         Caption         =   "Append data to the database"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   2130
         Visible         =   0   'False
         Width           =   2715
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Import From"
      Height          =   3705
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5535
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         Height          =   2790
         Left            =   360
         TabIndex        =   6
         Tag             =   "Select directory to import from"
         Top             =   300
         Width           =   3195
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   390
         TabIndex        =   5
         Tag             =   "Select drive"
         Top             =   3225
         Width           =   3225
      End
      Begin VB.FileListBox File1 
         Height          =   3210
         Left            =   3660
         Pattern         =   "ihr001.MDB;IHR001X.MDB;IHRWFC.MDB"
         TabIndex        =   4
         Top             =   315
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmImportDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkImpFromSQL_Click()
    Frame1.Visible = True
    Frame2.Visible = False
End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)  '19Aug99 js
End Sub


Private Sub cmdOK_Click()

If chkImpFromSQL Then
    Call ExportFromSQL2Access
    Exit Sub
End If

If Frame2.Visible Then
    If optAccess Then
        Frame1.Visible = True
        Frame2.Visible = False
    Else
        If optSQL.Value Then
            frmImportDbSQL.DatabaseType = "SQL Server"
            If chkDiffServer.Value Then 'Ticket #19762
                frmImportDbSQL.DiffSQLServer = "Y"
            Else
                frmImportDbSQL.DiffSQLServer = "N"
            End If
        End If
        If optOracle.Value Then
            frmImportDbSQL.DatabaseType = "Oracle"
        End If
        Unload Me
        frmImportDbSQL.Show vbModal
    End If
    Exit Sub
End If
Dim dbSource As New ADODB.Connection, dbDestination As New ADODB.Connection
Dim catSource As New ADOX.Catalog, catDestination As New ADOX.Catalog
Dim sSourcePath As String, sDestinationPath As String
Dim FldName, FldNameNew
Dim rstSource As New ADODB.Recordset, rstDestination As New ADODB.Recordset
Dim rstLoadDataWrk As New ADODB.Recordset
Dim tblNameList
Dim iCounter As Integer, iTableCounter As Integer, iFieldCounter As Integer, iRecordCounter As Long
Dim SQLQ, tblName, frtLetter
Dim lRecordCounter As Long
Dim asDatabaseName(1) As String        'Jaddy 6/4/99
Dim fldValue
Dim tblType
Dim TitleSTR
Dim DataSTR
Dim ctlSTR
Dim xReplaceStr
Dim OLDtblName, OLDfldNameNew
Dim LSTRtblName, LSTRfldNameNew
Dim strTemp, posTemp, MyString

'On Error GoTo LocalErrorHandler:

'get the directory that the use wants to import from...
sSourcePath = Dir1.Path
sSourcePath = sSourcePath & IIf(Right$(sSourcePath, 1) = "\", "", "\")
sDestinationPath = glbDBDir & IIf(Right$(glbDBDir, 1) = "\", "", "\")

If sSourcePath = sDestinationPath And Not glbSQL And Not glbOracle Then
    MsgBox "Unable to copy data, source and destination paths are the same."
    Exit Sub
Else
    If MsgBox("This function will convert data. Are you sure you want to continue?", vbQuestion + vbYesNo) = vbNo Then
        Unload Me
        Exit Sub
    Else
        Screen.MousePointer = vbHourglass
    End If
End If
SSPanel1.Caption = "Please wait"
asDatabaseName(0) = "IHR001.mdb"
asDatabaseName(1) = "IHR001x.mdb"
If chkOracle = 1 Then
    Open sSourcePath & "SQL_Load_Files.BAT" For Output As #3
    Open sSourcePath & "UPDATE_LOAD_DATA.SQL" For Output As #4
    Print #4, "SET DEFINE ?"
    Print #4, ""
End If
OLDtblName = ""
OLDfldNameNew = ""

LSTRtblName = ""
LSTRfldNameNew = ""

For iCounter = 0 To UBound(asDatabaseName) 'number of database consts above...
     If Dir$(sSourcePath & asDatabaseName(iCounter)) <> "" Then
        dbSource.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Jet OLEDB:Database Password=petman;Data Source=" & sSourcePath & asDatabaseName(iCounter)
        If iCounter = 0 Then
            Set dbDestination = gdbAdoIhr001
        Else
            Set dbDestination = gdbAdoIhr001X
        End If
        Set catSource.ActiveConnection = dbSource
        Set catDestination.ActiveConnection = dbDestination
        tblNameList = ""
        If glbOracle Then
            Dim rsTABLES As New ADODB.Recordset
            rsTABLES.Open "USER_TABLES", dbDestination, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
            Do Until rsTABLES.EOF
                tblName = rsTABLES("TABLE_NAME")
                If InStr("INFO_HR_TABLES,HRTABDES,", tblName & ",") = 0 Then
                    tblNameList = tblNameList & tblName & ","
                End If
                rsTABLES.MoveNext
            Loop
        Else
            For iTableCounter = 0 To catDestination.Tables.count - 1
                tblName = catDestination.Tables(iTableCounter).name
                If InStr("INFO_HR_TABLES,HRTABDES,", tblName & ",") = 0 Then
                    tblNameList = tblNameList & tblName & ","
                End If
            Next
        End If
        For iTableCounter = 0 To catSource.Tables.count - 1
            tblName = catSource.Tables(iTableCounter).name
            If catSource.Tables(iTableCounter).Type = "TABLE" Then
                If InStr(UCase(tblNameList), UCase(tblName) & ",") > 0 Then
                    rstSource.Open tblName, dbSource, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
                    SSPanel1.Caption = asDatabaseName(iCounter) & ": " & tblName
                    If chkOracle = 0 Then
                        If Not chkAPP Then
                            dbDestination.Execute "DELETE FROM " & tblName
                        End If
                        ' Mcn.CursorLocation = adUseClient
                        rstDestination.CursorLocation = adUseServer
                        rstDestination.Open tblName, dbDestination, adOpenKeyset, adLockOptimistic, adCmdTableDirect
                    End If
                    iRecordCounter = 0
                    TitleSTR = ""
                    
                    While Not rstSource.EOF
                        iRecordCounter = iRecordCounter + 1
                        SSPanel1.Caption = asDatabaseName(iCounter) & ": " & tblName & ":" & iRecordCounter & " Records Were Read."
                        If chkOracle = 1 Then
                            If iRecordCounter = 1 Then
                                Open sSourcePath & "CTL_" & UCase(tblName) & ".ctl" For Output As #1
                                Open sSourcePath & "DATA_" & UCase(tblName) & ".csv" For Output As #2
                                ctlSTR = "LOAD DATA" & Chr(13) & Chr(10)
                                ctlSTR = ctlSTR & "INFILE 'DATA_" & UCase(tblName) & ".CSV'" & Chr(13) & Chr(10)
                                ctlSTR = ctlSTR & "BADFILE 'BAD_" & UCase(tblName) & ".BAD'" & Chr(13) & Chr(10)
                                ctlSTR = ctlSTR & "REPLACE INTO TABLE " & tblName & Chr(13) & Chr(10)
                                ctlSTR = ctlSTR & "FIELDS TERMINATED BY ',' OPTIONALLY ENCLOSED BY '|'" & Chr(13) & Chr(10)
                                ctlSTR = ctlSTR & "( " & Chr(13) & Chr(10)
                            End If
                        Else
                            rstDestination.AddNew
                        End If
                        DataSTR = ""

                        For iFieldCounter = 0 To rstSource.Fields.count - 1
                           If (Not Right(rstSource.Fields(iFieldCounter).name, 3) = "_ID" _
                               And (Not rstSource.Fields(iFieldCounter).name = "ID" Or tblName = "HRNEWHIRE") _
                               And rstSource.Fields(iFieldCounter).name <> "PH_APPRAISAL_LINK" _
                               ) _
                               Or rstSource.Fields(iFieldCounter).name = "ED_PAYROLL_ID" _
                               Or rstSource.Fields(iFieldCounter).name = "AU_PAYROLL_ID" Then
                               
                                FldName = rstSource.Fields(iFieldCounter).name
                                FldNameNew = FldName
                                If (UCase(tblName) = "HREMP_FLAGS" Or UCase(tblName) = "TERM_HREMP_FLAGS") And FldName = "EF_FUREAS19" Then
                                    FldNameNew = "EF_FTREAS19"
                                End If
                                If tblName = "HR_FOLLOW_UP" And (FldName = "EF_ADMIN" Or FldName = "EF_ADMIN_TABL") Then
                                    FldNameNew = Left(FldName, 8) & "By" & Mid(FldName, 9)
                                End If
                                If tblName = "HR_EMAIL" And FldName = "EM_EMPNBR" Then
                                    FldNameNew = "EM_USERID"
                                End If
                                If tblName = "HRPASDEP" And FldName = "PD_EMPNBR" Then
                                    FldNameNew = "PD_USERID"
                                End If
                                If tblName = "HRPASDEP" And FldName = "PD_ADMINBBY_TABL" Then
                                    FldNameNew = "PD_ADMINBY_TABL"
                                End If
                                If (tblName = "HREEO" Or UCase(tblName) = "TERM_HREEO") And FldName = "EO_DISABLE-YN" Then
                                     FldNameNew = "EO_DISABLE_YN"
                                End If
                                
                                If UCase(FldName) = "LTIME" And UCase(tblName) = "HREARN" Then
                                    fldValue = Left(rstSource(FldName) & "", 8)
                                ElseIf UCase(FldName) = "CC_WCBNBR" Then
                                    fldValue = Left(Trim(rstSource(FldName)), 9)
                                ElseIf FldName = "ED_EXPYEAR" And rstSource(FldName) = 0 Then
                                    fldValue = Null
                                ElseIf rstSource(FldName).Type = 7 Then
                                    fldValue = rstSource(FldName)
                                    If IsDate(fldValue) Then
                                        If Year(fldValue) < 1900 Or Year(fldValue) > 2050 Then
                                            fldValue = CVDate(Format(fldValue, "mmm, dd ") & "19" & Format(fldValue, "yy"))
                                        End If
                                    End If
                                ElseIf tblName = "HR_SALARY_HISTORY" And FldName = "SH_LDATE" Then
                                    If Not IsDate(rstSource("SH_LDATE")) Then
                                        fldValue = rstSource("SH_EDATE")
                                    Else
                                        fldValue = rstSource("SH_LDATE")
                                    End If
                                Else
                                    fldValue = rstSource(FldName)
                                End If
                                
                                If chkOracle = 1 Then
                                    If iRecordCounter = 1 Then
                                        TitleSTR = TitleSTR & FldNameNew & ","
                                        ctlSTR = ctlSTR & FldNameNew & "," & Chr(13) & Chr(10)
                                    End If
                                    
                                    Select Case TypeName(fldValue)
                                    Case "Boolean"
                                        If fldValue Then fldValue = 1 Else fldValue = 0
                                    Case "String"
                                        If Len(fldValue) > 250 Then
                                            fldValue = Replace(fldValue, "'", "''")
                                            fldValue = Replace(fldValue, Chr(10), "~")
                                            fldValue = Replace(fldValue, Chr(13), "`")
                                            If Len(fldValue) > 2500 Then
                                                fldValue = Left(fldValue, 2500) & "' || '" & Mid(fldValue, 2501)
                                            End If
                                            xReplaceStr = "INFOHR_LOADING_REPLACE_LONG_" & iRecordCounter
                                            
                                            SQLQ = "UPDATE " & tblName & " SET " & FldNameNew & "='" & fldValue & "' WHERE " & FldNameNew & "='" & xReplaceStr & "';"
                                            Print #4, SQLQ
                                            Print #4, "COMMIT;"
                                            Print #4, ""
                                            LSTRtblName = tblName
                                            LSTRfldNameNew = FldNameNew
                                            fldValue = xReplaceStr
                                        End If
                                        If InStr(fldValue, Chr(10)) <> 0 Or InStr(fldValue, Chr(13)) <> 0 Then
                                            fldValue = Replace(fldValue, Chr(10), "~")
                                            fldValue = Replace(fldValue, Chr(13), "`")
                                            If (OLDtblName <> tblName Or OLDfldNameNew <> FldNameNew) And LSTRtblName = "" Then
                                                SQLQ = "UPDATE " & tblName & " SET " & FldNameNew & "=REPLACE(REPLACE(" & FldNameNew & ",'~',CHR(10)),'`',CHR(13));"
                                                Print #4, SQLQ
                                                Print #4, "COMMIT;"
                                                Print #4, ""
                                                OLDtblName = tblName
                                                OLDfldNameNew = FldNameNew
                                            End If
                                        End If
                                        If fldValue = "Report_Master_Education_Seminars" Then
                                            fldValue = "Report_Master_Edu_Seminars"
                                        End If
                                        fldValue = "|" & fldValue & "|"
                                    Case "Date"
                                        fldValue = Format(fldValue, "DD-MMM-YYYY")
                                    Case "Null"
                                        fldValue = "||"
                                    End Select
                                    DataSTR = DataSTR & fldValue & ","
                                Else
                                    If Not (tblName = "HRTABL" And FldNameNew = "TB_POINT") Then
                                        If FldNameNew = "CO_COMMENTS" Or FldNameNew = "CL_COMMENTS" Or FldNameNew = "PH_COMMENTS" Then
                                            fldValue = Left(fldValue, 4000)
                                        End If
                                        If tblName = "HRGLDIST" And FldNameNew = "GL_PERCENT" Then
                                            If Not IsNull(fldValue) Then
                                            rstDestination(FldNameNew) = fldValue
                                            End If
                                        Else
                                            rstDestination(FldNameNew) = fldValue
                                        End If
                                    End If
                                End If
                            End If
                            
                        Next
                        If chkOracle = 1 Then
                            If iRecordCounter = 1 Then
                                Print #1, CutRight1Letter(ctlSTR, 3) & ")"
                                Print #3, "SQLLDR " & SQLUserName & "/" & SQLUserPassword & "@%1 CONTROL=CTL_" & UCase(tblName) & ".ctl LOG=LOG_" & UCase(tblName) & ".log"
                                'Print #2, CutRight1Letter(TitleSTR)
                            End If
                            Print #2, CutRight1Letter(DataSTR)
                        Else
                            rstDestination.Update
                        End If
                        rstSource.MoveNext
                        DoEvents
                    Wend
                    If chkOracle = 1 Then
                        If LSTRtblName <> "" Then
                            SQLQ = "UPDATE " & LSTRtblName & " SET " & LSTRfldNameNew & "=REPLACE(REPLACE(" & LSTRfldNameNew & ",'~',CHR(10)),'`',CHR(13));"
                            Print #4, SQLQ
                            Print #4, "COMMIT;"
                            Print #4, ""
                            LSTRtblName = ""
                            LSTRfldNameNew = ""
                        End If
'
                        If iRecordCounter > 0 Then
                            Close #1
                            Close #2
                        End If
                    Else
                        rstDestination.Close
                    End If
                    rstSource.Close
                End If
            End If

        Next 'iTableCounter
        dbSource.Close
        dbDestination.Close
    Else
        MsgBox sSourcePath & asDatabaseName(iCounter) & " does not exist as a source file, skipping file."
    End If
Next
If chkOracle = 1 Then
    Close #3
    Close #4
End If
MsgBox "Converted successfully, Please click OK to close the application."
End
Screen.MousePointer = vbDefault
LocalErrorHandlerExit:
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

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Dir1_GotFocus()
    Call SetPanHelp(ActiveControl)  '19Aug99 js
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Drive1_GotFocus()
    Call SetPanHelp(ActiveControl)  '19Aug99 js
End Sub



Private Sub Form_Load()
glbOnTop = "FRMIMPORTDB"
Frame1.Visible = Not glbSQL
Frame2.Visible = glbSQL
If (glbSQL) Then
    chkImpFromSQL.Visible = True
End If

End Sub

Function CutRight1Letter(xStr, Optional xlen As Integer)
If TypeName(xStr) = "String" Then
    If Len(xStr) > 0 Then
        If xlen <> 0 Then
            CutRight1Letter = Left(xStr, Len(xStr) - xlen)
        Else
            CutRight1Letter = Left(xStr, Len(xStr) - 1)
        End If
    End If
End If
End Function

Private Sub ExportFromSQL2Access()
Dim dbSource As New ADODB.Connection, dbDestination As New ADODB.Connection
Dim catSource As New ADOX.Catalog ', catDestination As New ADOX.Catalog
Dim sSourcePath As String, sDestinationPath As String
Dim FldName, FldNameNew
Dim rstSource As New ADODB.Recordset, rstDestination As New ADODB.Recordset
Dim rstLoadDataWrk As New ADODB.Recordset
Dim tblNameList
Dim iCounter As Integer, iTableCounter As Integer, iFieldCounter As Integer, iRecordCounter As Long
Dim SQLQ, tblName, frtLetter
Dim lRecordCounter As Long
Dim asDatabaseName(1) As String        'Jaddy 6/4/99
Dim fldValue
Dim tblType
Dim TitleSTR
Dim DataSTR
Dim ctlSTR
Dim xReplaceStr
Dim OLDtblName, OLDfldNameNew
Dim LSTRtblName, LSTRfldNameNew
Dim strTemp, posTemp, MyString

'On Error GoTo LocalErrorHandler:

'get the directory that the use wants to import from...
sSourcePath = Dir1.Path
sSourcePath = sSourcePath & IIf(Right$(sSourcePath, 1) = "\", "", "\")
sDestinationPath = glbDBDir & IIf(Right$(glbDBDir, 1) = "\", "", "\")

If sSourcePath = sDestinationPath And Not glbSQL And Not glbOracle Then
    MsgBox "Unable to copy data, source and destination paths are the same."
    Exit Sub
Else
    If MsgBox("This function will convert data. Are you sure you want to continue?", vbQuestion + vbYesNo) = vbNo Then
        Unload Me
        Exit Sub
    Else
        Screen.MousePointer = vbHourglass
    End If
End If
SSPanel1.Caption = "Please wait"
asDatabaseName(0) = "IHR001.mdb"
asDatabaseName(1) = "IHR001x.mdb"

OLDtblName = ""
OLDfldNameNew = ""

LSTRtblName = ""
LSTRfldNameNew = ""

For iCounter = 0 To UBound(asDatabaseName) 'number of database consts above...
     If Dir$(sSourcePath & asDatabaseName(iCounter)) <> "" Then
        dbSource.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Jet OLEDB:Database Password=petman;Data Source=" & sSourcePath & asDatabaseName(iCounter)
        If iCounter = 0 Then
            Set dbDestination = gdbAdoIhr001
        Else
            Set dbDestination = gdbAdoIhr001X
        End If
        Set catSource.ActiveConnection = dbSource
        tblNameList = ""

        For iTableCounter = 0 To catSource.Tables.count - 1
            tblName = catSource.Tables(iTableCounter).name
            If InStr("INFO_HR_TABLES,HRTABDES,", tblName & ",") = 0 Then
                tblNameList = tblNameList & tblName & ","
            End If
        Next

        For iTableCounter = 0 To catSource.Tables.count - 1
            tblName = catSource.Tables(iTableCounter).name
            If catSource.Tables(iTableCounter).Type = "TABLE" Then
                If InStr(UCase(tblNameList), UCase(tblName) & ",") > 0 Then
                    rstSource.Open tblName, dbSource, adOpenDynamic, adLockOptimistic, adCmdTableDirect ',   'adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
                    SSPanel1.Caption = asDatabaseName(iCounter) & ": " & tblName
                    'If chkOracle = 0 Then
                        'If Not chkAPP Then
                            'dbDestination.Execute "DELETE FROM " & tblName
                            dbSource.Execute "DELETE FROM " & tblName
                        'End If
                        ' Mcn.CursorLocation = adUseClient
                        rstDestination.CursorLocation = adUseServer
                        rstDestination.Open tblName, dbDestination, adOpenKeyset, adLockOptimistic, adCmdTableDirect
                    'End If
                    iRecordCounter = 0
                    TitleSTR = ""
                    
                    'While Not rstSource.EOF
                    While Not rstDestination.EOF
                        iRecordCounter = iRecordCounter + 1
                        SSPanel1.Caption = asDatabaseName(iCounter) & ": " & tblName & ":" & iRecordCounter & " Records Were Read."

                        'rstDestination.AddNew
                        rstSource.AddNew

                        DataSTR = ""

                        'For iFieldCounter = 0 To rstSource.Fields.count - 1
                        For iFieldCounter = 0 To rstSource.Fields.count - 1
                           If (Not Right(rstSource.Fields(iFieldCounter).name, 3) = "_ID" _
                               And (Not rstSource.Fields(iFieldCounter).name = "ID" Or tblName = "HRNEWHIRE") _
                               And rstSource.Fields(iFieldCounter).name <> "PH_APPRAISAL_LINK" _
                               ) _
                               Or rstSource.Fields(iFieldCounter).name = "ED_PAYROLL_ID" _
                               Or rstSource.Fields(iFieldCounter).name = "AU_PAYROLL_ID" Then
                               
                                FldName = rstSource.Fields(iFieldCounter).name
                                FldNameNew = FldName
                                If (UCase(tblName) = "HREMP_FLAGS" Or UCase(tblName) = "TERM_HREMP_FLAGS") And FldName = "EF_FUREAS19" Then
                                    FldNameNew = "EF_FTREAS19"
                                End If
                                If tblName = "HR_FOLLOW_UP" And (FldName = "EF_ADMIN" Or FldName = "EF_ADMIN_TABL") Then
                                    FldNameNew = Left(FldName, 8) & "By" & Mid(FldName, 9)
                                End If
                                If tblName = "HR_EMAIL" And FldName = "EM_EMPNBR" Then
                                    FldNameNew = "EM_USERID"
                                End If
                                If tblName = "HRPASDEP" And FldName = "PD_EMPNBR" Then
                                    FldNameNew = "PD_USERID"
                                End If
                                If (tblName = "HREEO" Or UCase(tblName) = "TERM_HREEO") And FldName = "EO_DISABLE-YN" Then
                                     FldNameNew = "EO_DISABLE-YN"
                                     FldName = "EO_DISABLE_YN"
                                End If
                                If tblName = "HREMP_FLAGS" And FldName = "EF_FUREAS19" Then
                                    FldName = "EF_FTREAS19"
                                    FldNameNew = "EF_FUREAS19"
                                End If
                                
                                If UCase(FldName) = "LTIME" And UCase(tblName) = "HREARN" Then
                                    fldValue = Left(rstDestination(FldName) & "", 8)
                                ElseIf UCase(FldName) = "CC_WCBNBR" Then
                                    fldValue = Left(Trim(rstDestination(FldName)), 9)
                                ElseIf FldName = "ED_EXPYEAR" And rstDestination(FldName) = 0 Then
                                    fldValue = Null
                                ElseIf rstDestination(FldName).Type = 7 Then
                                    fldValue = rstDestination(FldName)
                                    If IsDate(fldValue) Then
                                        If Year(fldValue) < 1900 Or Year(fldValue) > 2050 Then
                                            fldValue = CVDate(Format(fldValue, "mmm, dd ") & "19" & Format(fldValue, "yy"))
                                        End If
                                    End If
                                ElseIf tblName = "HR_SALARY_HISTORY" And FldName = "SH_LDATE" Then
                                    If Not IsDate(rstDestination("SH_LDATE")) Then
                                        fldValue = rstDestination("SH_EDATE")
                                    Else
                                        fldValue = rstDestination("SH_LDATE")
                                    End If
                                Else
                                    fldValue = rstDestination(FldName)
                                End If
                                

                                If Not (tblName = "HRTABL" And FldNameNew = "TB_POINT") Then
                                    If FldNameNew = "CO_COMMENTS" Or FldNameNew = "CL_COMMENTS" Then
                                        fldValue = Left(fldValue, 4000)
                                    End If
                                    If tblName = "HRGLDIST" And FldNameNew = "GL_PERCENT" Then
                                        If Not IsNull(fldValue) Then
                                        rstSource(FldNameNew) = fldValue
                                        End If
                                    ElseIf (tblName = "HR_SECURE_COMMENTS" Or tblName = "HR_SECURE_FOLLOW_UP" Or tblName = "HR_SECURE_ATTENDANCE" Or tblName = "HR_SECURE_DOCUMENT_TYPE") And FldNameNew = "ACCESSABLE" Then    'Release 8.1
                                        If fldValue Then rstSource(FldNameNew) = 0 Else rstSource(FldNameNew) = 0
                                    ElseIf (tblName = "HR_SECURE_COMMENTS" Or tblName = "HR_SECURE_FOLLOW_UP" Or tblName = "HR_SECURE_ATTENDANCE" Or tblName = "HR_SECURE_DOCUMENT_TYPE") And FldNameNew = "MAINTAINABLE" Then  'Release 8.1
                                        If fldValue Then rstSource(FldNameNew) = 1 Else rstSource(FldNameNew) = 0
                                    Else
                                        If Not IsNull(fldValue) Then
                                            rstSource(FldNameNew) = fldValue
                                        End If
                                    End If
                                End If

                            End If
                            
                        Next

                        rstSource.Update

                        rstDestination.MoveNext
                        DoEvents
                    Wend
                    rstDestination.Close
                    rstSource.Close
                End If
            End If

        Next 'iTableCounter
        dbSource.Close
        dbDestination.Close
    Else
        MsgBox sSourcePath & asDatabaseName(iCounter) & " does not exist as a source file, skipping file."
    End If
Next
If chkOracle = 1 Then
    Close #3
    Close #4
End If
MsgBox "Converted successfully, Please click OK to close the application."
End
Screen.MousePointer = vbDefault
LocalErrorHandlerExit:
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

Private Sub optAccess_Click()
    Call setByDBType 'Ticket #19762
End Sub

Private Sub optOracle_Click()
    Call setByDBType 'Ticket #19762
End Sub

Private Sub optSQL_Click()
    Call setByDBType 'Ticket #19762
End Sub

Private Sub setByDBType() 'Ticket #19762
    If optSQL Then
        chkAPP.Visible = True
        chkDiffServer.Visible = True
    Else
        chkAPP.Visible = False
        chkDiffServer.Visible = False
    End If
End Sub
