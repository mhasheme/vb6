VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmSFDownload 
   Caption         =   "Download File from FTP"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   10035
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdLoadXML 
      Appearance      =   0  'Flat
      Caption         =   "Load XML File"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   5280
      Width           =   2685
   End
   Begin VB.CommandButton cmdDownToPath 
      Appearance      =   0  'Flat
      Caption         =   "Download to 'Copy to Path'"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   4800
      Width           =   2685
   End
   Begin VB.Frame fraXMLLocation 
      Height          =   3255
      Left            =   360
      TabIndex        =   11
      Top             =   720
      Width           =   9735
      Begin VB.TextBox txtPDFilename 
         Enabled         =   0   'False
         Height          =   288
         Left            =   5880
         MaxLength       =   200
         TabIndex        =   15
         Tag             =   "00-File Name"
         Top             =   480
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00FFFFFF&
         Height          =   2565
         Left            =   1920
         TabIndex        =   13
         Tag             =   "Path"
         Top             =   555
         Width           =   3372
      End
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Tag             =   "Disk Drive"
         Top             =   195
         Width           =   3372
      End
      Begin VB.Label lblPDFilename 
         BackStyle       =   0  'Transparent
         Caption         =   "XML File Location"
         Height          =   195
         Left            =   5880
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Label lblPath 
         BackStyle       =   0  'Transparent
         Caption         =   "Copy to Path"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test XML Download"
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   6960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      Caption         =   "Close"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   5760
      Width           =   2685
   End
   Begin VB.CommandButton cmdStart 
      Appearance      =   0  'Flat
      Caption         =   "Download"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   4320
      Width           =   2685
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   7725
      Width           =   10035
      _Version        =   65536
      _ExtentX        =   17701
      _ExtentY        =   741
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
      Begin VB.CommandButton cmdRecalPenMaster 
         Appearance      =   0  'Flat
         Caption         =   "Not Use"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   2445
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8340
         Top             =   180
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
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      Height          =   285
      Index           =   0
      Left            =   1965
      TabIndex        =   0
      Tag             =   "41-Effective date "
      Top             =   360
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1300
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   1
      Tag             =   "41-Effective date "
      Top             =   360
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1300
   End
   Begin VB.Label lblDet 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   3720
      TabIndex        =   9
      Top             =   360
      Width           =   195
   End
   Begin VB.Label lblDet 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Range:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   8
      Top             =   360
      Width           =   1275
   End
End
Attribute VB_Name = "frmSFDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Property Get ChangeAction() As UpdateStateEnum
ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
'RelateMode = Reports
'RelateMode = RelateEMP
RelateMode = MassChanges
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = False 'True  'False
End Property

Public Property Get Addable() As Boolean
Addable = False 'True  'False
End Property
Public Property Get Updateble() As Boolean
Updateble = True
End Property
Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = False 'True
End Property

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDownToPath_Click()
Dim I As Integer
Dim xFlag As Boolean
Dim xFilePath As String
Dim rsEmp As New ADODB.Recordset, rsJH As New ADODB.Recordset, rsSH As New ADODB.Recordset
Dim rsTABL As New ADODB.Recordset, rsJOB As New ADODB.Recordset, rsDEPT As New ADODB.Recordset
Dim rsPYMST As New ADODB.Recordset
Dim RecCNT As Long, buf As String
Dim ImportStatus As CSVImportStatus
Dim NotImpExists As Long
Dim FTPStatus As FtpErrorEnum
Dim xRemoteFile As String
Dim xLocalFile As String
Dim xDays As Integer
Dim xDATE
Dim xFileAmt As Integer
Dim xDefPath As String

On Error GoTo Err_Line 'Ticket #27476 Franks 08/31/2015

    xDefPath = Dir1.Path
    
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "From Date is required"
        dlpDate(0).SetFocus
        Exit Sub
    End If
    If Not IsDate(dlpDate(1).Text) Then
        MsgBox "To Date is required"
        dlpDate(1).SetFocus
        Exit Sub
    End If
    

    If Not ReadFTPData(True) Then '(clpPAYP.Text) Then
        MsgBox "Please complete the FTP Setup screen before attempting to use the HRSoft Interface.", vbInformation + vbOKOnly, "Setup Data Not Entered"
        End
    Else
        xFilePath = xDefPath & "\"
    End If

    '''MsgBox "step 1: export path = " & xFilePath

    xDays = DateDiff("D", CVDate(dlpDate(0).Text), CVDate(dlpDate(1).Text))
    xFileAmt = 0
    For I = 0 To xDays
    ' get the file from the PayWeb SFTP site
    'If Not TestMode Then
        xDATE = DateAdd("D", I, CVDate(dlpDate(0).Text))
        If glbPayWebData.UserName = "woodbridge-prd" Then 'Ticket #24829 Franks 12/20/2013
            xRemoteFile = "WB_HireFeed_" & Format(xDATE, "yyyymmdd") & ".xml"
        Else
            xRemoteFile = "WB_HireFeed_UAT_" & Format(xDATE, "yyyymmdd") & ".xml" '"franktest.csv"
        End If
        xLocalFile = xFilePath & xRemoteFile
        
        ''xRemoteFile = "WB_HireFeed_UAT_20130724.xml" '"franktest.csv"
        ''FTPStatus = modFTP.FTPGetFile(glbPayWebData.Host, glbPayWebData.Username, glbPayWebData.Password, xFilePath & "franktest.csv", "franktest.csv") ' "PYMST.CSV")
        'FTPStatus = modFTP.FTPGetFile(glbPayWebData.Host, glbPayWebData.Username, glbPayWebData.Password, xFilePath & xRemoteFile, "outgoing/" & xRemoteFile) ' "PYMST.CSV")
        FTPStatus = modFTP.FTPGetFile(glbPayWebData.Host, glbPayWebData.UserName, glbPayWebData.Password, xLocalFile, "outgoing/" & xRemoteFile) ' "PYMST.CSV")

        '''MsgBox "step 3: check if the file has been downloaded "

        Select Case FTPStatus
            Case FtpErrorEnum.PSCPNotFound
                MsgBox "PSCP runtime utility not found.  Please call info:HR support for assistance.", vbInformation + vbOKOnly, "PSCP.EXE Not Found"
                Exit Sub
            Case FtpErrorEnum.PSCPTimedOut
                If MsgBox("SCP session timed out.  View log?", vbCritical + vbYesNo, "Timed Out") = vbYes Then
                    Shell "notepad " & glbIHRREPORTS & "SCPLOG.TXT", vbNormalFocus
                End If
                Exit Sub
            Case FtpErrorEnum.PSCPError
                '''MsgBox "step 4: PSCPError " & FtpErrorEnum.PSCPError
                
                'if we need to see the error message about "This computer must be initialized on "
                'then uncomment the follow 3 lines
                'If MsgBox("PSCP returned an error.  View log?", vbCritical + vbYesNo, "SCP Error") = vbYes Then
                '    Shell "notepad " & glbIHRREPORTS & "SCPLOG.TXT", vbNormalFocus
                'End If
                
                ''Ticket #25150 Franks the following 6 lines caused error "Permission denied" at Woodbridge
                ''If Not IsPSCP_Initialized Then  'Ticket #25088 Franks 02/18/2014
                ''    If MsgBox("PSCP returned an error.  View log?", vbCritical + vbYesNo, "SCP Error") = vbYes Then
                ''        Shell "notepad " & glbIHRREPORTS & "SCPLOG.TXT", vbNormalFocus
                ''    End If
                ''    Exit Sub
                ''End If
                
                GoTo next_day
        End Select
        '''''MsgBox "step 5: write download log to 'HRSF_DOWNLOAD_LOG' table "
        '''write download log
        ''Call UptDownloadLog(xRemoteFile, xLocalFile, xDATE)
        '''''MsgBox "step 6: update 'HRSF_XML_IMPORT' table  "
        ''Call UptXMLImportTable(xRemoteFile, xLocalFile, xDATE)
        '''''MsgBox "step 7:  success!"
        
        xFileAmt = xFileAmt + 1
next_day:
    'End If
    Next I
    Screen.MousePointer = vbHourglass
    Screen.MousePointer = vbDefault 'xFileAmt
    buf = ImportStatus.ImportedOK & " records imported successfully" & vbCrLf
    If xFileAmt = 0 Then
        MsgBox "No file found. "
    ElseIf xFileAmt = 1 Then
        MsgBox "One file downloaded successfully. "
    Else
        MsgBox xFileAmt & " files downloaded successfully."
    End If

'Ticket #27476 Franks 08/31/2015
    Exit Sub
Err_Line:
    'MsgBox Err.Description
    
End Sub

Private Sub cmdLoadXML_Click()
Dim I As Integer
Dim xFlag As Boolean
Dim xFilePath As String
Dim rsEmp As New ADODB.Recordset, rsJH As New ADODB.Recordset, rsSH As New ADODB.Recordset
Dim rsTABL As New ADODB.Recordset, rsJOB As New ADODB.Recordset, rsDEPT As New ADODB.Recordset
Dim rsPYMST As New ADODB.Recordset
Dim RecCNT As Long, buf As String
Dim ImportStatus As CSVImportStatus
Dim NotImpExists As Long
Dim FTPStatus As FtpErrorEnum
Dim xRemoteFile As String
Dim xLocalFile As String
Dim xDays As Integer
Dim xDATE
Dim xFileAmt As Integer
Dim xDefPath As String

    xDefPath = Dir1.Path
    
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "From Date is required"
        dlpDate(0).SetFocus
        Exit Sub
    End If
    If Not IsDate(dlpDate(1).Text) Then
        MsgBox "To Date is required"
        dlpDate(1).SetFocus
        Exit Sub
    End If
    

    If Not ReadFTPData(True) Then '(clpPAYP.Text) Then
        MsgBox "Please complete the FTP Setup screen before attempting to use the HRSoft Interface.", vbInformation + vbOKOnly, "Setup Data Not Entered"
        End
    Else
        xFilePath = xDefPath & "\"
    End If
    
    '''MsgBox "step 1: export path = " & xFilePath

    xDays = DateDiff("D", CVDate(dlpDate(0).Text), CVDate(dlpDate(1).Text))
    xFileAmt = 0
    For I = 0 To xDays
    ' get the file from the PayWeb SFTP site
    'If Not TestMode Then
        xDATE = DateAdd("D", I, CVDate(dlpDate(0).Text))
        If glbPayWebData.UserName = "woodbridge-prd" Then 'Ticket #24829 Franks 12/20/2013
            xRemoteFile = "WB_HireFeed_" & Format(xDATE, "yyyymmdd") & ".xml"
        Else
            xRemoteFile = "WB_HireFeed_UAT_" & Format(xDATE, "yyyymmdd") & ".xml" '"franktest.csv"
        End If
        xFilePath = xDefPath & "\"
        xLocalFile = xFilePath & xRemoteFile
        
        ''''xRemoteFile = "WB_HireFeed_UAT_20130724.xml" '"franktest.csv"
        ''''FTPStatus = modFTP.FTPGetFile(glbPayWebData.Host, glbPayWebData.Username, glbPayWebData.Password, xFilePath & "franktest.csv", "franktest.csv") ' "PYMST.CSV")
        '''FTPStatus = modFTP.FTPGetFile(glbPayWebData.Host, glbPayWebData.Username, glbPayWebData.Password, xFilePath & xRemoteFile, "outgoing/" & xRemoteFile) ' "PYMST.CSV")
        ''FTPStatus = modFTP.FTPGetFile(glbPayWebData.Host, glbPayWebData.UserName, glbPayWebData.Password, xLocalFile, "outgoing/" & xRemoteFile) ' "PYMST.CSV")
        ''
        '''''MsgBox "step 3: check if the file has been downloaded "
        ''
        ''Select Case FTPStatus
        ''    Case FtpErrorEnum.PSCPNotFound
        ''        MsgBox "PSCP runtime utility not found.  Please call INFO:HR support for assistance.", vbInformation + vbOKOnly, "PSCP.EXE Not Found"
        ''        Exit Sub
        ''    Case FtpErrorEnum.PSCPTimedOut
        ''        If MsgBox("SCP session timed out.  View log?", vbCritical + vbYesNo, "Timed Out") = vbYes Then
        ''            Shell "notepad " & glbIHRREPORTS & "SCPLOG.TXT", vbNormalFocus
        ''        End If
        ''        Exit Sub
        ''    Case FtpErrorEnum.PSCPError
        ''
        ''
        ''        GoTo next_day
        ''End Select
        '''MsgBox "step 5: write download log to 'HRSF_DOWNLOAD_LOG' table "
        'write download log
        
        If Dir(xLocalFile) = "" Then
            'MsgBox "FILE not Found :" & Chr(10) & "[" & xRemoteFile & "]"
        Else
            Call UptDownloadLog(xRemoteFile, xLocalFile, xDATE)
            '''MsgBox "step 6: update 'HRSF_XML_IMPORT' table  "
            Call UptXMLImportTable(xRemoteFile, xLocalFile, xDATE)
            '''MsgBox "step 7:  success!"
            
            xFileAmt = xFileAmt + 1
        End If
next_day:
    'End If
    Next I
    Screen.MousePointer = vbHourglass
    Screen.MousePointer = vbDefault 'xFileAmt
    buf = ImportStatus.ImportedOK & " records imported successfully" & vbCrLf
    If xFileAmt = 0 Then
        MsgBox "No file found. "
    ElseIf xFileAmt = 1 Then
        MsgBox "One file imported successfully. "
    Else
        MsgBox xFileAmt & " files imported successfully."
    End If

End Sub

Private Sub cmdStart_Click()
Dim I As Integer
Dim xFlag As Boolean
Dim xFilePath As String
Dim rsEmp As New ADODB.Recordset, rsJH As New ADODB.Recordset, rsSH As New ADODB.Recordset
Dim rsTABL As New ADODB.Recordset, rsJOB As New ADODB.Recordset, rsDEPT As New ADODB.Recordset
Dim rsPYMST As New ADODB.Recordset
Dim RecCNT As Long, buf As String
Dim ImportStatus As CSVImportStatus
Dim NotImpExists As Long
Dim FTPStatus As FtpErrorEnum
Dim xRemoteFile As String
Dim xLocalFile As String
Dim xDays As Integer
Dim xDATE
Dim xFileAmt As Integer
Dim xDefPath As String

On Error GoTo Err_Line 'Ticket #27476 Franks 08/31/2015

    If glbCompSerial = "S/N - 2379W" Then  'Ticket #26912 Franks 06/22/2015 Town of LaSalle
        If glbFrmCaption$ = "Download File from FTP" Then Call TownofLaSalleDownload
        If glbFrmCaption$ = "Upload File To FTP" Then Call TownofLaSalleUpload
        Exit Sub
    End If
    
    '''Debug.Print FileDateTime("C:\A\IHR.exe")
    ''If App.Path = "C:\SSWORK\IHR80" Then
    ''    xDefPath = "C:\HR\HRSOFT\NewHireInterfaceINFOHR"
    ''Else
    ''    'xDefPath = "H:\HR\HRSOFT\NewHireInterfaceINFOHR"
    ''    'Ticket #25599 Franks 06/05/2014 - use infohr working folder
    ''    xDefPath = glbIHRREPORTS
    ''End If
    
    '''Ticket #25604 Franks 06/18/2014
    ''xDefPath = getXMLFileLocation
    ''If Len(Dir$(xDefPath, vbDirectory)) = 0 Then
    ''    MsgBox "Invalid XML File Location Path: " & Chr(10) & "  " & xDefPath & Chr(10) & Chr(10) & "Please go to HRSoft/Setup/XML File Location to setup a valid location."
    ''    Exit Sub
    ''End If
    'Ticket #25927 Frank 08/25/2014 - let user to change the Copy to Path if they need
    xDefPath = Dir1.Path
    
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "From Date is required"
        dlpDate(0).SetFocus
        Exit Sub
    End If
    If Not IsDate(dlpDate(1).Text) Then
        MsgBox "To Date is required"
        dlpDate(1).SetFocus
        Exit Sub
    End If
    

    If Not ReadFTPData(True) Then '(clpPAYP.Text) Then
        MsgBox "Please complete the FTP Setup screen before attempting to use the HRSoft Interface.", vbInformation + vbOKOnly, "Setup Data Not Entered"
        End
    Else
        '''xFilePath = glbSystemData.Path & clpPAYP.Text & "\"
        '''If Dir(xFilePath, vbDirectory) = "" Then
        '''   MkDir xFilePath
        '''End If
        ''If Dir$(xDefPath, vbDirectory) = "NewHireInterfaceINFOHR" Then
        ''    'MsgBox "There is no " & glbIHRREPORTS & "T10 Form folder."
        ''    'Exit Function
        ''    xFilePath = xDefPath & "\"
        ''Else
        ''    xFilePath = glbIHRREPORTS
        ''End If
        
        xFilePath = xDefPath & "\"
    End If

    '''MsgBox "step 1: export path = " & xFilePath

    xDays = DateDiff("D", CVDate(dlpDate(0).Text), CVDate(dlpDate(1).Text))
    xFileAmt = 0
    For I = 0 To xDays
    ' get the file from the PayWeb SFTP site
    'If Not TestMode Then
        xDATE = DateAdd("D", I, CVDate(dlpDate(0).Text))
        If glbPayWebData.UserName = "woodbridge-prd" Then 'Ticket #24829 Franks 12/20/2013
            xRemoteFile = "WB_HireFeed_" & Format(xDATE, "yyyymmdd") & ".xml"
        Else
            xRemoteFile = "WB_HireFeed_UAT_" & Format(xDATE, "yyyymmdd") & ".xml" '"franktest.csv"
        End If
        xLocalFile = xFilePath & xRemoteFile
        
        '''MsgBox "step 2: export file = " & xLocalFile

        ''xRemoteFile = "WB_HireFeed_UAT_20130724.xml" '"franktest.csv"
        ''FTPStatus = modFTP.FTPGetFile(glbPayWebData.Host, glbPayWebData.Username, glbPayWebData.Password, xFilePath & "franktest.csv", "franktest.csv") ' "PYMST.CSV")
        'FTPStatus = modFTP.FTPGetFile(glbPayWebData.Host, glbPayWebData.Username, glbPayWebData.Password, xFilePath & xRemoteFile, "outgoing/" & xRemoteFile) ' "PYMST.CSV")
        FTPStatus = modFTP.FTPGetFile(glbPayWebData.Host, glbPayWebData.UserName, glbPayWebData.Password, xLocalFile, "outgoing/" & xRemoteFile) ' "PYMST.CSV")

        '''MsgBox "step 3: check if the file has been downloaded "

        Select Case FTPStatus
            Case FtpErrorEnum.PSCPNotFound
                MsgBox "PSCP runtime utility not found.  Please call info:HR support for assistance.", vbInformation + vbOKOnly, "PSCP.EXE Not Found"
                Exit Sub
            Case FtpErrorEnum.PSCPTimedOut
                If MsgBox("SCP session timed out.  View log?", vbCritical + vbYesNo, "Timed Out") = vbYes Then
                    Shell "notepad " & glbIHRREPORTS & "SCPLOG.TXT", vbNormalFocus
                End If
                Exit Sub
            Case FtpErrorEnum.PSCPError
                '''MsgBox "step 4: PSCPError " & FtpErrorEnum.PSCPError
                
                'if we need to see the error message about "This computer must be initialized on "
                'then uncomment the follow 3 lines
                'If MsgBox("PSCP returned an error.  View log?", vbCritical + vbYesNo, "SCP Error") = vbYes Then
                '    Shell "notepad " & glbIHRREPORTS & "SCPLOG.TXT", vbNormalFocus
                'End If
                
                ''Ticket #25150 Franks the following 6 lines caused error "Permission denied" at Woodbridge
                ''If Not IsPSCP_Initialized Then  'Ticket #25088 Franks 02/18/2014
                ''    If MsgBox("PSCP returned an error.  View log?", vbCritical + vbYesNo, "SCP Error") = vbYes Then
                ''        Shell "notepad " & glbIHRREPORTS & "SCPLOG.TXT", vbNormalFocus
                ''    End If
                ''    Exit Sub
                ''End If
                
                GoTo next_day
        End Select
        '''MsgBox "step 5: write download log to 'HRSF_DOWNLOAD_LOG' table "
        'write download log
        Call UptDownloadLog(xRemoteFile, xLocalFile, xDATE)
        '''MsgBox "step 6: update 'HRSF_XML_IMPORT' table  "
        Call UptXMLImportTable(xRemoteFile, xLocalFile, xDATE)
        '''MsgBox "step 7:  success!"
        
        xFileAmt = xFileAmt + 1
next_day:
    'End If
    Next I
    Screen.MousePointer = vbHourglass
'    panProgress.Visible = True
'    gdbAdoIhr001.Execute "DELETE FROM PAYWEB_PYMST"
'    rsPYMST.Open "SELECT * FROM PAYWEB_PYMST", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    rsEMP.Open "SELECT * FROM HREMP", glbDBConnections("IHR001"), adOpenKeyset, adLockOptimistic
'    rsJH.Open "SELECT * FROM HR_JOB_HISTORY", glbDBConnections("IHR001"), adOpenKeyset, adLockOptimistic
'    rsSH.Open "SELECT * FROM HR_SALARY_HISTORY", glbDBConnections("IHR001"), adOpenKeyset, adLockOptimistic
'    rsTABL.Open "SELECT * FROM HRTABL", glbDBConnections("IHR001"), adOpenKeyset, adLockOptimistic
'    rsJOB.Open "SELECT * FROM HRJOB", glbDBConnections("IHR001"), adOpenKeyset, adLockOptimistic
'    rsDEPT.Open "SELECT * FROM HRDEPT", glbDBConnections("IHR001"), adOpenKeyset, adLockOptimistic
'    panProgress.FloodType = 0
'    ImportStatus = ImportCSV(rsPYMST, xFilePath & "PYMST.CSV", panProgress, 1, True)
'    panProgress.FloodType = 1
'    rsPYMST.Requery
'    ' create the standard table codes
'    CreateTABL rsTABL, "SDRC", "IMP", "IMPORT FROM PAYROLL"
'    CreateJOB rsJOB, "IMPORT", "IMPORT FROM PAYROLL"
'    Do Until rsPYMST.EOF
'        rsEMP.Filter = "ED_EMPNBR=" & rsPYMST("EMPNBR")
'        If Not rsEMP.EOF Then
'            NotImpExists = NotImpExists + 1
'            GoTo SkipEmployee
'        End If
'        ProcessOneRecord rsPYMST, rsJH, rsSH, rsEMP, rsDEPT
'SkipEmployee:
'        panProgress.FloodPercent = rsPYMST.AbsolutePosition / rsPYMST.RecordCount * 100
'        rsPYMST.MoveNext
'        DoEvents
'    Loop
'    rsPYMST.Close
'    rsEMP.Close
'    rsJH.Close
'    rsSH.Close
'    rsTABL.Close
'    rsJOB.Close
'    rsDEPT.Close
    Screen.MousePointer = vbDefault 'xFileAmt
    buf = ImportStatus.ImportedOK & " records imported successfully" & vbCrLf
    'If ImportStatus.RecLengthErrors > 0 Then Buf = Buf & ImportStatus.RecLengthErrors & " records skipped due to record length errors"
    'If NotImpExists > 0 Then Buf = Buf & NotImpExists & " records skipped due to existing employee with same number"
    'MsgBox Buf, vbInformation + vbOKOnly, "Import Complete"
    If xFileAmt = 0 Then
        MsgBox "No file found. "
    ElseIf xFileAmt = 1 Then
        MsgBox "One file downloaded successfully. "
    Else
        MsgBox xFileAmt & " files downloaded successfully."
    End If

'Ticket #27476 Franks 08/31/2015
    Exit Sub
Err_Line:
    'MsgBox Err.Description
    
End Sub

Private Sub Command1_Click()
Dim I As Integer
Dim xFlag As Boolean
Dim xFilePath As String
Dim rsEmp As New ADODB.Recordset, rsJH As New ADODB.Recordset, rsSH As New ADODB.Recordset
Dim rsTABL As New ADODB.Recordset, rsJOB As New ADODB.Recordset, rsDEPT As New ADODB.Recordset
Dim rsPYMST As New ADODB.Recordset
Dim RecCNT As Long, buf As String
Dim ImportStatus As CSVImportStatus
Dim NotImpExists As Long
Dim FTPStatus As FtpErrorEnum
Dim xRemoteFile As String
Dim xLocalFile As String
Dim xDays As Integer
Dim xDATE
Dim xFileAmt As Integer
Dim xDefPath As String

    '''Debug.Print FileDateTime("C:\A\IHR.exe")
    ''If App.Path = "C:\SSWORK\IHR80" Then
    ''    xDefPath = "C:\HR\HRSOFT\NewHireInterfaceINFOHR"
    ''Else
    ''    'xDefPath = "H:\HR\HRSOFT\NewHireInterfaceINFOHR"
    ''    'Ticket #25599 Franks 06/05/2014 - use infohr working folder
    ''    xDefPath = glbIHRREPORTS
    ''End If
    
    'Ticket #25604 Franks 06/18/2014
    xDefPath = getXMLFileLocation
    If Len(Dir$(xDefPath, vbDirectory)) = 0 Then
        MsgBox "Invalid XML File Location Path: " & Chr(10) & "  " & xDefPath & Chr(10) & Chr(10) & "Please go to HRSoft/Setup/XML File Location to setup a valid location."
        Exit Sub
    End If
    
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "From Date is required"
        dlpDate(0).SetFocus
        Exit Sub
    End If
    If Not IsDate(dlpDate(1).Text) Then
        MsgBox "To Date is required"
        dlpDate(1).SetFocus
        Exit Sub
    End If
    

    If Not ReadFTPData(True) Then '(clpPAYP.Text) Then
        MsgBox "Please complete the FTP Setup screen before attempting to use the HRSoft Interface.", vbInformation + vbOKOnly, "Setup Data Not Entered"
        End
    Else
        '''xFilePath = glbSystemData.Path & clpPAYP.Text & "\"
        '''If Dir(xFilePath, vbDirectory) = "" Then
        '''   MkDir xFilePath
        '''End If
        ''If Dir$(xDefPath, vbDirectory) = "NewHireInterfaceINFOHR" Then
        ''    'MsgBox "There is no " & glbIHRREPORTS & "T10 Form folder."
        ''    'Exit Function
        ''    xFilePath = xDefPath & "\"
        ''Else
        ''    xFilePath = glbIHRREPORTS
        ''End If
        
        xFilePath = xDefPath & "\"
    End If

    '''MsgBox "step 1: export path = " & xFilePath

    xDays = DateDiff("D", CVDate(dlpDate(0).Text), CVDate(dlpDate(1).Text))
    xFileAmt = 0
    For I = 0 To xDays
    ' get the file from the PayWeb SFTP site
    'If Not TestMode Then
        xDATE = DateAdd("D", I, CVDate(dlpDate(0).Text))
        If glbPayWebData.UserName = "woodbridge-prd" Then 'Ticket #24829 Franks 12/20/2013
            xRemoteFile = "WB_HireFeed_" & Format(xDATE, "yyyymmdd") & ".xml"
        Else
            xRemoteFile = "WB_HireFeed_UAT_" & Format(xDATE, "yyyymmdd") & ".xml" '"franktest.csv"
        End If
        xLocalFile = xFilePath & xRemoteFile
        
        '''MsgBox "step 2: export file = " & xLocalFile

        ''xRemoteFile = "WB_HireFeed_UAT_20130724.xml" '"franktest.csv"
        ''FTPStatus = modFTP.FTPGetFile(glbPayWebData.Host, glbPayWebData.Username, glbPayWebData.Password, xFilePath & "franktest.csv", "franktest.csv") ' "PYMST.CSV")
        'FTPStatus = modFTP.FTPGetFile(glbPayWebData.Host, glbPayWebData.Username, glbPayWebData.Password, xFilePath & xRemoteFile, "outgoing/" & xRemoteFile) ' "PYMST.CSV")
        FTPStatus = modFTP.FTPGetFile(glbPayWebData.Host, glbPayWebData.UserName, glbPayWebData.Password, xLocalFile, "outgoing/" & xRemoteFile) ' "PYMST.CSV")

        '''MsgBox "step 3: check if the file has been downloaded "

        Select Case FTPStatus
            Case FtpErrorEnum.PSCPNotFound
                MsgBox "PSCP runtime utility not found.  Please call info:HR support for assistance.", vbInformation + vbOKOnly, "PSCP.EXE Not Found"
                Exit Sub
            Case FtpErrorEnum.PSCPTimedOut
                If MsgBox("SCP session timed out.  View log?", vbCritical + vbYesNo, "Timed Out") = vbYes Then
                    Shell "notepad " & glbIHRREPORTS & "SCPLOG.TXT", vbNormalFocus
                End If
                Exit Sub
            Case FtpErrorEnum.PSCPError
                '''MsgBox "step 4: PSCPError " & FtpErrorEnum.PSCPError
                
                'if we need to see the error message about "This computer must be initialized on "
                'then uncomment the follow 3 lines
                If MsgBox("PSCP returned an error.  View log?", vbCritical + vbYesNo, "SCP Error") = vbYes Then
                    Shell "notepad " & glbIHRREPORTS & "SCPLOG.TXT", vbNormalFocus
                End If
                Exit Sub
                
                ''Ticket #25150 Franks the following 6 lines caused error "Permission denied" at Woodbridge
                ''If Not IsPSCP_Initialized Then  'Ticket #25088 Franks 02/18/2014
                ''    If MsgBox("PSCP returned an error.  View log?", vbCritical + vbYesNo, "SCP Error") = vbYes Then
                ''        Shell "notepad " & glbIHRREPORTS & "SCPLOG.TXT", vbNormalFocus
                ''    End If
                ''    Exit Sub
                ''End If
                
                GoTo next_day
        End Select
        '''MsgBox "step 5: write download log to 'HRSF_DOWNLOAD_LOG' table "
        'write download log
        Call UptDownloadLog(xRemoteFile, xLocalFile, xDATE)
        '''MsgBox "step 6: update 'HRSF_XML_IMPORT' table  "
        Call UptXMLImportTable(xRemoteFile, xLocalFile, xDATE)
        '''MsgBox "step 7:  success!"
        
        xFileAmt = xFileAmt + 1
next_day:
    'End If
    Next I
    Screen.MousePointer = vbHourglass
'    panProgress.Visible = True
'    gdbAdoIhr001.Execute "DELETE FROM PAYWEB_PYMST"
'    rsPYMST.Open "SELECT * FROM PAYWEB_PYMST", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    rsEMP.Open "SELECT * FROM HREMP", glbDBConnections("IHR001"), adOpenKeyset, adLockOptimistic
'    rsJH.Open "SELECT * FROM HR_JOB_HISTORY", glbDBConnections("IHR001"), adOpenKeyset, adLockOptimistic
'    rsSH.Open "SELECT * FROM HR_SALARY_HISTORY", glbDBConnections("IHR001"), adOpenKeyset, adLockOptimistic
'    rsTABL.Open "SELECT * FROM HRTABL", glbDBConnections("IHR001"), adOpenKeyset, adLockOptimistic
'    rsJOB.Open "SELECT * FROM HRJOB", glbDBConnections("IHR001"), adOpenKeyset, adLockOptimistic
'    rsDEPT.Open "SELECT * FROM HRDEPT", glbDBConnections("IHR001"), adOpenKeyset, adLockOptimistic
'    panProgress.FloodType = 0
'    ImportStatus = ImportCSV(rsPYMST, xFilePath & "PYMST.CSV", panProgress, 1, True)
'    panProgress.FloodType = 1
'    rsPYMST.Requery
'    ' create the standard table codes
'    CreateTABL rsTABL, "SDRC", "IMP", "IMPORT FROM PAYROLL"
'    CreateJOB rsJOB, "IMPORT", "IMPORT FROM PAYROLL"
'    Do Until rsPYMST.EOF
'        rsEMP.Filter = "ED_EMPNBR=" & rsPYMST("EMPNBR")
'        If Not rsEMP.EOF Then
'            NotImpExists = NotImpExists + 1
'            GoTo SkipEmployee
'        End If
'        ProcessOneRecord rsPYMST, rsJH, rsSH, rsEMP, rsDEPT
'SkipEmployee:
'        panProgress.FloodPercent = rsPYMST.AbsolutePosition / rsPYMST.RecordCount * 100
'        rsPYMST.MoveNext
'        DoEvents
'    Loop
'    rsPYMST.Close
'    rsEMP.Close
'    rsJH.Close
'    rsSH.Close
'    rsTABL.Close
'    rsJOB.Close
'    rsDEPT.Close
    Screen.MousePointer = vbDefault 'xFileAmt
    buf = ImportStatus.ImportedOK & " records imported successfully" & vbCrLf
    'If ImportStatus.RecLengthErrors > 0 Then Buf = Buf & ImportStatus.RecLengthErrors & " records skipped due to record length errors"
    'If NotImpExists > 0 Then Buf = Buf & NotImpExists & " records skipped due to existing employee with same number"
    'MsgBox Buf, vbInformation + vbOKOnly, "Import Complete"
    If xFileAmt = 0 Then
        MsgBox "No file found. "
    ElseIf xFileAmt = 1 Then
        MsgBox "One file downloaded successfully. "
    Else
        MsgBox xFileAmt & " files downloaded successfully."
    End If


'Call UptXMLImportTable("WB_HireFeed_UAT_20130812.xml", "C:\A\WB_HireFeed_UAT_20130812.xml", CVDate("08/12/2013"))

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Activate()
glbOnTop = "frmSFDownload"
Call set_Buttons
End Sub

Private Sub Form_Load()
    glbOnTop = "frmSFDownload"
    
    If glbCompSerial = "S/N - 2379W" Then  'Ticket #26912 Franks 06/22/2015 Town of LaSalle
        Me.Caption = glbFrmCaption$
        If glbFrmCaption$ = "Download File from FTP" Then Call TownofLaSalleScreenDownload
        If glbFrmCaption$ = "Upload File To FTP" Then Call TownofLaSalleScreenUpload
    Else 'WFC
        'dlpDate(0) = CVDate(Date) '("Jan 1, 2009")
        'dlpDate(1) = CVDate(Date) '("Dec 31, 2009")
        
        'If App.Path = "C:\SSWORK\HRSoftInt" Then
        If glbUserID = "3142" Or glbUserID = "999999999" Then
            Command1.Visible = True
        End If
        
        fraXMLLocation.BorderStyle = 0
        Call XMLFileLocationScreen("DISP")
        
        'Ticket #27515 Franks 09/08/2015
        Call WFCDateRange
    End If
        
End Sub

Private Sub UptDownloadLog(xFileName, xLocFile, xFileDate)
Dim rsLog As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM HRSF_DOWNLOAD_LOG WHERE SF_FILENAME = '" & xFileName & "' "
    rsLog.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsLog.EOF Then
        rsLog.AddNew
        rsLog("SF_FILENAME") = xFileName
        rsLog("SF_PROCESSED") = 0
    End If
    rsLog("SF_FILEDATE") = xFileDate ' FileDateTime(xLocFile)
    rsLog("SF_LDATE") = Date
    rsLog("SF_LTIME") = Time$
    rsLog("SF_LUSER") = glbUserID
    rsLog.Update
    rsLog.Close
End Sub

Private Sub UptXMLImportTable(xFileName, xLocFile, xFileDate)
Dim nodelist As MSXML2.IXMLDOMNodeList
Dim nodelistCh1 As MSXML2.IXMLDOMNodeList
Dim nodelistCh2 As MSXML2.IXMLDOMNodeList
Dim nodelistCh3 As MSXML2.IXMLDOMNodeList
'Dim doc As New MSXML2.DOMDocument
Dim doc As New MSXML2.DOMDocument60 'Ticket #28670 Franks 06/02/2016
Dim success As Boolean
Dim Node As MSXML2.IXMLDOMNode
Dim NodeChild1 As MSXML2.IXMLDOMNode
Dim NodeChild2 As MSXML2.IXMLDOMNode
Dim NodeChild3 As MSXML2.IXMLDOMNode
Dim xVal
Dim rsXMLTbl As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim xJobNo, xCandID, xTemp
Dim xJobCountry, xCategory, xEEO, xUnion, xLoc, xHireMgr, xBusi, xPlant, xDiv, xPosCode
Dim I As Integer, L As Integer

'set
'success = doc.Load("C:\A\WB_HireFeed_UAT_20130812.xml")
success = doc.Load(xLocFile)

If success Then
    'Set nodelist = doc.selectNodes("/feed/job/jobdata")
    '------------ hire ---------- begin
    ''GoTo unhires
    Set nodelist = doc.SelectNodes("/feed/hires/job")
    If Not (nodelist Is Nothing) Then
        For Each Node In nodelist 'hires/job
            'Debug.Print Node.SelectSingleNode("@number").Text
            xJobNo = Node.SelectSingleNode("@number").Text
            'name = Node.Text
            If Node.HasChildNodes Then
                'Debug.Print displaynode; Node.childNodes
                Set nodelistCh1 = Node.ChildNodes
                For Each NodeChild1 In nodelistCh1
                    ''Debug.Print NodeChild.nodeName & " - " & NodeChild.nodeValue
                    'Debug.Print NodeChild.Text
                    'If Not NodeChild.nodeName = "language" Then
                    If NodeChild1.nodeName = "jobdata" Then 'hires/job/jobdata
                        xJobCountry = "": xCategory = "": xEEO = "": xUnion = "": xLoc = "": xHireMgr = "": xBusi = "": xPlant = "": xDiv = "": xPosCode = ""
                        If NodeChild1.HasChildNodes Then
                            Set nodelistCh2 = NodeChild1.ChildNodes
                            For Each NodeChild2 In nodelistCh2
                                'Debug.Print NodeChild2.SelectSingleNode("@name").Text & " - " & NodeChild2.Text
                                'Debug.Print NodeChild2.Text
                                If NodeChild2.SelectSingleNode("@name").Text = "JOBLOCATIONCOUNTRY" Then
                                    xJobCountry = NodeChild2.Text
                                    xJobCountry = getCountryFromDesc(xJobCountry)
                                End If
                                If NodeChild2.SelectSingleNode("@name").Text = "EMPLOYEECATEGORY" Then xCategory = NodeChild2.Text
                                If NodeChild2.SelectSingleNode("@name").Text = "EEO" Then xEEO = NodeChild2.Text
                                If NodeChild2.SelectSingleNode("@name").Text = "UNION" Then xUnion = NodeChild2.Text
                                If NodeChild2.SelectSingleNode("@name").Text = "SITELOCATION" Then xLoc = NodeChild2.Text
                                If NodeChild2.SelectSingleNode("@name").Text = "HIRINGMANAGER" Then xHireMgr = NodeChild2.Text
                                If NodeChild2.SelectSingleNode("@name").Text = "BUSUNITPLANTDIVISION" Then
                                    xTemp = NodeChild2.Text
                                    'get BUSUNIT
                                    L = Len(xTemp): I = InStr(1, xTemp, "-")
                                    xBusi = Trim(Left(xTemp, I - 1))
                                    xTemp = Trim(Mid(xTemp, I + 1, L)) 'remove BUSUNIT
                                    'get PLANT
                                    L = Len(xTemp): I = InStr(1, xTemp, "-")
                                    xPlant = Trim(Left(xTemp, I - 1))
                                    'get DIVISION
                                    xDiv = Trim(Mid(xTemp, I + 1, L)) 'remove PLANT and the rest is DIVISION
                                End If
                                If NodeChild2.SelectSingleNode("@name").Text = "POSITIONCODE" Then xPosCode = NodeChild2.Text

                            Next
                        End If
                    End If
                    If NodeChild1.nodeName = "candidate" Then 'hires/job/candidate
                        'Debug.Print NodeChild.SelectSingleNode("@id").Text '
                        xCandID = NodeChild1.SelectSingleNode("@id").Text
                        
                        SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE SF_CANDIDATE = " & xCandID & " "
                        SQLQ = SQLQ & "AND SF_JOB_NUMBER = " & xJobNo & " "
                        SQLQ = SQLQ & "AND SF_FILEDATE = " & Date_SQL(xFileDate) & " "
                         SQLQ = SQLQ & "AND SF_XML_GROUP = 'Hires' "
                        If rsXMLTbl.State <> 0 Then rsXMLTbl.Close
                        rsXMLTbl.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If rsXMLTbl.EOF Then
                            rsXMLTbl.AddNew
                            rsXMLTbl("SF_XML_GROUP") = "Hires"
                            rsXMLTbl("SF_FILEDATE") = CVDate(xFileDate)
                            rsXMLTbl("SF_JOB_NUMBER") = xJobNo
                            rsXMLTbl("SF_CANDIDATE") = xCandID
                        End If
                        If Len(xJobCountry) > 0 Then rsXMLTbl("SF_WORKCOUNTRY") = Left(xJobCountry, 30)
                        If Len(xCategory) > 0 Then
                            rsXMLTbl("SF_PT") = Left(xCategory, 30)
                            xTemp = Left(xCategory, 30)
                            xTemp = getHRTABLCodeFromDesc("EDPT", xTemp)
                            If Len(xTemp) > 0 Then
                                rsXMLTbl("SF_PTCODE") = xTemp
                            End If
                        End If
                        If Len(xEEO) > 0 Then rsXMLTbl("SF_EEO") = Left(xEEO, 20)
                        'If Len(xUnion) > 0 Then rsXMLTbl("SF_ORG") = Left(xUnion, 30) '?
                        If Len(xUnion) > 0 Then
                            rsXMLTbl("SF_UNION") = Left(xUnion, 30)
                            xTemp = Left(xUnion, 30)
                            xTemp = getHRTABLCodeFromDesc("EDOR", xTemp)
                            If Len(xTemp) > 0 Then
                                rsXMLTbl("SF_ORG") = xTemp
                            End If
                        End If
                        If Len(xLoc) > 0 Then rsXMLTbl("SF_SITE_LOCATION") = Left(xLoc, 50)
                        If Len(xHireMgr) > 0 Then rsXMLTbl("SF_HIRING_MANAGER") = Left(xHireMgr, 60)
                        If Len(xBusi) > 0 Then rsXMLTbl("SF_BUSUNIT") = Left(xBusi, 30)
                        If Len(xPlant) > 0 Then rsXMLTbl("SF_PLANT") = Left(xPlant, 30)
                        If Len(xDiv) > 0 Then
                            xDiv = Left(xDiv, 4)
                            rsXMLTbl("SF_DIV") = xDiv
                            'adding the other fields based on Div
                            SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & xDiv & "' "
                            'If xDiv = "1094" Then
                            '    Debug.Print ""
                            'End If
                            If rsTemp.State <> 0 Then rsTemp.Close
                            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                            If Not rsTemp.EOF Then
                                If Not IsNull(rsTemp("DV_LOC")) Then rsXMLTbl("SF_LOC") = rsTemp("DV_LOC")
                                If Not IsNull(rsTemp("DV_SECTION")) Then rsXMLTbl("SF_SECTION") = rsTemp("DV_SECTION")
                                If Not IsNull(rsTemp("DV_REGION")) Then rsXMLTbl("SF_REGION") = rsTemp("DV_REGION")
                                If Not IsNull(rsTemp("DV_ADMINBY")) Then rsXMLTbl("SF_ADMINBY") = rsTemp("DV_ADMINBY")
                            End If
                        End If
                        
                        'If Len(xPosCode) > 0 Then rsXMLTbl("SF_POSITIONCODE") = Left(xPosCode, 6)
                        If Len(xPosCode) > 0 Then
                            rsXMLTbl("SF_JOBCODE") = Left(xPosCode, 6)
                            rsXMLTbl("SF_POSITIONCODE") = Left(getWFCPosFromJobSec(xPosCode, xDiv), 25)
                        End If
                        
                        'xml - candidate loop - begin
                        If NodeChild1.HasChildNodes Then
                            Set nodelistCh2 = NodeChild1.ChildNodes
                            For Each NodeChild2 In nodelistCh2
                                'Debug.Print NodeChild2.nodeName 'lastname
                                'Debug.Print NodeChild2.Text 'Smitch
                                If NodeChild2.nodeName = "lastname" Then rsXMLTbl("SF_SURNAME") = Left(Trim(NodeChild2.Text), 40)
                                If NodeChild2.nodeName = "firstname" Then rsXMLTbl("SF_FNAME") = Left(Trim(NodeChild2.Text), 40)
                                If NodeChild2.nodeName = "address" Then
                                    xTemp = Left(Trim(NodeChild2.Text), 40)
                                    If Len(xTemp) > 0 Then rsXMLTbl("SF_ADDR1") = xTemp
                                End If
                                If NodeChild2.nodeName = "apt" Then
                                    xTemp = Left(Trim(NodeChild2.Text), 40)
                                    If Len(xTemp) > 0 Then rsXMLTbl("SF_ADDR2") = xTemp
                                End If
                                If NodeChild2.nodeName = "city" Then
                                    xTemp = Left(Trim(NodeChild2.Text), 30)
                                    If Len(xTemp) > 0 Then rsXMLTbl("SF_CITY") = xTemp
                                End If
                                If NodeChild2.nodeName = "provincestate" Then
                                    xTemp = Left(Trim(NodeChild2.Text), 30)
                                    If Len(xTemp) > 0 Then
                                        rsXMLTbl("SF_PROV") = xTemp
                                        xTemp = getProvCodeFromDesc(xTemp)
                                        If Len(xTemp) > 0 Then
                                            rsXMLTbl("SF_HR_PROV") = xTemp
                                        End If
                                    End If
                                End If
                                If NodeChild2.nodeName = "postalcode" Then
                                    xTemp = Left(Trim(NodeChild2.Text), 15)
                                    If Len(xTemp) > 0 Then rsXMLTbl("SF_PCODE") = xTemp
                                End If
                                If NodeChild2.nodeName = "country" Then
                                    xTemp = Left(Trim(NodeChild2.Text), 30)
                                    xTemp = getCountryFromDesc(xTemp)
                                    If Len(xTemp) > 0 Then rsXMLTbl("SF_COUNTRY") = xTemp
                                End If
                                If NodeChild2.nodeName = "phone" Then
                                    xTemp = Left(Trim(NodeChild2.Text), 25)
                                    xTemp = Replace(xTemp, "-", "")
                                    xTemp = Replace(xTemp, ".", "")
                                    xTemp = Replace(xTemp, "(", "")
                                    xTemp = Replace(xTemp, ")", "")
                                    xTemp = Replace(xTemp, " ", "")
                                    If Len(xTemp) > 0 Then rsXMLTbl("SF_PHONE") = xTemp
                                End If
                                If NodeChild2.nodeName = "cellphone" Then
                                    xTemp = Left(Trim(NodeChild2.Text), 25)
                                    xTemp = Replace(xTemp, "-", "")
                                    xTemp = Replace(xTemp, ".", "")
                                    xTemp = Replace(xTemp, "(", "")
                                    xTemp = Replace(xTemp, ")", "")
                                    xTemp = Replace(xTemp, " ", "")
                                    If Len(xTemp) > 0 Then rsXMLTbl("SF_BUSNBR") = xTemp
                                End If
                                If NodeChild2.nodeName = "candidateprofile" Then
                                    If NodeChild2.HasChildNodes Then
                                        Set nodelistCh3 = NodeChild2.ChildNodes
                                        For Each NodeChild3 In nodelistCh3
                                            'Debug.Print NodeChild3.SelectSingleNode("@name").Text & " - " & NodeChild3.Text
                                            If NodeChild3.SelectSingleNode("@name").Text = "SALARYFREQUENCY" Then
                                                xTemp = Left(Trim(NodeChild3.Text), 10)
                                                If Len(xTemp) > 0 Then rsXMLTbl("SF_SALARYFREQUENCY") = xTemp
                                            End If
                                            If NodeChild3.SelectSingleNode("@name").Text = "VETERAN" Then
                                                xTemp = Left(Trim(NodeChild3.Text), 3)
                                                If Len(xTemp) > 0 Then rsXMLTbl("SF_VETERAN") = xTemp
                                            End If
                                            If NodeChild3.SelectSingleNode("@name").Text = "EMPLOYMENTSTATUS" Then
                                                xTemp = Left(Trim(NodeChild3.Text), 30)
                                                If Len(xTemp) > 0 Then
                                                    rsXMLTbl("SF_EMP") = xTemp
                                                    xTemp = getHRTABLCodeFromDesc("EDEM", xTemp)
                                                    If Len(xTemp) > 0 Then
                                                        rsXMLTbl("SF_EMPCODE") = xTemp
                                                    End If
                                                End If
                                            End If
                                            If NodeChild3.SelectSingleNode("@name").Text = "VIETNAMVET" Then
                                                xTemp = Left(Trim(NodeChild3.Text), 3)
                                                If Len(xTemp) > 0 Then rsXMLTbl("SF_VIETNAMVET") = xTemp
                                            End If
                                            If NodeChild3.SelectSingleNode("@name").Text = "DISABILITY" Then
                                                xTemp = Left(Trim(NodeChild3.Text), 3)
                                                If Len(xTemp) > 0 Then rsXMLTbl("SF_DISABILITY") = xTemp
                                            End If
                                            If NodeChild3.SelectSingleNode("@name").Text = "SALARY" Then
                                                xTemp = NodeChild3.Text
                                                If IsNumeric(xTemp) Then
                                                    rsXMLTbl("SF_SALARY") = xTemp
                                                End If
                                            End If
                                            If NodeChild3.SelectSingleNode("@name").Text = "STARTDATE" Then
                                                xTemp = Trim(NodeChild3.Text)
                                                If Len(xTemp) > 0 Then
                                                    xTemp = CVDate(Format(xTemp, "yyyy-mm-dd"))
                                                    If IsDate(xTemp) Then
                                                        rsXMLTbl("SF_STARTDATE") = CVDate(xTemp)
                                                    End If
                                                End If
                                            End If
                                            If NodeChild3.SelectSingleNode("@name").Text = "HIRETYPE" Then
                                                xTemp = Left(Trim(NodeChild3.Text), 10)
                                                If Len(xTemp) > 0 Then rsXMLTbl("SF_HIRETYPE") = xTemp
                                            End If
                                            If NodeChild3.SelectSingleNode("@name").Text = "EMPLOYEENUMBER" Then
                                                xTemp = NodeChild3.Text
                                                If IsNumeric(xTemp) Then
                                                    rsXMLTbl("SF_EMPNBR") = xTemp
                                                End If
                                                If Len(xTemp) > 0 Then rsXMLTbl("SF_PAYROLL_ID") = Left(Trim(xTemp), 15)
                                            End If
                                            If NodeChild3.SelectSingleNode("@name").Text = "MIDDLENAME" Then
                                                xTemp = Left(Trim(NodeChild3.Text), 30)
                                                If Len(xTemp) > 0 Then rsXMLTbl("SF_MIDNAME") = xTemp ' Left(Trim(NodeChild3.Text), 30)
                                            End If
                                        Next
                                    End If
                                
                                End If
                                If NodeChild2.nodeName = "eeodata" Then
                                    If NodeChild2.HasChildNodes Then
                                        Set nodelistCh3 = NodeChild2.ChildNodes
                                        For Each NodeChild3 In nodelistCh3
                                            'Debug.Print NodeChild3.SelectSingleNode("@name").Text & " - " & NodeChild3.Text
                                            If NodeChild3.SelectSingleNode("@name").Text = "ETHNICITY" Then
                                                xTemp = Left(Trim(NodeChild3.Text), 30)
                                                If Len(xTemp) > 0 Then rsXMLTbl("SF_ETHNICITY") = xTemp
                                            End If
                                            If NodeChild3.SelectSingleNode("@name").Text = "GENDER" Then
                                                xTemp = Left(Trim(NodeChild3.Text), 6)
                                                If Len(xTemp) > 0 Then rsXMLTbl("SF_GENDER") = xTemp
                                            End If
                                            If NodeChild3.SelectSingleNode("@name").Text = "RACE" Then
                                                xTemp = Left(Trim(NodeChild3.Text), 30)
                                                If Len(xTemp) > 0 Then rsXMLTbl("SF_RACE") = xTemp
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                        End If
                        'xml - candidate loop - end
                        rsXMLTbl("SF_LDATE") = Date
                        rsXMLTbl("SF_LTIME") = Time$
                        rsXMLTbl("SF_LUSER") = glbUserID
                        rsXMLTbl.Update
                        rsXMLTbl.Close
                    End If
                Next
            End If
        Next Node
    End If
    '------------ hire ---------- end
    
    '------------ unhire ---------- begin
unhires:
    Set nodelist = doc.SelectNodes("/feed/unhires")
    If Not (nodelist Is Nothing) Then
        For Each Node In nodelist 'hires/job
            If Node.HasChildNodes Then
                Set nodelistCh1 = Node.ChildNodes
                For Each NodeChild1 In nodelistCh1
                    'Debug.Print NodeChild1.nodeName 'lastname
                    'Debug.Print NodeChild1.Text 'Smitch
                    If NodeChild1.nodeName = "candidate" Then
                        xCandID = NodeChild1.SelectSingleNode("@id").Text
                        
                        SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE SF_CANDIDATE = " & xCandID & " "
                        'SQLQ = SQLQ & "AND SF_JOB_NUMBER = " & xJobNo & " "
                        SQLQ = SQLQ & "AND SF_FILEDATE = " & Date_SQL(xFileDate) & " "
                        SQLQ = SQLQ & "AND SF_XML_GROUP = 'Unhires' "
                        If rsXMLTbl.State <> 0 Then rsXMLTbl.Close
                        rsXMLTbl.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If rsXMLTbl.EOF Then
                            rsXMLTbl.AddNew
                            rsXMLTbl("SF_XML_GROUP") = "Unhires"
                            rsXMLTbl("SF_FILEDATE") = CVDate(xFileDate)
                            rsXMLTbl("SF_CANDIDATE") = xCandID
                            rsXMLTbl("SF_HIRETYPE") = "UNHIRES"
                        End If
                        
                        'xml - candidate loop - begin
                        If NodeChild1.HasChildNodes Then
                            Set nodelistCh2 = NodeChild1.ChildNodes
                            For Each NodeChild2 In nodelistCh2
                                'Debug.Print NodeChild2.nodeName 'lastname
                                'Debug.Print NodeChild2.Text 'Smitch
                                If NodeChild2.nodeName = "lastname" Then rsXMLTbl("SF_SURNAME") = Left(Trim(NodeChild2.Text), 40)
                                If NodeChild2.nodeName = "firstname" Then rsXMLTbl("SF_FNAME") = Left(Trim(NodeChild2.Text), 40)
                                If NodeChild2.nodeName = "field" Then
                                    If NodeChild2.SelectSingleNode("@name").Text = "BUSUNITPLANTDIVISION" Then
                                        xTemp = NodeChild2.Text
                                        'get BUSUNIT
                                        L = Len(xTemp): I = InStr(1, xTemp, "-")
                                        xBusi = Trim(Left(xTemp, I - 1))
                                        xTemp = Trim(Mid(xTemp, I + 1, L)) 'remove BUSUNIT
                                        'get PLANT
                                        L = Len(xTemp): I = InStr(1, xTemp, "-")
                                        xPlant = Trim(Left(xTemp, I - 1))
                                        'get DIVISION
                                        xDiv = Trim(Mid(xTemp, I + 1, L)) 'remove PLANT and the rest is DIVISION
                                        If Len(xBusi) > 0 Then rsXMLTbl("SF_BUSUNIT") = Left(xBusi, 30)
                                        If Len(xPlant) > 0 Then rsXMLTbl("SF_PLANT") = Left(xPlant, 30)
                                        If Len(xDiv) > 0 Then
                                            'rsXMLTbl("SF_DIV") = Left(xDiv, 4)
                                            xDiv = Left(xDiv, 4)
                                            rsXMLTbl("SF_DIV") = xDiv
                                            'adding the other fields based on Div
                                            SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & xDiv & "' "
                                            'If xDiv = "1094" Then
                                            '    Debug.Print ""
                                            'End If
                                            If rsTemp.State <> 0 Then rsTemp.Close
                                            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                                            If Not rsTemp.EOF Then
                                                If Not IsNull(rsTemp("DV_LOC")) Then rsXMLTbl("SF_LOC") = rsTemp("DV_LOC")
                                                If Not IsNull(rsTemp("DV_SECTION")) Then rsXMLTbl("SF_SECTION") = rsTemp("DV_SECTION")
                                                If Not IsNull(rsTemp("DV_REGION")) Then rsXMLTbl("SF_REGION") = rsTemp("DV_REGION")
                                                If Not IsNull(rsTemp("DV_ADMINBY")) Then rsXMLTbl("SF_ADMINBY") = rsTemp("DV_ADMINBY")
                                            End If
                                        End If
                                    End If
                                    If NodeChild2.SelectSingleNode("@name").Text = "STARTDATE" Then
                                        xTemp = Trim(NodeChild2.Text)
                                        If Len(xTemp) > 0 Then
                                            xTemp = CVDate(Format(xTemp, "yyyy-mm-dd"))
                                            If IsDate(xTemp) Then
                                                rsXMLTbl("SF_STARTDATE") = CVDate(xTemp)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                        rsXMLTbl("SF_LDATE") = Date
                        rsXMLTbl("SF_LTIME") = Time$
                        rsXMLTbl("SF_LUSER") = glbUserID
                        rsXMLTbl.Update
                        rsXMLTbl.Close
                    End If
                Next
            End If
        Next
    End If
    '------------ unhire ---------- end
    Set nodelist = Nothing
    
    SQLQ = "UPDATE HRSF_XML_IMPORT SET SF_UPT_DEMO = 0 WHERE SF_UPT_DEMO IS NULL"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRSF_XML_IMPORT SET SF_UPT_STATUS = 0 WHERE SF_UPT_STATUS IS NULL"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRSF_XML_IMPORT SET SF_UPT_POSITION = 0 WHERE SF_UPT_POSITION IS NULL"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRSF_XML_IMPORT SET SF_UPT_SALARY = 0 WHERE SF_UPT_SALARY IS NULL"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRSF_XML_IMPORT SET SF_UPT_PROCESSED = 0 WHERE SF_UPT_PROCESSED IS NULL"
    gdbAdoIhr001.Execute SQLQ
End If

End Sub



Private Function getProvCodeFromDesc(xDesc)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = ""
    If Not IsNull(xDesc) Then
        'SQLQ = "SELECT * FROM HRPROV WHERE DESCR = '" & xDesc & "' "
        'SQLQ = "SELECT * FROM HRPROV WHERE '" & xDesc & "' LIKE DESCR "
        SQLQ = "SELECT * FROM HRPROV WHERE DESCR LIKE UPPER('%" & xDesc & "%') "
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            retVal = rsTemp("CODE")
        End If
    End If
    getProvCodeFromDesc = retVal
End Function

Private Function getCountryFromDesc(xDesc)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = ""
    If Not IsNull(xDesc) Then
        If xDesc = "United States" Then xDesc = "U.S.A."
        xDesc = UCase(xDesc)
        SQLQ = "SELECT * FROM HR_DIVISION WHERE '" & xDesc & "' LIKE DV_COUNTRY "
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            retVal = rsTemp("DV_COUNTRY")
        End If
    End If
    getCountryFromDesc = retVal
End Function

Private Function getXMLFileLocation()
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim xPath
Dim retVal As String
    retVal = ""
    SQLQ = "SELECT * FROM HRSF_XMLFILE_LOCATION"
    rs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
        rs.AddNew
        xPath = glbIHRREPORTS
        If Right(xPath, 1) = "\" Then
            xPath = Left(xPath, Len(xPath) - 1)
        End If
        rs("SF_LOCATION") = xPath
        rs.Update
    Else
        xPath = rs("SF_LOCATION")
    End If
    rs.Close
    retVal = xPath
    getXMLFileLocation = retVal
End Function

Private Sub XMLFileLocationScreen(xType)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim xPath
Dim xDefPath


If xType = "DISP" Then
    xDefPath = getXMLFileLocation
    If Len(Dir$(xDefPath, vbDirectory)) = 0 Then
        MsgBox "Invalid XML File Location Path: " & Chr(10) & "  " & xDefPath & Chr(10) & Chr(10) & "Please go to HRSoft/Setup/XML File Location to setup a valid location."
        'Exit Sub
    End If
    
    txtPDFilename.Text = ""
    SQLQ = "SELECT * FROM HRSF_XMLFILE_LOCATION"
    rs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
        rs.AddNew
        xPath = glbIHRREPORTS
        If Right(xPath, 1) = "\" Then
            xPath = Left(xPath, Len(xPath) - 1)
        End If
        rs("SF_LOCATION") = xPath
        rs.Update
    End If
    txtPDFilename.Text = rs("SF_LOCATION")
    If Len(Dir$(txtPDFilename.Text, vbDirectory)) > 0 Then
        Drive1.Drive = Left(txtPDFilename.Text, 2)
        Dir1.Path = txtPDFilename.Text
    Else
        xPath = glbIHRREPORTS
        If Right(xPath, 1) = "\" Then
            xPath = Left(xPath, Len(xPath) - 1)
        End If
        txtPDFilename.Text = xPath
        Drive1.Drive = Left(xPath, 2)
        Dir1.Path = xPath
    End If
    rs.Close

End If

End Sub

Private Sub TownofLaSalleScreenDownload() 'Ticket #26912 Franks 06/22/2015
    lblDet(0).Visible = False
    lblDet(1).Visible = False
    dlpDate(0).Visible = False
    dlpDate(1).Visible = False
    fraXMLLocation.Visible = False
    cmdDownToPath.Visible = False
    cmdLoadXML.Visible = False
    
    'cmdClose.Top = cmdDownToPath.Top
    
    cmdStart.Top = 720
    cmdClose.Top = cmdStart.Top + 480
End Sub
Private Sub TownofLaSalleScreenUpload() 'Ticket #26912 Franks 06/23/2015
    lblDet(0).Visible = False
    lblDet(1).Visible = False
    dlpDate(0).Visible = False
    dlpDate(1).Visible = False
    fraXMLLocation.Visible = False
    cmdDownToPath.Visible = False
    cmdLoadXML.Visible = False
    
    cmdStart.Top = 720
    cmdClose.Top = cmdStart.Top + 480 'cmdDownToPath.Top
        
    cmdStart.Caption = "Upload"
End Sub

Private Sub TownofLaSalleUpload() 'Ticket #26912 Franks 06/23/2015
Dim I As Integer
Dim xFlag As Boolean
Dim xFilePath As String
Dim rsEmp As New ADODB.Recordset, rsJH As New ADODB.Recordset, rsSH As New ADODB.Recordset
Dim rsTABL As New ADODB.Recordset, rsJOB As New ADODB.Recordset, rsDEPT As New ADODB.Recordset
Dim rsPYMST As New ADODB.Recordset
Dim RecCNT As Long, buf As String
Dim ImportStatus As CSVImportStatus
Dim NotImpExists As Long
Dim FTPStatus As FtpErrorEnum
Dim xRemoteFile As String
Dim xLocalFile As String
Dim xDays As Integer
Dim xDATE
Dim xFileAmt As Integer
Dim xDefPath As String
Dim a As Integer, msg As String, DtTm As Variant
Dim Title$, DgDef, Response%

    'Msg$ = "This program will export employee information into " & AppPath & " " & Chr(10)
    'Msg$ = Msg$ & "The file name is 'EmployeeExpTownOfLasalle.xls' " & Chr(10)
    msg$ = "This program will export employee information into 'LasalleEmployeeExp.csv', " ' & Chr(10)
    msg$ = msg$ & "then upload it to System 24/7 FTP site " & Chr(10) & Chr(10)
    msg$ = msg$ & "Notes: not include Leave of Absence employees " & Chr(10) & Chr(10)
    msg$ = msg$ & "Are you sure you want to do it?"

    Title$ = "info:HR - Data Export/Upload"
    DgDef = 4 + 256
    Response% = MsgBox(msg, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then    ' Evaluate response
        Exit Sub
    End If
    
    If Not TownofLaSalleEmployeeExport Then
        Exit Sub
    End If
    
    '---- begin to upload file to ftp
    If Not ReadFTPData(True) Then '(clpPAYP.Text) Then
        MsgBox "Please complete the FTP Setup screen before attempting to use this function.", vbInformation + vbOKOnly, "Setup Data Not Entered"
        Exit Sub
    Else
        
        xFilePath = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") ' xDefPath & "\"
    End If
    
        
    xRemoteFile = "LasalleEmployeeExp.csv"
    xLocalFile = xFilePath & xRemoteFile
        
        
    'xRemoteFile = "TownOfLasalleCourseScore.csv" '???
    xLocalFile = xFilePath & xRemoteFile
    
    locDownloadFalg = False
    
    Screen.MousePointer = vbHourglass

    '----------create winscriptfile - begin
    'Call TownofLaSalleWinscpScripts("Upload", xRemoteFile, xFilePath)
    Call TownofLaSalleWinscpScripts("Upload", xRemoteFile, xFilePath)
    '----------create winscriptfile - end
    Screen.MousePointer = vbDefault
    
    Exit Sub
        '''MsgBox "step 2: export file = " & xLocalFile

        ''''xRemoteFile = "WB_HireFeed_UAT_20130724.xml" '"franktest.csv"
        ''''FTPStatus = modFTP.FTPGetFile(glbPayWebData.Host, glbPayWebData.Username, glbPayWebData.Password, xFilePath & "franktest.csv", "franktest.csv") ' "PYMST.CSV")
        ''FTPStatus = modFTP.FTPGetFile(glbPayWebData.Host, glbPayWebData.UserName, glbPayWebData.Password, xLocalFile, "outgoing/" & xRemoteFile) ' "PYMST.CSV")
        'FTPStatus = modFTP.FTPSendFile(glbPayWebData.Host, glbPayWebData.UserName, glbPayWebData.Password, xLocalFile, "outgoing/" & xRemoteFile)
        FTPStatus = modFTP.FTPSendFile(glbPayWebData.Host, glbPayWebData.UserName, glbPayWebData.Password, xLocalFile, xRemoteFile)
        
        '''MsgBox "step 3: check if the file has been downloaded "

        Select Case FTPStatus
            Case FtpErrorEnum.PSCPNotFound
                MsgBox "PSCP runtime utility not found.  Please call info:HR support for assistance.", vbInformation + vbOKOnly, "PSCP.EXE Not Found"
                Exit Sub
            Case FtpErrorEnum.PSCPTimedOut
                If MsgBox("SCP session timed out.  View log?", vbCritical + vbYesNo, "Timed Out") = vbYes Then
                    Shell "notepad " & glbIHRREPORTS & "SCPLOG.TXT", vbNormalFocus
                End If
                Exit Sub
            Case FtpErrorEnum.PSCPError
                '''MsgBox "step 4: PSCPError " & FtpErrorEnum.PSCPError
                
                'if we need to see the error message about "This computer must be initialized on "
                'then uncomment the follow 3 lines
                'If MsgBox("PSCP returned an error.  View log?", vbCritical + vbYesNo, "SCP Error") = vbYes Then
                '    Shell "notepad " & glbIHRREPORTS & "SCPLOG.TXT", vbNormalFocus
                'End If
                
                ''Ticket #25150 Franks the following 6 lines caused error "Permission denied" at Woodbridge
                ''If Not IsPSCP_Initialized Then  'Ticket #25088 Franks 02/18/2014
                ''    If MsgBox("PSCP returned an error.  View log?", vbCritical + vbYesNo, "SCP Error") = vbYes Then
                ''        Shell "notepad " & glbIHRREPORTS & "SCPLOG.TXT", vbNormalFocus
                ''    End If
                ''    Exit Sub
                ''End If
                MsgBox "File uploaded successfully."
        End Select
    
End Sub

Private Function TownofLaSalleWinscpScript_NotUsed(xType, xFileName, xPath)  'Ticket #26912 Franks 07/17/2015
Dim xWinscpScriptFile As String
Dim buf, xTemp, xWin247str
Dim xlsFileMat
Dim xComFile, xBatFileName, xDFilename
Dim xTmpPath
    
    xTmpPath = GetShortName(glbIHRREPORTS)
    'xTmpPath = GetShortName("C:\SSWORK\IHR80\")
    
    'create bat file - begin
    xBatFileName = xTmpPath & "WinscpSc.bat"
    If (Dir(xBatFileName)) <> "" Then
        Kill xBatFileName
        Call Pause(0.5)
    End If
    If (Dir(xBatFileName)) = "" Then
        Open xBatFileName For Output As #1
        buf = Left(xTmpPath, 2)
        Print #1, buf
        buf = "CD\"
        Print #1, buf
        buf = "CD " & xTmpPath
        Print #1, buf
        buf = "winscp.com /script=WinscpSc.txt"
        Print #1, buf
        Close #1
        Call Pause(1)
    End If
    'create bat file - end

    xWinscpScriptFile = xTmpPath & "WinscpSc.txt"
    
    xPath = xPath & IIf(Right(xPath, 1) = "\", "", "\")
    xPath = GetShortName(xPath)
    
    'the following code worked well, but it is hard code.
    'xWin247str = "open ftp://lasalle.systems24-7.com|lasalle:$#tol1#$@lasalle.systems24-7.com:21 -passive -explicitssl -explicittls"
    xWin247str = "open ftp://" & glbPayWebData.UserName & ":" & glbPayWebData.Password & "@" & glbPayWebData.Host & ":21 -passive -explicitssl -explicittls"
    
    If (Dir(xWinscpScriptFile)) <> "" Then
        Kill xWinscpScriptFile
        Call Pause(0.5)
    End If
    If xType = "Upload" Then
        Open xWinscpScriptFile For Output As #1
        buf = "option batch abort"
        Print #1, buf
        
        buf = xWin247str 'open ftp
        Print #1, buf
        
        buf = "option confirm off"
        Print #1, buf
        
        'xtemp = "C:\INFOHR\WinSCP\scripts\license2.txt"
        'buf = "put " & xPath & xFileName 'upload xFileName to ftp
        buf = "put " & xFileName
        Print #1, buf
        
        buf = "close"
        Print #1, buf
        buf = "close"
        Print #1, buf
        
        Close #1
        '------------------------ end of create file
        
        ''Shell "cmd /c C:\INFOHR\WinSCP\scripts\winscp.com /script=C:\INFOHR\WinSCP\scripts\winscpupload.txt"
        'xtemp = "cmd /c " & glbIHRREPORTS & "winscp.com /script=" & xWinscpScriptFile
        'Shell xtemp
        
        Call Pause(3)
        
        DoEvents
        
        xComFile = xBatFileName
        xTemp = "cmd /c " & xComFile
        Shell xTemp
        
        Call Pause(2)
        
        If (Dir(xWinscpScriptFile)) <> "" Then
            Kill xWinscpScriptFile
        End If
        
        MsgBox "   Finished.   "
    End If
    If xType = "Download" Then
        xDFilename = xTmpPath & xFileName
        If (Dir(xDFilename)) <> "" Then ' download file
            Kill xDFilename
        End If
        Open xWinscpScriptFile For Output As #1
        buf = "option batch abort"
        Print #1, buf
        
        buf = xWin247str 'open ftp
        Print #1, buf
        
        buf = "option confirm off"
        Print #1, buf
        
        'xtemp = "C:\INFOHR\WinSCP\scripts\license2.txt"
        'buf = "get " & " " & xFileName
        'buf = "get " & " Export\" & xFileName
        buf = "get " & " Export\" & xFileName & " " & xFileName
        
        Print #1, buf
        
        buf = "close"
        Print #1, buf
        buf = "close"
        Print #1, buf
        
        Close #1
        '------------------------ end of create file
        
        'xComFile = xTmpPath & "A.bat"
        xComFile = xBatFileName
        xTemp = "cmd /c " & xComFile
        Shell xTemp
        
        Call Pause(3)
        If (Dir(xWinscpScriptFile)) <> "" Then
            Kill xWinscpScriptFile
        End If
        
        If (Dir(xDFilename)) <> "" Then
            locDownloadFalg = True
        Else
            'not file found, let's wait for another 5''
            Call Pause(5)
            If (Dir(xDFilename)) <> "" Then
                locDownloadFalg = True
            End If
        End If
        If Not locDownloadFalg Then
            MsgBox "No file was downloaded, please try again."
        End If
        'check if found this file, then
    End If

End Function

Private Function TownofLaSalleWinscpScripts(xType, xFileName, xPath) 'Ticket #26912 Franks 07/17/2015
Dim xWinscpScriptFile As String
Dim buf, xTemp, xWin247str
Dim xlsFileMat
Dim xComFile, xBatFileName, xDFilename
Dim xTmpPath
Dim xLMSDownLoadLog As String

    xTmpPath = GetShortName(glbIHRREPORTS)
    'xTmpPath = GetShortName("C:\SSWORK\IHR80\")
    
    'create bat file - begin
    'Ticket #27742 Franks 11/10/2015 - do not delete this file if it exists
    'xBatFileName = xTmpPath & "WinscpSc.bat"
    'If (Dir(xBatFileName)) <> "" Then
    '    Kill xBatFileName
    '    Call Pause(0.5)
    'End If
    If xType = "Download" Then xBatFileName = xTmpPath & "WinscpScD.bat"
    If xType = "Upload" Then xBatFileName = xTmpPath & "WinscpScU.bat"
    
    If (Dir(xBatFileName)) = "" Then
        Open xBatFileName For Output As #1
        buf = Left(xTmpPath, 2)
        Print #1, buf
        buf = "CD\"
        Print #1, buf
        buf = "CD " & xTmpPath
        Print #1, buf
        'buf = "winscp.com /script=WinscpSc.txt"
        'Ticket #27742 Franks 11/10/2015
        If xType = "Download" Then buf = "winscp.com /script=WinscpScD.txt"
        If xType = "Upload" Then buf = "winscp.com /script=WinscpScU.txt"
        Print #1, buf
        Close #1
        Call Pause(1)
    End If
    'create bat file - end

    'xWinscpScriptFile = xTmpPath & "WinscpSc.txt"
    'Ticket #27742 Franks 11/10/2015
    If xType = "Download" Then xWinscpScriptFile = xTmpPath & "WinscpScD.txt"
    If xType = "Upload" Then xWinscpScriptFile = xTmpPath & "WinscpScU.txt"

    xPath = xPath & IIf(Right(xPath, 1) = "\", "", "\")
    xPath = GetShortName(xPath)
    
    ''the following code worked well, but it is hard code.
    ''xWin247str = "open ftp://lasalle.systems24-7.com|lasalle:$#tol1#$@lasalle.systems24-7.com:21 -passive -explicitssl -explicittls"
    'xWin247str = "open ftp://" & glbPayWebData.UserName & ":" & glbPayWebData.Password & "@" & glbPayWebData.Host & ":21 -passive -explicitssl -explicittls"
    ''open ftp://lasalle.systems24-7.com|lasalle:$#tol1#$@lasalle.systems24-7.com:21/ -passive -explicitssl -explicittls -certificate="*"
    xWin247str = "open ftp://" & glbPayWebData.UserName & ":" & glbPayWebData.Password & "@" & glbPayWebData.Host & ":21/ -passive -explicitssl -explicittls -certificate=" & Chr(34) & "*" & Chr(34)

    'Ticket #27742 Franks 11/10/2015
    'If (Dir(xWinscpScriptFile)) <> "" Then
    '    Kill xWinscpScriptFile
    '    Call Pause(0.5)
    'End If
    If xType = "Upload" Then
        If (Dir(xWinscpScriptFile)) = "" Then
            Open xWinscpScriptFile For Output As #1
            buf = "option batch abort"
            Print #1, buf
            
            buf = xWin247str 'open ftp
            Print #1, buf
            
            buf = "option confirm off"
            Print #1, buf
            
            'xtemp = "C:\INFOHR\WinSCP\scripts\license2.txt"
            'buf = "put " & xPath & xFileName 'upload xFileName to ftp
            buf = "put " & xFileName
            Print #1, buf
            
            buf = "close"
            Print #1, buf
            buf = "close"
            Print #1, buf
            
            Close #1
            '------------------------ end of create file
        End If
        
        ''Shell "cmd /c C:\INFOHR\WinSCP\scripts\winscp.com /script=C:\INFOHR\WinSCP\scripts\winscpupload.txt"
        'xtemp = "cmd /c " & glbIHRREPORTS & "winscp.com /script=" & xWinscpScriptFile
        'Shell xtemp
        
        Call Pause(3)
        
        DoEvents
        
        xComFile = xBatFileName
        xTemp = "cmd /c " & xComFile
        Shell xTemp
        
        Call Pause(2)
        
        'If (Dir(xWinscpScriptFile)) <> "" Then
        '    Kill xWinscpScriptFile
        'End If
        
        MsgBox "   Finished.   "
    End If
    If xType = "Download" Then
        xDFilename = xTmpPath & xFileName
        If (Dir(xDFilename)) <> "" Then ' download file
            Kill xDFilename
        End If
        
        If (Dir(xWinscpScriptFile)) = "" Then
            Open xWinscpScriptFile For Output As #1
            buf = "option batch abort"
            Print #1, buf
            
            buf = xWin247str 'open ftp
            Print #1, buf
            
            buf = "option confirm off"
            Print #1, buf
            
            'xtemp = "C:\INFOHR\WinSCP\scripts\license2.txt"
            'buf = "get " & " " & xFileName
            'buf = "get " & " Export\" & xFileName
            buf = "get " & " Export\" & xFileName & " " & xFileName
            
            Print #1, buf
            
            buf = "close"
            Print #1, buf
            buf = "close"
            Print #1, buf
            
            Close #1
        End If
        '------------------------ end of create file
        
        'xComFile = xTmpPath & "A.bat"
        xComFile = xBatFileName
        xTemp = "cmd /c " & xComFile
        Shell xTemp
        
        Call Pause(3)
        'If (Dir(xWinscpScriptFile)) <> "" Then
        '    Kill xWinscpScriptFile
        'End If
        
        If (Dir(xDFilename)) <> "" Then
            locDownloadFalg = True
        Else
            'not file found, let's wait for another 5''
            Call Pause(5)
            If (Dir(xDFilename)) <> "" Then
                locDownloadFalg = True
            End If
        End If
        
        'Ticket #28931 - Franks 07/22/2016 - begin
        xLMSDownLoadLog = Replace(xDFilename, "LasalleCourseScore.csv", "LasalleLMSLog.csv")
        If Not locDownloadFalg Then
            buf = Now & " - No file was downloaded, please try again."
        Else
            buf = Now & " - " & xDFilename & " was downloaded."
        End If
        'Open xLMSDownLoadLog For Output As #3
        Open xLMSDownLoadLog For Append As #3
        Print #3, buf
        Close #3
        'Ticket #28931 - Franks 07/22/2016 - end
        If Not locDownloadFalg Then
            MsgBox "No file was downloaded, please try again."
        End If
        'check if found this file, then
    End If

End Function

Private Function TownofLaSalleEmployeeExport() 'Ticket #26912 Franks 06/23/2015
    Dim rsHREmp As New ADODB.Recordset
    Dim rsPos As New ADODB.Recordset
    Dim SQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim I, totNum
    Dim xRow As Integer
    Dim Response%, msg$, DgDef
    Dim xDesc
    Dim AppPath
    Dim buf
    
    TownofLaSalleEmployeeExport = False
    
    AppPath = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\")

    On Error GoTo EmployeeExportForTownofLasalle_Err
    
    SQLQ = "SELECT * FROM HREMP "
    'only export employees who are not Leave of Absence
    SQLQ = SQLQ & "WHERE ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'EDEM' AND TB_USR3 = 0) "
    SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME, ED_EMPNBR"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsHREmp.EOF Then
        MsgBox "There is no record found to be exported"
        Exit Function
    End If
    
    
    ''Initialise/Open Excel Report file
    'xlsFileTmp = AppPath & "EmployeeExpTownOfLasalleTmp.xls"
    xlsFileMat = AppPath & "LasalleEmployeeExp.csv" '"EmployeeExpTownOfLasalle.xls"

    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
    Open xlsFileMat For Output As #1
    
    Screen.MousePointer = HOURGLASS
    
    'Print header line
    buf = """Employee Number"",""First Name"",""Last_Name"",""access(Basic or Admin)"",""store number"",""department"",""position"""
    Print #1, buf
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0

    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: I = 0
        rsHREmp.MoveFirst

        xRow = 4
        Do While Not rsHREmp.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            
            DoEvents
        
            'Employee #, Name, Previous Vac
            buf = """" & rsHREmp("ED_EMPNBR") & """"
            buf = buf & ",""" & rsHREmp("ED_FNAME") & """"
            buf = buf & ",""" & rsHREmp("ED_SURNAME") & """"
            If IsNull(rsHREmp("ED_ADMINBY")) Then
                buf = buf & ","
            Else
                buf = buf & ",""" & GetTABLDesc("EDAB", rsHREmp("ED_ADMINBY")) & """"
            End If
            buf = buf & ",""" & "7770" & """" 'store number
            xDesc = getDivDescPub(rsHREmp("ED_DIV"))
            If Len(xDesc) = 0 Then
                buf = buf & ","
            Else
                buf = buf & ",""" & xDesc & """"
            End If
            
            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE NOT JH_CURRENT = 0 AND JH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " "
            If rsPos.State <> 0 Then rsPos.Close
            rsPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsPos.EOF Then '
                xDesc = getPosMasterValueByField(rsPos("JH_JOB"), "JB_DESCR")
                buf = buf & ",""" & xDesc & """"
            Else
                buf = buf & ","
            End If
            Print #1, buf
            
            'exSheet.Cells(xRow, 1) = rsHREmp("ED_EMPNBR")
            'exSheet.Cells(xRow, 2) = rsHREmp("ED_FNAME")
            'exSheet.Cells(xRow, 3) = rsHREmp("ED_SURNAME")
            'If IsNull(rsHREmp("ED_ADMINBY")) Then
            '    xDesc = ""
            'Else
            '    xDesc = GetTABLCode("EDAB", rsHREmp("ED_ADMINBY"))
            'End If
            'exSheet.Cells(xRow, 4) = xDesc
            'exSheet.Cells(xRow, 5) = getDivDescPub(rsHREmp("ED_DIV"))
            '
            'SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE NOT JH_CURRENT = 0 AND JH_EMPNBR = " & rsHREmp("ED_EMPNBR") & " "
            'If rsPos.State <> 0 Then rsPos.Close
            'rsPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
            'If Not rsPos.EOF Then '
            '    exSheet.Cells(xRow, 6) = getPosMasterValueByField(rsPos("JH_JOB"), "JB_DESCR")
            'End If
            
            xRow = xRow + 1
            
            rsHREmp.MoveNext
        Loop
        
        ''Print Solid line under the last row and Totals of Vacation Balance Accrual
        'exSheet.Range("H" & xRow & ":H" & xRow).Borders(xlEdgeBottom).LineStyle = xlSolid
        'exSheet.Cells(xRow + 1, 8).Formula = "=Round(SUM(H9:H" & xRow - 1 & "),2)"

    End If
    
    rsHREmp.Close
    'Set rsHREmp = Nothing
    Close #1

    MDIMain.panHelp(0).FloodPercent = 100
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
    Screen.MousePointer = DEFAULT
    
    TownofLaSalleEmployeeExport = True
    
Exit Function

EmployeeExportForTownofLasalle_Err:


End Function
Private Sub TownofLaSalleDownload() 'Ticket #26912 Franks 06/23/2015
Dim I As Integer
Dim xFlag As Boolean
Dim xFilePath As String
Dim rsEmp As New ADODB.Recordset, rsJH As New ADODB.Recordset, rsSH As New ADODB.Recordset
Dim rsTABL As New ADODB.Recordset, rsJOB As New ADODB.Recordset, rsDEPT As New ADODB.Recordset
Dim rsPYMST As New ADODB.Recordset
Dim RecCNT As Long, buf As String
Dim ImportStatus As CSVImportStatus
Dim NotImpExists As Long
Dim FTPStatus As FtpErrorEnum
Dim xRemoteFile As String
Dim xLocalFile As String
Dim xDays As Integer
Dim xDATE
Dim xFileAmt As Integer
Dim xDefPath As String
Dim a As Integer, msg As String, DtTm As Variant
Dim Title$, DgDef, Response%

    msg$ = "This program will download employee Course Score file from System 24/7 FTP site " ' & Chr(10)
    msg$ = msg$ & "then update info:HR Continuing Education screen " & Chr(10) & Chr(10)
    msg$ = msg$ & "Are you sure you want to do it?"

    Title$ = "info:HR - Data Export/Download"
    DgDef = 4 + 256
    Response% = MsgBox(msg, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then    ' Evaluate response
        Exit Sub
    End If
    
    xDefPath = glbIHRREPORTS 'Dir1.Path
    
    If Not ReadFTPData(True) Then '(clpPAYP.Text) Then
        MsgBox "Please complete the FTP Setup screen before attempting to use this function.", vbInformation + vbOKOnly, "Setup Data Not Entered"
        Exit Sub
    'Else
    '    xFilePath = xDefPath & "\"
    End If

    
    'xRemoteFile = "TownOfLasalleCourseScore.csv" '???
    xRemoteFile = "LasalleCourseScore.csv"
    xLocalFile = xFilePath & xRemoteFile
    
    locDownloadFalg = False
    
    Screen.MousePointer = vbHourglass

    '----------create winscriptfile - begin
    'Call TownofLaSalleWinscpScripts("Upload", xRemoteFile, xFilePath)
    Call TownofLaSalleWinscpScripts("Download", xRemoteFile, xFilePath)
    '----------create winscriptfile - end
    
    'xLocalFile = xFilePath & "TownOfLasalleCourseScore.csv"
    xLocalFile = xDefPath & xRemoteFile
    Call TownofLaSalleCourseTaken(xLocalFile, xDefPath)
    Screen.MousePointer = vbDefault
    
    MsgBox "   Finished.   "
End Sub
Private Sub TownofLaSalleCourseTaken(ImportFile, xPath) 'Ticket #26912 Franks 06/24/2015
Dim rsEmp As New ADODB.Recordset
Dim rsTermEmp As New ADODB.Recordset 'Ticket #28931 Franks 07/22/2016
Dim rsCRSMaster As New ADODB.Recordset
Dim rsCRSEmp As New ADODB.Recordset
Dim SQLQ As String
Dim xdata As String
Dim xCNT, xTotRec
Dim xEmpNo, xCourseID, xCourseName, xCompDate, xScore, xExpiryDate, xCoursNameM
Dim xTemp
Dim xErrRec As Long
Dim ImportErrFile

    '--- get the total records
    Call Pause(2)
    Open ImportFile For Input As #1
    xTotRec = 0
    Do While Not EOF(1)
      Line Input #1, xdata
      xTotRec = xTotRec + 1
    Loop
    Close #1
        
    '--- any reject of record, copy the contents to the same CSV file
    'o   CSV filename will be "LasalleCourse_Import_Errors_YYMMDD.CSV".
    xErrRec = 0
    ImportErrFile = "LasalleCourse_Import_Errors_" & Format(Date, "yymmdd") & ".csv"
    ImportErrFile = xPath & ImportErrFile

    If (Dir(ImportErrFile)) <> "" Then Kill ImportErrFile
    Open ImportErrFile For Output As #2
    'Print header line
    buf = """Employee Number"",""First Name"",""Last_Name"",""Course Name"",""Course ID"",""Score"",""Completed Date"",""Expiry Date"",""Reason"""
    Print #2, buf

    
    '-- update Course table
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
        
    Open ImportFile For Input As #1
    xCNT = 0
    Do While Not EOF(1)
        MDIMain.panHelp(0).FloodPercent = (xCNT / xTotRec) * 100
        DoEvents
        Line Input #1, xdata
        'xEmpNo, xCourseID, xCourseName, xCompDate, xScore, xExpiryDate
        xEmpNo = CSVGet(xdata, 1)
        xCourseName = Trim(CSVGet(xdata, 4))
        xCourseID = CSVGet(xdata, 5)
        'xCompDate = CSVGet(xdata, 6)
        'xScore = CSVGet(xdata, 7)
        xScore = CSVGet(xdata, 6)
        xCompDate = CSVGet(xdata, 7)
        xExpiryDate = CSVGet(xdata, 8)
                
        If Not IsNumeric(xEmpNo) Then
            'update the error log file
            GoTo next_rec
        End If
        If Len(xCourseID) = 0 Then
            'update the error log file
            GoTo next_rec
        End If
        If Not IsDate(xCompDate) Then
            'update the error log file
            GoTo next_rec
        End If
        
        If Not IsDate(xCompDate) Then
            'update the error log file
            GoTo next_rec
        End If
        
        'Ticket #30505 Franks 09/06/2017 - begin
        xCompDate = DateConForLaSalle(xCompDate)
        If IsDate(xExpiryDate) Then
            xExpiryDate = DateConForLaSalle(xExpiryDate)
        End If
        'Ticket #30505 Franks 09/06/2017 - end
        
        'find the matching employee
        SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
        If rsEmp.State <> 0 Then rsEmp.Close
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If rsEmp.EOF Then
            'Ticket #28931 Franks 07/22/2016 - begin
            'check if this emp # can be found in term_hremp, if found then no error msg
            SQLQ = "SELECT * FROM TERM_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
            If rsTermEmp.State <> 0 Then rsTermEmp.Close
            rsTermEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTermEmp.EOF Then
                GoTo next_rec
            End If
            'Ticket #28931 Franks 07/22/2016 -  end
            buf = xdata & "," & "Cannot find this Employee Number"
            Print #2, buf
            xErrRec = xErrRec + 1
            GoTo next_rec
        End If
        If Not rsEmp.EOF Then
            'check if this xCourseID exist in Course Code Master
            SQLQ = "SELECT * from HR_COURSECODE_MASTER WHERE ES_CRSCODE = '" & xCourseID & "' "
            If rsCRSMaster.State <> 0 Then rsCRSMaster.Close
            rsCRSMaster.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If rsCRSMaster.EOF Then
                'update the error log file
                buf = xdata & "," & "Cannot find '" & xCourseID & "' in Course Code Master"
                Print #2, buf
                xErrRec = xErrRec + 1
                GoTo next_rec
            End If
            
            'check the if Course Name = Course Code Desc
            xTemp = GetTABLCodePub("ESCD", xCourseID)
            If Len(xTemp) = 0 Then
                ''update the error log file 'not found this code
                'buf = xdata & "," & "Cannot find '" & xCourseID & "' in Course Code Master"
                'Print #2, buf
                'xErrRec = xErrRec + 1
                GoTo next_rec
            End If
            'Ticket #28305 Franks 03/10/2016 - Do Not compare these, comment out the codes below
            ''If Not (xTemp = xCourseName) Then
            ''    '"   Course ID and Course Name must match the Course Code Master's Course Code and description.
            ''    'o   If a match is not found, copy the contents of the course to another CSV file. The last column of the CSV file would identify the reason for the reject
            ''    'update the error log file
            ''    buf = xdata & "," & "Course Name('" & xCourseName & "') not match info:HR Course Code Description('" & xTemp & "')"
            ''    Print #2, buf
            ''    xErrRec = xErrRec + 1
            ''    GoTo next_rec
            ''End If
            
            'check if Score is valid
            xTemp = GetTABLCodePub("ESRT", xScore)
            If Len(xTemp) = 0 Then '???
                ''update the error log file 'not found this code
                'buf = xdata & "," & "Invalid Score ('" & xScore & "') in info:HR Results Code list"
                'Print #2, buf
                'xErrRec = xErrRec + 1
                'GoTo next_rec
                'Ticket #27742 Franks 11/10/2015 - add this code automatically if not found in the Course Results List
                CheckHRTABLCode "ESRT", xScore, Left(Trim(xScore & " - Import"), 50)
            End If
            
            'update Course table
            SQLQ = "SELECT * FROM HREDSEM WHERE ES_EMPNBR = " & xEmpNo & " "
            SQLQ = SQLQ & "AND ES_CRSCODE = '" & xCourseID & "' "
            SQLQ = SQLQ & "AND ES_DATCOMP = " & Date_SQL(xCompDate) & " "
            If rsCRSEmp.State <> 0 Then rsCRSEmp.Close
            rsCRSEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsCRSEmp.EOF Then
                rsCRSEmp.AddNew
                rsCRSEmp("ES_EMPNBR") = xEmpNo
                rsCRSEmp("ES_CRSCODE") = xCourseID
                rsCRSEmp("ES_START") = CVDate(xCompDate)
                rsCRSEmp("ES_DATCOMP") = CVDate(xCompDate)
            End If
            rsCRSEmp("ES_RESULTS") = xScore
            If IsDate(xExpiryDate) Then
                rsCRSEmp("ES_RENEW") = CVDate(xExpiryDate)
            End If
            rsCRSEmp("ES_COURSE") = Left(xCourseName, 125) ' Left(xCourseName, 60) 'course name
            
            'update fields from Course Master - begin
            'rsCRSEmp("") = ""
            rsCRSEmp("ES_CTYPE") = rsCRSMaster("ES_CTYPE")  'Course Type
            rsCRSEmp("ES_COORDINATED") = rsCRSMaster("ES_COORDINATED") 'Co-Ordinated By
            rsCRSEmp("ES_COMPANYNAME") = rsCRSMaster("ES_COMPANYNAME") '
            rsCRSEmp("ES_TRAINNER") = rsCRSMaster("ES_TRAINNER") 'txtTrainerName.Text 'Trainer Name
            rsCRSEmp("ES_HOURS") = rsCRSMaster("ES_HOURS") 'txtCourseHRS.Text 'Course Hours
            rsCRSEmp("ES_TBEMP") = rsCRSMaster("ES_TBEMP") 'medEECont(0).Text 'Employee $
            rsCRSEmp("ES_OTHER") = rsCRSMaster("ES_OTHER") 'medEECont(2).Text 'Other Expenses $
            rsCRSEmp("ES_TBCO") = rsCRSMaster("ES_TBCO") 'medEECont(1).Text 'Employer $
            rsCRSEmp("ES_ACCOM") = rsCRSMaster("ES_ACCOM") 'medEECont(3).Text 'Accommodation $
            rsCRSEmp("ES_LEARNING") = rsCRSMaster("ES_LEARNING") 'medEECont(4).Text 'Learning Material $
            rsCRSEmp("ES_EMPCUR") = rsCRSMaster("ES_EMPCUR") 'clpEmpCur.Text 'Currency
            rsCRSEmp("ES_OTCUR") = rsCRSMaster("ES_OTCUR") 'clpOherCur.Text 'Currency
            rsCRSEmp("ES_EMPLOYCUR") = rsCRSMaster("ES_EMPLOYCUR") 'clpEmployerCur.Text 'Currency
            rsCRSEmp("ES_ACOMCUR") = rsCRSMaster("ES_ACOMCUR") 'clpAcomCur.Text 'Currency
            rsCRSEmp("ES_LEARNINGCUR") = rsCRSMaster("ES_LEARNINGCUR") 'clpLearnCur.Text 'Currency
            rsCRSEmp("ES_TOTCUR") = rsCRSMaster("ES_TOTCUR") 'clpTotCur.Text 'Currency
            rsCRSEmp("ES_COORDINATED") = rsCRSMaster("ES_COORDINATED") 'Coordinated By
            rsCRSEmp("ES_CEUTYPE") = rsCRSMaster("ES_CEUTYPE") 'CEU Type
            rsCRSEmp("ES_METHODUSED") = rsCRSMaster("ES_METHODUSED") 'Method Used
            'update fields from Course Master - end
            
            rsCRSEmp("ES_LDATE") = Date
            rsCRSEmp("ES_LTIME") = Time$
            rsCRSEmp("ES_LUSER") = "SYSTEM247" 'glbUserID
            rsCRSEmp.Update
        End If

next_rec:
        xCNT = xCNT + 1
    Loop
    Close #1
    Close #2
    MDIMain.panHelp(0).FloodPercent = 100
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "

End Sub

Private Sub WFCDateRange() 'Ticket #27515 Franks 09/08/2015
'vbSunday 1
'vbMonday 2
'vbTuesday 3
'vbWednesday 4
'vbThursday 5
'vbFriday 6
'vbSaturday 7

    '"   If today is Monday, From and To Date = previous Friday.
    If Weekday(Date) = 2 Then
        dlpDate(0).Text = DateAdd("d", -3, Date)
        dlpDate(1).Text = DateAdd("d", -3, Date)
    End If
    '"   If today is Tuesday, Wednesday, Thursday or Friday, Saturday, From and To Date = today -1.
    If Weekday(Date) >= 3 And Weekday(Date) <= 7 Then
        dlpDate(0).Text = DateAdd("d", -1, Date)
        dlpDate(1).Text = DateAdd("d", -1, Date)
    End If
    '"   If today is Sunday, From Date = today -2 and To Date = today.
    If Weekday(Date) = 1 Then
        dlpDate(0).Text = DateAdd("d", -2, Date)
        dlpDate(1).Text = DateAdd("d", 0, Date)
    End If

End Sub

Private Function DateConForLaSalle(xDATE) 'Ticket #30505 Franks 09/06/2017
Dim retVal
    retVal = xDATE
    If IsDate(xDATE) Then
        retVal = Trim(xDATE)
        If Len(retVal) > 10 Then
            retVal = Left(retVal, 10)
            retVal = CVDate(retVal)
        End If
    End If
    DateConForLaSalle = retVal
End Function
