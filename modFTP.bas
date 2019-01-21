Attribute VB_Name = "modFTP"
Option Explicit
' how long we wait for FTP to finish, in seconds, before giving up
Private Const FTPTimeout = 1200 '600

Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As Any, lpProcessInformation As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const CREATE_ALWAYS = 2
Private Const TRUNCATE_EXISTING = 5
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&
Private Const STARTF_USESHOWWINDOW = 1
Private Const STILL_ACTIVE = 259
Private Const SW_HIDE = 0
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type
Public Enum FtpErrorEnum
    success = 0
    PSCPNotFound
    PSCPTimedOut
    PSCPError
End Enum

Type PayWebSetupData
    Host As String
    Username As String
    Password As String
End Type

Type CSVImportStatus
    ImportedOK As Long
    RecLengthErrors As Long
End Type

Global glbPayWebData As PayWebSetupData
Global IsPSCP_Initialized As Boolean 'Ticket #25088 Franks 02/18/2014

' upload a file to a given SFTP/SCP site.  returns Success if the file was uploaded OK, nonzero otherwise.
' if it returns PSCPError, the log from the FTP session is held in the INFO:HR path, under the file'
' 'SCPLOG.TXT'.
Public Function FTPSendFile(FTPAddress As String, Username As String, Password As String, LocalFilename As String, RemoteFilename As String) As FtpErrorEnum
    Dim CmdLine As String, OldPointer As Integer
    Dim OutputFilename As String, OutputFileHandle
    Dim Buf As String, NbrLines As Long
    
    OldPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    ' make sure the PSCP (securecp utility) file is found
    If Dir(glbIHRREPORTS & "PSCP.EXE") = "" Then
        FTPSendFile = PSCPNotFound
        Screen.MousePointer = OldPointer
        Exit Function
    End If
    
    CmdLine = glbIHRREPORTS & "PSCP -batch -C -q -pw " & Password & " """ & LocalFilename & """ " & Username & "@" & FTPAddress & ":" & RemoteFilename
    OutputFilename = glbIHRREPORTS & "SCPLOG.TXT"
    frmFTP.Show
    If ExecCmdWithRedirect(CmdLine, OutputFilename, FTPTimeout) = False Then
        Unload frmFTP
        FTPSendFile = PSCPTimedOut
        Screen.MousePointer = OldPointer
        Exit Function
    End If
    Unload frmFTP
    OutputFileHandle = FreeFile()
    Open OutputFilename For Input As #OutputFileHandle
    Do Until EOF(OutputFileHandle)
        Line Input #OutputFileHandle, Buf
        If Len(Buf) > 0 Then NbrLines = NbrLines + 1
    Loop
    Close #OutputFileHandle
    If NbrLines > 0 Then
        FTPSendFile = PSCPError
        Open OutputFilename For Append As #OutputFileHandle
        Buf = LocalFilename & vbCrLf & FTPAddress & ":" & RemoteFilename
        Write #OutputFileHandle, Buf
        Close #1
        Screen.MousePointer = OldPointer
        Exit Function
    End If
    FTPSendFile = success
    Screen.MousePointer = OldPointer
End Function

' download a file from a given SFTP/SCP site.  returns Success if the file was uploaded OK, nonzero otherwise.
' if it returns PSCPError, the log from the FTP session is held in the INFO:HR path, under the file'
' 'SCPLOG.TXT'.
Public Function FTPGetFile(FTPAddress As String, Username As String, Password As String, LocalFilename As String, RemoteFilename As String) As FtpErrorEnum
    Dim CmdLine As String, OldPointer As Integer
    Dim OutputFilename As String, OutputFileHandle
    Dim Buf As String, NbrLines As Long
    Dim IniReq As Boolean
    Dim msg As String
    OldPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    ' make sure the PSCP (securecp utility) file is found
    If Dir(glbIHRREPORTS & "PSCP.EXE") = "" Then
        FTPGetFile = PSCPNotFound
        Screen.MousePointer = OldPointer
        Exit Function
    End If
    
    IsPSCP_Initialized = True 'Ticket #25088 Franks 02/18/2014
    
    ''CmdLine = glbIHRREPORTS & "PSCP -batch -C -q -pw " & Password & " " & Username & "@" & FTPAddress & ":" & RemoteFilename & " """ & LocalFilename & """"
    'CmdLine = "C:\A\PSCP -pw udols-AscHb2S4gM woodbridge-uat@ftp01.workstreaminc.com:outgoing/WB_HireFeed_UAT_20130723.xml C:\A\WB_HireFeed_UAT_20130723.xml"
    'CmdLine = glbIHRREPORTS & "PSCP -pw udols-AscHb2S4gM " & Password & " " & Username & "@" & FTPAddress & ":" & RemoteFilename & " """ & LocalFilename & """"
    CmdLine = glbIHRREPORTS & "PSCP -batch -C -q -pw " & Password & " " & Username & "@" & FTPAddress & ":" & RemoteFilename & " """ & LocalFilename & """"

    OutputFilename = glbIHRREPORTS & "SCPLOG.TXT"
    frmFTP.Show
    If ExecCmdWithRedirect(CmdLine, OutputFilename, FTPTimeout) = False Then
        Unload frmFTP
        FTPGetFile = PSCPTimedOut
        Screen.MousePointer = OldPointer
        Exit Function
    End If
    Unload frmFTP
    IniReq = False
    OutputFileHandle = FreeFile()
    Open OutputFilename For Input As #OutputFileHandle
    Do Until EOF(OutputFileHandle)
        Line Input #OutputFileHandle, Buf
        If Len(Buf) > 0 Then NbrLines = NbrLines + 1
        If InStr(Buf, "fingerprint") <> 0 Then
            IniReq = True
        End If
        msg = msg & Buf & vbNewLine
    Loop
    Close #OutputFileHandle
'    MsgBox msg

    If IniReq Then
        msg = msg & vbNewLine
        msg = msg & "This computer must be initialized on " & FTPAddress & vbNewLine
        msg = msg & "Please do the followings to register:" & vbNewLine
        msg = msg & vbNewLine
        msg = msg & "1) Click on START - RUN menu found on Windows toolbar. " & vbNewLine
        msg = msg & "2) Enter this command to OPEN: " & glbIHRREPORTS & "PSCP -pw " & Password & " " & Username & "@" & FTPAddress & ":" & RemoteFilename & " """ & LocalFilename & """" & vbNewLine
        msg = msg & "3) When prompted, answer ""y"" for the question ""Store key in cache? (y/n)""""" & vbNewLine
        msg = msg & "4) Redo the import in info:HR Payweb after the process finished."
        OutputFileHandle = FreeFile()
        Open OutputFilename For Output As #OutputFileHandle
        Print #OutputFileHandle, msg
        Close #OutputFileHandle '
        IsPSCP_Initialized = False 'Ticket #25088 Franks 02/18/2014
    End If
    If NbrLines > 0 Then
        FTPGetFile = PSCPError
        Screen.MousePointer = OldPointer
        Exit Function
    End If
    FTPGetFile = success
    Screen.MousePointer = OldPointer
End Function

' start a process, redirecting output to Filename.  Returns true if it completed, false if it timed out.
Private Function ExecCmdWithRedirect(CmdLine As String, OutputFilename As String, Timeout As Single) As Boolean
    Dim proc As PROCESS_INFORMATION, ret As Long, bSuccess As Long
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES, hFile As Long
    Dim bytesread As Long, mybuff As String
    Dim I As Integer, StartTimer As Single
    Dim lpExitCode As Long
    
    mybuff = String(256, Chr$(65))
    ' set up the security attributes (inherit handle must be 1)
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    ' set up the start parameters for the process
    start.cb = Len(start)
    start.dwFlags = STARTF_USESTDHANDLES + STARTF_USESHOWWINDOW
    ' create/open the output file
    hFile = CreateFile(OutputFilename, GENERIC_WRITE, 0&, sa, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0&)
    If hFile < 0 Then
        MsgBox "CreateFile failed.  Error: " & Err.LastDllError
        Exit Function
    End If
    ' assign the output file to the process
    start.hStdOutput = hFile
    start.hStdError = hFile
    start.wShowWindow = SW_HIDE
    ' launch the process
    ret& = CreateProcessA(0&, CmdLine, sa, sa, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    If ret <> 1 Then
        MsgBox "CreateProcess failed. Error: " & Err.LastDllError
        Exit Function
    End If
    
    StartTimer = Timer
    Do
        GetExitCodeProcess proc.hProcess, lpExitCode
        DoEvents
    Loop While lpExitCode = STILL_ACTIVE And Timer - StartTimer < Timeout
    If lpExitCode = STILL_ACTIVE Then
        ' timed out
        ExecCmdWithRedirect = False
    Else
        ' completed
        ExecCmdWithRedirect = True
    End If
    ret& = CloseHandle(proc.hProcess)
    ret& = CloseHandle(proc.hThread)
    ret& = CloseHandle(hFile)
End Function

Public Function ReadFTPData(Optional CurrentFTP As Boolean)  '(CCode)
Dim SQLQ As String
    Dim rsFTPSetup As New ADODB.Recordset
    ' get FTP data from the PayWeb Setup table
    'rsFTPSetup.Open "SELECT HOST, USERNAME, PASSWORD FROM PAYWEB_SETUP WHERE CLIENTCODE='" & CCode & "'", glbDBConnections("IHRPWeb")
    'rsFTPSetup.Open "SELECT HOST, USERNAME, PASSWORD FROM HRSF_FTP_SETUP ", gdbAdoIhr001, adOpenStatic
    SQLQ = "SELECT HOST, USERNAME, PASSWORD FROM HRSF_FTP_SETUP "
    If Not IsMissing(CurrentFTP) Then
        If CurrentFTP Then
            SQLQ = SQLQ & "WHERE NOT (CURRENT_FTP = 0) "
        End If
    End If
    rsFTPSetup.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If rsFTPSetup.EOF Then
        glbPayWebData.Host = ""
        glbPayWebData.Username = ""
        glbPayWebData.Password = ""
        ReadFTPData = False
    Else
        glbPayWebData.Host = rsFTPSetup("HOST")
        glbPayWebData.Username = rsFTPSetup("USERNAME")
        glbPayWebData.Password = rsFTPSetup("PASSWORD")
        ReadFTPData = True
    End If
    rsFTPSetup.Close
End Function

