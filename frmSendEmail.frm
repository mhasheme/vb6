VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSendEmail 
   Caption         =   "                                "
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Left            =   7920
      Top             =   2160
   End
   Begin Threed.SSPanel sspMain 
      Align           =   1  'Align Top
      Height          =   6315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6645
      _Version        =   65536
      _ExtentX        =   11721
      _ExtentY        =   11139
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
      Begin VB.CommandButton cmdAddress 
         Caption         =   "Address Book"
         Height          =   315
         Left            =   2160
         TabIndex        =   12
         Top             =   60
         Width           =   1395
      End
      Begin VB.TextBox txtFrom 
         Enabled         =   0   'False
         Height          =   315
         Left            =   900
         TabIndex        =   1
         Top             =   600
         Width           =   5655
      End
      Begin VB.TextBox txtCC 
         Height          =   315
         Left            =   900
         TabIndex        =   3
         Top             =   1320
         Width           =   5655
      End
      Begin MSWinsockLib.Winsock wskMain 
         Left            =   8040
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   315
         Left            =   1140
         TabIndex        =   9
         Top             =   60
         Width           =   915
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   915
      End
      Begin VB.TextBox txtBody 
         Height          =   3795
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2160
         Width           =   6495
      End
      Begin VB.TextBox txtSubject 
         Height          =   315
         Left            =   900
         TabIndex        =   4
         Top             =   1680
         Width           =   5655
      End
      Begin VB.TextBox txtTo 
         Height          =   315
         Left            =   900
         TabIndex        =   2
         Top             =   960
         Width           =   5655
      End
      Begin VB.Label lblFrom 
         Caption         =   "From:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   660
         Width           =   855
      End
      Begin VB.Label lblCC 
         Caption         =   "CC..."
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1380
         Width           =   315
      End
      Begin VB.Line linSep1 
         X1              =   0
         X2              =   6660
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lblSubject 
         Caption         =   "Subject:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label lblTo 
         Caption         =   "To..."
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   315
      End
      Begin VB.Line linSep2 
         X1              =   0
         X2              =   6660
         Y1              =   2040
         Y2              =   2040
      End
   End
End
Attribute VB_Name = "frmSendEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Buf As String           ' Buffer to hold incoming data before processing
Dim Logging As Boolean      ' Flag to indicate if we're logging or not
Dim EmailAddress As String  ' User's Email Address
Dim SMTPServer As String    ' SMTP Server
Dim SMTPUsername As String  ' User's SMTP Username - NOT CURRENTLY USED
Dim SMTPPassword As String  ' User's SMTP Password - NOT CURRENTLY USED
Dim SMTPPort As Integer
Dim State As String         ' What we're waiting for (so we can tell the user if we timeout)

' Timeouts (in seconds)
Private Const ConnectTimeout = 45
Private Const CommandTimeout = 45

Private Sub SendData(DataToSend As String)
    wskMain.SendData DataToSend
    Log DataToSend
End Sub

Private Sub Log(Text As String)
    'Ticket #24629 - SMTP Log is now maintained in a table. So commenting the following out
'    'Ticket #24583 Franks 11/08/2013, add Linamar to here
'    If glbCompSerial = "S/N - 2382W" Or glbLinamar Then 'Ticket #19871 Franks 02/16/2011
'        Exit Sub
'    End If
'    If Logging = True Then Print #1, Text;
    
    'Ticket #24629 - Use Table instead to log
    Call SMTP_Log("Email", Replace(Text, vbCrLf, ""))
End Sub

Private Sub FillEmailInfo()
    Dim rsEmail As New adodb.Recordset
    
    If (glbWFC Or glbBurlTech) And Not glbWFCEmailTest Then  'And MDIMain.mnu_File_EmailSetup.Visible = False Then lost condition afther menu items removing, should check
        If glbWFC Then
            ' Don't read settings from HR_EMAIL, use hardcoded settings.
            SMTPServer = glbSMTPServerIP
            'EmailAddress = frmETERM.GetEmpData(glbEmpNbr, "ED_EMAIL")
            rsEmail.Open "SELECT * FROM HR_EMAIL WHERE EM_USERID='" & Replace(glbUserID, "'", "''") & "'", gdbAdoIhr001
            If Not rsEmail.EOF Then
                EmailAddress = rsEmail("EM_ADDRESS")
            End If
            rsEmail.Close
            If EmailAddress = "" Then
                MsgBox "Error retrieving email address, aborting email send.", vbCritical + vbOKOnly, "Error Building Address"
                Unload Me
            End If
        End If
        If glbBurlTech Then
            SMTPServer = glbSMTPServerIP
            rsEmail.Open "SELECT * FROM HR_EMAIL WHERE EM_USERID='" & Replace(glbUserID, "'", "''") & "'", gdbAdoIhr001
            If Not rsEmail.EOF Then
                EmailAddress = rsEmail("EM_ADDRESS")
            End If
            rsEmail.Close
            If EmailAddress = "" Then
                EmailAddress = "ses@burltech.com"
            End If
            'If EmailAddress = "" Then
            '    MsgBox "Error retrieving email address, aborting email send.", vbCritical + vbOKOnly, "Error Building Address"
            '    Unload Me
            'End If
        End If
    Else
        rsEmail.Open "SELECT * FROM HR_EMAIL WHERE EM_USERID='" & Replace(glbUserID, "'", "''") & "'", gdbAdoIhr001
        If rsEmail.EOF Then
            MsgBox "You have not been set up for email sending.  Please use the Setup->Security->Email Setup menu option to set up your account for emailing.", vbCritical + vbOKOnly, "No Email Setup Found"
            rsEmail.Close
        Else
            EmailAddress = rsEmail("EM_ADDRESS")
            If (glbBurlTech) And glbWFCEmailTest Then
                SMTPServer = glbSMTPServerIP
            Else
                SMTPServer = rsEmail("EM_SERVER")
            End If
            SMTPUsername = Format(rsEmail("EM_USERNAME"), "@")
            SMTPPassword = Format(rsEmail("EM_PASSWORD"), "@")
            If Not IsNull(rsEmail("EM_PORT")) Then
                SMTPPort = rsEmail("EM_PORT")
            End If
            rsEmail.Close
        End If
    End If
End Sub

Private Function GetWFCEmailAddress(EmpNbr) As String
    Dim rsEMP As New adodb.Recordset
    Dim FName As String
        
    rsEMP.Open "SELECT ED_FNAME, ED_SURNAME FROM HREMP WHERE ED_EMPNBR=" & EmpNbr, gdbAdoIhr001
    If rsEMP.EOF Then
        GetWFCEmailAddress = ""
        rsEMP.Close
        Exit Function
    End If
    
    FName = IIf(InStr(rsEMP("ED_FNAME"), " ") > 0, Left(rsEMP("ED_FNAME"), InStr(rsEMP("ED_FNAME"), " ") - 1), rsEMP("ED_FNAME"))
    GetWFCEmailAddress = LCase(FName & "_" & rsEMP("ED_SURNAME") & "@woodbridgegroup.com")
    rsEMP.Close
End Function

Private Sub cmdAddress_Click()
    frmOutLookAddress.Show 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub StartSMTPLog()
    'Ticket #24629 - SMTP Log is now maintained in a table. So commenting the following out.
'    'Ticket #24583 Franks 11/08/2013, add Linamar to here
'    If glbCompSerial = "S/N - 2382W" Or glbLinamar Then 'Samuel Ticket #19871 Franks 02/16/2011
'        Exit Sub
'    End If
'    Logging = True
'    Open App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "SMTPLOG.TXT" For Output As #1
'    'MsgBox "Now logging SMTP transactions to " & App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "SMTPLOG.TXT"
End Sub


Public Sub cmdSend_Click()

    
    
    Screen.MousePointer = vbHourglass
     'Ticket #21468 BACI, Ticket #25846 Pacific Sands, Ticket #30372 Peterborough FHT
    If glbCompSerial = "S/N - 2431W" Or glbCompSerial = "S/N - 2442W" Or glbCompSerial = "S/N - 2484W" Then
    'If glbCompSerial = "S/N - 2431W" Or glbCompSerial = "S/N - 9999W" Then '9999 for test
        SendMail
        Exit Sub
    End If
    
    ''Ticket #24475 Franks 10/22/2013
    'If glbSamuel Then
    '    If Left(txtSubject.Text, 26) = "info:HR Termination Notice" Then
    '    'If UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then
    '        Call SendMail(True)
    '        Exit Sub
    '    End If
    'End If
    
    'Ticket #18307
    wskMain.Close   'Since the email sending window is not being closed, the user can resend the email.
                    'If the wskmain is not closed it gives an error. So I am closing everytime Send is clicked.
    
    ' dkoala - 04/09/2001 - Always log SMTP transactions.
    StartSMTPLog
    With wskMain
        .RemoteHost = SMTPServer
        .RemotePort = 25
        ' Set up timeout timer
        Log "Connecting to " & .RemoteHost & " port " & .RemotePort & vbCrLf
        State = "Connecting to SMTP server"
        tmrTimeout.Interval = CLng(ConnectTimeout) * CLng(1000)
        tmrTimeout.Enabled = True
        .Connect
    End With
    
    ' Actual sending happens in wskMain_Connect procedure.
End Sub

Private Sub Form_Load()
    ' Fill the form-scope variables with email sending info from HR_EMAIL
    Call FillEmailInfo
    txtFrom = EmailAddress
End Sub

Private Sub Form_Resize()
    ' Resize all the controls to match the form's new size
    txtBody.Width = Me.Width - 255
    If IIf(Me.Height < 2500, 2500, Me.Height) - 2750 > 0 Then
        txtBody.Height = IIf(Me.Height < 2500, 2500, Me.Height) - 2750 '2340
    End If
    sspMain.Height = Me.Height
    linSep1.x2 = Me.Width
    linSep2.x2 = Me.Width
    txtTo.Width = Me.Width - 1110
    txtCC.Width = Me.Width - 1110
    txtSubject.Width = Me.Width - 1110
End Sub

Private Sub lblCC_Click()
    frmOutLookAddress.Show 1
End Sub

Private Sub lblTo_Click()
    frmOutLookAddress.Show 1
End Sub

Private Sub tmrTimeout_Timer()
    ' Timed out waiting for server, tell the user and abort
    tmrTimeout.Enabled = False
    Screen.MousePointer = vbDefault
    Log "ERROR - Timed out.  State: " & State & IIf(Buf <> "", "  Buf: " & Buf, "")
    
    'Ticket #24629 - SMTP Log is now maintained in a table. So commenting the following out.
'    If Logging Then Close #1: Logging = False

    MsgBox "Timed out waiting for SMTP server response." & vbCrLf & "State: " & State & vbCrLf & IIf(Buf <> "", "  Buf: " & Buf, "")
    If glbWFC = False Then
        Unload Me
    Else
        ' MC - dkostka - 05/03/2001 - Changed tag from "DONE" to "ERROR", so we know if it worked
        Me.Tag = "ERROR"
    End If
End Sub

Private Sub wskMain_Close()
'    wskMain.Close
End Sub

Private Sub wskMain_Connect()
    Dim i As Byte
    Dim MailFrom As String, MailTo As String, MailSubject As String
    Dim TopTo As Byte, TopCC As Byte
    Dim ToList(0 To 10) As String, CCList(0 To 10) As String
        
    ' Cancel the timeout timer, we're connected onw
    tmrTimeout.Enabled = False
    
    MailFrom = EmailAddress
    MailSubject = txtSubject.Text
    
    ' Parse the email address textboxes into arrays
    If txtTo.Text <> "" Then Call ParseEmailAddresses(ToList, TopTo, txtTo.Text)
    If txtCC.Text <> "" Then Call ParseEmailAddresses(CCList, TopCC, txtCC.Text)
    
    ' Send the email!
    With wskMain
        If WaitForStatus("2", "Waiting for hello from server") = False Then Exit Sub
        SendData "HELO infohr" & vbCrLf
        ' Say hello
        If WaitForStatus("2", "Sent HELO") = False Then Exit Sub
        ' Tell the server who we are
        SendData "MAIL From: " & MailFrom & vbCrLf
        If WaitForStatus("2", "Sent MAIL FROM") = False Then Exit Sub
        ' Send To addresses
        If TopTo > 0 Then
            For i = 0 To TopTo - 1
                SendData "RCPT To: " & ToList(i) & vbCrLf
                If WaitForStatus("250", "Sent RCPT TO (To)") = False Then Exit Sub
            Next i
        End If
        ' Send CC addresses
        If TopCC > 0 Then
            For i = 0 To TopCC - 1
                SendData "RCPT To: " & CCList(i) & vbCrLf
                If WaitForStatus("250", "Sent RCPT TO (CC)") = False Then Exit Sub
            Next i
        End If
        ' Tell the mail server we're going to send the message now
        SendData "DATA" & vbCrLf
        If WaitForStatus("354", "Sent DATA") = False Then Exit Sub
        ' Send header
        SendData "From: " & MailFrom & vbCrLf
        SendData "To: " & IIf(txtTo.Text = "" And txtCC.Text = "", "(undisclosed recipients)", txtTo.Text) & vbCrLf
        If txtCC.Text <> "" Then SendData "CC: " & txtCC.Text & vbCrLf
        SendData "Subject: " & MailSubject & vbCrLf
        SendData vbCrLf
        ' Send message body
        SendData txtBody.Text & vbCrLf
        ' Tell the server that's all there is, there is no more
        SendData "." & vbCrLf
        If WaitForStatus("2", "Sent message body") = False Then Exit Sub
        ' Tell the server we're done
        SendData "QUIT" & vbCrLf
        If WaitForStatus("2", "Sent QUIT") = False Then Exit Sub
    End With
    Screen.MousePointer = vbDefault
    
    'Ticket #24629 - SMTP Log is now maintained in a table. So commenting the following out.
'    ' Close the log file if we're logging
'    If Logging Then Close #1: Logging = False
    
    If Not (glbWFC Or glbCompSerial = "S/N - 2382W") Then
        '2382W - Samuel
        MsgBox "Mail sent successfully.", vbInformation + vbOKOnly, "Mail Sent"
'        If glbCompSerial = "S/N - 2378W" Then
            Me.Tag = "DONE"
'        Else
'            Unload Me
'        End If
    Else
        Me.Tag = "DONE"
        wskMain.Close
        If glbBurlTech Then
            Unload Me
        End If
    End If
End Sub

Private Function SendMail(Optional xhtml As Boolean = False) As Boolean
    On Error Resume Next
    
    Dim iMsg As Object
    Dim iConf As Object
    
    Dim Flds As Variant
    Dim strbody As String
    
    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
    
    iConf.Load -1 ' CDO Source Defaults
    Set Flds = iConf.Fields
    With Flds
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = EmailAddress 'TxtUser.Text  '"infoHR1@gobaci.com" ' "hrsstest@gmail.com"
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPPassword '"infoHRpass" '"bill@work"
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer '"smtp.gmail.com"
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 20
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    
     'Ticket #21468 BACI
    If glbCompSerial = "S/N - 2431W" Then
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
    Else
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort '465
    End If
    .Update
    End With
    
    strbody = txtBody.Text
    
    
    With iMsg
        Set .Configuration = iConf
        .To = txtTo.Text
        .CC = txtCC.Text
        .BCC = ""
        .ReplyTo = EmailAddress
        ' Note: The reply address is not working if you use this Gmail example
        ' It will use your Gmail address automatic. But you can add this line
        ' to change the reply address .ReplyTo = "Reply@something.com"
        .From = EmailAddress
        .Subject = txtSubject.Text
        If Not xhtml Then
            .TextBody = strbody
        Else
            ''Ticket #21701 Franks 03/13/2013
            'If xhtml Then
                .HTMLBody = strbody '<font color=>
            'Else
            '    .TextBody = strbody
            'End If
        End If
        .Send
    End With
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MsgBox "Failed to Send Mail" & Err.Number & " - " & Err.Source & " - " & Err.Description
        Log "Failed to Send Mail" & Err.Number & " - " & Err.Source & " - " & Err.Description
        SendMail = False
        Me.Tag = "Failed"
        Exit Function
    End If
     
         Me.Tag = "DONE"
          Unload Me
     SendMail = True
End Function

Private Sub wskMain_DataArrival(ByVal bytesTotal As Long)
    Dim DataIn As String
    
    ' Only get data if the connection is still open.  Sometimes we get this event just as the connection
    '   closes, and we don't want to trigger an error in this case.
    If wskMain.State = 1 Or wskMain.State = 7 Then
        ' Add the data coming in to the buffer
        wskMain.GetData DataIn, vbString
        Buf = Buf & DataIn
        Log DataIn
    End If
End Sub

Private Function WaitForStatus(StatusCode As String, CommandState As String) As Boolean
    Dim Status As String
    ' Set the form scope variable State to let the timeout procedure know what we were doing
    '   if the process times out.
    State = CommandState
    ' Set the timer control up to interrupt if the process times out
    tmrTimeout.Interval = CLng(CommandTimeout) * CLng(1000)
    tmrTimeout.Enabled = True
    ' Wait for a code to come in
    Do While InStr(Buf, vbLf) = 0
        DoEvents
    Loop
    ' Disable the timeout timer, we have data now
    tmrTimeout.Enabled = False
    ' Store the status portion of the line we received
    Status = Left(Buf, Len(StatusCode))
    WaitForStatus = True
    ' Check if the status we got is the one we were expecting
    If Status <> StatusCode Then
        ' Uh-oh, we got something different.  Abort and tell the user what went wrong.
        Log "ERROR - Expecting " & StatusCode & " got " & Status & ".  State: " & State & "  Buf: " & Buf
        
        'Ticket #24629 - SMTP Log is now maintained in a table. So commenting the following out.
'        If Logging Then Close #1: Logging = False
        
        MsgBox "Error talking to SMTP server." & vbCrLf & "Was expecting Status Code " & StatusCode & ", got " & Status & " instead.  Aborting." & vbCrLf & "Buf: " & Buf, vbCritical + vbOKOnly, "Error Sending Mail"
        If wskMain.State = 1 Then
            SendData vbCrLf & "QUIT" & vbCrLf
            wskMain.Close
        End If
        WaitForStatus = False
        Screen.MousePointer = vbDefault
    End If
    ' Clear the buffer to prepare for more data
    Buf = ""
    If glbWFC = True Then
        Me.Tag = "DONE"
    Else
        If glbCompSerial = "S/N - 2378W" Then
            Me.Tag = "DONE"
        End If
    End If
End Function

' Turn a list of seperated email addresses (seperated with one or more commas, semicolons, or spaces) into
'   an array of addresses.
Public Sub ParseEmailAddresses(ByRef List() As String, ByRef Top As Byte, FullString As String)
    Dim ToString As String, StrPos As Integer, OneChar As String, Buf As String
    
    ToString = FullString
    StrPos = 1
    ' Loop through the string finding each email address seperated by spaces and adding it to the list
    Do
        OneChar = Mid(ToString, StrPos, 1)
        ' Take the one character we just got off of the input string
        ToString = Right(ToString, Len(ToString) - 1)
        ' Is the next character a seperator?
        If OneChar <> " " And OneChar <> "," And OneChar <> ";" Then
            ' No, add the new character to the list
            Buf = Buf & OneChar
        Else
            ' Yes, add the buffer to the list, and increment the 'top of list' pointer
            ' If they have two seperators in a row, this will be a blank string (people use
            '   comma-space).  In this case, don't write out the string, but DO clear the buffer.
            If Trim(Buf) <> "" Then
                List(Top) = Buf
                Top = Top + 1
            End If
            Buf = ""
        End If
    Loop While Len(ToString) > 0
    ' We will have one address left in the buffer here, as we only dump out to the array on seperators,
    '   and there's no seperator at the end of the string.  Dump the last address into the array.
    If Trim(Buf) <> "" Then
        List(Top) = Buf
        Top = Top + 1
    End If
    Buf = ""
End Sub

'Private Sub wskMain_SendComplete()
'    wskMain.Close
'End Sub
