VERSION 5.00
Begin VB.Form frmSPLASH 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2130
   ClientLeft      =   555
   ClientTop       =   1035
   ClientWidth     =   4170
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
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
   Icon            =   "FXsplash.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2130
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picSplash 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H80000008&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   240
      Picture         =   "FXsplash.frx":1CCA
      ScaleHeight     =   1215
      ScaleWidth      =   3255
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox picHRBlack 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   360
      Picture         =   "FXsplash.frx":46B7
      ScaleHeight     =   945
      ScaleWidth      =   1305
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   7560
      Top             =   4560
   End
   Begin VB.PictureBox picEllipse 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   -15
      Picture         =   "FXsplash.frx":80DD
      ScaleHeight     =   1725
      ScaleMode       =   0  'User
      ScaleWidth      =   4215
      TabIndex        =   5
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label lblDemoExpiry 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "System Expires on Month Day, Year"
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   105
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   3960
   End
   Begin VB.Label lblInfoHR 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "INFO:HR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label lblHRname 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HR SYSTEMS STRATEGIES INC."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   300
      Left            =   2040
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   4110
   End
   Begin VB.Shape shpCoverUp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   2115
      Left            =   -60
      Top             =   4380
      Width           =   9615
   End
End
Attribute VB_Name = "frmSPLASH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ht As String
Dim NumInter As Integer    'number of intervals
Dim SIncr As Integer    ' number for increment/decriment
Dim MaxSize As Integer  'max expansion size for frmSPLASH

Private Sub Form_Load()
    Dim strT As String
    Dim ValidLicense As Boolean
    
    'Ticket #17626 - United Way of the Lower Mainland
    'Ticket #18286 - Extend the Expiry to April 10 2010.
    'Begin
'    Dim Newdate               'See FXDEMO.FRM
'    Newdate = CVDate("10 Apr. 2010")
'    'MsgBox Newdate
'    If Date >= Newdate Then
'        'MsgBox "info:HR system has expired!"
'        MsgBox "info:HR System expired on April 10th, 2010." & vbCrLf & vbCrLf & "Please call info:HR Support at 1-800-567-4254 for more information.", vbOKOnly, "info:HR System Expired!"
'        End
'    End If
    'End
    
    ' dkostka - 06/26/2001 - Added option to assist in debugging client-side problems
    If UCase(Command) = "/DEBUG" Then MsgBox "Calling Demo_Setup...", , "Debug Information"
    Call Demo_Setup
    ' dkostka - 06/26/2001 - Added option to assist in debugging client-side problems
    If UCase(Command) = "/DEBUG" Then MsgBox "Setting up timer...", , "Debug Information"
    
    'Check if the info:HR System Expires for the Hosted Solution
    If glbHosted Then
        'Get the License Key
        strT = ""
        strT = INIRead("Options", "LicenseKey", glbHostFile)
        If Len(strT) > 0 Then
            'Call function to check the validity of the License
            ValidLicense = Check_License_Key(strT)
            
            'If not valid license then end the program
            If Not ValidLicense Then End
        Else
            'Blank license key not allowed. For unlimited info:HR license Jerry wants to have “Dec. 31, 2028” date
            MsgBox "Invalid License Key." & vbCrLf & "Please call info:HR Support for assistance.", vbOKOnly, "Invalid License Key"
            End
        End If
    End If
    
    Timer1.Interval = 100  ' Set interval.
   ' Timer1.Interval = 1 'for test jaddy
    NumInter = 0   'number of intervals
    SIncr = 1000 'increment/decriment for screen splays window
    MaxSize = 2500
    'Picture1.BackColor = 12632256

    ' Application starts here (Load event of Startup form).

End Sub

Private Sub Timer1_Timer()
    
    ' MaxSize set in form load, reset in second run through
    ' Dpass and Dcheck set in form load for reviewing number passes
    ' window gets bigger, then when shrinks displays logo,
    ' then gets real big and loads next form in grey
    NumInter = NumInter + 1

    Select Case NumInter
        Case 1
            If UCase(Command) = "/DS" Then
                'User wants the Login screen for Data Source setup
                lblDemoExpiry.Visible = True
                lblDemoExpiry.Caption = "Loading Login screen for Data Source setup..."
            
                picHRBlack.Visible = False
                lblHRname.Visible = False
                'Unload frmSPLASH
                Load frmSECURITY
                Timer1.Interval = 0
                Exit Sub
            Else
                'Check if SSO Enabled
                glbSSO = ""
                glbSSO = get_Single_Sign_On_Setting
                
                If glbSSO = "" Then
                    'Timer1.Interval = 0
                    'End
                End If
                
                If glbSSO = "YES" Then
                    'Single Sign On Enabled
                    'Get the Windows User ID
                    glbUserID = GetCurrentWinUser
                    
                    'Check if valid Windows User ID in info:HR, get Password
                    glbSSOPwd = ""
                    glbSSOPwd = GetPassword_CurrentWindows_User(glbUserID)
                    
                    'If valid User ID do not show the login screen
                    If Len(glbSSOPwd) > 0 Then
                        'OK click of frmSecurity form
                        Call Load_infoHR_Directly
                        
                        'Timer1.Interval = 0
                        Exit Sub
                    Else
                        'Ticket #22682 - Release 8.0 bug fix. If the Windows User ID is not found in info:HR's
                        'User IDs then we are prompting Login screen. Since 'glbUserID' still have Windows User ID,
                        'the first time the user tries to log in - it says 'Security record not found' as it is
                        'using Windows User ID in 'glbUserID' and password entered on the Login screen to check for the
                        'validity of info:HR User. So I am clearing the value in 'glbUserID' since this user was not
                        'found in info:HR. At the later stage if 'glbUserID' is blank we are assigning the value (User ID)
                        'entered on the Login screen which gets matched with the Password entered on the Login screen and
                        'that fixes this issue.
                        glbUserID = ""
                    End If
                Else
                    'Let NumInter increment to show the login screen
                End If
            End If
        Case 3
            If InStr(UCase(Command), "/FAST") > 0 Then
                picHRBlack.Visible = False
                lblHRname.Visible = False
                'Unload frmSPLASH
                frmSPLASH.Hide
                Load frmSECURITY
                'frmSECURITY.Show
                'frmSPLASH.Hide
                Timer1.Interval = 0
                Exit Sub
            End If
        'Case 10
            'Hemu = Commented
            'picSplash.Visible = True
        'Case 25
            'Hemu - Commented
            'lblInfoHR.Visible = True
        Case 10
            picHRBlack.Visible = False
            lblHRname.Visible = False
            ' dkostka - 06/26/2001 - Added option to assist in debugging client-side problems
            If UCase(Command) = "/DEBUG" Then MsgBox "Unloading myself...", , "Debug Information"
            'Unload frmSPLASH
            frmSPLASH.Hide
            ' dkostka - 06/26/2001 - Added option to assist in debugging client-side problems
            If UCase(Command) = "/DEBUG" Then MsgBox "Loading frmSECURITY...", , "Debug Information"
            Load frmSECURITY
            'frmSECURITY.Show
            'frmSPLASH.Hide
            ' dkostka - 06/26/2001 - Added option to assist in debugging client-side problems
            If UCase(Command) = "/DEBUG" Then MsgBox "Stopping time...", , "Debug Information"
            Timer1.Interval = 0
            Exit Sub
    End Select  '  stop and load the frmSECURITY form.

End Sub

Public Sub Demo_Setup()
    DemoSystem = False
    DemoMaxEmp% = 0
    
    'Ticket #17626 - United Way of the Lower Mainland
    'Ticket #18286 - Extend the Expiry to April 10 2010.
'    DemoSystem = True
'    DemoMaxEmp% = 100
'    lblDemoExpiry.Visible = True
'    lblDemoExpiry.Caption = "System Expires on April 10th, 2010"
End Sub

Private Function get_Single_Sign_On_Setting()
    'Open the database
    If Not gdbIhr001_Opn() Then ' can we open the database?
        'MsgBox "Check which database you are trying to access."
        
        Screen.MousePointer = DEFAULT
        
        get_Single_Sign_On_Setting = ""
        
        Exit Function
    Else
        
        Dim rsPrefer As New ADODB.Recordset
        rsPrefer.Open "SELECT * FROM HRPREFERENCE WHERE HP_FUN_NAME = 'SSO_INFOHR'", gdbAdoIhr001, adOpenStatic
        If Not rsPrefer.EOF Then
            If rsPrefer("HP_ENABLED") Then
                get_Single_Sign_On_Setting = "YES"
            Else
                get_Single_Sign_On_Setting = "NO"
            End If
        Else
            get_Single_Sign_On_Setting = "NO"
        End If
        rsPrefer.Close
        Set rsPrefer = Nothing
        
        'Ticket #24352 - PIPEDA - At the same time get if Encryption is turned-ON
        gsDB_CONNECT_ENCRYPT = Get_CompanyPreference_Value("DB_CONNNECT_ENCRYPT")
    End If
End Function

Private Function GetPassword_CurrentWindows_User(xCurWinUser)
    Dim Secure_Snap As New ADODB.Recordset
    Dim xPassword As String
    Dim SQLQ As String
    
    GetPassword_CurrentWindows_User = ""
    
    'Retrieve User record with Password
    SQLQ = "Select * from HR_SECURE_BASIC"
    SQLQ = SQLQ & " where USERID = '" & Replace(xCurWinUser, "'", "''") & "'"
    Secure_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not Secure_Snap.EOF Then
        xPassword = Secure_Snap("PassWord")
        If gsMultiLang = "YES" Then
            GetPassword_CurrentWindows_User = DecryptPasswordMultiLang(xPassword)
        Else
            GetPassword_CurrentWindows_User = DecryptPassword(xPassword)
        End If
    Else
        GetPassword_CurrentWindows_User = ""
    End If
    Secure_Snap.Close
    Set Secure_Snap = Nothing
    
End Function

Private Sub Load_infoHR_Directly()
Dim EEID As Variant, tries As Integer
Dim Msg As String, SQLQ
Dim PARCO_Snap As New ADODB.Recordset
Dim a%
Dim strLockedMsg As String
Dim txtPWord As String
Dim rsAT_Multi As New ADODB.Recordset

Screen.MousePointer = HOURGLASS

Call setPreference

SQLQ = "SELECT PC_SERIAL FROM HRPARCO WHERE PC_CO = '001'"
PARCO_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
glbCompSerial = PARCO_Snap("PC_SERIAL")
glbLinamar = glbCompSerial = "S/N - 2309W"
PARCO_Snap.Close

glbWSIBModule = isWSIBModule

'WFC 7.2
If glbSQL And glbCompSerial = "S/N - 2282W" Then 'Mississauga
    glbWFC = True
    If gdbAdoIhrWFC.State = adStateOpen Then gdbAdoIhrWFC.Close
    gdbAdoIhrWFC.Mode = adModeReadWrite
    gdbAdoIhrWFC.Open glbAdoIHRDB
End If

'Ticket #29206 - Added this piece that Frank added on Security form (Login screen) OK click. Clients using the SSO
'will not go on to the Login screen and therefore the glbCRWPrintSetup was remaining as False. This was then causing
'as issue as users will not see hte Printer Setup button. So I am adding this piece here as well so the right value is
'assigned to glbCRWPrintSetup.
If glbWFC Then 'Ticket #25436 Franks 05/06/2014 -
    glbCRWPrintSetup = False
Else
    glbCRWPrintSetup = True
End If

'Ticket #18491 - Burlington Technologies Inc.
'Logic to increase Employee License from 400 to 600 for 2 months and then set it back to 400
If glbCompSerial = "S/N - 2351W" Then
    SQLQ = "SELECT PC_MAX_EMPLOYEES FROM HRPARCO WHERE PC_CO = '001'"
    PARCO_Snap.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If PARCO_Snap("PC_MAX_EMPLOYEES") = 400 And Date < CVDate("15 Jul. 2010") Then
        PARCO_Snap("PC_MAX_EMPLOYEES") = 600
        PARCO_Snap.Update
        MsgBox "The Employee License has been increased to 600 and it will reset back to 400 on July 15th 2010." & vbCrLf & vbCrLf & "For more information please call info:HR Support at 1-800-567-4254.", vbOKOnly, "info:HR: Employee License Extension"
    ElseIf PARCO_Snap("PC_MAX_EMPLOYEES") = 600 And Date >= CVDate("15 Jul. 2010") Then
        PARCO_Snap("PC_MAX_EMPLOYEES") = 400
        PARCO_Snap.Update
        MsgBox "The Employee License extension has expired and has been reset back to 400 from 600." & vbCrLf & vbCrLf & "For more information please call info:HR Support at 1-800-567-4254.", vbOKOnly, "info:HR: Employee License Extension Expired!"
    End If
    PARCO_Snap.Close
    Set PARCO_Snap = Nothing
End If

If Len(glbSSOPwd) > 0 Then
    txtPWord = glbSSOPwd
End If

If modSecurity_Check(glbUserID, txtPWord) Then
    '~~~~~~~~~~~~~~~ADDED BY RAUBREY 4/11/97~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    glbNDepts = Dept_Secure() ' set up Department frmSECURITY matrix
    If glbNDepts = 0 Then
        Msg = "There are no Departments assigned to your User ID."
        Msg = Msg & Chr(10) & Chr(10) & "To correct this problem:"
        Msg = Msg & Chr(10) & "Have the System Administrator sign on"
        Msg = Msg & Chr(10) & "using their User ID and edit your"
        Msg = Msg & Chr(10) & "'Security' setup."
        MsgBox Msg, MB_OK + MB_ICONSTOP
        Screen.MousePointer = DEFAULT
        ' dkostka - 03/12/2002 - if the login fails remove all info from the security collections
        '   if we don't do this, if they re-try the login it errors out because the keys already exist.
        Do Until gSec_Upd_Master_Table.count = 0
          gSec_Upd_Master_Table.Remove 1
        Loop
        Do Until gSec_Inq_Master_Table.count = 0
          gSec_Inq_Master_Table.Remove 1
        Loop
        Exit Sub
    End If
    
    'If Department Security is "ALL" only, glbDeptAllRight = True
    'Ticket #12204
    glbDeptAllRight = True
    If InStr(1, glbSeleDeptUn, "ED_DEPTNO=") > 0 Then
        glbDeptAllRight = False
    End If
    If InStr(1, glbSeleDeptUn, "ED_ORG=") > 0 Then
        glbDeptAllRight = False
    End If
    If InStr(1, glbSeleDeptUn, "ED_DIV=") > 0 Then
        glbDeptAllRight = False
    End If
    If InStr(1, glbSeleDeptUn, "ED_SECTION=") > 0 Then
        glbDeptAllRight = False
    End If
    'Ticket #18235
    If InStr(1, glbSeleDeptUn, "ED_ADMINBY=") > 0 Then
        glbDeptAllRight = False
    End If
    
    If glbSQL And glbCompSerial = "S/N - 2282W" Then 'Mississauga
        'Dim rsAT_Multi As New ADODB.Recordset
        If InStr(1, glbSeleSection, "TB_KEY=") > 0 Then
            Call SetWFCSerialNo
        Else
            glbWFCFullRights = True 'No Plant in glbSeleSection
        End If
        If glbAdv Then 'etpath.dat exists
            glbAdv = False
            If Len(glbPlantCode) > 0 Then
                SQLQ = "SELECT PARA_TYPE FROM APPLICATION_PARAMETER WHERE PARA_TYPE = 'Integration' AND PARA_CATEGORY = 'Advanced Tracker' "
                SQLQ = SQLQ & "AND PARA_CATEGORY2='Integration Selection' AND PARA_NAME='Section' "
                SQLQ = SQLQ & "AND PARA_VALUE = '" & glbPlantCode & "' "
                rsAT_Multi.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsAT_Multi.EOF Then  '"GREN"
                    glbAdv = True
                End If
            End If
            If glbPlantCode = "GREN" Then
                glbAdv = True
            End If
        End If
        glbSMTPServerIP = "mail.woodbridgegroup.com"
        'glbWFCTermEmail = "termnotice@woodbridgegroup.com"
        glbWFCTermEmail = "termnotice@woodbridgegroup.com;InfoHR_Term_Notice@woodbridgegroup.com" 'Ticket #14409
        Call WFCTermEmailInfo("termnotice")
    End If
    
    If glbCompSerial = "S/N - 2335W" Then 'Mitchell Plastics Ticket #20982 Franks 12/14/2011
        'Dim rsAT_Multi As New ADODB.Recordset
        If InStr(1, glbSeleSection, "TB_KEY=") > 0 Then
            'Call SetWFCSerialNo
            glbPlantCode = Replace(Trim(Mid(glbSeleSection, 12, 4)), "'", "")
        Else
            glbWFCFullRights = True 'No Plant in glbSeleSection
        End If
        If glbAdv Then 'etpath.dat exists
            glbAdv = False
            If Len(glbPlantCode) > 0 Then
                SQLQ = "SELECT PARA_TYPE FROM APPLICATION_PARAMETER WHERE PARA_TYPE = 'Integration' AND PARA_CATEGORY = 'Advanced Tracker' "
                SQLQ = SQLQ & "AND PARA_CATEGORY2='Integration Selection' AND PARA_NAME='Section' "
                SQLQ = SQLQ & "AND PARA_VALUE = '" & glbPlantCode & "' "
                rsAT_Multi.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsAT_Multi.EOF Then
                    glbAdv = True
                End If
            End If
        End If
    End If
    
    'Ticket #22342 Franks 10/23/2012
    If glbCompSerial = "S/N - 2443W" Then 'Walters Inc
        If InStr(1, glbSeleDiv, "DIV=") > 0 Then 'glbSeleDiv
            glbPlantCode = Replace(Trim(Mid(glbSeleDiv, 9, 4)), "'", "")
        End If
    End If
    
    If glbCompSerial = "S/N - 2351W" Then
        glbSMTPServerIP = "10.4.10.5" 'mail.burltech.com
    End If
    
'    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    tries = 0
    'panHelpEntry.Caption = "Getting Version Info..." 'Added by Bryan 11/07/05 Ticket #8855
    setCompInfo "001"
    'panHelpEntry.Caption = "Done." 'Added by Bryan 11/07/05 Ticket #8855
    
'    If gsSECURED_PSW Then 'Ticket #12707
'        Dim snapSec As New ADODB.Recordset
'        Dim xInt As Long
'
'        SQLQ = "SELECT * FROM HR_SECURE_BASIC "
'        SQLQ = SQLQ & "Where (USERID = '" & glbUserID & "')"
'        snapSec.Open SQLQ, gdbAdoIhr001, adOpenStatic
'        If Not snapSec.EOF Then
'            If IsDate(snapSec("PS_EXPIR_DATE")) Then
'                glbTempFlag = False
'                If Len(txtPWord.Text) < 8 Then
'                        Screen.MousePointer = DEFAULT
'                        Msg = "You current Password less than 8 characters. "
'                        Msg = Msg & Chr(10) & "Please click OK button to change the password."
'                        MsgBox Msg
'                        frmSPassCh.fdFrameName = "fraLogonPsw"
'                        Load frmSPassCh
'                        frmSPassCh.Show 1
'                        If Not glbTempFlag Then
'                            Exit Sub
'                        End If
'                Else
'                    xInt = DateDiff("D", Date, snapSec("PS_EXPIR_DATE"))
'                    If xInt <= 0 Then 'Expired already
'                        Screen.MousePointer = DEFAULT
'                        Msg = "The password expired on " & CVDate(snapSec("PS_EXPIR_DATE")) & " "
'                        Msg = Msg & Chr(10) & "Please click OK button to change the password."
'                        MsgBox Msg
'                        frmSPassCh.fdFrameName = "fraLogonPsw"
'                        Load frmSPassCh
'                        frmSPassCh.Show 1
'                        If Not glbTempFlag Then
'                            Exit Sub
'                        End If
'                    ElseIf xInt > 0 And xInt < 6 Then 'Give warning before 5 days expired
'                        Msg = "The password will expire on " & CVDate(snapSec("PS_EXPIR_DATE")) & " "
'                        Msg = Msg & Chr(10) & "Do you want to change the password? "
'                        a% = MsgBox(Msg, 36, "Confirm")
'                        If a% = 6 Then
'                            Screen.MousePointer = DEFAULT
'                            frmSPassCh.fdFrameName = "fraLogonPsw"
'                            Load frmSPassCh
'                            frmSPassCh.Show 1
'                            If Not glbTempFlag Then
'                                Exit Sub
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        End If
'        snapSec.Close
'    End If
    
    If glbCompSerial = "S/N - 2242W" Then 'CCAC London #9014
        glbBatchNumber = ""
        Do While True
            glbBatchNumber = InputBox("Please enter Batch Number.", "London CCAC Batch Number", "")
            If Not (Len(glbBatchNumber) > 0 And IsNumeric(glbBatchNumber)) Then
                MsgBox "Batch Number should be number. "
            Else
                Exit Do
            End If
        Loop
    End If
    
    If glbCompSerial = "S/N - 2288W" Then 'Musashi - Ticket #12690
        'Get list of Union Group user does not have access to.
        'The value is stored in glbNoAccessGrp
        Call Get_No_Access_Group_List
    End If
    
    Unload frmSPLASH
    MDIMain.Show
    Screen.MousePointer = DEFAULT

'Else
'    txtLogonID = ""
'    txtPWord = ""
'    txtLogonID.SetFocus
'    MsgBox "Security record not found or password incorrect."
'
'    gblTries = gblTries + 1
'    If gblTries > 3 Then
'        If glbCompSerial = "S/N - 2407W" Then 'Ticket #18406 - Farmers' Mutual Insurance
'            strLockedMsg = "Your login has been locked. Please contact your System Administrator to unlock you login."
'            MsgBox "Yup - 3 strikes and you are Out!" & vbCrLf & vbCrLf & strLockedMsg
'
'            'Call procedure to lock the password
'            Call Lock_Password(glbUserID)
'        Else
'            MsgBox "Yup - 3 strikes and you are Out!"
'        End If
'
'        Unload frmSECURITY
'        Screen.MousePointer = DEFAULT
'        End
'    End If
End If

Screen.MousePointer = DEFAULT

End Sub

Private Function Check_License_Key(licenseKey As String) As Boolean
    Dim ExpiryDate As Date
    Dim UnlimitedDate As Date
    Dim newKey As String
    Dim X As Boolean
    
    Check_License_Key = True
    
    If Not ValidKey(licenseKey) Then
        newKey = InputBox("Invalid Licence Key found. Please enter the new key. If you do not have a new key, please call info:HR Support at 1-800-567-4254.", "Checksum Failed")
        
        If newKey = "" Then Check_License_Key = False ' Hit cancel
        
        If Not ValidKey(newKey) Then
            MsgBox "Invalid License Key entered." & vbCrLf & "Please call info:HR Support for assistance.", vbOKOnly, "Invalid License Key"
            Check_License_Key = False
        ElseIf DateDiff("d", DateAdd("d", 120, KeyToDate(newKey)), Now) > 0 Then
            MsgBox "License Key entered has already expired." & vbCrLf & "Please call info:HR Support for assistance.", vbOKOnly, "info:HR System Expired"
            Check_License_Key = False
        Else
            X = INIWrite("Options", "LicenseKey", newKey, glbHostFile)
            'WriteRegistrySetting HKEY_LOCAL_MACHINE, "Software\HR Systems\Options", "DemoData", newKey
            MsgBox "License Key accepted" & vbCrLf & "info:HR System will now expire on " & Format(DateAdd("d", 120, KeyToDate(newKey)), "mmmm d, yyyy")
            ExpiryDate = DateAdd("d", 120, KeyToDate(newKey))
        End If
    End If
    
    ExpiryDate = DateAdd("d", 120, KeyToDate(licenseKey))
    
    If DateDiff("d", ExpiryDate, Now) > 0 Then
        newKey = InputBox("info:HR System has expired." & vbCrLf & "Please call info:HR Support at 1-800-567-4254 if you require an extension, and enter the key supplied below.", "info:HR System Expired")
        
        If newKey = "" Then Check_License_Key = False ' Hit cancel
        
        If Not ValidKey(newKey) Then
            MsgBox "Invalid License Key entered." & vbCrLf & "Please call info:HR Support for assistance.", vbOKOnly, "Invalid License Key"
            Check_License_Key = False
        ElseIf DateDiff("d", DateAdd("d", 120, KeyToDate(newKey)), Now) > 0 Then
            MsgBox "License Key entered has already expired." & vbCrLf & "Please call info:HR Support for assistance.", vbOKOnly, "Invalid License Key"
            Check_License_Key = False
        Else
            X = INIWrite("Options", "LicenseKey", newKey, glbHostFile)
            'WriteRegistrySetting HKEY_LOCAL_MACHINE, "Software\HR Systems\Options", "DemoData", newKey
            MsgBox "License Key accepted." & vbCrLf & "info:HR System will now expire on " & Format(DateAdd("d", 120, KeyToDate(newKey)), "mmmm d, yyyy")
            ExpiryDate = DateAdd("d", 120, KeyToDate(newKey))
        End If
    End If
    
    'Do not display the expiry date if unlimited license Dec. 31, 2028.
    UnlimitedDate = Format("12/31/2028", "mm/dd/yyyy")
    If CVDate(ExpiryDate) <> CVDate(UnlimitedDate) Then
        lblDemoExpiry.Visible = True
        lblDemoExpiry.Caption = "info:HR System expires on " & Format(ExpiryDate, "mmm d, yyyy")
        glbLicenseKey = ExpiryDate
    Else
        glbLicenseKey = Format("12/31/2028", "mm/dd/yyyy")
    End If
    
End Function

Function ValidKey(LicKey As String) As Boolean
    If Len(LicKey) <> 6 Then Exit Function
    If Asc(Mid(LicKey, 1, 1)) <> Asc(Mid(LicKey, 2, 1)) - 2 Then Exit Function
    If Asc(Mid(LicKey, 3, 1)) <> Asc(Mid(LicKey, 4, 1)) - 2 Then Exit Function
    If Asc(Mid(LicKey, 5, 1)) <> Asc(Mid(LicKey, 6, 1)) - 2 Then Exit Function
    ValidKey = True
End Function

Function KeyToDate(LicKey As String)
    Dim Year As Integer, month As Byte, Day As Byte, MonthString As String

    Year = Asc(Mid(LicKey, 1, 1)) - 35
    'Ticket #19589 Frank 01/04/2011
    'the problem was: 01/04/2011 going to 01/04/2001
    'add 10 year for 201? year
    
    'Hemu - Commented the line below - for some reason it's not working anymore - it adds extra 10years
    'to client's license. I tested with 1/4/2011 and it worked fine - it did not change to 1/4/2001
    'Year = Year + 10
    
    Year = Int("2" & Right("000" & Year, 3))
    
    month = Asc(Mid(LicKey, 5, 1)) - 35
    Day = Asc(Mid(LicKey, 3, 1)) - 35
    ' Turn the month into words so we don't have to worry about date formats
    Select Case month
        Case 1
            MonthString = "January "
        Case 2
            MonthString = "February "
        Case 3
            MonthString = "March "
        Case 4
            MonthString = "April "
        Case 5
            MonthString = "May "
        Case 6
            MonthString = "June "
        Case 7
            MonthString = "July "
        Case 8
            MonthString = "August "
        Case 9
            MonthString = "September "
        Case 10
            MonthString = "October "
        Case 11
            MonthString = "November "
        Case 12
            MonthString = "December "
        Case Else
            MsgBox "Tampering with license data detected. info:HR will now exit.  Please call HR Systems Support for assistance.", vbExclamation + vbOKOnly, "Checksum Failed"
            End
    End Select
    KeyToDate = CVDate(MonthString & Day & ", " & Year)
    
End Function


