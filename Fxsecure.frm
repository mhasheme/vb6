VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmSECURITY 
   Appearance      =   0  'Flat
   Caption         =   "info:HR Security"
   ClientHeight    =   5970
   ClientLeft      =   900
   ClientTop       =   1290
   ClientWidth     =   8445
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Fxsecure.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5970
   ScaleWidth      =   8445
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CMDIni 
      Left            =   8880
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   9
      Top             =   4935
      Width           =   8445
      _Version        =   65536
      _ExtentX        =   14896
      _ExtentY        =   952
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      BevelInner      =   1
      Alignment       =   1
      Begin VB.OptionButton optDefault 
         Caption         =   "Default Data Source"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5482
         TabIndex        =   7
         Top             =   480
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.OptionButton optinfoHR 
         Caption         =   "info:HR System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1162
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optVolunteer 
         Caption         =   "Volunteer System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3232
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog CMDialog1 
         Left            =   7380
         Top             =   465
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         FontSize        =   0
         MaxFileSize     =   256
      End
      Begin VB.TextBox txtLogonID 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1020
         TabIndex        =   0
         Top             =   120
         Width           =   1635
      End
      Begin VB.TextBox txtPWord 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3660
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdSOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Height          =   375
         Left            =   5565
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   75
         Width           =   855
      End
      Begin VB.CommandButton ctrlHelp 
         Appearance      =   0  'Flat
         Caption         =   "&Help"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7365
         TabIndex        =   4
         Top             =   75
         Width           =   855
      End
      Begin VB.CommandButton ctrlExit 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   6465
         TabIndex        =   3
         Top             =   75
         Width           =   855
      End
      Begin VB.Label lblEEID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   150
         Width           =   540
      End
      Begin VB.Label lblPWord 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   2820
         TabIndex        =   10
         Top             =   150
         Width           =   690
      End
   End
   Begin Threed.SSPanel panHelpEntry 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   5475
      Width           =   8445
      _Version        =   65536
      _ExtentX        =   14896
      _ExtentY        =   873
      _StockProps     =   15
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   1
      Begin Threed.SSPanel PanFName 
         Height          =   375
         Left            =   3930
         TabIndex        =   13
         Top             =   60
         Width           =   4485
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         Caption         =   "Path = D:\IHR\IHR001.mdb      "
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         FloodColor      =   8421504
         Alignment       =   1
      End
   End
   Begin Threed.SSPanel Panel3D1 
      Height          =   4320
      Left            =   337
      TabIndex        =   12
      Top             =   240
      Width           =   7770
      _Version        =   65536
      _ExtentX        =   13705
      _ExtentY        =   7620
      _StockProps     =   15
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   2
      Begin VB.PictureBox picEllipse 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   1720
         Picture         =   "Fxsecure.frx":1CCA
         ScaleHeight     =   1725
         ScaleMode       =   0  'User
         ScaleWidth      =   4215
         TabIndex        =   14
         Top             =   1080
         Width           =   4215
      End
      Begin Threed.SSPanel panHRIco 
         Height          =   2145
         Left            =   1510
         TabIndex        =   15
         Top             =   870
         Width           =   4650
         _Version        =   65536
         _ExtentX        =   8202
         _ExtentY        =   3784
         _StockProps     =   15
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         BevelInner      =   2
         Alignment       =   1
         Autosize        =   1
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLocdb 
         Caption         =   "&Location Of Data Bases"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu mnuIHRDB 
            Caption         =   "&INFO:HR Database "
         End
         Begin VB.Menu mnuBUPDB 
            Caption         =   "&Backup Database "
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCompDB 
            Caption         =   "&Temp. Compact Location"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuSep 
         Caption         =   ""
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCompact 
         Caption         =   "&Compact and Backup Database"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecover 
         Caption         =   "&Recover Backup Database"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "&Data Source"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnu_SysInfo 
         Caption         =   "&System Information"
      End
      Begin VB.Menu mnu_sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About &info:HR"
      End
      Begin VB.Menu mnuHowStart 
         Caption         =   "&How to Start info:HR"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmSECURITY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Const KEY_READ = &H20019
Const REG_SZ = 1
Const ERROR_MORE_DATA = 234
Dim gblTries As Integer  'global frmSECURITY tries As Integer
Dim gdbAdoIHRDS As New ADODB.Connection
Dim rsIHRDS As New ADODB.Recordset
Dim xAdoIHRDB, xSQLDriver

Private Function chkExportzip()
Dim UnzipEx
chkExportzip = False

UnzipEx = UCase(App.Path & "\export.zip")
If Dir(UnzipEx) = "" Then
   'MsgBox "FILE not Found :" & Chr(10) & "[" & ImportFile & "]", , "Payroll Import"
    Exit Function
End If

chkExportzip = True

End Function

Private Function Chng_Ini(Ini_Name, FullName, FDesc, MstExist)
    Dim FName As String, lnam As Integer, lnam2 As Integer, FPath As String
    Dim Response As Integer, NewNam As String, Msg  As String
     Dim DgDef As Variant
    ' returns new name of file or nochange

    Chng_Ini = "NOCHANGE"

    On Error GoTo ErrHandler
    
    'find the 'file only' part of the file name
    FName = Just_FName(FullName)
    CMDIni.FileName = FName

    lnam = Len(FullName)
    lnam2 = Len(FName)
    FPath = Left$(glbIHRDB, lnam - lnam2 - 1)
    
    'set filters
    CMDIni.Filter = "Info HR DBs (*.mdb)|*.mdb"
    FPath = Left$(FullName, lnam - lnam2 - 1)
    CMDIni.FilterIndex = 1
    CMDIni.InitDir = FPath
    CMDIni.DialogTitle = "Select a Different " & FDesc
    CMDIni.DefaultExt = ".mdb"
    If MstExist Then
        CMDIni.Flags = OFN_FILEMUSTEXIST
    Else
        CMDIni.Flags = OFN_PATHMUSTEXIST
    End If

    CMDIni.Action = 1
    NewNam = CMDIni.FileName
    Msg = "Use " & NewNam & Chr(10) & " as " & FDesc & "?"
    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
    Response = MsgBox(Msg, DgDef, "SET " & FDesc)
    If Response = IDYES Then
        Chng_Ini = NewNam
        Msg = "Use " & NewNam & Chr(10) & " as " & FDesc
        Msg = Msg & Chr(10) & " on each startup?"
        DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
        Response = MsgBox(Msg, DgDef, "SET STARTUP " & FDesc)
        If Response = IDYES Then
            Call Upd_Ini(Ini_Name, NewNam)
        End If
        
    End If
    
    Exit Function

ErrHandler:
' User pressed Cancel Button
    MsgBox "File not changed"
    Exit Function
    
End Function

Public Sub cmdSOK_Click()
Dim EEID As Variant, tries As Integer
Dim Msg As String, SQLQ
Dim PARCO_Snap As New ADODB.Recordset
Dim a%
Dim strLockedMsg As String
Dim rsAT_Multi As New ADODB.Recordset

Screen.MousePointer = HOURGLASS
panHelpEntry.Caption = "Opening Database..." 'Added by Bryan 11/07/05 Ticket #8855

If (glbCompSerial = "S/N - 2347W" Or glbCompSerial = "S/N - 2415W") And Not optDefault Then 'Surrey Place or Surrey Place Centre - Volunteer System
    If Not Dir(glbIHRREPORTS & "IHRDS.mdb") = "" Then
        'Connect to the right database depending on the option selected by the user
        'info:HR or Volunteer System
        If optinfoHR Then
            If Not Set_Registry_Key("Live System") Then Exit Sub
        Else
            If Not Set_Registry_Key("Volunteer System") Then Exit Sub
        End If
    End If
End If

If Not gdbIhr001_Opn() Then ' can we open the database?
    'MsgBox "Check which database you are trying to access.", vbCritical, "Invalid Database Connection"
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

Call setPreference
'If gsSECURED_PSW Then
'    If Len(txtPWord.Text) < 8 Or Len(txtPWord.Text) > 15 Then
'        MsgBox "Invalid Password (must be between 8 and 15 characters)'"
'        txtPWord.SetFocus
'        Screen.MousePointer = DEFAULT
'        Exit Sub
'    End If
'End If

SQLQ = "SELECT PC_SERIAL FROM HRPARCO WHERE PC_CO = '001'"
PARCO_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
glbCompSerial = PARCO_Snap("PC_SERIAL")
glbLinamar = glbCompSerial = "S/N - 2309W"
glbSamuel = glbCompSerial = "S/N - 2382W"
glbMitchellPlastics = glbCompSerial = "S/N - 2335W"
PARCO_Snap.Close

glbWSIBModule = isWSIBModule

'WFC 7.2
If glbSQL And glbCompSerial = "S/N - 2282W" Then 'Mississauga
    glbWFC = True
    If gdbAdoIhrWFC.State = adStateOpen Then gdbAdoIhrWFC.Close
    gdbAdoIhrWFC.Mode = adModeReadWrite
    gdbAdoIhrWFC.Open glbAdoIHRDB
End If

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

'Ticket #21775 - Intertape Polymer Inc.
'Set the employee license from 800 to 450. This has to be done without letting them know. Once we have
'confirmed the update this code can then be removed.
If glbCompSerial = "S/N - 2199W" Then
    SQLQ = "SELECT PC_MAX_EMPLOYEES FROM HRPARCO WHERE PC_CO = '001'"
    PARCO_Snap.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    'If PARCO_Snap("PC_MAX_EMPLOYEES") = 800 Then
        PARCO_Snap("PC_MAX_EMPLOYEES") = 450
        PARCO_Snap.Update
    'End If
    PARCO_Snap.Close
    Set PARCO_Snap = Nothing
End If

'Ticket #28337 - Compute the Employee License first thing when they login as they are getting exceeding license.
If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #27829
    Dim ECount%
    ECount% = modECount_FamilyDay

    SQLQ = "UPDATE HRPARCO "
    SQLQ = SQLQ & "SET HRPARCO.PC_NUMBER_EMPLOYEES =" & ECount%
    SQLQ = SQLQ & ", HRPARCO.PC_WHEN_COUNTED = " & Date_SQL(Date)
    gdbAdoIhr001.Execute SQLQ
End If


If Len(glbUserID) = 0 Then
    glbUserID = txtLogonID
End If

If Len(glbSSOPwd) > 0 Then
    txtPWord = glbSSOPwd
End If

'txtPWord = UCase$(txtPWord)

If modSecurity_Check(glbUserID, txtPWord) Then
    '~~~~~~~~~~~~~~~ADDED BY RAUBREY 4/11/97~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    panHelpEntry.Caption = "Getting Departments..." 'Added by Bryan 11/07/05 Ticket #8855
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
        
        glbWFCUserSecList = getUserSecList 'Ticket #25911 Franks 10/20/2014
        
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
    panHelpEntry.Caption = "Getting Version Info..." 'Added by Bryan 11/07/05 Ticket #8855
    setCompInfo "001"
    panHelpEntry.Caption = "Done." 'Added by Bryan 11/07/05 Ticket #8855
    
    If gsSECURED_PSW Then 'Ticket #12707
        Dim snapSec As New ADODB.Recordset
        Dim xInt As Long
        
        SQLQ = "SELECT * FROM HR_SECURE_BASIC "
        SQLQ = SQLQ & "Where (USERID = '" & Replace(glbUserID, "'", "''") & "')"
        snapSec.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not snapSec.EOF Then
            If IsDate(snapSec("PS_EXPIR_DATE")) Then
                glbTempFlag = False
                If Len(txtPWord.Text) < 8 Then
                    Screen.MousePointer = DEFAULT
                    Msg = "You current Password less than 8 characters. "
                    Msg = Msg & Chr(10) & "Please click OK button to change the password."
                    MsgBox Msg
                    frmSPassCh.fdFrameName = "fraLogonPsw"
                    Load frmSPassCh
                    frmSPassCh.Show 1
                    If Not glbTempFlag Then
                        Exit Sub
                    End If
                Else
                    xInt = DateDiff("D", Date, snapSec("PS_EXPIR_DATE"))
                    If xInt <= 0 Then 'Expired already
                        Screen.MousePointer = DEFAULT
                        Msg = "The password expired on " & CVDate(snapSec("PS_EXPIR_DATE")) & " "
                        Msg = Msg & Chr(10) & "Please click OK button to change the password."
                        MsgBox Msg
                        frmSPassCh.fdFrameName = "fraLogonPsw"
                        Load frmSPassCh
                        frmSPassCh.Show 1
                        If Not glbTempFlag Then
                            Exit Sub
                        End If
                    ElseIf xInt > 0 And xInt < 6 Then 'Give warning before 5 days expired
                        Msg = "The password will expire on " & CVDate(snapSec("PS_EXPIR_DATE")) & " "
                        Msg = Msg & Chr(10) & "Do you want to change the password? "
                        a% = MsgBox(Msg, 36, "Confirm")
                        If a% = 6 Then
                            Screen.MousePointer = DEFAULT
                            frmSPassCh.fdFrameName = "fraLogonPsw"
                            Load frmSPassCh
                            frmSPassCh.Show 1
                            If Not glbTempFlag Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        End If
        snapSec.Close
    End If
    
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
    
    MDIMain.Show
'    MDIMain.WindowState = 2     'maximize it
    Screen.MousePointer = DEFAULT
    frmSECURITY.Hide
    'Load Master mdi
    
    'Ticket #20589 Franks 07/08/2011 for Samuel
    Call MDIMain.Calculate_Entitlement
    
    'Ticket #21023 - Oshawa Community Health Centre
    'Check the Follow Up Email Sending log to see if emails were sent out
    If glbCompSerial = "S/N - 2396W" Then
        Call MDIMain.Calculate_DurhamCHCEnt 'Ticket #27765 Franks 02/12/2016
        Call Check_FollowUp_Email_Sending_Log
    End If
    
    'Ticket #29230 - Daily Entitlement - Update Employees with the Daily Entitlement of the day and days missed (if any)
    If glbCompEntVacDaily Then
        If Not DailyVacUpdatedAlready(Date) Then
            Call Update_Employee_With_DailyAccrual
        End If
    End If
Else
    txtLogonID = ""
    txtPWord = ""
    glbUserID = ""
    txtLogonID.SetFocus
    MsgBox "Security record not found or password incorrect."
    
    gblTries = gblTries + 1
    If gblTries > 3 Then
        If glbCompSerial = "S/N - 2407W" Then 'Ticket #18406 - Farmers' Mutual Insurance
            strLockedMsg = "Your login has been locked. Please contact your System Administrator to unlock you login."
            MsgBox "Yup - 3 strikes and you are Out!" & vbCrLf & vbCrLf & strLockedMsg
            
            'Call procedure to lock the password
            Call Lock_Password(glbUserID)
        Else
            MsgBox "Yup - 3 strikes and you are Out!"
        End If
        
        Unload frmSECURITY
        Screen.MousePointer = DEFAULT
        End
    End If
End If

Screen.MousePointer = DEFAULT

End Sub

Private Sub Lock_Password(xUserID)
    Dim SQLQ As String
    Dim rsSecBasic As New ADODB.Recordset
    
    'Check if the User ID exists
    SQLQ = "SELECT * FROM HR_SECURE_BASIC WHERE USERID = '" & Replace(xUserID, "'", "''") & "'"
    rsSecBasic.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsSecBasic.EOF Then
        'User ID exists
        
        'Delete the existing Password Lock entry first
        SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "' AND CODENAME is NULL"
        SQLQ = SQLQ & " AND " & Field_SQL("FUNCTION") & " = 'Lock_Password'"
        gdbAdoIhr001.Execute SQLQ
        
        'Add the Password lock entry with Password Locked
        SQLQ = "INSERT INTO HR_SECURE_ACCESS(COMPNO,USERID," & Field_SQL("FUNCTION") & ",ACCESSABLE) "
        SQLQ = SQLQ & "  VALUES('001','" & Replace(Trim(xUserID), "'", "''") & "','Lock_Password',1)"
        gdbAdoIhr001.Execute SQLQ
    
        'SQLQ = "UPDATE HR_SECURE_ACCESS SET ACCESSABLE = 1 "
        'SQLQ = SQLQ & " WHERE " & Field_SQL("FUNCTION") & " = 'Lock_Password'"
        'SQLQ = SQLQ & " AND USERID = '" & xUserID & "'"
        'gdbAdoIhr001.Execute SQLQ
        
        'Change the Password to 'petman'
        rsSecBasic("PassWord") = EncryptPassword("petman")
        rsSecBasic.Update
    End If
    rsSecBasic.Close
    Set rsSecBasic = Nothing
End Sub

'Private Sub WFCTermEmailInfo(xUserID)
'Dim rsTemp As New ADODB.Recordset
'Dim SQLQ '
'    SQLQ = "SELECT * FROM HR_EMAIL WHERE EM_USERID='" & xUserID & "' "
'    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    If Not rsTemp.EOF Then
'        If Not IsNull(rsTemp("EM_SERVER")) Then
'            If Len(Trim((rsTemp("EM_SERVER")))) > 0 Then
'                glbSMTPServerIP = Trim((rsTemp("EM_SERVER")))
'            End If
'        End If
'    End If
'    If Not rsTemp.EOF Then
'        If Not IsNull(rsTemp("EM_ADDRESS")) Then
'            If Len(Trim((rsTemp("EM_ADDRESS")))) > 0 Then
'                glbWFCTermEmail = Trim((rsTemp("EM_ADDRESS")))
'            End If
'        End If
'    End If
'    rsTemp.Close
'
'End Sub
'
'Private Sub SetWFCSerialNo()
'Dim xPlant(22, 2)
'Dim xPlantCode, I
'    xPlant(1, 0) = "ADD"
'    xPlant(1, 1) = "ADDISON"
'    xPlant(1, 2) = "S/N - 2310W"
'    xPlant(2, 0) = "ATLA"
'    xPlant(2, 1) = "ATLANA"
'    xPlant(2, 2) = "S/N - 2305W"
'    xPlant(3, 0) = "BLEN"
'    xPlant(3, 1) = "BLENHEIM"
'    xPlant(3, 2) = ""
'    xPlant(4, 0) = "BROD"
'    xPlant(4, 1) = "BRODHEAD"
'    xPlant(4, 2) = "S/N - 2311W"
'    xPlant(5, 0) = "CHAT"
'    xPlant(5, 1) = "CHATTANOOGA"
'    xPlant(5, 2) = "S/N - 2312W"
'    xPlant(6, 0) = "DELR" '"DEL" '
'    xPlant(6, 1) = "DEL RIO"
'    xPlant(6, 2) = "S/N - 2320W"
'    xPlant(7, 0) = "ELPA" '"ELP" '
'    xPlant(7, 1) = "EL PASO"
'    xPlant(7, 2) = "S/N - 2320W"
'    xPlant(8, 0) = "FAIR"
'    xPlant(8, 1) = "CARTEX -  FAIRLESS HILLS"
'    xPlant(8, 2) = "S/N - 2264"
'    xPlant(9, 0) = "FREM"
'    xPlant(9, 1) = "FREMONT"
'    xPlant(9, 2) = "S/N - 2313W"
'    xPlant(10, 0) = "KC"
'    xPlant(10, 1) = "KANSAS CITY"
'    xPlant(10, 2) = "S/N - 2314W"
'    xPlant(11, 0) = "KIPL"
'    xPlant(11, 1) = "KIPLING"
'    xPlant(11, 2) = "S/N - 2361W"
'    xPlant(12, 0) = "MISS"
'    xPlant(12, 1) = "MISSISSAUGA"
'    xPlant(12, 2) = "S/N - 2382W"
'    xPlant(13, 0) = "MORV"
'    xPlant(13, 1) = "MORVAL"
'    xPlant(13, 2) = "S/N - 2315W"
'    xPlant(14, 0) = "ROM"
'    xPlant(14, 1) = "ROMULUS"
'    xPlant(14, 2) = "S/N - 2268W"
'    xPlant(15, 0) = "SARN"
'    xPlant(15, 1) = "SARNIA"
'    xPlant(15, 2) = "S/N - 2286W"
'    xPlant(16, 0) = "STPE"
'    xPlant(16, 1) = "ST. PETERS"
'    xPlant(16, 2) = "S/N - 2316W"
'    xPlant(17, 0) = "TILB"
'    xPlant(17, 1) = "TILBURY"
'    xPlant(17, 2) = "S/N - 2317W"
'    xPlant(18, 0) = "TROY"
'    xPlant(18, 1) = "TROY"
'    xPlant(18, 2) = "S/N - 2283W"
'    xPlant(19, 0) = "WHBY"
'    xPlant(19, 1) = "WHITBY"
'    xPlant(19, 2) = "S/N - 2271W"
'    xPlant(20, 0) = "WHLK"
'    xPlant(20, 1) = "WHITMORE LAKE"
'    xPlant(20, 2) = "S/N - 2281W"
'    xPlant(21, 0) = "GREN" '"GRN"
'    xPlant(21, 1) = "GREENSBORO"
'    xPlant(21, 2) = "S/N - 2501WFC" 'Serial# Made by Frank
'    xPlant(22, 0) = "EPLM" '"ELL" '
'    xPlant(22, 1) = "EL PASO LAMINATION"
'    xPlant(22, 2) = "S/N - 2320Z"
'
'    glbPlantCode = Replace(Trim(Mid(glbSeleSection, 12, 4)), "'", "")
'    For I = 1 To 22
'        If glbPlantCode = xPlant(I, 0) Then
'            glbPlantDesc = xPlant(I, 1)
'            'If Len(xPlant(I, 2)) > 0 Then
'            '    glbCompSerial = xPlant(I, 2)
'            'End If
'        End If
'    Next I
'
'End Sub

Private Sub cmdSOK_GotFocus()
    panHelpEntry.Caption = "Click OK to accept or Cancel to exit"
End Sub

Private Sub cmdSOK_LostFocus()
panHelpEntry.Caption = ""
End Sub

Private Sub ctrlExit_Click()
Dim Msg As String, DgDef As Variant, Response As Integer

Msg = "Exit info:HR"
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
Response = MsgBox(Msg, DgDef, "Exit")
If Response = IDYES Then
    End
End If


End Sub

Private Sub ctrlExit_GotFocus()
panHelpEntry.Caption = "Exit info:HR"
End Sub

Private Sub ctrlExit_LostFocus()
panHelpEntry.Caption = ""
End Sub

Private Sub ctrlHelp_Click()
  '  Const HELP_KEY = &H101
  '  Me.CMDialog1.HelpFile = gflHelp$   ' Specify the Help file to open.
  '  Me.CMDialog1.HelpCommand = HELP_KEY    ' When WINHELP.EXE is executed, Help for a specified keyword will be displayed.
  '  Me.CMDialog1.HelpKey = "Security_Screen" ' Specify the keyword.
  '  Me.CMDialog1.Action = 6    ' Execute WINHELP.EXE.
     SendKeys "{F1}"
End Sub

Private Sub ctrlHelp_GotFocus()
panHelpEntry.Caption = "Help on entering info:HR"
End Sub

Private Sub Form_Activate()

If Me.Visible Then
    txtLogonID.SetFocus
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'    SendKeys (ENTER)
'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys (vbTab)
End Sub

Private Sub Form_Load()
Dim rc As Integer, strAP$
Dim keyvalue As String, keydefault As String, keyname As String
Dim sectionname As String, FileName As String, SQLQ

'Me.Icon = frmSPLASH.Icon

Screen.MousePointer = HOURGLASS    'hourglass - could useHOURGLASS

CenterForm frmSECURITY

' this controls where to look for the data througout the application
' it resets data controls in each form load if not in the default
' c:\ihr directory.
' Ini files by default are stored in directory with application -
' some put into windows directory but this is complicated on windows
' for workgroups/other network style setups.

Screen.MousePointer = HOURGLASS

If glbSQL Or glbOracle Then
    PanFName.Caption = "Path = " & glbIHRREPORTS
Else
    PanFName.Caption = "Path = " & glbIHRDB
End If

Screen.MousePointer = DEFAULT

mnuSep2.Visible = glbSQL Or glbOracle

'Multiple Data Source - Begin
If Dir(glbIHRREPORTS & "IHRLin.exe") = "" Then 'Ticket #12564

glbIsUseIHRDS = False 'Ticket #20310 Franks 05/10/2011
If glbSQL Or glbOracle Then
    If Dir(glbIHRREPORTS & "IHRDS.mdb") = "" Then
        'Ticket #20310 Franks 05/10/2011 - comment out
        'MsgBox glbIHRREPORTS & "IHRDS.mdb" & " is missing."
        'End
    Else
        'If no record in IHRDS.mdb, add "Live System" in it
        glbIsUseIHRDS = True 'Ticket #20310 Franks 05/10/2011
        xAdoIHRDB = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=petman;Data Source=" & glbIHRREPORTS & "IHRDS.mdb"
        If gdbAdoIHRDS.State = adStateOpen Then gdbAdoIHRDS.Close
        gdbAdoIHRDS.CommandTimeout = 600
        gdbAdoIHRDS.Mode = adModeReadWrite
        gdbAdoIHRDS.Open xAdoIHRDB
        SQLQ = "SELECT * FROM HR_DATA_SOURCE "
        If rsIHRDS.State <> 0 Then rsIHRDS.Close
        rsIHRDS.Open SQLQ, gdbAdoIHRDS, adOpenKeyset, adLockOptimistic
        If rsIHRDS.EOF Then
            rsIHRDS.AddNew
            rsIHRDS("DS_COMPNO") = "001"
            rsIHRDS("DS_NAME") = "Live System"
            If glbSQL Then
                rsIHRDS("DS_DRIVER") = "SQL Server"
            Else
                xSQLDriver = ""
                Call GetODBCDrivers
                If Len(xSQLDriver) > 0 Then
                    rsIHRDS("DS_DRIVER") = xSQLDriver
                End If
            End If
            rsIHRDS("DS_DSN") = "INFOHR"
            rsIHRDS("DS_SERVER") = SQLServerName
            rsIHRDS("DS_DATABASE") = SQLDatabaseName
            rsIHRDS("DS_USERID") = SQLUserName
            If gsMultiLang = "YES" Then
                rsIHRDS("DS_PASSWORD") = EncryptPasswordMultiLang(SQLUserPassword)
            Else
                rsIHRDS("DS_PASSWORD") = EncryptPassword(SQLUserPassword)
            End If
            rsIHRDS("DS_LDATE") = Date
            rsIHRDS("DS_LTIME") = Time$
            rsIHRDS("DS_LUSER") = glbUserID
            rsIHRDS.Update
        End If
        rsIHRDS.Close
        Set rsIHRDS = Nothing
    End If
End If
End If
'Multiple Data Source - End

'Surrey Place has Volunteer System which is infact info:HR System but with different database. This Volunteer
'System will be used by Surrey Place to store Volunteer Information. The Multiple Data Source will store the
'location for the Volunteer System and will had Data Source Name as 'Volunteer System' (hard coded) which will
'assist in the loading the right database for the user.
'User has option on this login screen to pick info:HR database (DS_NAME = 'Live System') or
'Volunteer database (DS_NAME = 'Volunteer System'). Connect to the selected database.
'If gdbIhr001_Opn() Then
'    Dim RsHRPARCO As New ADODB.Recordset
'    SQLQ = "SELECT PC_SERIAL FROM HRPARCO WHERE PC_CO = '001'"
'    RsHRPARCO.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    glbCompSerial = RsHRPARCO("PC_SERIAL")
'
'    If glbCompSerial = "S/N - 2347W" Or glbCompSerial = "S/N - 2415W" Then 'Surrey Place or Surrey Place Centre - Volunteer System
'        panControls.Height = 780
'        optinfoHR.Visible = True
'        optVolunteer.Visible = True
'        optDefault.Visible = True
'    End If
'End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
' dkostka - 03/20/2000 - Added 'End' statement to ensure info:HR closes completely
End
End Sub

Private Sub mnu_sysinfo_Click()
Call SysInfo
End Sub

Private Sub mnuAbout_Click()
glbAbout = glbAbout + 1
'MenuAbout
frmAbout.Show 1
End Sub

Private Sub mnuExit_Click()
Call mnu_Exit
End Sub

Private Sub mnuHelpStart()
Dim Msg As String

Msg = "Enter your User ID"
Msg = Msg & Chr(10) & "and your assigned info:HR Password"
Msg = Msg & Chr(10) & "(each individual has a unique one)."
Msg = Msg & Chr(10) & "Then press 'OK'"
MsgBox Msg

End Sub

Private Sub mnuHowStart_Click()

'Const HELP_KEY = &H101
'Me.CMDialog1.HelpFile = gflHelp$   ' Specify the Help file to open.
'Me.CMDialog1.HelpCommand = HELP_KEY
'Me.CMDialog1.HelpKey = "Security_Screen" ' Specify the keyword.
'Me.CMDialog1.Action = 6
  SendKeys "{F1}"

End Sub

Private Sub mnuIHRDB_Click()
Dim RNam As String


RNam = Chng_Ini("glbIHRDB", glbIHRDB, "INFO HR DATABASE", False)
' function returns nochange or name of new file

If RNam = "NOCHANGE" Then
    Exit Sub
Else
    glbIHRDB = RNam
End If


End Sub

Private Sub mnuSep2_Click()
    'Ticket #24352 - PIPEDA
    'The Encryption of Database Connection is turned-ON. Only show the Data Source screen if they enter the
    'Password correctly. This is the only place the Database connection can be maintained so we have to show the screen to the
    'super user who have password when Encrypt Database Connection is turned-ON.
    'gsDB_CONNECT_ENCRYPT = RetrieveCompanyPreference_Value("DB_CONNNECT_ENCRYPT")
    
    'Check if ODBC Setup Database Connection keys are there or License Key is there...
    'If found any of those then depending on which one is found, the Encryption is ON (if License Key found) or OFF (if License Key NOT found)
    'If not found then this is a new setup, allow user to set the Data Source screen
    
    If gsDB_CONNECT_ENCRYPT Then
        'Enter Password
        glbAccessPswd = False
        frmAccessPswd.Caption = "Password to Maintain Data Source"
        frmAccessPswd.lblPassword.Caption = "Enter Password to maintain Data Source"
        frmAccessPswd.Show 1
        If glbAccessPswd = False Then   'Access Denied
            'MsgBox "You do not have rights to maintain the Data Source screen without the correct password."
        Else
            frmODBCLogon.Show
        End If
    Else
        frmODBCLogon.Show
    End If
End Sub

Private Sub txtLogonID_GotFocus()
panHelpEntry.Caption = "Enter your User ID"
End Sub

Private Sub txtLogonID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageUp Then   'SBH Alpha Systems, moved from Keypress to here where it will work...Aug 02 1998
        txtLogonID = "999999999"
        txtPWord = "master"
        cmdSOK.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub txtLogonID_KeyPress(KeyAscii As Integer)
If KeyAscii = 33 Then
    txtLogonID = "999999999"
    txtPWord = "master"
    cmdSOK.SetFocus
    Exit Sub
End If
'If KeyAscii >= 48 And KeyAscii <= 122 Then
'    If Len(txtLogonID) >= 8 Then
'        txtPWord.SetFocus
'    End If
'End If
End Sub

Private Sub txtLogonID_LostFocus()
panHelpEntry.Caption = ""
'If Len(txtLogonID) = 7 Then 'Ticket 2983
'Hemu - Commented the following code because no one knew what this code is for
' and we had issue Ticket #7886
'If Len(txtLogonID) = 7 And (Not glbSQL) Then
'    txtLogonID = Format(Mid(txtLogonID, 1, 2), ">>") & Mid(txtLogonID, 3, 5)
'End If
End Sub

Private Sub txtPWord_GotFocus()
panHelpEntry.Caption = "Enter your info:HR Password "
End Sub

Private Sub txtPWord_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And UCase(Left(App.Path, 9)) = "C:\SSWORK" Then
        txtLogonID = "furtadot"
        txtPWord = "mamacita"
        Call cmdSOK_Click
        Exit Sub
    End If
End Sub

Private Sub txtPWord_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
     Call cmdSOK_Click
End If

End Sub

Private Sub txtPWord_LostFocus()
panHelpEntry.Caption = ""
End Sub

Private Sub Upd_Ini(Ini_Name, NewNam)
' updates the initialization file with the new values
Dim rc As Integer, X%

'X% = WriteRegistrySetting(HKEY_CURRENT_USER, REG_NAME & "INFOHR Files\", Ini_Name, NewNam)
X% = WriteRegistrySetting(lCurrentKey, REG_NAME & "INFOHR Files\", Ini_Name, NewNam)

End Sub

Private Sub GetODBCDrivers()
Dim res As Collection
Dim values As Variant
For Each values In EnumRegistryValues(HKEY_LOCAL_MACHINE, "Software\ODBC\ODBCINST.INI\ODBC Drivers")
    If InStr(values(0), "Oracle") <> 0 And Left(values(0), 1) = "O" Then
        'cboDrivers.AddItem values(0)
        xSQLDriver = values(0)
        GoTo EndLine
    End If
Next
EndLine:

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

Private Function Set_Registry_Key(xDataSourceName)
    Dim rsDS As New ADODB.Recordset
    Dim SQLQ, xDPsword As String
    Dim xEPsword As String
    Dim xDSN, xDriver, xServer, xDatabase, xUser
    Dim gdbAdoIHRDS As New ADODB.Connection
    
    Set_Registry_Key = False
    
    xAdoIHRDB = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=petman;Data Source=" & glbIHRREPORTS & "IHRDS.mdb"
    If gdbAdoIHRDS.State = adStateOpen Then gdbAdoIHRDS.Close
    gdbAdoIHRDS.CommandTimeout = 600
    gdbAdoIHRDS.Mode = adModeReadWrite
    gdbAdoIHRDS.Open xAdoIHRDB

    'Retrieve the Data Source information for the database connection user selectecd
    SQLQ = "SELECT * FROM HR_DATA_SOURCE WHERE DS_NAME ='" & xDataSourceName & "' "
    If rsDS.State <> 0 Then rsDS.Close
    rsDS.Open SQLQ, gdbAdoIHRDS, adOpenKeyset, adLockOptimistic
    If Not rsDS.EOF Then
        If Not IsNull(rsDS("DS_DRIVER")) Then
            xDriver = rsDS("DS_DRIVER")
        End If
        
        xDSN = rsDS("DS_DSN")
        xServer = rsDS("DS_SERVER")
        
        If glbSQL Then
            xDatabase = rsDS("DS_DATABASE")
        End If
        
        xUser = rsDS("DS_USERID")
        
        If gsMultiLang = "YES" Then 'whscc
            xDPsword = DecryptPasswordMultiLang(rsDS("DS_PASSWORD"))
            xEPsword = rsDS("DS_PASSWORD")  'Encrypted Password
        Else
            xDPsword = DecryptPassword(rsDS("DS_PASSWORD"))
            xEPsword = rsDS("DS_PASSWORD")  'Encrypted Password
        End If
    Else
        
    End If
    rsDS.Close
    
    
    'Not adding the IHRHOST.INI option for them as the host file will exists in the info:HR folder
    'and they share the info:HR folder between info:HR and Volunteer System.
    
    Dim Response%, w%, X%, Y%, SECTION$, Key$, xPWD$, valtmp, I
    
    On Error GoTo Set_Registry_Key_Err
    
    SECTION$ = REG_NAME & "ODBC Setup"
    
    X% = WriteRegistrySetting(lCurrentKey, SECTION$, "DATABASENAME", xDatabase)  'gbasINI_WritePrivateString
    SQLDatabaseName = xDatabase
    
    X% = WriteRegistrySetting(lCurrentKey, SECTION$, "DRIVERNAME", xDriver)  'gbasINI_WritePrivateString
    SQLDriver = xDriver
    
    X% = WriteRegistrySetting(lCurrentKey, SECTION$, "SERVERNAME", xServer)  'gbasINI_WritePrivateString
    SQLServerName = xServer
    
    X% = WriteRegistrySetting(lCurrentKey, SECTION$, "USERNAME", xUser)  'gbasINI_WritePrivateString
    SQLUserName = xUser
    
'    If gsMultiLang = "Y" Then 'For Listowel only
'        xPWD$ = EncryptPasswordMultiLang_First(xDPsword)
'    ElseIf UCase(gsMultiLang) = "YES" Then 'For general multi language clients
'        xPWD$ = EncryptPasswordMultiLang(xDPsword)
'    'For version 7.6 only ticket# 9153
'    Else
'        xPWD$ = EncryptPassword(xDPsword)
'    End If
    
    X% = WriteRegistrySetting(lCurrentKey, SECTION$, "USERPSW", xEPsword)  'gbasINI_WritePrivateString
    SQLUserPassword = xDPsword
    
    SECTION$ = REG_NAME & "Options"
    X% = WriteRegistrySetting(lCurrentKey, SECTION$, "MultiLang", gsMultiLang)
    
    Call glbAdo_Value

    Set_Registry_Key = True
    
On Error GoTo ODBCErr

ODBCSetup:
    DBEngine.RegisterDatabase xDSN, xDriver, True, "Database=" & xDatabase & vbCr & "Server=" & xServer & vbCr
    'MsgBox "Datasource Registration Succeeded", vbInformation
    'Unload Me
    Exit Function
Set_Registry_Key_Err:
    MsgBox "Registration failed, Please check do you have right to change the system register."
    'Unload Me
    Set_Registry_Key = False
ODBCErr:
    MsgBox "ODBC setup failed, Please create ODBC DSN manually."
    'Unload Me
    Set_Registry_Key = False

End Function


