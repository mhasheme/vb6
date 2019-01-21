VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSCodes 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Codes Security"
   ClientHeight    =   5955
   ClientLeft      =   465
   ClientTop       =   1410
   ClientWidth     =   9765
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5955
   ScaleWidth      =   9765
   Begin VB.Frame frmDetail 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   360
      TabIndex        =   14
      Top             =   3870
      Width           =   9165
      Begin VB.CheckBox chkMaintain 
         Caption         =   "Maintain"
         DataField       =   "Maintainable"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   5790
         TabIndex        =   1
         Tag             =   "40-Maintain -y/n"
         Top             =   60
         Width           =   1335
      End
      Begin VB.CheckBox chkInquire 
         Caption         =   "Inquire"
         DataField       =   "Accessable"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   7290
         TabIndex        =   2
         Tag             =   "40-Inquire -y/n"
         Top             =   30
         Width           =   1245
      End
      Begin VB.Label lblFunction 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         DataField       =   "Function"
         DataSource      =   "Data1"
         Height          =   195
         Left            =   60
         TabIndex        =   15
         Top             =   90
         Width           =   975
      End
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9765
      _Version        =   65536
      _ExtentX        =   17224
      _ExtentY        =   873
      _StockProps     =   15
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
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin VB.Label lblPosl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   135
         Width           =   660
      End
      Begin VB.Label lblUSERID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ABCD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1320
         TabIndex        =   11
         Top             =   125
         Width           =   630
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Descr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3030
         TabIndex        =   10
         Top             =   120
         Width           =   630
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   13
      Top             =   5295
      Width           =   9765
      _Version        =   65536
      _ExtentX        =   17224
      _ExtentY        =   1164
      _StockProps     =   15
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
      Begin VB.CommandButton cmdRemoveAll 
         Appearance      =   0  'Flat
         Caption         =   "&Remove All"
         Height          =   360
         Left            =   4800
         TabIndex        =   17
         Tag             =   "Grant All"
         Top             =   195
         Width           =   1545
      End
      Begin VB.CommandButton cmdGrantInquire 
         Appearance      =   0  'Flat
         Caption         =   "Grant All &Inquire"
         Height          =   360
         Left            =   6420
         TabIndex        =   16
         Tag             =   "Grant All"
         Top             =   195
         Width           =   1545
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   3480
         TabIndex        =   7
         Tag             =   "Print Codes Security"
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdGrantAll 
         Appearance      =   0  'Flat
         Caption         =   "&Grant All"
         Height          =   360
         Left            =   8040
         TabIndex        =   8
         Tag             =   "Grant All"
         Top             =   195
         Width           =   1545
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   870
         TabIndex        =   4
         Tag             =   "Edit the information "
         Top             =   180
         Width           =   765
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   45
         TabIndex        =   3
         Tag             =   "Close and exit this screen"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1725
         TabIndex        =   5
         Tag             =   "Save the changes made"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2565
         TabIndex        =   6
         Tag             =   "Cancel the changes made"
         Top             =   180
         Width           =   795
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   405
         Left            =   4200
         Top             =   180
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   714
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         BoundReportHeading=   "RGELIST"
         BoundReportFooter=   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fscodes.frx":0000
      Height          =   3285
      Left            =   240
      OleObjectBlob   =   "fscodes.frx":0014
      TabIndex        =   0
      Top             =   480
      Width           =   9330
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_Return 
         Caption         =   "&Return to Security"
      End
   End
End
Attribute VB_Name = "frmSCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fglbEditMode%
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim fmode As String

Private Sub chkMaintain_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub
Private Sub chkInquire_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub
Private Sub cmdCancel_Click()

On Error GoTo Can_Err

Data1.Recordset.CancelUpdate
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
Call ST_UPD_MODE(False)  ' reset screen's attributes
Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdCancel_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdGrantAll_Click()
Dim Msg, Title$, DgDef, Response%
Dim xTemplate As String

Msg = "Would you like to grant all securities to all codes?"
Title$ = "Grant all?"   ' zzz
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON3  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.

If Response = IDYES Then
    'Ticket #20585 - If Template then update users with this template as well.
    'If User and with no template then update that user's profile.
    'if User and with Template then do not update user's profile.
    'Get the Template Name of this User ID
    xTemplate = Get_Template(glbSecUSERID)
    
    If xTemplate = "TEMPLATE" Then
        'Update all users with this template. After the changes are saved
    ElseIf xTemplate = "" Then
        'User - User with no template - don't do anything let system update user's profile
    ElseIf xTemplate <> "TEMPLATE" Then
        'User with template - do not allow to save these changes.
        MsgBox "Security change cannot be saved. This user's security profile is based on the '" & xTemplate & "' template.", vbExclamation, "Template based User Security Profile"
    End If
    
    'if Template or User
    If xTemplate = "TEMPLATE" Or xTemplate = "" Then
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute "update HR_SECURE_ACCESS set ACCESSABLE=1,Maintainable=1 where USERID='" & Replace(glbSecUSERID, "'", "''") & "' and CODENAME IS NOT NULL"
        gdbAdoIhr001.CommitTrans
        'Data1.Refresh
    End If
    
    Data1.Refresh
    
    If xTemplate = "TEMPLATE" Then
        'Call procedure to Update all users with this template.
        Call Update_Users_withthis_Template(glbSecUSERID)
    End If
    
End If
End Sub

Private Sub cmdGrantAll_GotFocus()
    'Hemu - 05/13/2003 Begin
    Call SetPanHelp(ActiveControl)
    'Hemu - 05/13/2003 End
End Sub

Private Sub cmdGrantInquire_Click()
Dim Msg, Title$, DgDef, Response%
Dim xTemplate As String

Msg = "Would you like to grant all Inquire to all codes?"
Title$ = "Grant all Inquire?"   ' zzz
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON3  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.

If Response = IDYES Then
    'Ticket #20585 - If Template then update users with this template as well.
    'If User and with no template then update that user's profile.
    'if User and with Template then do not update user's profile.
    'Get the Template Name of this User ID
    xTemplate = Get_Template(glbSecUSERID)
    
    If xTemplate = "TEMPLATE" Then
        'Update all users with this template. After the changes are saved
    ElseIf xTemplate = "" Then
        'User - User with no template - don't do anything let system update user's profile
    ElseIf xTemplate <> "TEMPLATE" Then
        'User with template - do not allow to save these changes.
        MsgBox "Security change cannot be saved. This user's security profile is based on the '" & xTemplate & "' template.", vbExclamation, "Template based User Security Profile"
    End If
    
    'if Template or User
    If xTemplate = "TEMPLATE" Or xTemplate = "" Then
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute "update HR_SECURE_ACCESS set ACCESSABLE=1 where USERID='" & Replace(glbSecUSERID, "'", "''") & "' and CODENAME IS NOT NULL"
        gdbAdoIhr001.CommitTrans
    End If
    
    Data1.Refresh
    
    If xTemplate = "TEMPLATE" Then
        'Call procedure to Update all users with this template.
        Call Update_Users_withthis_Template(glbSecUSERID)
    End If
End If

End Sub

Private Sub cmdModify_Click()
Dim SQLQ As String

If Not gSec_Upd_Security Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Call ST_UPD_MODE(True)

On Error GoTo Edit_Err

chkMaintain.SetFocus

Exit Sub
Edit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdEdit", "HRJOBEVL", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim x%
Dim xID
Dim xTemplate As String

On Error GoTo OK_Err

'Ticket #20585 - If Template then update users with this template as well.
'If User and with no template then update that user's profile.
'if User and with Template then do not update user's profile.
'Get the Template Name of this User ID
xTemplate = Get_Template(glbSecUSERID)

If xTemplate = "TEMPLATE" Then
    'Update all users with this template. After the changes are saved
ElseIf xTemplate = "" Then
    'User - User with no template - don't do anything let system update user's profile
ElseIf xTemplate <> "TEMPLATE" Then
    'User with template - do not allow to save these changes.
    MsgBox "Security change cannot be saved. This user's security profile is based on the '" & xTemplate & "' template.", vbExclamation, "Template based User Security Profile"
End If

'if Template or User
If xTemplate = "TEMPLATE" Or xTemplate = "" Then
    Data1.Recordset("Maintainable") = IIf(chkMaintain, 1, 0)
    Data1.Recordset.UpdateBatch
End If

If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

If xTemplate = "TEMPLATE" Then
    '????Ticket #24808 - User's based on this Template does not need their Profile to be updated as we are now retrieving Template profile for the users
    'Call procedure to Update all users with this template.
    'Call Update_Users_withthis_Template(glbSecUSERID)
End If

Call ST_UPD_MODE(False)

fglbEditMode% = False

Me.vbxTrueGrid.SetFocus

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRJOBEVL", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdOK_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub cmdPrint_Click()
Dim RHeading As String, xReport, x%

cmdPrint.Enabled = False

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

Me.vbxCrystal.WindowTitle = "Codes Security Report"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 8 '10
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RGTABSEC.rpt"
    Me.vbxCrystal.SelectionFormula = "{HR_SECURE_BASIC.USERID}='" & Replace(lblUSERID, "'", "''") & "' "

Me.vbxCrystal.Action = 1

cmdPrint.Enabled = True

End Sub

Private Sub cmdPrint_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdRemoveAll_Click()
Dim Msg, Title$, DgDef, Response%
Dim xTemplate As String

Msg = "Would you like to remove all securities to all codes?"
Title$ = "Remove all?"   ' zzz
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON3  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.

If Response = IDYES Then
    'Ticket #20585 - If Template then update users with this template as well.
    'If User and with no template then update that user's profile.
    'if User and with Template then do not update user's profile.
    'Get the Template Name of this User ID
    xTemplate = Get_Template(glbSecUSERID)
    
    If xTemplate = "TEMPLATE" Then
        'Update all users with this template. After the changes are saved
    ElseIf xTemplate = "" Then
        'User - User with no template - don't do anything let system update user's profile
    ElseIf xTemplate <> "TEMPLATE" Then
        'User with template - do not allow to save these changes.
        MsgBox "Security change cannot be saved. This user's security profile is based on the '" & xTemplate & "' template.", vbExclamation, "Template based User Security Profile"
    End If
    
    'if Template or User
    If xTemplate = "TEMPLATE" Or xTemplate = "" Then
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute "UPDATE HR_SECURE_ACCESS set ACCESSABLE=0,Maintainable=0 where USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND CODENAME IS NOT NULL"
        gdbAdoIhr001.CommitTrans
    End If
    
    Data1.Refresh
    
    If xTemplate = "TEMPLATE" Then
        'Call procedure to Update all users with this template.
        Call Update_Users_withthis_Template(glbSecUSERID)
    End If
End If

End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRJOBEVL", "SELECT")

End Sub

Private Sub Form_Load()
glbOnTop = Me.name
Dim SQLQ
Dim xTemplate  As String

Screen.MousePointer = HOURGLASS
lblUSERID.Caption = glbSecUSERID
lblEEName.Caption = glbSecEEName

'Ticket #29024 - Had to comment this so this form can be shown Modal form from frmSecurity
'frmSCodes.Show

Data1.ConnectionString = glbAdoIHRDB

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbSecUSERID)

SQLQ = "SELECT * FROM HR_SECURE_ACCESS"
If xTemplate = "" Or xTemplate = "TEMPLATE" Then
    SQLQ = SQLQ & " WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND CODENAME IS NOT NULL "
Else
    '????Ticket #24808 -  Retrieve template's security profile
    SQLQ = SQLQ & " WHERE USERID='" & Replace(xTemplate, "'", "''") & "' AND CODENAME IS NOT NULL "
End If
SQLQ = SQLQ & " ORDER BY " & Upper_SQL(Field_SQL("FUNCTION"))

Data1.RecordSource = SQLQ
Data1.Refresh

Call INIData

If vbxTrueGrid.Visible Then Me.vbxTrueGrid.SetFocus

Call ST_UPD_MODE(False)

'Ticket #20585 - Enable/Disable Edit, Grant All, Remove All and Grant All Inquire buttons based
'on the type of user
xTemplate = Get_Template(glbSecUSERID)
If xTemplate = "" Or xTemplate = "TEMPLATE" Then
    'User without Template or Template
    cmdGrantAll.Enabled = True
    cmdGrantInquire.Enabled = True
    cmdRemoveAll.Enabled = True
Else
    'User with Template
    cmdGrantAll.Enabled = False
    cmdGrantInquire.Enabled = False
    cmdRemoveAll.Enabled = False
    cmdModify.Enabled = False
End If

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
Set frmSCodes = Nothing
End Sub

Private Sub mnu_Return_Click()
   Unload Me
End Sub


Private Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

glbOHSEdit% = TF

fUPMode = TF    ' update mode
cmdOK.Enabled = TF
cmdModify.Enabled = FT
cmdCancel.Enabled = TF
cmdClose.Enabled = FT
cmdPrint.Enabled = FT
cmdGrantAll.Enabled = FT
'chkMaintain.Enabled = TF
'chkInquire.Enabled = TF
frmDetail.Enabled = TF
vbxTrueGrid.Enabled = FT

End Sub

Private Sub vbxTrueGrid_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT * FROM HR_SECURE_ACCESS"
        SQLQ = SQLQ & " WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND CODENAME IS NOT NULL "
        SQLQ = SQLQ & " ORDER BY " & Upper_SQL(Field_SQL(vbxTrueGrid.Columns(ColIndex).DataField)) & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
    If cmdOK.Enabled Then
        cmdOK.SetFocus
    Else
        cmdClose.SetFocus
    End If
End If

End Sub

Private Sub INIData()
Dim rsTD As New ADODB.Recordset
Dim rsSR As New ADODB.Recordset
Dim xStr As String
Dim SQLQ
Dim xTemplate As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbSecUSERID)

rsTD.Open "HRTABDES", gdbAdoIhr001, adOpenStatic
Do Until rsTD.EOF
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        rsSR.Open "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' and CODENAME='" & rsTD("TD_NAME") & "' AND CODENAME IS NOT NULL", gdbAdoIhr001, adOpenStatic, adLockOptimistic
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        rsSR.Open "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xTemplate, "'", "''") & "' and CODENAME='" & rsTD("TD_NAME") & "' AND CODENAME IS NOT NULL", gdbAdoIhr001, adOpenStatic, adLockOptimistic
    End If
    If rsSR.EOF Then
        '????Ticket #24808 -  Only insert default records for Template Users or Normal users
        If xTemplate = "" Or xTemplate = "TEMPLATE" Then
            SQLQ = "INSERT INTO HR_SECURE_ACCESS(COMPNO,USERID," & Field_SQL("FUNCTION") & ",ACCESSABLE,Maintainable,CODENAME) "
            SQLQ = SQLQ & " VALUES('001','" & Replace(glbSecUSERID, "'", "''") & "','" & Replace(lStr(rsTD("TD_DESC")), "'", "''") & "',0,0,'" & rsTD("TD_NAME") & "')"
            gdbAdoIhr001.Execute SQLQ
        End If
    Else
'        rsSR.Delete
        xStr = rsTD("TD_DESC")
        xStr = lStr(xStr)
        If rsSR("FUNCTION") <> xStr Then
            rsSR("FUNCTION") = xStr
            rsSR.Update
        End If
    End If
    rsSR.Close
    rsTD.MoveNext
Loop
rsTD.Close

Data1.Refresh
End Sub

Private Sub Update_Users_withthis_Template(xTemplate)
    Dim SQLQ As String
    Dim rsSecBasic As New ADODB.Recordset
    
    'Retrieve all users associated with this changed Template
    SQLQ = "SELECT USERID, SECURE_TEMPLATE FROM HR_SECURE_BASIC WHERE SECURE_TEMPLATE = '" & xTemplate & "'"
    SQLQ = SQLQ & " ORDER BY USERID"
    rsSecBasic.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsSecBasic.EOF
        If Not IsNull(rsSecBasic("USERID")) Then
            'Update each user with this changed Template
            Call SpecificFunction_Template_Based_Security_Profile_Update(rsSecBasic("USERID"), xTemplate, "Change", "CODES")
        End If
        rsSecBasic.MoveNext
    Loop
    rsSecBasic.Close
    Set rsSecBasic = Nothing
    
End Sub


