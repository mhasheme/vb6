VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSIHRLin 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Codes Security"
   ClientHeight    =   6360
   ClientLeft      =   465
   ClientTop       =   1410
   ClientWidth     =   10020
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
   ScaleHeight     =   6360
   ScaleWidth      =   10020
   Begin VB.Frame frmDetail 
      BorderStyle     =   0  'None
      Height          =   4515
      Left            =   180
      TabIndex        =   0
      Top             =   540
      Width           =   8595
      Begin VB.CommandButton cmdGrantAll 
         Appearance      =   0  'Flat
         Caption         =   "&Grant All"
         Height          =   360
         Left            =   6000
         TabIndex        =   1
         Top             =   4020
         Width           =   1305
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   0
         Left            =   420
         TabIndex        =   11
         Top             =   300
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Annual Employee Attendance Sheet"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   1
         Left            =   420
         TabIndex        =   12
         Top             =   600
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Benefits Eligibility Report / New Eligibility Report"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   2
         Left            =   420
         TabIndex        =   13
         Top             =   900
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Course Log Report"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   3
         Left            =   420
         TabIndex        =   15
         Top             =   1500
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Employee Performance Evaluation"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   4
         Left            =   420
         TabIndex        =   16
         Top             =   1800
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Employee Preparation && Worksheet 1 && 2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   5
         Left            =   420
         TabIndex        =   17
         Top             =   2100
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Human Resources Report"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   6
         Left            =   420
         TabIndex        =   18
         Top             =   2400
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "New Employee Data Infomation Report"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   7
         Left            =   420
         TabIndex        =   19
         Top             =   2700
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Payroll Change Notice"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   8
         Left            =   420
         TabIndex        =   20
         Top             =   3000
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Vacation Request Form"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   9
         Left            =   420
         TabIndex        =   21
         Top             =   3600
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Pay Period Table"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   10
         Left            =   420
         TabIndex        =   22
         Top             =   4080
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Reset Upload Log"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   11
         Left            =   420
         TabIndex        =   14
         Top             =   1200
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Emergency Leave Report"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10020
      _Version        =   65536
      _ExtentX        =   17674
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
         TabIndex        =   5
         Top             =   120
         Width           =   630
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
         TabIndex        =   4
         Top             =   125
         Width           =   630
      End
      Begin VB.Label lblPosl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   135
         Width           =   660
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   6
      Top             =   5700
      Width           =   10020
      _Version        =   65536
      _ExtentX        =   17674
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
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2565
         TabIndex        =   10
         Tag             =   "Cancel the changes made"
         Top             =   180
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1725
         TabIndex        =   9
         Tag             =   "Save the changes made"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   45
         TabIndex        =   8
         Tag             =   "Close and exit this screen"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   870
         TabIndex        =   7
         Tag             =   "Edit the information "
         Top             =   180
         Width           =   765
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
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_Return 
         Caption         =   "&Return to Security"
      End
   End
End
Attribute VB_Name = "frmSIHRLin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fglbEditMode%
Dim fUPMode As Integer, fglbEmptyNew As Integer

Private Sub cmdCancel_Click()

On Error GoTo Can_Err

Call Display_Values
Call ST_UPD_MODE(False)  ' reset screen's attributes

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
Dim x%
For x% = 0 To 11
    chkLSecurity(x%).Value = 1
Next x%
End Sub

Private Sub cmdModify_Click()
Dim SQLQ As String

If Not gSec_Upd_Security Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Call ST_UPD_MODE(True)

On Error GoTo Edit_Err

chkLSecurity(0).SetFocus

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

Private Sub cmdOK_Click()
Dim x%
Dim xID
Dim xTemplate As String

On Error GoTo OK_Err

Call ST_UPD_MODE(False)

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

    'Redisplay the security settings
    Call Display_Values
End If

'Template or User only
If xTemplate = "TEMPLATE" Or xTemplate = "" Then
    Call UpdSecAccess
End If

If xTemplate = "TEMPLATE" Then
    '????Ticket #24808 - User's based on this Template does not need their Profile to be updated as we are now retrieving Template profile for the users
    'Call procedure to Update all users with this template.
    'Call Update_Users_withthis_Template(glbSecUSERID)
End If

fglbEditMode% = False

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


Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRJOBEVL", "SELECT")


End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim x%
Dim xTemplate  As String

glbOnTop = Me.name

Screen.MousePointer = HOURGLASS

lblUSERID.Caption = glbSecUSERID
lblEEName.Caption = glbSecEEName
frmSIHRLin.Show

Me.Caption = lStr("Linamar Security - ") & lblEEName

Data1.ConnectionString = glbAdoIHRDB

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbSecUSERID)

If xTemplate = "" Or xTemplate = "TEMPLATE" Then
    Data1.RecordSource = "select * from LN_SECURE_ACCESS where userid='" & Replace(glbSecUSERID, "'", "''") & "'"
Else
    '????Ticket #24808 -  Retrieve template's security profile
    Data1.RecordSource = "select * from LN_SECURE_ACCESS where userid='" & Replace(xTemplate, "'", "''") & "'"
End If
Data1.Refresh

'Call INIData
Call Display_Values


Call ST_UPD_MODE(False)

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
Set frmSIHRLin = Nothing

End Sub

Private Sub mnu_File_Exit_Click()
    Call ApplicationEnd
End Sub

Private Sub mnu_F_PrintSetup_Click()
MDIMain.vbxCommonDlg.Action = 5

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
frmDetail.Enabled = TF
'vbxTrueGrid.Enabled = FT
End Sub

'Private Sub INIData()
'Dim rsTD As New ADODB.Recordset
'Dim rsSR As New ADODB.Recordset
'Dim xStr As String
'Dim SQLQ
'rsTD.Open "HRTABDES", gdbAdoIhr001, adOpenStatic
'Do Until rsTD.EOF
'    rsSR.Open "select * from HR_SECURE_ACCESS where userid='" & glbSecUSERID & " ' and CODENAME='" & rsTD("TD_NAME") & "' AND CODENAME IS NOT NULL", gdbAdoIhr001, adOpenStatic, adLockOptimistic
'    If rsSR.EOF Then
'        SQLQ = "INSERT INTO HR_SECURE_ACCESS(COMPNO,USERID,FUNCTION,ACCESSABLE,Maintainable,CODENAME) "
'        SQLQ = SQLQ & " VALUES('001','" & glbSecUSERID & "','" & rsTD("TD_DESC") & "',0,0,'" & rsTD("TD_NAME") & "')"
'        gdbAdoIhr001.Execute SQLQ
'    Else
'        xStr = rsTD("TD_DESC")
'        xStr = lStr(xStr)
'        If rsSR("FUNCTION") <> xStr Then
'            rsSR("FUNCTION") = xStr
'            rsSR.Update
'        End If
'    End If
'    rsSR.Close
'    rsTD.MoveNext
'Loop
'rsTD.Close
'Data1.Refresh
'End Sub


Private Sub Display_Values()
Dim rsSR As New ADODB.Recordset
Dim x%, SQLQ
Dim xTemplate As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbSecUSERID)

If xTemplate = "" Or xTemplate = "TEMPLATE" Then
    SQLQ = "select * from LN_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "'"
Else
    '????Ticket #24808 -  Retrieve template's security profile
    SQLQ = "select * from LN_SECURE_ACCESS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
End If

rsSR.Open SQLQ, gdbAdoIhr001, adOpenStatic
Call ResetAll
Do Until rsSR.EOF
    If UCase(rsSR("FUNCTION")) = UCase("Annual_Attendance") Then chkLSecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Benefits_Eligibility") Then chkLSecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Course_Log") Then chkLSecurity(2) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Employee_Performance") Then chkLSecurity(3) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Employee_Preparation") Then chkLSecurity(4) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Human_Resources") Then chkLSecurity(5) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("New_Employee") Then chkLSecurity(6) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Payroll_Change") Then chkLSecurity(7) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Vacation_Request") Then chkLSecurity(8) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Pay_Period_Table") Then chkLSecurity(9) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Reset_Upload_Log") Then chkLSecurity(10) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Emergency Leave") Then chkLSecurity(11) = rsSR("ACCESSABLE")
    rsSR.MoveNext
Loop

End Sub

Private Sub UpdSecAccess()
Dim SQLQ

SQLQ = "DELETE LN_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "'"
gdbAdoIhr001.Execute SQLQ

Call AddSecAccess

End Sub

Private Sub AddSecAccess()
Dim SQLQ, sqlI

sqlI = "INSERT INTO LN_SECURE_ACCESS(COMPNO,USERID,[FUNCTION],ACCESSABLE) "
sqlI = sqlI & " VALUES('001','" & Replace(Trim(lblUSERID), "'", "''") & "',"

SQLQ = sqlI & "'Annual_Attendance'," & IIf(chkLSecurity(0), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Benefits_Eligibility'," & IIf(chkLSecurity(1), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Course_Log'," & IIf(chkLSecurity(2), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Emergency Leave'," & IIf(chkLSecurity(11), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Employee_Performance'," & IIf(chkLSecurity(3), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Employee_Preparation'," & IIf(chkLSecurity(4), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Human_Resources'," & IIf(chkLSecurity(5), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'New_Employee'," & IIf(chkLSecurity(6), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Payroll_Change'," & IIf(chkLSecurity(7), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Vacation_Request'," & IIf(chkLSecurity(8), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Pay_Period_Table'," & IIf(chkLSecurity(9), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Reset_Upload_Log'," & IIf(chkLSecurity(10), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
End Sub

Private Sub ResetAll()
Dim x%

For x% = 0 To 11
    chkLSecurity(x%).Value = 0
Next x%

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
            Call SpecificFunction_Template_Based_Security_Profile_Update(rsSecBasic("USERID"), xTemplate, "Change", "CUSTOMFEATURE")
        End If
        rsSecBasic.MoveNext
    Loop
    rsSecBasic.Close
    Set rsSecBasic = Nothing
    
End Sub

