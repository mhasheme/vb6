VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmSIHRWHSCC 
   Caption         =   "WHSCC Security"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmDetail 
      BorderStyle     =   0  'None
      Height          =   2715
      Left            =   480
      TabIndex        =   9
      Top             =   600
      Width           =   6555
      Begin VB.CommandButton cmdGrantAll 
         Appearance      =   0  'Flat
         Caption         =   "&Grant All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   10
         Top             =   2160
         Width           =   1305
      End
      Begin Threed.SSCheck chkUMSecurity 
         Height          =   225
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkUMSecurity 
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkUISecurity 
         Height          =   225
         Index           =   0
         Left            =   1320
         TabIndex        =   14
         Top             =   600
         Width           =   2325
         _Version        =   65536
         _ExtentX        =   4101
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Advance Sick Leave"
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
         Font3D          =   3
      End
      Begin Threed.SSCheck chkUISecurity 
         Height          =   225
         Index           =   2
         Left            =   1320
         TabIndex        =   15
         Top             =   840
         Width           =   2445
         _Version        =   65536
         _ExtentX        =   4313
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Union Sick Bank"
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
         Font3D          =   3
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inquire"
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
         Height          =   195
         Index           =   11
         Left            =   1200
         TabIndex        =   17
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maintain"
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
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Basic"
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
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   480
      End
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7980
      _Version        =   65536
      _ExtentX        =   14076
      _ExtentY        =   873
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
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin VB.Label lblPosl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
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
         Height          =   195
         Left            =   120
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   120
         Width           =   2310
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   3990
      Width           =   7980
      _Version        =   65536
      _ExtentX        =   14076
      _ExtentY        =   1164
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
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
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
         Left            =   1140
         TabIndex        =   2
         Tag             =   "Edit the information "
         Top             =   0
         Width           =   765
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
         Left            =   315
         TabIndex        =   1
         Tag             =   "Close and exit this screen"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
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
         Left            =   2010
         TabIndex        =   3
         Tag             =   "Save the changes made"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
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
         Left            =   2835
         TabIndex        =   4
         Tag             =   "Cancel the changes made"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Menu munu_file 
      Caption         =   "&File"
      Begin VB.Menu menu_file_return 
         Caption         =   "&Return to Security"
      End
   End
End
Attribute VB_Name = "frmSIHRWHSCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fglbEditMode%
Dim fUPMode As Integer, fglbEmptyNew As Integer

Private Sub chkUMSecurity_Click(Index As Integer, Value As Integer)
    If chkUMSecurity(Index) Then
        chkUISecurity(Index) = chkUMSecurity(Index)
    End If
End Sub

Private Sub cmdCancel_Click()

On Error GoTo Can_Err

Call Display_Values

Call ST_UPD_MODE(False)  ' reset screen's attributes

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "WHSCC", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdGrantAll_Click()
Dim x%
    For x% = 0 To 2
        If x% <> 1 Then
            chkUMSecurity(x%).Value = -1
            chkUISecurity(x%).Value = -1
        End If
    Next x%

'    chkLSecurity(0).Value = -1
End Sub

Private Sub cmdModify_Click()
Dim SQLQ As String

If Not gSec_Upd_Security Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Call ST_UPD_MODE(True)

On Error GoTo Edit_Err

chkUMSecurity(0).SetFocus

Exit Sub
Edit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdEdit", "WHSCC", "Edit")
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim x%

glbOnTop = Me.name

Screen.MousePointer = HOURGLASS

lblUSERID.Caption = glbSecUSERID
lblEEName.Caption = glbSecEEName

frmSIHRWHSCC.Show

Me.Caption = lStr("WHSCC Security - ") & lblEEName

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

Private Sub menu_file_return_Click()
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
End Sub

Private Sub ResetAll()
Dim x%
    For x% = 0 To 2
        If x% <> 1 Then
            chkUMSecurity(x%).Value = 0
            chkUISecurity(x%).Value = 0
        End If
    Next x%

'    chkLSecurity(0).Value = 0

End Sub

Private Sub Display_Values()
Dim rsSR As New ADODB.Recordset
Dim x%, SQLQ
Dim xTemplate  As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbSecUSERID)

    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        SQLQ = "select * from HR_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND LEFT([FUNCTION],4)='WHSC'"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        SQLQ = "select * from HR_SECURE_ACCESS WHERE USERID='" & Replace(xTemplate, "'", "''") & "' AND LEFT([FUNCTION],4)='WHSC'"
    End If
    rsSR.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    Call ResetAll
    
    Do Until rsSR.EOF
        If UCase(rsSR("FUNCTION")) = UCase("WHSCC_ASL") Then
            chkUMSecurity(0) = rsSR("Maintainable")
            chkUISecurity(0) = rsSR("ACCESSABLE")
        End If
'        If UCase(rsSR("FUNCTION")) = UCase("WHSCC_BUDPOS") Then
'            chkUMSecurity(1) = rsSR("Maintainable")
'            chkUISecurity(1) = rsSR("ACCESSABLE")
'        End If
        If UCase(rsSR("FUNCTION")) = UCase("WHSCC_USB") Then
            chkUMSecurity(2) = rsSR("Maintainable")
            chkUISecurity(2) = rsSR("ACCESSABLE")
        End If
'        If UCase(rsSR("FUNCTION")) = UCase("WHSCC_PLAN_ESTABLISMNET_REPORT") Then chkLSecurity(0) = rsSR("ACCESSABLE")

        rsSR.MoveNext
    Loop

End Sub

Private Sub UpdSecAccess()
Dim SQLQ

SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND LEFT([FUNCTION],4)='WHSC'"
gdbAdoIhr001.Execute SQLQ

Call AddSecAccess

End Sub

Private Sub AddSecAccess()
Dim SQLQ, sqlI, sqlA

sqlI = "INSERT INTO HR_SECURE_ACCESS(COMPNO,USERID,[FUNCTION],ACCESSABLE) "
sqlI = sqlI & " VALUES('001','" & Replace(Trim(lblUSERID), "'", "''") & "',"
sqlA = "INSERT INTO HR_SECURE_ACCESS(COMPNO,USERID,[FUNCTION],Maintainable,ACCESSABLE) "
sqlA = sqlA & " VALUES('001','" & Replace(Trim(lblUSERID), "'", "''") & "',"

SQLQ = sqlA & "'WHSCC_ASL'," & IIf(chkUMSecurity(0), 1, 0) & "," & IIf(chkUISecurity(0), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'SQLQ = sqlA & "'WHSCC_BUDPOS'," & IIf(chkUMSecurity(1), 1, 0) & "," & IIf(chkUISecurity(1), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ

SQLQ = sqlA & "'WHSCC_USB'," & IIf(chkUMSecurity(2), 1, 0) & "," & IIf(chkUISecurity(2), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'SQLQ = sqlI & "'WHSCC_PLAN_ESTABLISMNET_REPORT'," & IIf(chkLSecurity(0), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ

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

