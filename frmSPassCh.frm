VERSION 5.00
Begin VB.Form frmSPassCh 
   Caption         =   "Change password"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraLogonPsw 
      Height          =   2415
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdLCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdLGOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtLGVerNewPass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   14
         Tag             =   "00-Type new password"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtLGNewPass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   13
         Tag             =   "00-Type new password"
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtLGCurPass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   12
         Tag             =   "00-Type new password"
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "For Logon Screen->  Change Password"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   720
         TabIndex        =   18
         Top             =   2040
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label lblPswNewV 
         Caption         =   "Verify New password"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblPswNew 
         Caption         =   "New password"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblPswCur 
         Caption         =   "Current password"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraVerify 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtNewPass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   0
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   4
         Tag             =   "00-Type new password"
         Top             =   960
         Width           =   4215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "&Verify"
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtVerify 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   0
         PasswordChar    =   "*"
         TabIndex        =   1
         Tag             =   "00-Verify new password"
         Top             =   1560
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.CommandButton cmdEnt 
         Caption         =   "&OK"
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "For Setup -> Security -> Change Password"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   1920
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label lblNewPass 
         Caption         =   "Type new password"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label lblExtend 
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   480
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmSPassCh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mvarFrameName As String
Private Sub cmdCancel_Click()
Unload Me
End Sub

Sub cmdEnt_Click()

If Len(txtNewPass) > 15 Then
    MsgBox "Password must be 6 characters or under.", vbExclamation + vbOKOnly, "Password Change Cancelled"
    Exit Sub
End If
If Len(txtNewPass) > 0 Then
    Me.Height = 2760
    lblNewPass.Caption = "Type password again to verify."
'    lblExtend.Caption = "Verify Password Change"
'    txtNewPass = ""
    cmdVerify.Enabled = True
    cmdVerify.Visible = True
    cmdEnt.Visible = False
    txtVerify.Visible = True
    txtVerify.SetFocus
End If

End Sub

Private Sub cmdLCancel_Click()
Unload Me
End Sub

Private Sub cmdLGOK_Click()
Dim snapSec As New ADODB.Recordset
Dim SQLQ As String
Dim EncPswCur As String, EncPswNew As String
    If Len(txtLGCurPass) = 0 Then
        MsgBox "Please Type Current password"
        txtLGCurPass.SetFocus
        Exit Sub
    End If
    If Len(txtLGNewPass) = 0 Then
        MsgBox "Please Type New password"
        txtLGNewPass.SetFocus
        Exit Sub
    End If
    If Len(txtLGVerNewPass) = 0 Then
        MsgBox "Please Type Verify New password"
        txtLGVerNewPass.SetFocus
        Exit Sub
    End If
    
    'If Len(txtLGCurPass.Text) < 8 Or Len(txtLGCurPass.Text) > 15 Then
    '    MsgBox "Invalid Password (must be between 8 and 15 characters)'"
    '    txtLGCurPass.SetFocus
    '    Exit Sub
    'End If
    If Len(txtLGNewPass.Text) < 8 Or Len(txtLGNewPass.Text) > 15 Then
        MsgBox "Invalid Password (must be between 8 and 15 characters)'"
        txtLGNewPass.SetFocus
        Exit Sub
    End If
    If Len(txtLGVerNewPass.Text) < 8 Or Len(txtLGVerNewPass.Text) > 15 Then
        MsgBox "Invalid Password (must be between 8 and 15 characters)'"
        txtLGVerNewPass.SetFocus
        Exit Sub
    End If
    If txtLGNewPass.Text = txtLGCurPass.Text Then
        MsgBox "New password equal to Current Password"
        txtLGNewPass.SetFocus
        Exit Sub
    End If
    If txtLGNewPass.Text <> txtLGVerNewPass.Text Then
        MsgBox "Password verification failed"
        txtLGVerNewPass.SetFocus
        Exit Sub
    End If
    
    If gsMultiLang = "YES" Then 'whscc
        EncPswCur = EncryptPasswordMultiLang(txtLGCurPass.Text)
        EncPswNew = EncryptPasswordMultiLang(txtLGNewPass.Text)
    Else
        EncPswCur = EncryptPassword(txtLGCurPass.Text)
        EncPswNew = EncryptPassword(txtLGNewPass.Text)
    End If
    
    SQLQ = "SELECT * FROM HR_SECURE_BASIC "
    SQLQ = SQLQ & "Where (USERID = '" & Replace(glbUserID, "'", "''") & "')"
    snapSec.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If snapSec.EOF Then
        MsgBox "User security record not found."
        snapSec.Close
        Unload Me
    Else
        If Not (EncPswCur = snapSec("PassWord")) Then
            MsgBox "Current password incorrect."
            snapSec.Close
            Exit Sub
        End If
        If EncPswNew = snapSec("PS_OLDPW") Or EncPswNew = snapSec("PS_OLDPW2") Or EncPswNew = snapSec("PS_OLDPW3") Then
            MsgBox "This New Password already has been used before."
            Exit Sub
        End If
    End If
    snapSec.Close
                    
    Call modUpdPass(glbUserID, EncPswNew)

    Call UpdPswExpireDatac(glbUserID, EncPswCur)
    
    glbTempFlag = True
    
    Unload Me
            
End Sub

Private Sub cmdVerify_Click()
Dim OldPwd As String
If Len(txtVerify) > 0 Then
        If txtVerify <> txtNewPass Then
            MsgBox "Password verification failed"
            Unload Me
        Else
            If gsMultiLang = "Y" Then
                glbPassword$ = txtVerify
            ElseIf gsMultiLang = "YES" Then 'whscc
                glbPassword$ = EncryptPasswordMultiLang(txtVerify)
            Else
                glbPassword$ = EncryptPassword(txtVerify)
            End If
            
            If gsSECURED_PSW Then 'Ticket #12707
                Dim snapSec As New ADODB.Recordset

                    If Len(txtVerify.Text) < 8 Or Len(txtVerify.Text) > 15 Then
                        MsgBox "Invalid Password (must be between 8 and 15 characters)'"
                        txtVerify.SetFocus
                        Exit Sub
                    End If

                    SQLQ = "SELECT * FROM HR_SECURE_BASIC "
                    SQLQ = SQLQ & "Where (USERID = '" & Replace(glbUserID, "'", "''") & "')"
                    snapSec.Open SQLQ, gdbAdoIhr001, adOpenStatic
                    If Not (snapSec.BOF And snapSec.EOF) Then
                        OldPwd = snapSec("PassWord")
                        If glbPassword$ = snapSec("PassWord") Or glbPassword$ = snapSec("PS_OLDPW") Or glbPassword$ = snapSec("PS_OLDPW2") Or glbPassword$ = snapSec("PS_OLDPW3") Then
                            MsgBox "This password already has been used before."
                            Exit Sub
                        End If
                    End If
                    snapSec.Close

            End If

            Call modUpdPass(glbUserID, glbPassword$)
            
            If gsSECURED_PSW Then
                Call UpdPswExpireDatac(glbUserID, OldPwd)
            End If
            Unload Me
        End If
     End If
End Sub
Private Sub modUpdPass(USERID As String, strNPWord As String)
Dim SQLQ As String

On Error GoTo modUpdPass_Err

SQLQ = "UPDATE HR_SECURE_BASIC "
If glbOracle Then
    SQLQ = SQLQ & "SET PASSWORD = '" & strNPWord & "' "
    SQLQ = SQLQ & "WHERE USERID = '" & Replace(USERID, "'", "''") & "'"
Else
    SQLQ = SQLQ & "SET [PassWord] = '" & strNPWord & "' "
    SQLQ = SQLQ & "WHERE [USERID] = '" & Replace(USERID, "'", "''") & "'"
End If
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

If gsMultiLang = "Y" Then
    gdbAdoIhr001.BeginTrans
    SQLQ = "QRY_SETPASSWORD ('" & Replace(USERID, "'", "''") & "',80)"
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    
    SQLQ = "UPDATE HR_SECURE_BASIC SET PASSWORD = (SELECT PASSWORD FROM HRSECWRK WHERE WRKEMP=USERID AND USERID='" & Replace(USERID, "'", "''") & "') WHERE USERID='" & Replace(USERID, "'", "''") & "'"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans

    SQLQ = "DELETE FROM HRSECWRK WHERE WRKEMP=USERID AND USERID='" & Replace(USERID, "'", "''") & "'"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
End If


Exit Sub

modUpdPass_Err:
glbFrmCaption$ = "Module - Update Password"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update Password", "Password", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Sub

Private Sub Command1_Click()


End Sub

Private Sub Form_Activate()
    If mvarFrameName = "fraVerify" Then
        txtNewPass.SetFocus
    End If
    If mvarFrameName = "fraLogonPsw" Then
        txtLGCurPass.SetFocus
    End If
End Sub

Private Sub Form_Load()
    If mvarFrameName = "fraVerify" Then
        Me.Height = 2040
        Me.Width = 4605
        fraVerify.Left = 120
        fraVerify.Top = 120
        fraVerify.BorderStyle = 0
        fraVerify.Visible = True
    End If
    If mvarFrameName = "fraLogonPsw" Then
        Me.Height = 3000
        Me.Width = 4920
        fraLogonPsw.Left = 120
        fraLogonPsw.Top = 120
        fraLogonPsw.BorderStyle = 0
        fraLogonPsw.Visible = True
    End If
End Sub

Private Sub txtNewPass_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub
Private Sub txtVerify_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Public Property Let fdFrameName(ByVal vData As String)
    mvarFrameName = vData
End Property

