VERSION 5.00
Begin VB.Form frmAccessPswd 
   Caption         =   "Password to Save Changes"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
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
      Left            =   3135
      TabIndex        =   1
      Top             =   120
      Width           =   1200
   End
   Begin VB.TextBox txtPassword 
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
      Left            =   120
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Tag             =   "00-Enter password to save changes"
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
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
      Left            =   3135
      TabIndex        =   2
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label lblPassword 
      Caption         =   "Enter the Password to save the changes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmAccessPswd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Dim X
    X = MsgBox("Access Denied", vbOKOnly, "Security Access")
    glbAccessPswd = False
    Unload Me
End Sub

Sub cmdOK_Click()
Dim X
If Len(txtPassword) > 0 Then
    If UCase(txtPassword.Text) = "PETMAN" Then
        'X = MsgBox("Access Granted", vbOKOnly, "Security Access")
        glbAccessPswd = True
    Else
        X = MsgBox("Incorrect Password" & vbCrLf & "Access Denied", vbOKOnly, "Security Access")
        glbAccessPswd = False
    End If
Else
    X = MsgBox("Access Denied", vbOKOnly, "Security Access")
    glbAccessPswd = False
End If

Unload Me

End Sub

Private Sub txtNewPass_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Form_Load()
    If glbWFC Then
        Me.Caption = "Password"
    End If
End Sub
