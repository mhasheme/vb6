VERSION 5.00
Begin VB.Form frmConfPass 
   Caption         =   "Verify Password Changes"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
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
      Left            =   360
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Tag             =   "00-Verify Password Changes"
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "&Verify"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblNewPass 
      Caption         =   "Verify or Cancel Changes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmConfPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
glbConfPass = ""
Unload Me
End Sub

Private Sub cmdVerify_Click()
If Me.txtNewPass = glbConfPass Then
    Unload Me
Else
    MsgBox "Verification not correct!"
    glbConfPass = ""
    Unload Me
End If
End Sub


Private Sub txtNewPass_GotFocus()
 Call SetPanHelp(ActiveControl)
End Sub
