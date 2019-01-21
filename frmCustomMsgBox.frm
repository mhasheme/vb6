VERSION 5.00
Begin VB.Form frmCustomMsgBox 
   Caption         =   "Confirm"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   ControlBox      =   0   'False
   LinkTopic       =   "Confirm"
   MaxButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      Picture         =   "frmCustomMsgBox.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton UnButton 
      Caption         =   "Unknown"
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
      Left            =   4230
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton YesButton 
      Caption         =   "Yes"
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
      Left            =   1350
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton NoButton 
      Caption         =   "No"
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
      Left            =   2790
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Caption         =   "Message"
      Height          =   1095
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmCustomMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
glbMsgCustomVal = 0
End Sub

Private Sub NoButton_Click()
    glbMsgCustomVal = 2
    Unload Me
End Sub

Private Sub UnButton_Click()
    glbMsgCustomVal = 3
    Unload Me
End Sub

Private Sub YesButton_Click()
    glbMsgCustomVal = 1
    Unload Me
End Sub
