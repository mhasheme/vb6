VERSION 5.00
Begin VB.Form frmMsgYesNoUn 
   Caption         =   "Confirm"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   ControlBox      =   0   'False
   LinkTopic       =   "Confirm"
   MaxButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      Picture         =   "frmMsgYesNoUn.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton UnButton 
      Caption         =   "Unknown"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton YesButton 
      Caption         =   "Yes"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton NoButton 
      Caption         =   "No"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame fraWFC 
      Height          =   975
      Left            =   5520
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   3735
      Begin VB.OptionButton OptWFC 
         Caption         =   "All Benefits (30 and over hours per week)"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   3375
      End
      Begin VB.OptionButton OptWFC 
         Caption         =   "Life Only (20 and over hours per week)"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3495
      End
      Begin VB.OptionButton OptWFC 
         Caption         =   "No Benefits (under 20 hours per week)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Caption         =   "Message"
      Height          =   855
      Left            =   960
      TabIndex        =   4
      Top             =   360
      Width           =   4815
   End
End
Attribute VB_Name = "frmMsgYesNoUn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
glbMsgCustomVal = 0
End Sub

Private Sub NoButton_Click()
If NoButton.Caption = "Update" Then '"Process" Then 'Ticket #23247 Franks 02/25/2014
    If OptWFC(0).Value = False And OptWFC(1).Value = False And OptWFC(2).Value = False Then
        MsgBox "You must have at least one selected"
    Else
        If OptWFC(0).Value Then glbMsgCustomVal = 4
        If OptWFC(1).Value Then glbMsgCustomVal = 5
        If OptWFC(2).Value Then glbMsgCustomVal = 6
        Unload Me
    End If
Else
    glbMsgCustomVal = 2
    Unload Me
End If
End Sub

Private Sub UnButton_Click()
    glbMsgCustomVal = 3
    Unload Me
End Sub

Private Sub YesButton_Click()
    glbMsgCustomVal = 1
    Unload Me
End Sub
Public Sub WFCFrameSetup()
    fraWFC.Left = 1440
    fraWFC.Top = 600
    fraWFC.Visible = True
    fraWFC.BorderStyle = 0
    YesButton.Visible = False
    UnButton.Visible = False
    NoButton.Caption = "Update" ' "Process"
End Sub

Public Sub DailyEntitlementSetup()
    'Used by Daily Accruals
    frmMsgYesNoUn.YesButton.Font = "Microsoft Sans Serif"
    frmMsgYesNoUn.NoButton.Font = "Microsoft Sans Serif"
    frmMsgYesNoUn.UnButton.Font = "Microsoft Sans Serif"
    frmMsgYesNoUn.lblMsg.Font = "Microsoft Sans Serif"
    frmMsgYesNoUn.lblMsg.Width = 5200
    frmMsgYesNoUn.lblMsg.Top = 300
    frmMsgYesNoUn.lblMsg.Height = 1000
    
    frmMsgYesNoUn.YesButton.Left = 400
    frmMsgYesNoUn.NoButton.Left = 2700
    frmMsgYesNoUn.UnButton.Left = 5000
    
    frmMsgYesNoUn.YesButton.Width = 2000
    frmMsgYesNoUn.NoButton.Width = 2000
    frmMsgYesNoUn.UnButton.Width = 1000
End Sub

