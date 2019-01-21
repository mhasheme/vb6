VERSION 5.00
Begin VB.Form frmMsgDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirm"
   ClientHeight    =   2040
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      Picture         =   "frmMsgDialog.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Replace"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Keep"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblMsg3 
      Caption         =   "If Cancel is checked, undo the change."
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblMsg2 
      Caption         =   "If Continue is checked, save the record with the incorrect RA#1."
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label lblMsg 
      Caption         =   "Keep the Salary Effective Date or change the Salary Effective Date to equal the Position’s Start Date?"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmMsgDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim xMsgNum As Integer

Private Sub CancelButton_Click()
glbMsgCustomVal = 2
Unload Me
End Sub

Private Sub Form_Load()
If glbWFC Then 'Ticket #29069 Franks 08/18/2016
    xMsgNum = glbMsgCustomVal
    glbMsgCustomVal = 0
    Call WFCScreenSetup
Else
    glbMsgCustomVal = 0
End If
End Sub

Private Sub OKButton_Click()
    glbMsgCustomVal = 1
    Unload Me
End Sub

Private Sub WFCScreenSetup()
If xMsgNum = 11 Then
    lblMsg.Caption = "The Reporting Authority #1 entered does not hold the Position that matches the Employee's Position Master's Reporting Authority #1"
    'lblMsg2.Caption = "If <<Continue>> is checked, save the record with the incorrect RA#1."
    'lblMsg3.Caption = "If <<Cancel>> is checked, undo the change."
    lblMsg2.Caption = "Click Continue to save the record with the incorrect RA#1."
    lblMsg3.Caption = "Click Cancel to undo the change."
    lblMsg2.Visible = True
    lblMsg3.Visible = True
    
    OKButton.Caption = "Continue"
    CancelButton.Caption = "Cancel"
End If
If xMsgNum = 20 Then 'Ticket #29438 Franks 11/07/2016
    lblMsg.Caption = "Leaving the Interim/New Reporting Authority blank may cause a break in the organization chain. Do you want to proceed?  "
    lblMsg2.Caption = ""
    lblMsg3.Caption = ""
    
    OKButton.Caption = "Update"
    CancelButton.Caption = "Leave Blank"
End If
End Sub
