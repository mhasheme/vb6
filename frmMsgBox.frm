VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INFO:HR Message"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLongMsg 
      BackColor       =   &H8000000F&
      Height          =   2655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmMsgBox.frx":0000
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblQuestion 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2850
      Width           =   4395
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    glbMsgBoxResult = vbCancel
    Unload Me
End Sub

Private Sub cmdOK_Click()
    glbMsgBoxResult = vbOK
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim LongMsg As String
    
    Printer.Font.Size = 14
    Printer.Print vbCrLf & vbCrLf & vbCrLf & vbCrLf
    LongMsg = Replace(txtLongMsg.Text, vbCrLf, vbCrLf & "            ")
    LongMsg = Replace(LongMsg, "Continue?", "")
    Printer.Print "            " & LongMsg
    Printer.EndDoc
End Sub





