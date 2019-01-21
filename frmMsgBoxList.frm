VERSION 5.00
Begin VB.Form frmMsgBoxList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INFO:HR Message"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLongMsg 
      Height          =   6495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1080
      Width           =   9615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   7800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Height          =   495
      Left            =   4268
      TabIndex        =   1
      Top             =   7800
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblQuestion 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   90
      Width           =   9555
   End
End
Attribute VB_Name = "frmMsgBoxList"
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





