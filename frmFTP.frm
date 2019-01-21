VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFTP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Transfer in Progress..."
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.Animation aniMain 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1296
      _Version        =   393216
      Center          =   -1  'True
      FullWidth       =   305
      FullHeight      =   49
   End
   Begin VB.Label lblNoAni 
      Caption         =   "File transfer in progress, please wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   4575
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AniPlaying As Boolean
Private Sub Form_Load()
    'If Dir(glbSystemData.Path & "FILECOPY.AVI") <> "" Then
    '    aniMain.Open glbSystemData.Path & "FILECOPY.AVI"
    If Dir(glbIHRREPORTS & "FILECOPY.AVI") <> "" Then
        aniMain.Open glbIHRREPORTS & "FILECOPY.AVI"
        aniMain.Play
        AniPlaying = True
    Else
        lblNoAni.Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If AniPlaying Then
        aniMain.Stop
        aniMain.Close
    End If
End Sub
