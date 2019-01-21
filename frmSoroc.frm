VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmSoroc 
   Caption         =   "Termination Data"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2325
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Tag             =   "41-Termination Code "
      Top             =   840
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "TERM"
   End
   Begin INFOHR_Controls.DateLookup dlpTermDate 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Tag             =   "41-Date Terminated"
      Top             =   480
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   1665
      Width           =   6315
      _Version        =   65536
      _ExtentX        =   11139
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
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Tag             =   "Save the changes made"
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Termination Reason"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   810
      Width           =   1710
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Termination Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Tag             =   "41-Date Terminated"
      Top             =   480
      Width           =   1470
   End
End
Attribute VB_Name = "frmSoroc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    If Len(dlpTermDate.Text) < 1 Then
        MsgBox ("Termination Date is a required field")
        dlpTermDate.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(dlpTermDate.Text) Then
        MsgBox ("Termination Date is not a valid date.")
        dlpTermDate.SetFocus
        Exit Sub
    End If
    
    If Len(clpCode(1).Text) < 1 Then
        MsgBox ("Termination Reason is a required field")
        clpCode(1).SetFocus
        Exit Sub
    End If
    If clpCode(1).Caption = "Unassigned" Then
        MsgBox ("Termination Reason is not a valid field")
        clpCode(1).SetFocus
        Exit Sub
    End If
    glbSorocTermDate = dlpTermDate
    glbSorocTermReason = clpCode(1)
    glbChgTermDate = dlpTermDate
    glbChgTermReason = clpCode(1)
    Unload Me
End Sub

Private Sub Form_Load()
glbOnTop = Me.name
Call INI_Controls(Me)
End Sub
