VERSION 5.00
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmNewPosWFC 
   Caption         =   "Create New Position"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
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
      Height          =   345
      Left            =   1365
      TabIndex        =   4
      Top             =   2190
      Width           =   1125
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
      Height          =   345
      Left            =   3135
      TabIndex        =   3
      Top             =   2190
      Width           =   1125
   End
   Begin INFOHR_Controls.CodeLookup clpDIv 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpJobMaster 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Tag             =   "01-Job code"
      Top             =   360
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   13
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Tag             =   "00-Position Group  Code"
      Top             =   1320
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGC"
      Enabled         =   0   'False
   End
   Begin VB.Label lblPosGroup 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Group"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1365
      Width           =   1035
   End
   Begin VB.Label lblJob 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Job Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   405
      Width           =   675
   End
   Begin VB.Label lblTitle 
      Caption         =   "Division"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1035
   End
End
Attribute VB_Name = "frmNewPosWFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub clpJobMaster_LostFocus()
    Call getDataFromJobMaster(clpJobMaster.Text)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim xMsg As String

    If Len(Trim(clpJobMaster.Text)) = 0 Then
        MsgBox "Job Code cannot be blank"
        clpJobMaster.SetFocus
        Exit Sub
    ElseIf Len(clpJobMaster.Text) > 0 And clpJobMaster.Caption = "Unassigned" Then
        MsgBox "Job Code is invalid. Please enter a correct Job Code."
        clpJobMaster.SetFocus
        Exit Sub
    Else
        glbWFCNewPosJob = clpJobMaster.Text
    End If
    
    If Len(Trim(clpDIv.Text)) = 0 Then
        MsgBox lStr("Division") & " cannot be blank"
        clpDIv.SetFocus
        Exit Sub
    ElseIf Len(clpDIv.Text) > 0 And clpDIv.Caption = "Unassigned" Then
        MsgBox lStr("Division") & " is invalid."
        clpDIv.SetFocus
        Exit Sub
    Else
        glbWFCNewPosDiv = clpDIv.Text
    End If
    
    If Len(Trim(clpCode(0).Text)) = 0 Then
        MsgBox lStr("Position Group") & " cannot be blank"
        'clpCode(0).SetFocus
        Exit Sub
    ElseIf Len(clpCode(0).Text) > 0 And clpCode(0).Caption = "Unassigned" Then
        MsgBox lStr("Position Group") & " is invalid."
        'clpCode(0).SetFocus
        Exit Sub
    Else
        glbWFCNewPosStatus = clpCode(0).Text
    End If
    
    'Ticket #29069 Franks  08/17/2016 - remove this check
    '''check if there is duplicate record for the same Job and Div
    ''SQLQ = "SELECT * FROM HRJOB WHERE JB_JOBCODE = '" & clpJobMaster.Text & "' "
    ''SQLQ = SQLQ & "AND JB_DIV = '" & clpDIv.Text & "' "
    ''rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    ''If Not rsTemp.EOF Then
    ''    xMsg = "There is a Position Code with the same Job Code('" & clpJobMaster.Text & "') and same " & lStr("Division") & "('" & clpDIv.Text & "') already." & Chr(10)
    ''    xMsg = xMsg & Chr(10) & "    Position Code is: " & rsTemp("JB_CODE") & " "
    ''    xMsg = xMsg & Chr(10) & "    Description is: " & rsTemp("JB_DESCR") & " "
    ''    xMsg = xMsg & Chr(10) & Chr(10) & "Cannot add a Position with same Job Code and " & lStr("Division")
    ''    MsgBox xMsg
    ''    rsTemp.Close
    ''    Exit Sub
    ''Else
    ''    rsTemp.Close
    ''End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    lblTitle(0).Caption = lStr("Division")
    lblPosGroup.Caption = lStr("Position Group")
    
    Call INI_Controls(Me)
End Sub

Private Sub getDataFromJobMaster(xCode)
Dim SQLQ As String
Dim rs As New ADODB.Recordset
    If Len(xCode) > 0 Then
        SQLQ = "SELECT * FROM HRJOBMASTER WHERE JB_JOBCODE = '" & xCode & "' "
        rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rs.EOF Then
            If Not IsNull(rs("JB_GRPCD")) Then clpCode(0).Text = rs("JB_GRPCD")
        End If
        rs.Close
    End If
End Sub

