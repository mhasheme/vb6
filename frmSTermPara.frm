VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmSTermPara 
   Caption         =   "Termination Data"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2775
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEmpNum 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2480
      MaxLength       =   9
      TabIndex        =   2
      Tag             =   "11-New Employee Number"
      Top             =   1200
      Width           =   1185
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Tag             =   "41-Termination Code "
      Top             =   720
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
      Top             =   360
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   7
      Top             =   2115
      Width           =   6570
      _Version        =   65536
      _ExtentX        =   11589
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
         Left            =   120
         TabIndex        =   4
         Tag             =   "Save the changes made"
         Top             =   120
         Width           =   735
      End
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "SH_PAYP"
      Height          =   285
      Index           =   4
      Left            =   2160
      TabIndex        =   3
      Tag             =   "00-Enter pay period code"
      Top             =   1560
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDPP"
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NEW Pay Period"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   12
      Left            =   120
      TabIndex        =   10
      Top             =   1605
      Width           =   1200
   End
   Begin VB.Label lblEmpNum 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NEW Employee #"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1245
      Width           =   2340
   End
   Begin VB.Label lblEEName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Message to user"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   3900
      TabIndex        =   8
      Top             =   1222
      Visible         =   0   'False
      Width           =   1500
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
      TabIndex        =   6
      Top             =   765
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
      TabIndex        =   5
      Tag             =   "41-Date Terminated"
      Top             =   405
      Width           =   1470
   End
End
Attribute VB_Name = "frmSTermPara"
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
    
    If Len(txtEmpNum) = 0 Then
        MsgBox "The Employee number must be entered"
        txtEmpNum.SetFocus
        Exit Sub
    End If
    If lblEEName.Visible = True Then 'the employee number to change to already exixts
        MsgBox "The NEW Employee number already exits"
        txtEmpNum.SetFocus
        Exit Sub
    End If
    
    'Ticket #25553 - Pay Period/Company Code change causes Termination and New Hire
    If Not clpCode(4).ListChecker Then
        Exit Sub
    End If
    
    If Len(clpCode(4).Text) = 0 Then
        MsgBox lStr("Pay Period Code is required")
        clpCode(4).SetFocus
        Exit Sub
    End If
    
    glbSPCTermDate = dlpTermDate
    glbSPCTermReason = clpCode(1)
    glbSPCNewEmpNo = txtEmpNum
    
    'Ticket #25553 - Pay Period/Company Code change causes Termination and New Hire
    glbSPCPPay = Trim(clpCode(4).Text)
    
    Unload Me
End Sub

Private Sub Form_Load()
    Call INI_Controls(Me)
    
    'Ticket #25553 - Pay Period/Company Code change also causes New Hire in the new company and
    'termination in the existing company. Pay Period/Company Code change also causes the EI Code change that's
    'why putting the Pay Period here. Later in future they may ask us to disable this because there will
    'no longer be any more transfers, from JQP to HSM. The JQP is the older company, everyone moving to HSM now.
    Call setCaption(lblTitle(12))
    
    'Retrieve the current Pay Period of the employee
    clpCode(4).Text = GetSHData(glbLEE_ID, "SH_PAYP", "")
End Sub

Private Sub txtEmpNum_Change()
    Dim rsEMP As New ADODB.Recordset
    lblEEName = ""
    lblEEName.Visible = False
    If Len(txtEmpNum) > 0 Then
        rsEMP.Open "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR=" & txtEmpNum, gdbAdoIhr001, adOpenForwardOnly
        If Not rsEMP.EOF Then
            lblEEName = "This number already exists"
            lblEEName.Visible = True
        End If
    End If
End Sub

Private Sub txtEmpNum_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtEmpNum_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
