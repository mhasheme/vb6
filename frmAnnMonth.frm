VERSION 5.00
Begin VB.Form frmAnnMonth 
   Caption         =   "Year End for Anniversary Month"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6255
      Begin VB.ComboBox cmbAnnMonth 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2925
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "Select Anniversary Month"
         Top             =   990
         Width           =   1725
      End
      Begin VB.Label lblEffectDate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2925
         TabIndex        =   8
         Top             =   630
         Width           =   1020
      End
      Begin VB.Label lblEntitlPeriod 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vacation Entitlement Period"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2520
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Label lblAsOf 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1080
         TabIndex        =   6
         Top             =   630
         Width           =   1245
      End
      Begin VB.Label lblPeriod 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vacation Entitlement Period"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Anniversary Month"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   3
         Tag             =   "41-Date Terminated"
         Top             =   1050
         Width           =   1590
      End
   End
   Begin VB.CommandButton cmdOk 
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
      Height          =   375
      Left            =   1867
      TabIndex        =   0
      Top             =   2160
      Width           =   1200
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
      Height          =   375
      Left            =   3427
      TabIndex        =   1
      Top             =   2160
      Width           =   1200
   End
End
Attribute VB_Name = "frmAnnMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbAnnMonth_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdCancel_Click()
    glbAnnMonth = 999
    Unload Me
End Sub

Sub cmdOK_Click()
Dim Response%

glbAnnMonth = 0

If cmbAnnMonth.ListIndex <> 0 And cmbAnnMonth.ListIndex <> -1 Then
    'Check if the User is not trying to update for a Month that is not yet completed or not arrived yet as per the
    'system date.
    If cmbAnnMonth.ListIndex >= month(Now) And Year(CVDate(lblEffectDate)) >= Year(Now) Then
        Response% = MsgBox("You are trying to do an Anniversary Month update for the month of '" & cmbAnnMonth.Text & "' which is not completed yet." & vbCrLf & vbCrLf & "Are you sure you wish to proceed?", vbYesNo + vbExclamation, "info:HR - Year End for Anniversary Month")
        If Response% = IDNO Then
            cmbAnnMonth.SetFocus
            Exit Sub
        End If
    End If
    
    
    'Prompt to confirm
    'Response% = MsgBox("Year End will be done for the Anniversary Month of " & cmbAnnMonth.Text & "." & vbCrLf & "Do you wish to proceed?", vbYesNo + vbQuestion, "info:HR - Year End for Anniversary Month")
    Response% = MsgBox("For all employees whose Anniversary Month is '" & cmbAnnMonth.Text & "', their Outstanding Amount will be moved into the Previous Year column and the Current Year column is Zeroed Out and updated with the next year's entitlement." & vbCrLf & vbCrLf & "Do you wish to proceed?", vbYesNo + vbQuestion, "info:HR - Year End for Anniversary Month")
    
    If Response% = IDNO Then
        cmbAnnMonth.SetFocus
        Exit Sub
    End If
End If

Screen.MousePointer = HOURGLASS

glbAnnMonth = cmbAnnMonth.ListIndex

Screen.MousePointer = DEFAULT

Unload Me

End Sub

Private Sub comAnnMonthAdding()
    cmbAnnMonth.Clear
    cmbAnnMonth.AddItem ""
    cmbAnnMonth.AddItem "January"
    cmbAnnMonth.AddItem "February"
    cmbAnnMonth.AddItem "March"
    cmbAnnMonth.AddItem "April"
    cmbAnnMonth.AddItem "May"
    cmbAnnMonth.AddItem "June"
    cmbAnnMonth.AddItem "July"
    cmbAnnMonth.AddItem "August"
    cmbAnnMonth.AddItem "September"
    cmbAnnMonth.AddItem "October"
    cmbAnnMonth.AddItem "November"
    cmbAnnMonth.AddItem "December"
End Sub

Private Sub Form_Load()

'Ticket #22893 - WHSC customization

'Show the Entitlement Period and Effective Date of the rule being updated.
lblEntitlPeriod.Caption = frmSVacEnt.dlpDateRange(0) & " - " & frmSVacEnt.dlpDateRange(1)
lblEffectDate.Caption = frmSVacEnt.dlpAsOf

Call comAnnMonthAdding

End Sub

