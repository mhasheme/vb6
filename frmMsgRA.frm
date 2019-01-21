VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmMsgRA 
   Caption         =   "Reporting Authority to Replace With"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   1275
      Width           =   7350
      _Version        =   65536
      _ExtentX        =   12965
      _ExtentY        =   979
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
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
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
         Left            =   3606
         TabIndex        =   3
         Tag             =   "Save changes made"
         Top             =   0
         Width           =   960
      End
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
         Left            =   1929
         TabIndex        =   2
         Tag             =   "Save changes made"
         Top             =   0
         Width           =   960
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8490
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   3
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "ED_LUSER"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   5580
      MaxLength       =   25
      TabIndex        =   6
      Top             =   5250
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "ED_LTIME"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   5220
      MaxLength       =   25
      TabIndex        =   5
      Top             =   5250
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "ED_LDATE"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   4950
      MaxLength       =   25
      TabIndex        =   4
      Top             =   5250
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame frmBasic 
      BorderStyle     =   0  'None
      Height          =   4305
      Left            =   -90
      TabIndex        =   0
      Top             =   -30
      Width           =   8235
      Begin INFOHR_Controls.EmployeeLookup elpRept 
         Height          =   285
         Index           =   0
         Left            =   2640
         TabIndex        =   8
         Tag             =   "10-Reporting Authority"
         Top             =   555
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin VB.Label lblWFCTermNo 
         Caption         =   "TermNo"
         Height          =   135
         Left            =   4560
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblEmpNum 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NEW Reporting Authority #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   7
         Top             =   600
         Width           =   2220
      End
   End
End
Attribute VB_Name = "frmMsgRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Dim Response%

If elpRept(0).Text = "" Then
    'Ticket #29648 - Correcting the logic as this standard logic was commented out by Frank when adding the WFC logic. I have uncommented the standard logic and put Frank's logic under WFC after
    'discussing with him.
    If Not glbWFC Then
        Response% = MsgBox("No new Reporting Authority selected to replace this employee. " & vbCrLf & vbCrLf & "Are you sure you want to proceed?", vbQuestion + vbYesNo, "No new Reporting Authority")
        If Response% = IDNO Then
            Exit Sub
        End If
    Else
        glbMsgCustomVal = 20
        frmMsgDialog.Show 1
        'if glbMsgCustomVal = 1 then 'If <<Continue>> is checked, save the record with the incorrect RA#1.
        If glbMsgCustomVal = 1 Then 'If <<Update> is click, enter the Rept again.
            elpRept(0).SetFocus
            Exit Sub
        Else
            ' "Leave Blank" is click
        End If
    
        glbNewRept = ""
    End If
End If
    
    glbNewRept = ""
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim Response%
Dim xPosCode

If Not elpRept(0).ListChecker Then
    Exit Sub
End If

If glbWFC Then 'Ticket #29343 Franks 10/24/2016
    If elpRept(0).Text = "" Then
        'Response% = MsgBox("Leaving the Interim/New Reporting Authority blank may cause a break in the organization chain." & vbCrLf & vbCrLf & "Do you want to proceed?", vbQuestion + vbYesNo, "No Interim/New Reporting Authority")
        'If Response% = IDNO Then
        '    Exit Sub
        'End If
        glbMsgCustomVal = 20
        frmMsgDialog.Show 1
        'if glbMsgCustomVal = 1 then 'If <<Continue>> is checked, save the record with the incorrect RA#1.
        If glbMsgCustomVal = 1 Then 'If <<Update> is click, enter the Rept again.
            elpRept(0).SetFocus
            Exit Sub
        Else
            ' "Leave Blank" is click
        End If
        
        glbNewRept = ""
    Else
        If locWFCEmpID = lblWFCTermNo Then  'From WFC Termination
            'check if is the correct RA based on locWFCEmpID
            xPosCode = getEmpPosFromReptNo1(locWFCEmpID) 'Ticket #29484 Franks 11/29/2016

            If Len(xPosCode) > 0 Then
                If IsRept1PosNotMatchPosMaster(elpRept(0).Text, xPosCode) Then
                    glbMsgCustomVal = 11
                    frmMsgDialog.Show 1
                    'if glbMsgCustomVal = 1 then 'If <<Continue>> is checked, save the record with the incorrect RA#1.
                    If glbMsgCustomVal = 2 Then 'If <<Cancel>> is checked, undo the change.
                        'elpRept(0).Text = GetReportingAuth1EmpNoBasePosMaster(xPosCode)
                        Exit Sub
                    End If
                End If
            End If
            glbNewRept = elpRept(0).Text
        Else
            glbNewRept = elpRept(0).Text
        End If
    End If
Else
    If elpRept(0).Text = "" Then
        Response% = MsgBox("No new Reporting Authority selected to replace this employee. " & vbCrLf & vbCrLf & "Are you sure you want to proceed?", vbQuestion + vbYesNo, "No new Reporting Authority")
        If Response% = IDNO Then
            Exit Sub
        End If
        glbNewRept = ""
    Else
        glbNewRept = elpRept(0).Text
    End If
End If

end_line:
    Unload Me
End Sub

Private Sub elpRept_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Load()

MDIMain.panHelp(0).Caption = "info:HR Message"

If glbWFC Then 'Ticket #29343 Franks 10/24/2016
    lblEmpNum.Caption = "Interim/New Reporting Authority"
End If

Call INI_Controls(Me)

glbNewRept = ""

End Sub

