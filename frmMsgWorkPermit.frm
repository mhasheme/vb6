VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmMsgWorkPermit 
   Caption         =   "Work Permit Info."
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   5
      Top             =   1275
      Width           =   6885
      _Version        =   65536
      _ExtentX        =   12144
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   5250
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame frmBasic 
      BorderStyle     =   0  'None
      Height          =   4305
      Left            =   -90
      TabIndex        =   4
      Top             =   -30
      Width           =   8235
      Begin VB.TextBox txtPermit 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         MaxLength       =   30
         TabIndex        =   0
         Tag             =   "00-Visa/Work permit #"
         Top             =   360
         Width           =   3360
      End
      Begin INFOHR_Controls.DateLookup dlpPermitDate 
         Height          =   285
         Left            =   3165
         TabIndex        =   1
         Tag             =   "40-Visa/Work Permit Expiration Date"
         Top             =   725
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin VB.Label lblWorkPermitNo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Visa/Work Permit #"
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
         Left            =   240
         TabIndex        =   10
         Top             =   405
         Width           =   1680
      End
      Begin VB.Label lblWorkPermitDate 
         AutoSize        =   -1  'True
         Caption         =   "Visa/Work Permit Expiration Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   765
         Width           =   2865
      End
   End
End
Attribute VB_Name = "frmMsgWorkPermit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Dim Response%

If Len(Trim(txtPermit.Text)) = 0 Or Len(Trim(dlpPermitDate.Text)) = 0 Then
    Response% = MsgBox("You will not be able to save employee information without the 'Visa/Work Permit #' and 'Visa/Work Permit Expiration Date'. " & vbCrLf & vbCrLf & "Are you sure you want to proceed?", vbQuestion + vbYesNo, "No Work Permit Info.")
    If Response% = IDNO Then
        Exit Sub
    End If
End If
    
glbWorkVisaNo = ""
glbWorkExpDate = ""
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim Response%

If glbLinamar Then 'Ticket #28875 Franks 07/13/2016
    'Linamar only need the Visa/Work Permit Expiration Date
Else
    If Len(Trim(txtPermit.Text)) = 0 Or Len(Trim(dlpPermitDate.Text)) = 0 Then
        MsgBox "'Visa/Work Permit #' and 'Visa/Work Permit Expiration Date' cannot be blank."
        If Len(Trim(txtPermit.Text)) = 0 Then
            txtPermit.SetFocus
        Else
            dlpPermitDate.SetFocus
        End If
        Exit Sub
    End If
End If
If Len(dlpPermitDate) > 0 Then
    If Not IsDate(dlpPermitDate) Then
        MsgBox "Invalid Visa/Work Permit Expiration Date"
        dlpPermitDate.SetFocus
        Exit Sub
    End If
End If

glbWorkVisaNo = txtPermit
glbWorkExpDate = dlpPermitDate

end_line:
    Unload Me
End Sub

Private Sub elpRept_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Load()

MDIMain.panHelp(0).Caption = "info:HR Message"

Call INI_Controls(Me)

'Show whatever info already available
If Len(glbWorkVisaNo) > 0 Then
    txtPermit.Text = glbWorkVisaNo
End If
If Len(glbWorkExpDate) > 0 Then
    dlpPermitDate.Text = glbWorkExpDate
End If
If glbLinamar Then 'Ticket #28875 Franks 07/13/2016
    lblWorkPermitNo.Visible = False
    txtPermit.Visible = False
End If

End Sub

