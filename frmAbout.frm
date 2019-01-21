VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About info:HR"
   ClientHeight    =   6510
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label lblinfoHRExpiry 
      Alignment       =   2  'Center
      Caption         =   "info:HR Expiration Date: "
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   1725
      Left            =   660
      Picture         =   "frmAbout.frx":0000
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "General Information:  www.infohr.com"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5520
      Width           =   5295
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Fax         (416) 599-5031"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5160
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "   E-Mail         support@infohr.com"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Width           =   5295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "               1-800-567-4254"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   5295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "For support call  (416) 599-4747"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   5295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Crystal Reports Version 8.0"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   5295
   End
   Begin VB.Label lblDatabase 
      Alignment       =   2  'Center
      Caption         =   "Database"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   5295
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "Version"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   5295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "All Rights Reserved."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   5295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Copyright © info:HR 2017"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   5295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Toronto, ON"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "HR Systems Strategies Inc."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   5295
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Load()
Dim Version As String
Dim EXEDate As Date
Dim Msg As String

If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHR.EXE") <> "" Then
    EXEDate = FileDateTime(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHR.EXE")
ElseIf Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHRDEMO.EXE") <> "" Then
    EXEDate = FileDateTime(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHRDEMO.EXE")
End If

Version = App.Major & "." & App.Minor & "." & App.Revision
Msg = "Windows Version " & Version & "  " & Format(EXEDate, "mmmm d, yyyy")
lblVersion = Msg

If glbOracle Then
    Msg = "ORACLE Server 9i Product"
ElseIf glbSQL Then
    Msg = "MS SQL Server 2008/2012/2014 Product"
Else
    Msg = "MS Access 2002 Database Product"
End If
lblDatabase = Msg

'Show License Key/Expiry Date for Hosted Environment else show blank date
If glbHosted Then
    If glbLicenseKey = Format("12/31/2028", "mm/dd/yyyy") Then
        lblinfoHRExpiry.Caption = "info:HR Expiration Date: "
    Else
        lblinfoHRExpiry.Caption = "info:HR Expiration Date: " & Format(glbLicenseKey, "mmmm d, yyyy")
    End If
Else
    lblinfoHRExpiry.Caption = "info:HR Expiration Date: "
End If
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub
