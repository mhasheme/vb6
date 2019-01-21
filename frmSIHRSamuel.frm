VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSIHRSamuel 
   Caption         =   "Samuel Securities"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmDetail 
      BorderStyle     =   0  'None
      Caption         =   "C.A.R.S. Administration Report"
      Height          =   6795
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   8235
      Begin VB.CommandButton cmdGrantAll 
         Appearance      =   0  'Flat
         Caption         =   "&Grant All"
         Height          =   330
         Left            =   3000
         TabIndex        =   12
         Top             =   5760
         Width           =   1305
      End
      Begin VB.CommandButton cmdGrantInqu 
         Appearance      =   0  'Flat
         Caption         =   "Grant All &Inquire"
         Height          =   330
         Left            =   1500
         TabIndex        =   11
         Tag             =   "Grant All Basic"
         Top             =   5760
         Width           =   1320
      End
      Begin VB.CommandButton cmdRemoveAll 
         Appearance      =   0  'Flat
         Caption         =   "&Remove All"
         Height          =   330
         Left            =   240
         TabIndex        =   10
         Tag             =   "Grant All Basic"
         Top             =   5760
         Width           =   1200
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   0
         Left            =   375
         TabIndex        =   13
         Top             =   1680
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Profit Sharing"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkMSecurity 
         Height          =   225
         Index           =   0
         Left            =   375
         TabIndex        =   14
         Top             =   825
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkSecurity 
         Bindings        =   "frmSIHRSamuel.frx":0000
         Height          =   225
         Index           =   0
         Left            =   1350
         TabIndex        =   15
         Top             =   825
         Width           =   3885
         _Version        =   65536
         _ExtentX        =   6853
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Profit Sharing"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   1
         Left            =   375
         TabIndex        =   18
         Top             =   1980
         Width           =   2685
         _Version        =   65536
         _ExtentX        =   4736
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Red Circled Report"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   30
         Left            =   360
         TabIndex        =   23
         Top             =   4080
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Profit Sharing"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   31
         Left            =   360
         TabIndex        =   25
         Top             =   3720
         Width           =   3465
         _Version        =   65536
         _ExtentX        =   6112
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Position && Salary Historical IDL"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   32
         Left            =   360
         TabIndex        =   26
         Top             =   3360
         Width           =   3585
         _Version        =   65536
         _ExtentX        =   6324
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "NOC and Employee Dates Update"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   2
         Left            =   375
         TabIndex        =   27
         Top             =   2280
         Width           =   3525
         _Version        =   65536
         _ExtentX        =   6218
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Missing Reporting Authorities Report"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   3
         Left            =   375
         TabIndex        =   28
         Top             =   2550
         Width           =   3165
         _Version        =   65536
         _ExtentX        =   5583
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Samuel Audit Report"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   4
         Left            =   2520
         TabIndex        =   30
         Top             =   0
         Width           =   3165
         _Version        =   65536
         _ExtentX        =   5583
         _ExtentY        =   397
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkMSecurity 
         Height          =   225
         Index           =   1
         Left            =   375
         TabIndex        =   16
         Top             =   1080
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkSecurity 
         Bindings        =   "frmSIHRSamuel.frx":000B
         Height          =   225
         Index           =   1
         Left            =   1350
         TabIndex        =   17
         Top             =   1080
         Width           =   3885
         _Version        =   65536
         _ExtentX        =   6853
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Table Master Edit Links"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkMSecurity 
         Height          =   225
         Index           =   2
         Left            =   375
         TabIndex        =   31
         Top             =   5265
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkSecurity 
         Bindings        =   "frmSIHRSamuel.frx":0016
         Height          =   225
         Index           =   2
         Left            =   1350
         TabIndex        =   32
         Top             =   5265
         Width           =   3885
         _Version        =   65536
         _ExtentX        =   6853
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Hours && Rept Auth Setup"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   33
         Left            =   360
         TabIndex        =   36
         Top             =   4440
         Width           =   3585
         _Version        =   65536
         _ExtentX        =   6324
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Import of Pension Data"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Setup"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   35
         Top             =   4800
         Width           =   420
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maintain"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   375
         TabIndex        =   34
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inquire"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1380
         TabIndex        =   33
         Top             =   5040
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Show Custom Features Menu"
         Height          =   195
         Left            =   90
         TabIndex        =   29
         Top             =   0
         Width           =   2085
      End
      Begin VB.Label Label8 
         Caption         =   "Import"
         Height          =   375
         Left            =   90
         TabIndex        =   24
         Top             =   3120
         Width           =   1875
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Reports"
         Height          =   195
         Left            =   90
         TabIndex        =   22
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inquire"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1380
         TabIndex        =   21
         Top             =   600
         Width           =   600
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maintain"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   375
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maintenance"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   90
         TabIndex        =   19
         Top             =   360
         Width           =   930
      End
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10230
      _Version        =   65536
      _ExtentX        =   18045
      _ExtentY        =   873
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BevelInner      =   2
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Descr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3030
         TabIndex        =   3
         Top             =   120
         Width           =   630
      End
      Begin VB.Label lblUSERID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ABCD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1320
         TabIndex        =   2
         Top             =   125
         Width           =   630
      End
      Begin VB.Label lblPosl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   135
         Width           =   660
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   8115
      Width           =   10230
      _Version        =   65536
      _ExtentX        =   18045
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
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2835
         TabIndex        =   8
         Tag             =   "Cancel the changes made"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2010
         TabIndex        =   7
         Tag             =   "Save the changes made"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   315
         TabIndex        =   6
         Tag             =   "Close and exit this screen"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1140
         TabIndex        =   5
         Tag             =   "Edit the information "
         Top             =   0
         Width           =   765
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   405
         Left            =   4200
         Top             =   180
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   714
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmSIHRSamuel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fglbEditMode%
Dim fUPMode As Integer, fglbEmptyNew As Integer

Private Sub chkMSecurity_Click(Index As Integer, Value As Integer)
If chkMSecurity(Index).Value = True Then
    chkSecurity(Index).Value = True
End If
End Sub

Private Sub chkSecurity_Click(Index As Integer, Value As Integer)
    If chkSecurity(Index).Value = False Then
        chkMSecurity(Index).Value = False
    End If
End Sub

Private Sub cmdCancel_Click()

On Error GoTo Can_Err

Call Display_Values

Call ST_UPD_MODE(False)  ' reset screen's attributes


Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdGrantAll_Click()
Dim x%
For x% = 0 To 2  '0
    chkMSecurity(x%).Value = 1
Next x%
For x% = 0 To 2 '0
    chkSecurity(x%).Value = 1
Next x%

For x% = 0 To 4 '3 '1
    chkLSecurity(x%).Value = 1
Next x%
For x% = 30 To 33
    chkLSecurity(x%).Value = 1
Next x%
End Sub

Private Sub cmdGrantInqu_Click()
Dim x%

For x% = 0 To 2 '1 '0
    chkSecurity(x%).Value = 1
Next x%

For x% = 0 To 4 '3 '1
    chkLSecurity(x%).Value = 1
Next x%
For x% = 30 To 33
    chkLSecurity(x%).Value = 1
Next x%

End Sub

Private Sub cmdModify_Click()
Dim SQLQ As String

If Not gSec_Upd_Security Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Call ST_UPD_MODE(True)

On Error GoTo Edit_Err

chkMSecurity(0).SetFocus

Exit Sub
Edit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdEdit", "SamuelSecurity", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub cmdOK_Click()
Dim x%
Dim xID
Dim xTemplate As String

On Error GoTo OK_Err

Call ST_UPD_MODE(False)

'Ticket #20585 - If Template then update users with this template as well.
'If User and with no template then update that user's profile.
'if User and with Template then do not update user's profile.
'Get the Template Name of this User ID
xTemplate = Get_Template(glbSecUSERID)

If xTemplate = "TEMPLATE" Then
    'Update all users with this template. After the changes are saved
ElseIf xTemplate = "" Then
    'User - User with no template - don't do anything let system update user's profile
ElseIf xTemplate <> "TEMPLATE" Then
    'User with template - do not allow to save these changes.
    MsgBox "Security change cannot be saved. This user's security profile is based on the '" & xTemplate & "' template.", vbExclamation, "Template based User Security Profile"
    
    'Redisplay the security settings
    Call Display_Values
End If

'Template or User only
If xTemplate = "TEMPLATE" Or xTemplate = "" Then
    Call UpdSecAccess
End If

If xTemplate = "TEMPLATE" Then
    '????Ticket #24808 - User's based on this Template does not need their Profile to be updated as we are now retrieving Template profile for the users
    'Call procedure to Update all users with this template.
    'Call Update_Users_withthis_Template(glbSecUSERID)
End If

fglbEditMode% = False

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "Pension Security", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdRemoveAll_Click()
Dim x%

For x% = 0 To 2 '0
    chkMSecurity(x%).Value = 0
Next x%
For x% = 0 To 2 '0
    chkSecurity(x%).Value = 0
Next x%

For x% = 0 To 4 '3 '1
    chkLSecurity(x%).Value = 0
Next x%
For x% = 30 To 33
    chkLSecurity(x%).Value = 0
Next x%

End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim x%

glbOnTop = Me.name

Screen.MousePointer = HOURGLASS
lblUSERID.Caption = glbSecUSERID
lblEEName.Caption = glbSecEEName

'frmSIHRWFCPen.Show
Me.Caption = ("Samuel Securities - ") & lblEEName
'Data1.ConnectionString = glbAdoIHRDB
'Data1.RecordSource = "select * from HR_SECURE_ACCESS where USERID='" & glbSecUSERID & "' AND LEFT([FUNCTION],4)='SAM_'"
'Data1.Refresh

Call Display_Values

Call ST_UPD_MODE(False)

Screen.MousePointer = DEFAULT

End Sub

Private Sub Display_Values()
Dim rsSR As New ADODB.Recordset
Dim x%, SQLQ
Dim xTemplate  As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbSecUSERID)

If xTemplate = "" Or xTemplate = "TEMPLATE" Then
    SQLQ = "select * from HR_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND LEFT([FUNCTION],4)='SAM_'"
Else
    '????Ticket #24808 -  Retrieve template's security profile
    SQLQ = "select * from HR_SECURE_ACCESS WHERE USERID='" & Replace(xTemplate, "'", "''") & "' AND LEFT([FUNCTION],4)='SAM_'"
End If
rsSR.Open SQLQ, gdbAdoIhr001, adOpenStatic

Call ResetAll

Do Until rsSR.EOF
    'Show Custom Features Menu
    If UCase(rsSR("FUNCTION")) = UCase("SAM_Show_CustomFeatures") Then chkLSecurity(4) = rsSR("ACCESSABLE")
    
    'Maintain
    If UCase(rsSR("FUNCTION")) = UCase("SAM_Profit_Sharing_Upt") Then chkMSecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("SAM_Table_Master_Links_Upt") Then chkMSecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("SAM_Hours_ReptAuth_Upt") Then chkMSecurity(2) = rsSR("ACCESSABLE")
    
    'Inquire
    If UCase(rsSR("FUNCTION")) = UCase("SAM_Profit_Sharing_Inq") Then chkSecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("SAM_Table_Master_Links_Inq") Then chkSecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("SAM_Hours_ReptAuth_Inq") Then chkSecurity(2) = rsSR("ACCESSABLE")
    
    'report
    If UCase(rsSR("FUNCTION")) = UCase("SAM_Profit_Sharing_Rpt") Then chkLSecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("SAM_Red_Circled_Rpt") Then chkLSecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("SAM_Missing_Auth_Rpt") Then chkLSecurity(2) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("SAM_Tran_Change_Rpt") Then chkLSecurity(3) = rsSR("ACCESSABLE")

    'Import
    If UCase(rsSR("FUNCTION")) = UCase("SAM_Profit_Sharing_Imp") Then chkLSecurity(30) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("SAM_PosSal_IDL_Imp") Then chkLSecurity(31) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("SAM_NOC_Emp_Imp") Then chkLSecurity(32) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("SAM_Pension_Data_Imp") Then chkLSecurity(33) = rsSR("ACCESSABLE")
    
    rsSR.MoveNext
Loop

End Sub

Private Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

glbOHSEdit% = TF

fUPMode = TF    ' update mode
cmdOK.Enabled = TF
cmdModify.Enabled = FT
cmdCancel.Enabled = TF
cmdClose.Enabled = FT
frmDetail.Enabled = TF
End Sub

Private Sub UpdSecAccess()
Dim SQLQ

SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND LEFT([FUNCTION],4)='SAM_'"
gdbAdoIhr001.Execute SQLQ

Call AddSecAccess

End Sub

Private Sub AddSecAccess()
Dim SQLQ, sqlI

sqlI = "INSERT INTO HR_SECURE_ACCESS(COMPNO,USERID,[FUNCTION],ACCESSABLE) "
sqlI = sqlI & " VALUES('001','" & Replace(Trim(lblUSERID), "'", "''") & "',"

'Show Custom Features Menu
SQLQ = sqlI & "'SAM_Show_CustomFeatures'," & IIf(chkLSecurity(4), 1, 0) & ")" '3 -> 4
gdbAdoIhr001.Execute SQLQ

'Maintain
SQLQ = sqlI & "'SAM_Profit_Sharing_Upt'," & IIf(chkMSecurity(0), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'SAM_Table_Master_Links_Upt'," & IIf(chkMSecurity(1), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'SAM_Hours_ReptAuth_Upt'," & IIf(chkMSecurity(2), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
'Inquire
SQLQ = sqlI & "'SAM_Profit_Sharing_Inq'," & IIf(chkSecurity(0), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'SAM_Table_Master_Links_Inq'," & IIf(chkSecurity(1), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'SAM_Hours_ReptAuth_Inq'," & IIf(chkSecurity(2), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'report
SQLQ = sqlI & "'SAM_Profit_Sharing_Rpt'," & IIf(chkLSecurity(0), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'SAM_Red_Circled_Rpt'," & IIf(chkLSecurity(1), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'SAM_Missing_Auth_Rpt'," & IIf(chkLSecurity(2), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'SAM_Tran_Change_Rpt'," & IIf(chkLSecurity(3), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Import
SQLQ = sqlI & "'SAM_Profit_Sharing_Imp'," & IIf(chkLSecurity(30), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'SAM_PosSal_IDL_Imp'," & IIf(chkLSecurity(31), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'SAM_NOC_Emp_Imp'," & IIf(chkLSecurity(32), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'SAM_Pension_Data_Imp'," & IIf(chkLSecurity(33), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

End Sub

Private Sub ResetAll()
Dim x%

For x% = 0 To 1 '0
    chkMSecurity(x%).Value = 0
Next x%
For x% = 0 To 1 '0
    chkSecurity(x%).Value = 0
Next x%

For x% = 0 To 4 '3 '1
    chkLSecurity(x%).Value = 0
Next x%
For x% = 30 To 32
    chkLSecurity(x%).Value = 0
Next x%

End Sub

Private Sub Update_Users_withthis_Template(xTemplate)
    Dim SQLQ As String
    Dim rsSecBasic As New ADODB.Recordset
    
    'Retrieve all users associated with this changed Template
    SQLQ = "SELECT USERID, SECURE_TEMPLATE FROM HR_SECURE_BASIC WHERE SECURE_TEMPLATE = '" & xTemplate & "'"
    SQLQ = SQLQ & " ORDER BY USERID"
    rsSecBasic.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsSecBasic.EOF
        If Not IsNull(rsSecBasic("USERID")) Then
            'Update each user with this changed Template
            Call SpecificFunction_Template_Based_Security_Profile_Update(rsSecBasic("USERID"), xTemplate, "Change", "CUSTOMFEATURE")
        End If
        rsSecBasic.MoveNext
    Loop
    rsSecBasic.Close
    Set rsSecBasic = Nothing
    
End Sub

