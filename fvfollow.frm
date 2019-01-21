VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmvFOLOWUP 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Follow-Up Overview"
   ClientHeight    =   7455
   ClientLeft      =   165
   ClientTop       =   705
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7455
   ScaleWidth      =   11400
   WindowState     =   1  'Minimized
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   29
      Top             =   10545
      Width           =   15240
      _Version        =   65536
      _ExtentX        =   26882
      _ExtentY        =   820
      _StockProps     =   15
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
      Begin VB.CommandButton cmdManagerSec 
         Appearance      =   0  'Flat
         Caption         =   "&Filter by Manager's Security"
         Height          =   375
         Left            =   8280
         TabIndex        =   14
         Tag             =   "Filter by Reporting Authority 1 or 2"
         Top             =   0
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.CommandButton cmdComSelected 
         Appearance      =   0  'Flat
         Caption         =   "Mass Flag Comp."
         Height          =   375
         Left            =   4560
         TabIndex        =   45
         Tag             =   "Remove all messages listed above."
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelSelected 
         Appearance      =   0  'Flat
         Caption         =   "Mass Delete"
         Height          =   375
         Left            =   6240
         TabIndex        =   44
         Tag             =   "Remove all messages listed above."
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdShowSearch 
         Appearance      =   0  'Flat
         Caption         =   "&Search Criteria"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Tag             =   "Show parameters used in finding messages"
         Top             =   0
         Width           =   1605
      End
      Begin VB.CommandButton cmdMarkAll 
         Appearance      =   0  'Flat
         Caption         =   "&Flag All Comp."
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Tag             =   "Mark all messages as being completed."
         Top             =   0
         Width           =   1380
      End
      Begin VB.CommandButton cmdMassDelete 
         Appearance      =   0  'Flat
         Caption         =   "Delete &All "
         Height          =   375
         Left            =   3240
         TabIndex        =   22
         Tag             =   "Remove all messages listed above."
         Top             =   0
         Width           =   1215
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8760
         Top             =   165
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.Frame frmFind 
      BorderStyle     =   0  'None
      Height          =   1755
      Left            =   240
      TabIndex        =   40
      Top             =   4920
      Width           =   9735
      Begin VB.CheckBox chkMyFollow 
         Alignment       =   1  'Right Justify
         Caption         =   "Show My Follow-ups"
         Height          =   225
         Left            =   180
         TabIndex        =   7
         Top             =   1020
         Width           =   2595
      End
      Begin VB.CommandButton cmdEESort 
         Appearance      =   0  'Flat
         Caption         =   "Sort by Reason"
         Height          =   375
         Index           =   3
         Left            =   6120
         TabIndex        =   12
         Tag             =   "Change the sorting method of the Employee List"
         Top             =   1380
         Width           =   2600
      End
      Begin VB.CommandButton cmdEESort 
         Appearance      =   0  'Flat
         Caption         =   "&Sort by Emp #"
         Height          =   375
         Index           =   0
         Left            =   6120
         TabIndex        =   9
         Tag             =   "Change the sorting method of the Employee List"
         Top             =   120
         Width           =   2600
      End
      Begin VB.CommandButton cmdEESort 
         Appearance      =   0  'Flat
         Caption         =   "&Sort by Surname"
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   10
         Tag             =   "Change the sorting method of the Employee List"
         Top             =   540
         Width           =   2600
      End
      Begin VB.CommandButton cmdEESort 
         Appearance      =   0  'Flat
         Caption         =   "Sort by Effective Date"
         Height          =   375
         Index           =   2
         Left            =   6120
         TabIndex        =   11
         Tag             =   "Change the sorting method of the Employee List"
         Top             =   960
         Width           =   2600
      End
      Begin VB.TextBox txtEESearch 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         TabIndex        =   6
         Tag             =   "00-Search Field"
         Top             =   480
         Width           =   1965
      End
      Begin VB.CommandButton cmdFind 
         Appearance      =   0  'Flat
         Caption         =   "&Find"
         Height          =   375
         Left            =   4980
         TabIndex        =   8
         Tag             =   "Find Employee"
         Top             =   420
         Width           =   735
      End
      Begin INFOHR_Controls.DateLookup dlpEESearch 
         Height          =   315
         Left            =   2760
         TabIndex        =   42
         Top             =   450
         Visible         =   0   'False
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   556
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpEESearch 
         Height          =   345
         Left            =   2040
         TabIndex        =   43
         Top             =   450
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   609
         ShowUnassigned  =   1
         TABLName        =   "FURE"
         Object.Height          =   345
      End
      Begin VB.Label lblSearchBy 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Search by Surname"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   480
         Width           =   1935
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fvfollow.frx":0000
      Height          =   2025
      Left            =   0
      OleObjectBlob   =   "fvfollow.frx":0014
      TabIndex        =   0
      Top             =   240
      Width           =   9135
   End
   Begin Threed.SSCommand cmdTLAY 
      Height          =   435
      Left            =   360
      TabIndex        =   13
      Top             =   6630
      Visible         =   0   'False
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   767
      _StockProps     =   78
      Caption         =   "View Temporary Lay Offs"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   8280
      Top             =   6600
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   3
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
      Caption         =   "Adodc2"
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
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
      DataField       =   "EF_COMMENTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Tag             =   "60-Comments "
      Top             =   3720
      Width           =   8565
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EF_LDATE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2610
      MaxLength       =   25
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6300
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EF_LTIME"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   4290
      MaxLength       =   25
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6300
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EF_LUSER"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   5970
      MaxLength       =   25
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6300
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSCheck chkCompleted 
      DataField       =   "EF_COMPLETED"
      Height          =   195
      Left            =   5520
      TabIndex        =   4
      Tag             =   "00-Followup Completed"
      Top             =   3090
      Width           =   1620
      _Version        =   65536
      _ExtentX        =   2857
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "Completed            "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EF_ADMINBY"
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   3
      Tag             =   "00-Enter Administered By Code"
      Top             =   3120
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EF_FREAS"
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Tag             =   "01-Followup Reason Code"
      Top             =   2760
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "FURE"
   End
   Begin INFOHR_Controls.DateLookup dlpEDate 
      DataField       =   "EF_FDATE"
      Height          =   285
      Left            =   6480
      TabIndex        =   2
      Tag             =   "41-Effective Date of Followup"
      Top             =   2760
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin VB.Frame frmSearch 
      Caption         =   "Search Criteria"
      Height          =   1005
      Left            =   300
      TabIndex        =   30
      Top             =   5250
      Visible         =   0   'False
      Width           =   9795
      Begin VB.TextBox txtDaysForward 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6240
         TabIndex        =   34
         Tag             =   "11-Number of Days forward to search for."
         Top             =   270
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton cmdReScan 
         Appearance      =   0  'Flat
         Caption         =   "Re-Scan"
         Height          =   330
         Left            =   6960
         TabIndex        =   33
         Tag             =   "Search again given new search parameters"
         Top             =   270
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         Caption         =   "&Save Changes"
         Height          =   330
         Left            =   6960
         TabIndex        =   32
         Tag             =   "Save Search Parameters"
         Top             =   600
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton cmdHide 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Height          =   645
         Left            =   8520
         TabIndex        =   31
         Tag             =   "Hide search paramter window."
         Top             =   270
         Visible         =   0   'False
         Width           =   750
      End
      Begin Threed.SSCheck chkCompl 
         Height          =   330
         Left            =   120
         TabIndex        =   35
         Tag             =   "00-Display old Follow-Up messages as well."
         Top             =   270
         Visible         =   0   'False
         Width           =   3555
         _Version        =   65536
         _ExtentX        =   6271
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "Show both Completed and Incomplete"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCheck chkAutoScan 
         Height          =   330
         Left            =   120
         TabIndex        =   36
         Tag             =   "00-Scan for Follow-Ups each time enter system"
         Top             =   600
         Visible         =   0   'False
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   582
         _StockProps     =   78
         Caption         =   "Show Follow-Ups on Sign-On"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Days forward to search for"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3960
         TabIndex        =   39
         Top             =   315
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Label lblFrom 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3960
         TabIndex        =   38
         Top             =   660
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lblTo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5520
         TabIndex        =   37
         Top             =   660
         Visible         =   0   'False
         Width           =   1170
      End
   End
   Begin VB.Label lblAdminBy 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Administered By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   28
      Top             =   3090
      Width           =   1125
   End
   Begin VB.Label lblEDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Effective"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5640
      TabIndex        =   27
      Top             =   2775
      Width           =   780
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   26
      Top             =   3420
      Width           =   735
   End
   Begin VB.Label lblType 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   25
      Top             =   2760
      Width           =   660
   End
   Begin VB.Label lblName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   2010
      TabIndex        =   24
      Top             =   2445
      Width           =   630
   End
   Begin VB.Label lblEENUM 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EEID"
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
      Left            =   480
      TabIndex        =   23
      Top             =   2445
      Width           =   540
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EF_EMPNBR"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1500
      TabIndex        =   18
      Top             =   6420
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EF_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   450
      TabIndex        =   19
      Top             =   6420
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmvFOLOWUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew
Dim fUPMode As Integer ', fglbEmptyNew As Integer
Dim EESNameSort
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim rsGrid As ADODB.Recordset
'Dim FRS As ADODB.Recordset

Private Sub chkAutoScan_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkAutoScan_LostFocus()
If chkAutoScan.Value = True Then
    glbFOLLOWUPS% = True
Else
    glbFOLLOWUPS% = False
End If

End Sub

Private Sub chkCompl_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkCompl_LostFocus()
If chkCompl.Value = True Then
    glbFOLLOWUPSCOMP% = True
Else
    glbFOLLOWUPSCOMP% = False
End If

End Sub

Private Sub chkCompleted_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Function chkEComment()
Dim SQLQ As String, Msg As String, dd#
Dim rs As New ADODB.Recordset
Dim xTemplate As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbUserID)


chkEComment = False

On Error GoTo chkEComment_Err

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Reason code is a required field"
    clpCode(1).SetFocus
    Exit Function
End If

If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Reason code must be valid"
    clpCode(1).SetFocus
    Exit Function
Else
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        SQLQ = "SELECT MAINTAINABLE from HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(glbUserID, "'", "''") & "'"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        SQLQ = "SELECT MAINTAINABLE from HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    End If
    'SQLQ = "SELECT ACCESSABLE from HR_SECURE_FOLLOW_UP WHERE USERID='" & glbUserID & "'"
    SQLQ = SQLQ & " AND CODENAME='" & clpCode(1).Text & "'"
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF = False And rs.BOF = False Then
        If rs("MAINTAINABLE") = 0 Then
        'If rs("ACCESSABLE") = 0 Then
            MsgBox "You do not have Authority to 'Maintain' on '" & clpCode(1).Text & "' Reason code.", vbOKOnly + vbInformation, "Authorization failed"
            rs.Close
            Set rs = Nothing
            Exit Function
        End If
    Else
        MsgBox "You do not have Authority to 'Maintain' on '" & clpCode(1).Text & "' Reason code.", vbOKOnly + vbInformation, "Authorization failed"
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    rs.Close
    Set rs = Nothing
End If

If clpCode(2).Caption = "Unassigned" And Len(Trim(clpCode(2).Text)) > 0 Then
    MsgBox lStr("Administered By") & " type must be valid"
    clpCode(2).SetFocus
    Exit Function
End If

If Len(dlpEDate.Text) >= 1 Then
    If Not IsDate(dlpEDate.Text) Then
        MsgBox "Effective Date is not a valid date."
        dlpEDate.SetFocus
        Exit Function
    End If
Else
    MsgBox "Effective Date is required."
    dlpEDate.SetFocus
    Exit Function
End If

chkEComment = True

Exit Function

chkEComment_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkVFollow", "HR_EFOLLOWUP", "edit/Add")
Resume Next

End Function

Sub cmdCancel_Click()

On Error GoTo Can_Err
Dim bk

rsDATA.CancelUpdate

Call Display_Value

Call modSTUPD(True)  ' reset screen's attributes
Call SET_UP_MODE

Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
Resume Next

End Sub

'Private Sub cmdCancel_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Unload Me

End Sub

'Private Sub cmdClose_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String
Dim xID, x
If Not gSec_Upd_Follow_Ups Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

If Not FollowUp_Sec Then
    MsgBox "You do not have Authority to complete this Reason Code transaction.", vbInformation + vbOKOnly, "Authorization failure"
    Exit Sub
End If

On Error GoTo Del_Err


Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub
xID = Data1.Recordset("EF_FOLLOWUP_ID")
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute "delete from HR_FOLLOW_UP where EF_FOLLOWUP_ID=" & xID
gdbAdoIhr001.CommitTrans

Data1.Refresh

Set rsGrid = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If
Call modSTUPD(True)
Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_EFOLLOWUP", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub chkMyFollow_Click()
If EERetrieve() = 0 Then     ' get the info for this person
    Exit Sub
End If
End Sub

Private Sub cmdComSelected_Click()
Dim SQLQ As String, Msg$, Response%, Title$, DgDef As Variant
Dim x

On Error GoTo MAll_Err
If Not gSec_Upd_Follow_Ups Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

Msg$ = lStr("Mark the highlighted Follow-Up Records as completed?")
Title$ = "Mark the Records completed?"   '
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDYES Then    ' Evaluate response
    Screen.MousePointer = HOURGLASS
    Dim xFollowList
    xFollowList = ""
    
    If vbxTrueGrid.SelBookmarks.count = 0 Then vbxTrueGrid.SelBookmarks.Add Data1.Recordset.Bookmark
    For x = 0 To vbxTrueGrid.SelBookmarks.count - 1
        Data1.Recordset.Bookmark = vbxTrueGrid.SelBookmarks(x)
        xFollowList = xFollowList & Data1.Recordset("EF_FOLLOWUP_ID") & ","
    Next
    
    xFollowList = xFollowList & "-1"
    'Friesens - Ticket #16591
    If glbCompSerial = "S/N - 2279W" Then
        gdbAdoIhr001.Execute "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "' WHERE (EF_FREAS <> 'EDUC' or EF_COMPLETED = 1) AND EF_FOLLOWUP_ID IN (" & xFollowList & ")"
    Else
        gdbAdoIhr001.Execute "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "' WHERE EF_FOLLOWUP_ID IN (" & xFollowList & ")"
    End If
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    Data1.Refresh
    Screen.MousePointer = DEFAULT
End If

chkCompl.Visible = False
chkAutoScan.Visible = False
Label3.Visible = False
txtDaysForward.Visible = False
cmdSave.Visible = False
cmdHide.Visible = False
lblFrom.Visible = False
lblTo.Visible = False
cmdReScan.Visible = False

If Data1.Recordset.EOF And Data1.Recordset.BOF Then Unload Me

Exit Sub

MAll_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdMarkAll", "HR_EFOLLOWUP", "Mark All")

Resume Next
Unload Me
End Sub

Private Sub cmdDelSelected_Click()
Dim SQLQ As String, Msg$, Response%, Title$, DgDef As Variant
Dim x, xID
On Error GoTo MDel_Err
If Not gSec_Upd_Follow_Ups Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

Msg$ = lStr("Delete the highlighted Follow-Up Records?")
'Msg$ = Msg$ & Chr(10) & "listed?"
Title$ = lStr("Delete the highlighted Follow-Up Records")   ' zzz
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.

If Response% = IDYES Then    ' Evaluate response
    Screen.MousePointer = HOURGLASS
    Dim xFollowList
    xFollowList = ""

    For x = 0 To vbxTrueGrid.SelBookmarks.count - 1
        Data1.Recordset.Bookmark = vbxTrueGrid.SelBookmarks(x)
        xFollowList = xFollowList & Data1.Recordset("EF_FOLLOWUP_ID") & ","
    Next
    
    xFollowList = xFollowList & "-1"
    'Added by Bryan to allow deleting one record. Nov 8, 2006 Ticket# 12058
    If xFollowList = "-1" And Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
        xID = Data1.Recordset("EF_FOLLOWUP_ID")
        gdbAdoIhr001.BeginTrans
        'Friesens - Ticket #16591
        If glbCompSerial = "S/N - 2279W" Then
            gdbAdoIhr001.Execute "delete from HR_FOLLOW_UP where (EF_FREAS <> 'EDUC' or EF_COMPLETED = 1) AND EF_FOLLOWUP_ID=" & xID
        Else
            gdbAdoIhr001.Execute "delete from HR_FOLLOW_UP where EF_FOLLOWUP_ID=" & xID
        End If
        gdbAdoIhr001.CommitTrans
    Else
        'Data1.Refresh
        'Friesens - Ticket #16591
        If glbCompSerial = "S/N - 2279W" Then
            gdbAdoIhr001.Execute "DELETE FROM HR_FOLLOW_UP WHERE (EF_FREAS <> 'EDUC' or EF_COMPLETED = 1) AND EF_FOLLOWUP_ID IN (" & xFollowList & ")"
        Else
            gdbAdoIhr001.Execute "DELETE FROM HR_FOLLOW_UP WHERE EF_FOLLOWUP_ID IN (" & xFollowList & ")"
        End If
    End If
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    Data1.Refresh
    Screen.MousePointer = DEFAULT
    'Unload Me
End If
Call modSTUPD(True)
Exit Sub

MDel_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_EFOLLOWUP", "Delete")

Resume Next
Unload Me

End Sub

'Private Sub cmdDelete_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdEESort_Click(Index As Integer)
Dim xUserEmpNo

Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).Caption = "Refreshing Employee List - Stand by"
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "

EESNameSort = Index

txtEESearch.Text = ""
dlpEESearch.Text = ""
clpEESearch.Text = ""
dlpEESearch.Visible = False
clpEESearch.Visible = False
txtEESearch.Visible = False

Select Case EESNameSort
Case 0
    txtEESearch.Visible = True
    lblSearchBy.Caption = "Search by Emp. #"
Case 1
    txtEESearch.Visible = True
    lblSearchBy.Caption = "Search by Surname"
Case 2
    dlpEESearch.Visible = True
    lblSearchBy.Caption = "Search by Effective Date"
Case 3
    clpEESearch.Visible = True
    lblSearchBy.Caption = "Search by Reason"
End Select

If cmdManagerSec.Visible = True Then
    If cmdManagerSec.Caption = "&Filter with Default Security" Then
        xUserEmpNo = Get_EmployeeNo_of_UserID(glbUserID)
        If xUserEmpNo <> "" Then
            If EERetrieve_ReptAuth_Managers(xUserEmpNo) = 0 Then
                Exit Sub
            End If
        End If
    Else
        If EERetrieve() = 0 Then     ' get the info for this person
            Exit Sub
        End If          ' dpartment specific and populate the list
    End If
Else
    If EERetrieve() = 0 Then     ' get the info for this person
        Exit Sub
    End If          ' dpartment specific and populate the list
End If

Screen.MousePointer = DEFAULT

MDIMain.panHelp(0).Caption = " "

If Index = 2 Then
    dlpEESearch.SetFocus
ElseIf Index = 3 Then
    clpEESearch.SetFocus
Else
    txtEESearch.SetFocus
End If

End Sub

Private Sub cmdEESort_GotFocus(Index As Integer)
  Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim Sch As String, SQLQ As String
Dim bkmark

On Error GoTo Srch_Err


Data1.Refresh
If Not Data1.Recordset.EOF Then
    If EESNameSort = 2 Then
        Sch = dlpEESearch
    ElseIf EESNameSort = 3 Then
        Sch = clpEESearch
    Else
        Sch = txtEESearch
    End If
    If Not Len(Sch) > 0 Then
        MsgBox "To search you must enter something to search for."
        Exit Sub
    End If
    Sch = Replace(Sch, "'", "''")
    
    Select Case EESNameSort
    Case 0
        If Not IsNumeric(txtEESearch.Text) And Not glbLinamar Then
            Beep
            MsgBox "Employee Identification must be numeric"
            Exit Sub
        End If
        If glbLinamar Then
            SQLQ = "EMPNBR >= '" & Sch & "'"
        Else
            SQLQ = "EF_EMPNBR >= '" & Sch & "'"
        End If
    Case 1
        SQLQ = "ED_SURNAME  >= '" & Replace(Sch, "'", "''") & "'"
    Case 2
        If IsDate(dlpEESearch) Then
            SQLQ = "EF_FDATE  = #" & dlpEESearch & "#"
        Else
            Beep
            MsgBox "Invalid Date format!"
            Exit Sub
        End If
    Case 3
        If clpEESearch.Caption <> "Unassigned" Then
            SQLQ = "EF_FREAS  = '" & clpEESearch & "'"
        Else
            Beep
            MsgBox "Invalid Reason Code!"
            Exit Sub
        End If
    End Select
    Data1.Recordset.Find SQLQ
End If
If Data1.Recordset.EOF Then
    MsgBox "Employee not found"
    Data1.Refresh
End If

Exit Sub

Srch_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HREMP", "Find Next")
Call RollBack '28July99 jsEnd Sub
End Sub

Private Sub cmdFind_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdHide_Click()
frmFind.Visible = True
frmSearch.Visible = False
chkCompl.Visible = False
chkAutoScan.Visible = False
Label3.Visible = False
txtDaysForward.Visible = False
cmdSave.Visible = False
cmdHide.Visible = False
lblFrom.Visible = False
lblTo.Visible = False
cmdReScan.Visible = False

'cmdOK.Visible = True
'cmdCancel.Visible = True
'cmdClose.Visible = True
'cmdModify.Visible = True
'cmdDelete.Visible = True
cmdShowSearch.Visible = True
'cmdPrint.Visible = True
cmdMarkAll.Visible = True
cmdMassDelete.Visible = True

End Sub

Private Sub cmdHide_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdManagerSec_Click()
Dim xUserEmpNo
If cmdManagerSec.Caption = "&Filter by Manager's Security" Then
    cmdManagerSec.Caption = "&Filter with Default Security"
    
    xUserEmpNo = Get_EmployeeNo_of_UserID(glbUserID)
    If xUserEmpNo <> "" Then
        If EERetrieve_ReptAuth_Managers(xUserEmpNo) = 0 Then
            Exit Sub
        End If
    Else
        MsgBox "Cannot retrieve Follow Up records for User ID with no Employee Number associated.", vbOKOnly, "Follow Up Filter"
        cmdManagerSec.Caption = "&Filter by Manager's Security"
    End If
ElseIf cmdManagerSec.Caption = "&Filter with Default Security" Then
    cmdManagerSec.Caption = "&Filter by Manager's Security"
    If EERetrieve() = 0 Then
        Exit Sub
    End If
End If
End Sub

Private Sub cmdMarkAll_Click()
Dim SQLQ As String, Msg$, Response%, Title$, DgDef As Variant

On Error GoTo MAll_Err
If Not gSec_Upd_Follow_Ups Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

Msg$ = lStr("Mark all Follow-Up Records")
Msg$ = Msg$ & Chr(10) & "listed as completed?"
Title$ = "Mark all completed?"   ' zzz
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDYES Then    ' Evaluate response
    Screen.MousePointer = HOURGLASS
    
    Dim xFollowList
    xFollowList = ""
    
    If Data1.Recordset.RecordCount <> 0 Then Data1.Recordset.MoveFirst
    
    Do Until Data1.Recordset.EOF
        xFollowList = xFollowList & Data1.Recordset("EF_FOLLOWUP_ID") & ","
        Data1.Recordset.MoveNext
    Loop
    xFollowList = xFollowList & "-1"
    Data1.Refresh
    
    'Friesens - Ticket #16591
    If glbCompSerial = "S/N - 2279W" Then
        gdbAdoIhr001.Execute "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "' WHERE (EF_FREAS <> 'EDUC' or EF_COMPLETED = 1) AND EF_FOLLOWUP_ID IN (" & xFollowList & ")"
    Else
        gdbAdoIhr001.Execute "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "' WHERE EF_FOLLOWUP_ID IN (" & xFollowList & ")"
    End If
    
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    Data1.Refresh
    Screen.MousePointer = DEFAULT
End If

chkCompl.Visible = False
chkAutoScan.Visible = False
Label3.Visible = False
txtDaysForward.Visible = False
cmdSave.Visible = False
cmdHide.Visible = False
lblFrom.Visible = False
lblTo.Visible = False
cmdReScan.Visible = False

If Data1.Recordset.EOF And Data1.Recordset.BOF Then Unload Me

Exit Sub

MAll_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdMarkAll", "HR_EFOLLOWUP", "Mark All")

Resume Next
Unload Me

End Sub

Private Sub cmdMarkAll_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdMassDelete_Click()
Dim SQLQ As String, Msg$, Response%, Title$, DgDef As Variant

On Error GoTo MDel_Err
If Not gSec_Upd_Follow_Ups Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

Msg$ = lStr("Delete all Follow-Up Records")
Msg$ = Msg$ & Chr(10) & "listed?"
Title$ = lStr("Delete all Follow-Up Records")   ' zzz
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.

If Response% = IDYES Then    ' Evaluate response
    Screen.MousePointer = HOURGLASS
    Dim xFollowList
    xFollowList = ""
    If Data1.Recordset.RecordCount <> 0 Then Data1.Recordset.MoveFirst
    Do Until Data1.Recordset.EOF
        xFollowList = xFollowList & Data1.Recordset("EF_FOLLOWUP_ID") & ","
        Data1.Recordset.MoveNext
    Loop
    xFollowList = xFollowList & "-1"
    Data1.Refresh
    
    'Friesens - Ticket #16591
    If glbCompSerial = "S/N - 2279W" Then
        gdbAdoIhr001.Execute "DELETE FROM HR_FOLLOW_UP WHERE (EF_FREAS <> 'EDUC' or EF_COMPLETED = 1) AND EF_FOLLOWUP_ID IN (" & xFollowList & ")"
    Else
        gdbAdoIhr001.Execute "DELETE FROM HR_FOLLOW_UP WHERE EF_FOLLOWUP_ID IN (" & xFollowList & ")"
    End If
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    Data1.Refresh
    Screen.MousePointer = DEFAULT
    Unload Me
End If

Call modSTUPD(True)

Exit Sub

MDel_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_EFOLLOWUP", "Delete")

Resume Next
Unload Me

End Sub

Sub cmdModify_Click()
Dim x%
On Error GoTo Mod_Err
Dim SQLQ As String, Fdate As Variant, Tdate As Variant

    
Fdate = Date
Tdate = DateAdd("d", glbFOLLOWUPDAYS%, Fdate)
If IsDate(Tdate) Then
    If Year(Tdate) > 2077 Or Year(Tdate) < 1900 Then Tdate = Date
Else
    Tdate = Date
End If

Call modSTUPD(True)
'clpCode(1).SetFocus
Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdModify_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim xID
On Error GoTo Add_Err

If Not chkEComment() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)

rsDATA!EF_FDATE = dlpEDate.Text

gdbAdoIhr001.BeginTrans
Call Set_Control("U", Me, rsDATA)
rsDATA.Update
gdbAdoIhr001.CommitTrans

Data1.Refresh

Set rsGrid = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Call modSTUPD(True)
Call SET_UP_MODE
Me.vbxTrueGrid.SetFocus

Exit Sub

Add_Err:
If Err = 3022 Then
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_EFOLLOWUP", "Update")
Resume Next
Unload Me


End Sub

'Private Sub cmdOK_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String
Dim dscGroup$
    Dim TempCri As String
    Dim dtYYY%, dtMM%, dtDD%
    Dim Fdate, Tdate
    
    vbxCrystal.Reset
 '   cmdPrint.Enabled = False
    glbiOneWhere = False
    glbstrSelCri = ""
    Call glbCri_DeptUN("")
        
    If Not glbFOLLOWUPSCOMP Then
        Fdate = Date
        Tdate = DateAdd("d", glbFOLLOWUPDAYS%, Fdate)
        If IsDate(Tdate) Then
            If Year(Tdate) > 2077 Or Year(Tdate) < 1900 Then Tdate = Date
        Else
            Tdate = Date
        End If
        TempCri = "({HR_FOLLOW_UP.EF_FDATE} "
        
        ' Franks Jan 9,2002
        ' Keep the same logic as Re-Scan as Jerry request
        'dtYYY% = Year(FDate)
        'dtMM% = Month(FDate)
        'dtDD% = Day(FDate)
        'TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
        'dtYYY% = Year(TDate)
        'dtMM% = Month(TDate)
        'dtDD% = Day(TDate)
        'TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        dtYYY% = Year(Tdate)
        dtMM% = month(Tdate)
        dtDD% = Day(Tdate)
        TempCri = TempCri & " <= " & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        ' Franks Jan 9,2002
        If glbOracle Then
            TempCri = TempCri & " AND {HR_FOLLOW_UP.EF_COMPLETED} = 0"
        Else
            TempCri = TempCri & " AND {HR_FOLLOW_UP.EF_COMPLETED} = FALSE"
        End If
    End If
        
    If Len(TempCri) >= 1 Then
        If Not glbiOneWhere Then
            glbstrSelCri = TempCri
        Else
            glbstrSelCri = glbstrSelCri & " AND " & TempCri
        End If
        glbiOneWhere = True
    End If
    
    Call Cri_Sec
    
    'Follow Up security
    'glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP.ACCESSABLE} = True"
    
    'Release 8.0 - Ticket #22682: View Own security
    'If View Own not checked then do not retrieve follow ups of the User/Employee No
    If Len(glbUserEmpNo) > 0 And glbUserEmpNo <> 0 And Not gSec_FollUp_ViewOwn Then
        'Do not show user's Follow Up records based on the Employee # associated to the User.
        If Len(glbstrSelCri) > 0 Then
            glbstrSelCri = glbstrSelCri & " AND {HR_FOLLOW_UP.EF_EMPNBR} <> " & glbUserEmpNo
        Else
            glbstrSelCri = glbstrSelCri & " {HR_FOLLOW_UP.EF_EMPNBR} <> " & glbUserEmpNo
        End If
    End If
    
    
    If Len(glbstrSelCri) >= 0 Then
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    End If
    
    RHeading = "Followup Overview"
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgfoverv.rpt"
    dscGroup$ = "PgHeading" & "= '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
    'Me.vbxCrystal.Formulas(0) = dscGroup$
    
    'Franks 07/19/02 For keeping the same order as Grid #2621
    If lblSearchBy.Caption = "Search by Emp. #" Then
        Me.vbxCrystal.GroupCondition(0) = "GROUP1;{HR_FOLLOW_UP.EF_EMPNBR};ANYCHANGE;A"
        If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
            Me.vbxCrystal.GroupCondition(1) = "GROUP2;{HR_FOLLOW_UP.EF_FDATE};ANYCHANGE;A"
        Else
            Me.vbxCrystal.GroupCondition(1) = "GROUP2;{HR_FOLLOW_UP.EF_FDATE};ANYCHANGE;D"
        End If
        Me.vbxCrystal.GroupCondition(2) = "GROUP3;{HR_FOLLOW_UP.EF_COMPLETED};ANYCHANGE;A"
    End If
    If lblSearchBy.Caption = "Search by Surname" Then
        Me.vbxCrystal.GroupCondition(0) = "GROUP1;{HREMP.ED_SURNAME};ANYCHANGE;A"
        If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
            Me.vbxCrystal.GroupCondition(1) = "GROUP2;{HR_FOLLOW_UP.EF_FDATE};ANYCHANGE;A"
        Else
            Me.vbxCrystal.GroupCondition(1) = "GROUP2;{HR_FOLLOW_UP.EF_FDATE};ANYCHANGE;D"
        End If
        Me.vbxCrystal.GroupCondition(2) = "GROUP3;{HR_FOLLOW_UP.EF_COMPLETED};ANYCHANGE;A"
    End If
    If lblSearchBy.Caption = "Search by Effective Date" Then
        If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
            Me.vbxCrystal.GroupCondition(0) = "GROUP1;{HR_FOLLOW_UP.EF_FDATE};ANYCHANGE;A"
        Else
            Me.vbxCrystal.GroupCondition(0) = "GROUP1;{HR_FOLLOW_UP.EF_FDATE};ANYCHANGE;D"
        End If
        Me.vbxCrystal.GroupCondition(1) = "GROUP2;{HREMP.ED_SURNAME};ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(2) = "GROUP3;{HR_FOLLOW_UP.EF_COMPLETED};ANYCHANGE;A"
    End If
     'Franks 07/19/02 For keeping the same order as Grid #2621
    
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
        Me.vbxCrystal.DataFiles(1) = glbIHRDB
        Me.vbxCrystal.DataFiles(2) = glbIHRDB
    End If
    Me.vbxCrystal.Destination = 1
    Me.vbxCrystal.Action = 1
'    cmdPrint.Enabled = True
End Sub

Sub cmdView_Click()

Dim RHeading As String
Dim dscGroup$
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim Fdate, Tdate
    
    vbxCrystal.Reset

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

'    cmdPrint.Enabled = False
    glbiOneWhere = False
    glbstrSelCri = ""
    
    Call glbCri_DeptUN("")
        
    If Not glbFOLLOWUPSCOMP Then
        Fdate = Date
        Tdate = DateAdd("d", glbFOLLOWUPDAYS%, Fdate)
        If IsDate(Tdate) Then
            If Year(Tdate) > 2077 Or Year(Tdate) < 1900 Then Tdate = Date
        Else
            Tdate = Date
        End If
        TempCri = "({HR_FOLLOW_UP.EF_FDATE} "
        
        ' Franks Jan 9,2002
        ' Keep the same logic as Re-Scan as Jerry request
        'dtYYY% = Year(FDate)
        'dtMM% = Month(FDate)
        'dtDD% = Day(FDate)
        'TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
        'dtYYY% = Year(TDate)
        'dtMM% = Month(TDate)
        'dtDD% = Day(TDate)
        'TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        dtYYY% = Year(Tdate)
        dtMM% = month(Tdate)
        dtDD% = Day(Tdate)
        TempCri = TempCri & " <= " & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        ' Franks Jan 9,2002
        If glbOracle Then
            TempCri = TempCri & " AND {HR_FOLLOW_UP.EF_COMPLETED} = 0"
        Else
            TempCri = TempCri & " AND {HR_FOLLOW_UP.EF_COMPLETED} = FALSE"
        End If
    End If
    
    If Len(TempCri) >= 1 Then
        If Not glbiOneWhere Then
            glbstrSelCri = TempCri
        Else
            glbstrSelCri = glbstrSelCri & " AND " & TempCri
        End If
        glbiOneWhere = True
    End If
    
    Call Cri_Sec
    
    'Follow Up security
    'glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP.ACCESSABLE} = TRUE"
    
    'Release 8.0 - Ticket #22682: View Own security
    'If View Own not checked then do not retrieve follow ups of the User/Employee No
    If Len(glbUserEmpNo) > 0 And glbUserEmpNo <> 0 And Not gSec_FollUp_ViewOwn Then
        'Do not show user's Follow Up records based on the Employee # associated to the User.
        If Len(glbstrSelCri) > 0 Then
            glbstrSelCri = glbstrSelCri & " AND {HR_FOLLOW_UP.EF_EMPNBR} <> " & glbUserEmpNo
        Else
            glbstrSelCri = glbstrSelCri & " {HR_FOLLOW_UP.EF_EMPNBR} <> " & glbUserEmpNo
        End If
    End If
    
    If Len(glbstrSelCri) >= 0 Then
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    End If
    
    RHeading = "Followup Overview"
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgfoverv.rpt"
    dscGroup$ = "PgHeading" & "= '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
    'Me.vbxCrystal.Formulas(0) = dscGroup$
    
    'Franks 07/19/02 For keeping the same order as Grid #2621
    If lblSearchBy.Caption = "Search by Emp. #" Then
        Me.vbxCrystal.GroupCondition(0) = "GROUP1;{HR_FOLLOW_UP.EF_EMPNBR};ANYCHANGE;A"
        If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
            Me.vbxCrystal.GroupCondition(1) = "GROUP2;{HR_FOLLOW_UP.EF_FDATE};ANYCHANGE;A"
        Else
            Me.vbxCrystal.GroupCondition(1) = "GROUP2;{HR_FOLLOW_UP.EF_FDATE};ANYCHANGE;D"
        End If
        Me.vbxCrystal.GroupCondition(2) = "GROUP3;{HR_FOLLOW_UP.EF_COMPLETED};ANYCHANGE;A"
    End If
    If lblSearchBy.Caption = "Search by Surname" Then
        Me.vbxCrystal.GroupCondition(0) = "GROUP1;{HREMP.ED_SURNAME};ANYCHANGE;A"
        If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
            Me.vbxCrystal.GroupCondition(1) = "GROUP2;{HR_FOLLOW_UP.EF_FDATE};ANYCHANGE;A"
        Else
            Me.vbxCrystal.GroupCondition(1) = "GROUP2;{HR_FOLLOW_UP.EF_FDATE};ANYCHANGE;D"
        End If
        Me.vbxCrystal.GroupCondition(2) = "GROUP3;{HR_FOLLOW_UP.EF_COMPLETED};ANYCHANGE;A"
    End If
    If lblSearchBy.Caption = "Search by Effective Date" Then
        If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
            Me.vbxCrystal.GroupCondition(0) = "GROUP1;{HR_FOLLOW_UP.EF_FDATE};ANYCHANGE;A"
        Else
            Me.vbxCrystal.GroupCondition(0) = "GROUP1;{HR_FOLLOW_UP.EF_FDATE};ANYCHANGE;D"
        End If
        Me.vbxCrystal.GroupCondition(1) = "GROUP2;{HREMP.ED_SURNAME};ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(2) = "GROUP3;{HR_FOLLOW_UP.EF_COMPLETED};ANYCHANGE;A"
    End If
     'Franks 07/19/02 For keeping the same order as Grid #2621
    
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
        Me.vbxCrystal.DataFiles(1) = glbIHRDB
        Me.vbxCrystal.DataFiles(2) = glbIHRDB
    End If
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.Action = 1
'    cmdPrint.Enabled = True
End Sub

'Private Sub cmdPrint_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdReScan_Click()

Screen.MousePointer = HOURGLASS
If EERetrieve() = False Then
    glbFollwUpsFound% = False
End If
Screen.MousePointer = DEFAULT

End Sub

Private Sub cmdReScan_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdSave_Click()
    Dim Msg$, Title$, DgDef  As Variant
    Dim Response%, w%, x%, Y%, SECTION$, Key$, Value$
    
    w% = False
    
    Msg$ = "Save changes to your 'Startup File'?"
    Title$ = "Save changes to Initialization File"
    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2  ' Describe dialog.
    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
    
    If Response = IDYES Then    ' Evaluate response
        SECTION$ = "FOLLOWUPS"
        If Len(txtDaysForward) > 0 Then
            If IsNumeric(txtDaysForward) Then
                If CInt(txtDaysForward) >= 0 Or CInt(txtDaysForward) <= 365 Then
                    If glbHosted Then
                        SECTION$ = "FOLLOWUPS"
                        Key$ = "FOLLOWUPDAYS"
                        Value$ = CStr(txtDaysForward)
                        w% = INIWrite(SECTION$, Key$, Value$, glbHostFile)
                        'w% = WriteRegistrySetting(lCurrentKey, SECTION$, Key$, Value$)
                    Else
                        SECTION$ = REG_NAME & "FOLLOWUPS"
                        Key$ = "FOLLOWUPDAYS"
                        Value$ = CStr(txtDaysForward)
                        w% = WriteRegistrySetting(lCurrentKey, SECTION$, Key$, Value$)
                    End If
                End If
            End If
        End If
        
        If w Then
            If glbHosted Then
                SECTION$ = "FOLLOWUPS"
                
                Key$ = "FOLLOWUPS"
                If chkAutoScan Then Value$ = "Y" Else Value$ = "N"
                x% = INIWrite(SECTION$, Key$, Value$, glbHostFile)
                'x% = WriteRegistrySetting(lCurrentKey, SECTION$, Key$, Value$) 'gbasINI_WritePrivateString
                
                If chkCompl Then Value$ = "Y" Else Value$ = "N"
                Key$ = "SHOWCOMPLETED"
                Y% = INIWrite(SECTION$, Key$, Value$, glbHostFile)
                'Y% = WriteRegistrySetting(lCurrentKey, SECTION$, Key$, Value$)
            Else
                SECTION$ = REG_NAME & "FOLLOWUPS"
                
                Key$ = "FOLLOWUPS"
                If chkAutoScan Then Value$ = "Y" Else Value$ = "N"
                x% = WriteRegistrySetting(lCurrentKey, SECTION$, Key$, Value$) 'gbasINI_WritePrivateString
                
                If chkCompl Then Value$ = "Y" Else Value$ = "N"
                Key$ = "SHOWCOMPLETED"
                Y% = WriteRegistrySetting(lCurrentKey, SECTION$, Key$, Value$)
            End If
        End If
        
        If w And x And Y Then
            MsgBox "Initialization file updated."
        Else
            MsgBox "Problem updating your Initializaion file."
        End If
    End If
   
End Sub

Private Sub cmdSave_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdShowSearch_Click()
cmdTLAY.Visible = False
frmFind.Visible = False
frmSearch.Visible = True
txtDaysForward = glbFOLLOWUPDAYS%

If glbFOLLOWUPSCOMP Then
    chkCompl.Value = True
Else
    chkCompl.Value = False
End If

If glbFOLLOWUPS Then
    chkAutoScan.Value = True
Else
    chkAutoScan.Value = False
End If

chkCompl.Enabled = True
chkCompl.Visible = True
chkCompl.SetFocus

chkAutoScan.Enabled = True
chkAutoScan.Visible = True

Label3.Visible = True
txtDaysForward.Visible = True
cmdSave.Visible = True
cmdHide.Visible = True
lblFrom.Visible = True
lblTo.Visible = True
cmdReScan.Visible = True

'cmdOK.Visible = False
'cmdCancel.Visible = False
'cmdClose.Visible = False
'cmdModify.Visible = False
'cmdDelete.Visible = False
cmdShowSearch.Visible = False
'cmdPrint.Visible = False
cmdMarkAll.Visible = False
cmdMassDelete.Visible = False

End Sub

Function EERetrieve()

Dim SQLQ As String, Fdate As Variant, Tdate As Variant, countr%, x
Dim xTemplate As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbUserID)

EERetrieve = False

On Error GoTo EERError
Fdate = Date
Tdate = DateAdd("d", glbFOLLOWUPDAYS%, Fdate)
If IsDate(Tdate) Then
    If Year(Tdate) > 2077 Or Year(Tdate) < 1900 Then Tdate = Date
Else
    Tdate = Date
End If

If glbOracle Then
    SQLQ = "SELECT HR_FOLLOW_UP.* "
    SQLQ = SQLQ & ", HR_FOLLOW_UP.EF_EMPNBR AS EMPNBR, "
    SQLQ = SQLQ & " HRTABL.TB_DESC, HREMP.ED_SURNAME, HREMP.ED_FNAME "
    SQLQ = SQLQ & "  FROM HR_FOLLOW_UP,HRTABL,HREMP "
    
    SQLQ = SQLQ & " WHERE (HR_FOLLOW_UP.EF_FREAS_TABL = HRTABL.TB_NAME(+)) AND (HR_FOLLOW_UP.EF_FREAS = HRTABL.TB_KEY(+)) "
    SQLQ = SQLQ & " AND (HR_FOLLOW_UP.EF_EMPNBR = HREMP.ED_EMPNBR(+)) "
    SQLQ = SQLQ & " AND "
Else
    SQLQ = "SELECT HR_FOLLOW_UP.* ,"
    If glbLinamar Then
        SQLQ = SQLQ & "right(HR_FOLLOW_UP.EF_EMPNBR,3)+'-'+ left(HR_FOLLOW_UP.EF_EMPNBR,LEN(HR_FOLLOW_UP.EF_EMPNBR)-3) AS EMPNBR,"
    Else
        SQLQ = SQLQ & "LTRIM(STR(HR_FOLLOW_UP.EF_EMPNBR)) AS EMPNBR,"
    End If
    SQLQ = SQLQ & " HRTABL.TB_DESC, HREMP.ED_SURNAME, HREMP.ED_FNAME from (HR_FOLLOW_UP"
    SQLQ = SQLQ & " LEFT JOIN HRTABL ON (HR_FOLLOW_UP.EF_FREAS_TABL = HRTABL.TB_NAME) AND (HR_FOLLOW_UP.EF_FREAS = HRTABL.TB_KEY)) "
    SQLQ = SQLQ & " LEFT JOIN HREMP ON HR_FOLLOW_UP.EF_EMPNBR = HREMP.ED_EMPNBR"
    
    'Hemu
    SQLQ = SQLQ & " INNER JOIN HR_SECURE_FOLLOW_UP ON HR_FOLLOW_UP.EF_FREAS = HR_SECURE_FOLLOW_UP.CODENAME"
    
    SQLQ = SQLQ & " WHERE "
End If
If Not glbFOLLOWUPSCOMP Then
    SQLQ = SQLQ & " HR_FOLLOW_UP.EF_FDATE <= " & Date_SQL(Tdate)
    SQLQ = SQLQ & " AND HR_FOLLOW_UP.EF_COMPLETED = 0 "
Else
    SQLQ = SQLQ & " HR_FOLLOW_UP.EF_COMPNO  = '001' "
End If

'Hemu
If xTemplate = "" Or xTemplate = "TEMPLATE" Then
    SQLQ = SQLQ & " AND HR_SECURE_FOLLOW_UP.USERID ='" & Replace(glbUserID, "'", "''") & "' AND HR_SECURE_FOLLOW_UP.ACCESSABLE <> 0"
Else
    '????Ticket #24808 -  Retrieve template's security profile
    SQLQ = SQLQ & " AND HR_SECURE_FOLLOW_UP.USERID ='" & Replace(xTemplate, "'", "''") & "' AND HR_SECURE_FOLLOW_UP.ACCESSABLE <> 0"
End If

SQLQ = SQLQ & " AND " & glbSeleDeptUn
If glbNoNONE And glbNoEXEC Then        'Hemu -EXE
    SQLQ = SQLQ & " AND NOT(EF_FREAS = 'SREV' AND EF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG = 'NONE' OR ED_ORG = 'EXEC' )) "   'Hemu -EXE
ElseIf glbNoNONE Then
    SQLQ = SQLQ & " AND NOT(EF_FREAS = 'SREV' AND EF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG = 'NONE' )) "
ElseIf glbNoEXEC Then   'Hemu -EXE
    SQLQ = SQLQ & " AND NOT(EF_FREAS = 'SREV' AND EF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG = 'EXEC' )) "   'Hemu -EXE
End If
If glbLinamar Then
     SQLQ = SQLQ & " AND LEN(HR_FOLLOW_UP.EF_EMPNBR)> 3 "
End If

If chkMyFollow Then
    SQLQ = SQLQ & " AND EF_LUSER='" & glbUserID & "'"
End If

'Release 8.0 - Ticket #22682: View Own security
'If View Own not checked then do not retrieve follow ups of the User/Employee No
If Len(glbUserEmpNo) > 0 And glbUserEmpNo <> 0 And Not gSec_FollUp_ViewOwn Then
    'Do not show user's Follow Up records based on the Employee # associated to the User.
    SQLQ = SQLQ & " AND EF_EMPNBR <> " & glbUserEmpNo
End If

SQLQ = SQLQ & " ORDER BY "
Select Case EESNameSort
Case 0
    SQLQ = SQLQ & IIf(glbLinamar, "EMPNBR", "EF_EMPNBR")
    SQLQ = SQLQ & ","
    If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
        SQLQ = SQLQ & " EF_FDATE ASC, EF_COMPLETED "
    Else
        SQLQ = SQLQ & " EF_FDATE DESC, EF_COMPLETED "
    End If
Case 1
    SQLQ = SQLQ & "ED_SURNAME "
    SQLQ = SQLQ & ","
    If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
        SQLQ = SQLQ & " EF_FDATE ASC, EF_COMPLETED "
    Else
        SQLQ = SQLQ & " EF_FDATE DESC, EF_COMPLETED "
    End If
Case 2
    If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
        SQLQ = SQLQ & " EF_FDATE ASC, EF_COMPLETED "
    Else
        SQLQ = SQLQ & " EF_FDATE DESC, EF_COMPLETED "
    End If
    SQLQ = SQLQ & ","
    SQLQ = SQLQ & IIf(glbLinamar, "EMPNBR", "EF_EMPNBR")
Case 3
    SQLQ = SQLQ & " EF_FREAS "
    SQLQ = SQLQ & ","
    If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
        SQLQ = SQLQ & " EF_FDATE ASC, EF_COMPLETED "
    Else
        SQLQ = SQLQ & " EF_FDATE DESC, EF_COMPLETED "
    End If
End Select

'WriteFile (SQLQ)
Data1.RecordSource = SQLQ
Data1.Refresh

Set rsGrid = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

If Data1.Recordset.EOF And Not glbFollowUpsRemain Then
    Exit Function
End If
EERetrieve = True

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FollowUpRetrieve", "HR_EFOLLOW_UP", "Tdate is:" & Tdate)

Resume Next

Exit Function

End Function

Private Sub cmdTLAY_Click()
If gSec_Inq_Terminations Then
    Screen.MousePointer = HOURGLASS
    Unload frmTLAY
    glbTLAY = "Follow-Up"
    Load frmTLAY
    Screen.MousePointer = DEFAULT
Else
    MsgBox "You Do Not Have Authority For This Transaction"
End If
End Sub

Private Sub Form_Activate()
If glbLinamar Then
    Dim rsTB As New ADODB.Recordset
    Set rsTB = Data1.Recordset.Clone
    rsTB.Filter = "EF_FREAS ='TLAY' AND EF_FDATE<=" & Date_SQL(Date) & " AND EF_COMPLETED=0"
    If Not rsTB.EOF Then
        cmdTLAY.Visible = True
    Else
        cmdTLAY.Visible = False
    End If
    rsTB.Close
End If
Call SET_UP_MODE

glbOnTop = Me.name ' "FRMVFOLLOWUP"
'Me.cmdModify_Click
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = Me.name '"FRMVFOLLOWUP"
EESNameSort = 2 '1     'Ticket #28635 - DESC order by EF_FDATE

Data1.ConnectionString = glbAdoIHRDB
If EERetrieve() = False Then
    glbFollwUpsFound% = False
    Screen.MousePointer = DEFAULT
    If Not glbFollowUpsRemain Then
        Exit Sub
    End If
Else
    glbFollwUpsFound% = True
End If

Screen.MousePointer = HOURGLASS
frmvFOLOWUP.WindowState = MAXIMIZED
Me.Caption = lStr("Follow-ups Overview ")

Call modSTUPD(False)

If Not gSec_Upd_Follow_Ups Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
    cmdMarkAll.Enabled = False
    cmdMassDelete.Enabled = False
End If

If glbWHSCC Then
    cmdMarkAll.Visible = False
    cmdMassDelete.Visible = False
End If

If glbCompSerial = "S/N - 2279W" Then   'Friesens Corporation - Ticket #16189
    cmdManagerSec.Visible = True
End If

Call Display_Value
Call modDaysForward
Call setCaption(lblAdminBy)
Call setCaption(chkMyFollow)
Call setCaption(chkAutoScan)

Call INI_Controls(Me)

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Call EERetrieve

'Set FRS = Data1.Recordset.Clone
'vbxTrueGrid.FetchRowStyle = True
vbxTrueGrid.MarqueeStyle = 3

If Data1.Recordset.BOF Or Data1.Recordset.EOF Then
Else
    Me.cmdModify_Click
    If glbCompSerial = "S/N - 2279W" Then   'Friesens Corporation - Ticket #16189
        Call cmdManagerSec_Click
    End If
End If

Screen.MousePointer = DEFAULT
vbxTrueGrid.Refresh
End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub

Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)

End Sub

Private Sub Form_Unload(Cancel As Integer)

MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
Set frmvFOLOWUP = Nothing

End Sub

Private Sub lblEEID_Change()
    If Data1.Recordset.BOF Or Data1.Recordset.EOF Then
        lblName.Visible = False
        'frmDetails.Visible = False    'js-7Apr99
        lblType.Visible = False
        lblAdminBy.Visible = False
        clpCode(1).Visible = False     '
        clpCode(2).Visible = False     '
       ' lblCodeDesc(1).Visible = False '
       ' lblCodeDesc(2).Visible = False '
        Label1.Visible = False         '
        memComments.Visible = False    '
        lblEDate.Visible = False       '
         dlpEDate.Visible = False       '
        chkCompleted.Visible = False   '
        Exit Sub
    Else
        lblName.Visible = False
        'frmDetails.Visible = True
        If Not IsNull(Data1.Recordset("ED_SURNAME")) And Not IsNull(Data1.Recordset("ED_FNAME")) Then
            lblName = RTrim$(Data1.Recordset("ED_SURNAME")) & ", " & RTrim$(Data1.Recordset("ED_FNAME"))
            lblName.Visible = True
        End If
        lblEENUM = ShowEmpnbr(lblEEID)
        lblType.Visible = True       'js-7Apr99
        lblAdminBy.Visible = True
        clpCode(1).Visible = True     '
        clpCode(2).Visible = True     '
       ' lblCodeDesc(1).Visible = True '
       ' lblCodeDesc(2).Visible = True '
        Label1.Visible = True         '
        memComments.Visible = True    '
        lblEDate.Visible = True       '
        'dlpEDate.Visible = True       '
        chkCompleted.Visible = True   '
    End If
End Sub

Private Sub memComments_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub modDaysForward()
Dim Fdate As Variant, Tdate As Variant

Fdate = Date
Tdate = DateAdd("d", glbFOLLOWUPDAYS%, Fdate)
lblFrom = Format(Fdate, "Short Date")
lblTo = Format(Tdate, "Short Date")

End Sub

Private Sub modSTUPD(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

fUPMode = TF    ' update mode

'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'jaddy removed because this is not working for the empty record
'cmdShowSearch.Enabled = TF  'FT
'jaddy
'cmdPrint.Enabled = FT
cmdMarkAll.Enabled = TF  'FT
cmdMassDelete.Enabled = IIf(glbLinamar, False, TF) 'FT

clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
dlpEDate.Enabled = TF
chkCompleted.Enabled = TF
'chkCompl.Enabled = TF
'chkAutoScan.Enabled = TF
'cmdReScan.Enabled = TF
'cmdSave.Enabled = TF
'cmdHide.Enabled = TF
If Data1.Recordset.BOF Or Data1.Recordset.EOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
    cmdMassDelete.Enabled = False
End If
memComments.Enabled = TF
'vbxTrueGrid.Enabled = FT

End Sub

Private Sub txtDaysForward_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtDaysForward_LostFocus()
Dim Fdate As Variant, Tdate As Variant

If Len(txtDaysForward) <= 0 Then Exit Sub
If IsNumeric(txtDaysForward) Then glbFOLLOWUPDAYS% = CInt(txtDaysForward)
Call modDaysForward
End Sub

Private Sub txtEESearch_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
On Error GoTo Eh
'Friesens - Ticket #16189
If glbCompSerial = "S/N - 2279W" Then
    rsGrid.Bookmark = Bookmark
    If Not IsNull(rsGrid("EF_FREAS")) And rsGrid("EF_FREAS") <> "" Then
        'Disable changes to EDUC records only
        If rsGrid("EF_FREAS") = "EDUC" And rsGrid("EF_COMPLETED") = False Then
            'Grey the text
            RowStyle.ForeColor = vbGrayText
            
            'Disable the controls - to prevent changes to the records
            clpCode(1).Enabled = False
            clpCode(2).Enabled = False
            dlpEDate.Enabled = False
            chkCompleted.Enabled = False
            memComments.Enabled = False
        Else
            'Enable controls
            clpCode(1).Enabled = True
            clpCode(2).Enabled = True
            dlpEDate.Enabled = True
            chkCompleted.Enabled = True
            memComments.Enabled = True
        End If
    End If
End If
Eh:
    Exit Sub
End Sub

Private Sub vbxTrueGrid_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)

Dim SQLQ As String, Fdate As Variant, Tdate As Variant, countr%, x
Dim xTemplate As String

    '????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
    xTemplate = ""
    xTemplate = Get_Template(glbUserID)
        
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
        
    Fdate = Date
    Tdate = DateAdd("d", glbFOLLOWUPDAYS%, Fdate)
    If IsDate(Tdate) Then
        If Year(Tdate) > 2077 Or Year(Tdate) < 1900 Then Tdate = Date
    Else
        Tdate = Date
    End If
    
    If glbOracle Then
        SQLQ = "SELECT HR_FOLLOW_UP.* "
        SQLQ = SQLQ & ", HR_FOLLOW_UP.EF_EMPNBR AS EMPNBR, "
        SQLQ = SQLQ & " HRTABL.TB_DESC, HREMP.ED_SURNAME, HREMP.ED_FNAME "
        SQLQ = SQLQ & "  FROM HR_FOLLOW_UP,HRTABL,HREMP "
        
        SQLQ = SQLQ & " WHERE (HR_FOLLOW_UP.EF_FREAS_TABL = HRTABL.TB_NAME(+)) AND (HR_FOLLOW_UP.EF_FREAS = HRTABL.TB_KEY(+)) "
        SQLQ = SQLQ & " AND (HR_FOLLOW_UP.EF_EMPNBR = HREMP.ED_EMPNBR(+)) "
        SQLQ = SQLQ & " AND "
    Else
        SQLQ = "SELECT HR_FOLLOW_UP.* ,"
        If glbLinamar Then
            SQLQ = SQLQ & "right(HR_FOLLOW_UP.EF_EMPNBR,3)+'-'+ left(HR_FOLLOW_UP.EF_EMPNBR,LEN(HR_FOLLOW_UP.EF_EMPNBR)-3) AS EMPNBR,"
        Else
            SQLQ = SQLQ & "LTRIM(STR(HR_FOLLOW_UP.EF_EMPNBR)) AS EMPNBR,"
        End If
        SQLQ = SQLQ & " HRTABL.TB_DESC, HREMP.ED_SURNAME, HREMP.ED_FNAME from (HR_FOLLOW_UP"
        SQLQ = SQLQ & " LEFT JOIN HRTABL ON (HR_FOLLOW_UP.EF_FREAS_TABL = HRTABL.TB_NAME) AND (HR_FOLLOW_UP.EF_FREAS = HRTABL.TB_KEY)) "
        SQLQ = SQLQ & " LEFT JOIN HREMP ON HR_FOLLOW_UP.EF_EMPNBR = HREMP.ED_EMPNBR"
        
        'Hemu
        SQLQ = SQLQ & " INNER JOIN HR_SECURE_FOLLOW_UP ON HR_FOLLOW_UP.EF_FREAS = HR_SECURE_FOLLOW_UP.CODENAME"
        
        SQLQ = SQLQ & " WHERE "
    End If
    If Not glbFOLLOWUPSCOMP Then
        SQLQ = SQLQ & " HR_FOLLOW_UP.EF_FDATE <= " & Date_SQL(Tdate)
        SQLQ = SQLQ & " AND HR_FOLLOW_UP.EF_COMPLETED = 0 "
    Else
        SQLQ = SQLQ & " HR_FOLLOW_UP.EF_COMPNO  = '001' "
    End If

    'Hemu
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        SQLQ = SQLQ & " AND HR_SECURE_FOLLOW_UP.USERID ='" & Replace(glbUserID, "'", "''") & "' AND HR_SECURE_FOLLOW_UP.ACCESSABLE <> 0"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        SQLQ = SQLQ & " AND HR_SECURE_FOLLOW_UP.USERID ='" & Replace(xTemplate, "'", "''") & "' AND HR_SECURE_FOLLOW_UP.ACCESSABLE <> 0"
    End If
    
    SQLQ = SQLQ & " AND " & glbSeleDeptUn
    If glbNoNONE And glbNoEXEC Then        'Hemu -EXE
        SQLQ = SQLQ & " AND NOT(EF_FREAS = 'SREV' AND EF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG = 'NONE' OR ED_ORG = 'EXEC' )) "   'Hemu -EXE
    ElseIf glbNoNONE Then
        SQLQ = SQLQ & " AND NOT(EF_FREAS = 'SREV' AND EF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG = 'NONE' )) "
    ElseIf glbNoEXEC Then   'Hemu -EXE
        SQLQ = SQLQ & " AND NOT(EF_FREAS = 'SREV' AND EF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG = 'EXEC' )) "   'Hemu -EXE
    End If
    If chkMyFollow Then
        SQLQ = SQLQ & " AND EF_LUSER='" & glbUserID & "'"
    End If

    'Release 8.0 - Ticket #22682: View Own security
    'If View Own not checked then do not retrieve follow ups of the User/Employee No
    If Len(glbUserEmpNo) > 0 And glbUserEmpNo <> 0 And Not gSec_FollUp_ViewOwn Then
        'Do not show user's Follow Up records based on the Employee # associated to the User.
        SQLQ = SQLQ & " AND EF_EMPNBR <> " & glbUserEmpNo
    End If

    If (Not glbSQL And Not glbOracle) And ColIndex = 0 Then
        SQLQ = SQLQ & "ORDER BY EF_EMPNBR"
    Else
        SQLQ = SQLQ & "ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    End If
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
    
    Set rsGrid = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$
Dim SQLQ As String
Dim ButtonS As New ButtonsSetting

On Error GoTo Tab1_Err

Call Display_Value

'Friesens - Ticket #16189
If glbCompSerial = "S/N - 2279W" Then
    If Not IsNull(clpCode(1).Text) And clpCode(1).Text <> "" Then
        'Disable changes to EDUC records only
        If clpCode(1).Text = "EDUC" And chkCompleted.Value = False Then
            'Disable the controls - to prevent changes to the records
            clpCode(1).Enabled = False
            clpCode(2).Enabled = False
            dlpEDate.Enabled = False
            chkCompleted.Enabled = False
            memComments.Enabled = False
            
            ButtonS.Enabled("delete") = False
            ButtonS.Enabled("save") = False
            ButtonS.Enabled("cancel") = False
        Else
            'Enable controls
            clpCode(1).Enabled = True
            clpCode(2).Enabled = True
            dlpEDate.Enabled = True
            chkCompleted.Enabled = True
            memComments.Enabled = True
            
            ButtonS.Enabled("delete") = True
            ButtonS.Enabled("save") = True
            ButtonS.Enabled("cancel") = True
        End If
    End If
End If

If Data1.Recordset.EOF Then
    lblEENUM = ""
End If
Exit Sub

Tab1_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_EFOLLOWUP", "Add")
Resume Next

End Sub

Sub Display_Value()
Dim SQLQ As String, Fdate As Variant, Tdate As Variant, countr%, x

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me, rsDATA)
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Call SET_UP_MODE
    Exit Sub
End If

Fdate = Date
Tdate = DateAdd("d", glbFOLLOWUPDAYS%, Fdate)
If IsDate(Tdate) Then
    If Year(Tdate) > 2077 Or Year(Tdate) < 1900 Then Tdate = Date
Else
    Tdate = Date
End If

SQLQ = "SELECT HR_FOLLOW_UP.*,"
If glbLinamar Then
    SQLQ = SQLQ & "right(EF_EMPNBR,3)+'-'+ left(EF_EMPNBR,LEN(EF_EMPNBR)-3) AS EMPNBR"
ElseIf glbOracle Then
    SQLQ = SQLQ & "EF_EMPNBR AS EMPNBR "
Else
    SQLQ = SQLQ & "LTRIM(STR(EF_EMPNBR)) AS EMPNBR "
End If
SQLQ = SQLQ & " FROM HR_FOLLOW_UP "
SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & Data1.Recordset!EF_FOLLOWUP_ID

'Ticket #28635
SQLQ = SQLQ & " ORDER BY EF_FDATE DESC"

If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
Call Set_Control("R", Me, rsDATA)
Call SET_UP_MODE
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fglbNew = True
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = nothingrelate
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Follow_Ups
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = True
End Property

Public Property Get Deleteble() As Boolean
Deleteble = True
End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
If fglbNew Then
    UpdateState = NewRecord
    TF = True
ElseIf rsDATA.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = FollowUp_Sec() 'True
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
Call modSTUPD(TF)
End Sub

Private Function RollBack()
On Error GoTo rr
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    glbUserUploadMode = UploadFormWithoutCheck: Unload Me
End If
rr:
End Function

Private Function EERetrieve_ReptAuth_Managers(xUserEmpNo)

Dim SQLQ As String, Fdate As Variant, Tdate As Variant, countr%, x

EERetrieve_ReptAuth_Managers = False

On Error GoTo EERError
Fdate = Date
Tdate = DateAdd("d", glbFOLLOWUPDAYS%, Fdate)
If IsDate(Tdate) Then
    If Year(Tdate) > 2077 Or Year(Tdate) < 1900 Then Tdate = Date
Else
    Tdate = Date
End If

If glbOracle Then
    SQLQ = "SELECT HR_FOLLOW_UP.* "
    SQLQ = SQLQ & ", HR_FOLLOW_UP.EF_EMPNBR AS EMPNBR, "
    SQLQ = SQLQ & " HRTABL.TB_DESC, HREMP.ED_SURNAME, HREMP.ED_FNAME "
    SQLQ = SQLQ & "  FROM HR_FOLLOW_UP,HRTABL,HREMP "
    
    SQLQ = SQLQ & " WHERE (HR_FOLLOW_UP.EF_FREAS_TABL = HRTABL.TB_NAME(+)) AND (HR_FOLLOW_UP.EF_FREAS = HRTABL.TB_KEY(+)) "
    SQLQ = SQLQ & " AND (HR_FOLLOW_UP.EF_EMPNBR = HREMP.ED_EMPNBR(+)) "
    SQLQ = SQLQ & " AND "
Else
    SQLQ = "SELECT HR_FOLLOW_UP.* ,"
    If glbLinamar Then
        SQLQ = SQLQ & "right(HR_FOLLOW_UP.EF_EMPNBR,3)+'-'+ left(HR_FOLLOW_UP.EF_EMPNBR,LEN(HR_FOLLOW_UP.EF_EMPNBR)-3) AS EMPNBR,"
    Else
        SQLQ = SQLQ & "LTRIM(STR(HR_FOLLOW_UP.EF_EMPNBR)) AS EMPNBR,"
    End If
    SQLQ = SQLQ & " HRTABL.TB_DESC, HREMP.ED_SURNAME, HREMP.ED_FNAME from (HR_FOLLOW_UP"
    SQLQ = SQLQ & " LEFT JOIN HRTABL ON (HR_FOLLOW_UP.EF_FREAS_TABL = HRTABL.TB_NAME) AND (HR_FOLLOW_UP.EF_FREAS = HRTABL.TB_KEY)) "
    SQLQ = SQLQ & " LEFT JOIN HREMP ON HR_FOLLOW_UP.EF_EMPNBR = HREMP.ED_EMPNBR"
    SQLQ = SQLQ & " WHERE "
End If
If Not glbFOLLOWUPSCOMP Then
    SQLQ = SQLQ & " HR_FOLLOW_UP.EF_FDATE <= " & Date_SQL(Tdate)
    SQLQ = SQLQ & " AND HR_FOLLOW_UP.EF_COMPLETED = 0 "
Else
    SQLQ = SQLQ & " HR_FOLLOW_UP.EF_COMPNO  = '001' "
End If

SQLQ = SQLQ & " AND " & glbSeleDeptUn
If glbNoNONE And glbNoEXEC Then        'Hemu -EXE
    SQLQ = SQLQ & " AND NOT(EF_FREAS = 'SREV' AND EF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG = 'NONE' OR ED_ORG = 'EXEC' )) "   'Hemu -EXE
ElseIf glbNoNONE Then
    SQLQ = SQLQ & " AND NOT(EF_FREAS = 'SREV' AND EF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG = 'NONE' )) "
ElseIf glbNoEXEC Then   'Hemu -EXE
    SQLQ = SQLQ & " AND NOT(EF_FREAS = 'SREV' AND EF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG = 'EXEC' )) "   'Hemu -EXE
End If
If glbLinamar Then
     SQLQ = SQLQ & " AND LEN(HR_FOLLOW_UP.EF_EMPNBR)> 3 "
End If

If chkMyFollow Then
    SQLQ = SQLQ & " AND EF_LUSER='" & glbUserID & "'"
End If

If Len(xUserEmpNo) > 0 Then
    SQLQ = SQLQ & " AND ((" & xUserEmpNo & " IN (SELECT JH_REPTAU FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=EF_EMPNBR)) "
    SQLQ = SQLQ & " OR (" & xUserEmpNo & " IN (SELECT JH_REPTAU2 FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=EF_EMPNBR)) "
    SQLQ = SQLQ & " OR (" & xUserEmpNo & " IN (SELECT JH_REPTAU3 FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=EF_EMPNBR))) "
End If

'Release 8.0 - Ticket #22682: View Own security
'If View Own not checked then do not retrieve follow ups of the User/Employee No
If Len(glbUserEmpNo) > 0 And glbUserEmpNo <> 0 And Not gSec_FollUp_ViewOwn Then
    'Do not show user's Follow Up records based on the Employee # associated to the User.
    SQLQ = SQLQ & " AND EF_EMPNBR <> " & glbUserEmpNo
End If


SQLQ = SQLQ & " ORDER BY "
Select Case EESNameSort
Case 0
    SQLQ = SQLQ & IIf(glbLinamar, "EMPNBR", "EF_EMPNBR")
    SQLQ = SQLQ & ","
    If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
        SQLQ = SQLQ & " EF_FDATE ASC, EF_COMPLETED "
    Else
        SQLQ = SQLQ & " EF_FDATE DESC, EF_COMPLETED "
    End If
Case 1
    SQLQ = SQLQ & "ED_SURNAME "
    SQLQ = SQLQ & ","
    If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
        SQLQ = SQLQ & " EF_FDATE ASC, EF_COMPLETED "
    Else
        SQLQ = SQLQ & " EF_FDATE DESC, EF_COMPLETED "
    End If
Case 2
    If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
        SQLQ = SQLQ & " EF_FDATE ASC, EF_COMPLETED "
    Else
        SQLQ = SQLQ & " EF_FDATE DESC, EF_COMPLETED "
    End If
    SQLQ = SQLQ & ","
    SQLQ = SQLQ & IIf(glbLinamar, "EMPNBR", "EF_EMPNBR")
Case 3
    SQLQ = SQLQ & " EF_FREAS "
    SQLQ = SQLQ & ","
    If glbCompSerial = "S/N - 2370W" Then 'David Chapman's Ice Cream Limited
        SQLQ = SQLQ & " EF_FDATE ASC, EF_COMPLETED "
    Else
        SQLQ = SQLQ & " EF_FDATE DESC, EF_COMPLETED "
    End If
End Select

Data1.RecordSource = SQLQ
Data1.Refresh

Set rsGrid = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

If Data1.Recordset.EOF And Not glbFollowUpsRemain Then
    Exit Function
End If
EERetrieve_ReptAuth_Managers = True

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FollowUpRetrieve", "HR_EFOLLOW_UP", "Tdate is:" & Tdate)

Resume Next

Exit Function

End Function

Private Function Get_EmployeeNo_of_UserID(xUserID)
    Dim rsSecurity As New ADODB.Recordset
    Dim SQLQ
    
    SQLQ = "SELECT EMPNBR FROM HR_SECURE_BASIC WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
    rsSecurity.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsSecurity.EOF Then
        Get_EmployeeNo_of_UserID = rsSecurity("EMPNBR")
    Else
        Get_EmployeeNo_of_UserID = ""
    End If
    rsSecurity.Close
    Set rsSecurity = Nothing
End Function

Private Function FollowUp_Sec() As Boolean
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim retVal As Boolean
    Dim xTemplate As String
    
    '????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
    xTemplate = ""
    xTemplate = Get_Template(glbUserID)
    
    strSQL = "SELECT MAINTAINABLE FROM HR_SECURE_FOLLOW_UP WHERE "
    'strSQL = "SELECT ACCESSABLE FROM HR_SECURE_FOLLOW_UP WHERE "
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        strSQL = strSQL & "CODENAME='" & clpCode(1).Text & "' AND USERID='" & Replace(glbUserID, "'", "''") & "'"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        strSQL = strSQL & "CODENAME='" & clpCode(1).Text & "' AND USERID='" & Replace(xTemplate, "'", "''") & "'"
    End If
    rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF = False And rs.BOF = False Then
        retVal = Abs(rs("MAINTAINABLE"))
        'retVal = Abs(rs("ACCESSABLE"))
    Else
        retVal = False
    End If
    
    FollowUp_Sec = retVal
End Function

Private Sub Cri_Sec()
    Dim EECri As String
    Dim strSec As String
    
    strSec = buildSec_FollowUp
    If Len(strSec) >= 1 Then
        EECri = "{HR_FOLLOW_UP.EF_FREAS} " & Replace(Replace(strSec, "(", "["), ")", "]")
    End If
    
    If Len(EECri) >= 1 Then
        If Len(glbstrSelCri) > 0 Then
            glbstrSelCri = glbstrSelCri & " AND " & EECri
        Else
            glbstrSelCri = EECri
        End If
    End If
    
End Sub

