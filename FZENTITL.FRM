VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#60.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmREntitle 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Entitlements Report"
   ClientHeight    =   7155
   ClientLeft      =   255
   ClientTop       =   615
   ClientWidth     =   10080
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
   ScaleHeight     =   7155
   ScaleWidth      =   10080
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1995
      MaxLength       =   2
      TabIndex        =   11
      Tag             =   "00-Employee Position Shift"
      Top             =   4020
      Visible         =   0   'False
      Width           =   450
   End
   Begin Threed.SSCheck chkAtt 
      Height          =   255
      Left            =   60
      TabIndex        =   16
      Tag             =   "If X-Show Attendance Details"
      Top             =   5220
      Width           =   2835
      _Version        =   65536
      _ExtentX        =   5001
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "   Show Attendance Details"
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
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
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
      Left            =   2005
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Tag             =   "First Level of grouping records"
      Top             =   6450
      Width           =   2325
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
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
      Left            =   2005
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Tag             =   "First Level of grouping records"
      Top             =   6135
      Width           =   2325
   End
   Begin VB.ComboBox cmbHours 
      Appearance      =   0  'Flat
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
      ItemData        =   "FZENTITL.frx":0000
      Left            =   1980
      List            =   "FZENTITL.frx":0002
      TabIndex        =   7
      Tag             =   "Choose display in hours or days"
      Text            =   "Combo1"
      Top             =   2670
      Width           =   1215
   End
   Begin Threed.SSCheck chkEntitle 
      Height          =   225
      Left            =   60
      TabIndex        =   14
      Tag             =   "If X-Show Hourly Entitlements"
      Top             =   4740
      Width           =   2835
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "   Show Hourly Entitlements"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCheck chkCompTm 
      Height          =   225
      Left            =   60
      TabIndex        =   15
      Tag             =   "If X-Show Compensatory Time"
      Top             =   4980
      Width           =   2835
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "   Show Compensatory Time"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1680
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2010
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1350
      Width           =   7000
      _ExtentX        =   12356
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDOR"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1020
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   690
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   360
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   10
      Tag             =   "00-Enter Section Code"
      Top             =   3690
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1680
      TabIndex        =   9
      Tag             =   "00-Enter Administered By Code"
      Top             =   3360
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   8
      Tag             =   "00-Enter Region Code"
      Top             =   3030
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2340
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7075
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   0
      Top             =   6500
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowWidth     =   480
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   2
      WindowControls  =   -1  'True
      MarginTop       =   720
      MarginBottom    =   720
      PrintFileLinesPerPage=   60
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3330
      TabIndex        =   13
      Tag             =   "40-Date upto and including this date forward"
      Top             =   4360
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1255
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   12
      Tag             =   "40-Date from and including this date forward"
      Top             =   4360
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1255
   End
   Begin VB.Label lblFromTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vacation Entitl. Date"
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
      TabIndex        =   35
      Top             =   4410
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblShift 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
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
      TabIndex        =   34
      Top             =   4035
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      TabIndex        =   33
      Top             =   2010
      Width           =   630
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
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
      TabIndex        =   32
      Top             =   3690
      Width           =   540
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
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
      TabIndex        =   31
      Top             =   3030
      Width           =   510
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   30
      Top             =   3360
      Width           =   1125
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
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
      TabIndex        =   29
      Top             =   1020
      Width           =   615
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Final Sort"
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
      Index           =   1
      Left            =   90
      TabIndex        =   28
      Top             =   6480
      Width           =   660
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #1"
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
      Index           =   0
      Left            =   90
      TabIndex        =   27
      Top             =   6165
      Width           =   885
   End
   Begin VB.Label lblRepGrp 
      Caption         =   "Report Grouping"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   5925
      Width           =   1575
   End
   Begin VB.Label lblSelCri 
      BackStyle       =   0  'Transparent
      Caption         =   "Selection Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   90
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Display"
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
      Index           =   5
      Left            =   60
      TabIndex        =   24
      Top             =   2670
      Width           =   510
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      TabIndex        =   23
      Top             =   1695
      Width           =   450
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
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
      TabIndex        =   22
      Top             =   1350
      Width           =   420
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
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
      Index           =   2
      Left            =   60
      TabIndex        =   21
      Top             =   2340
      Width           =   1290
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      TabIndex        =   20
      Top             =   690
      Width           =   825
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
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
      TabIndex        =   19
      Top             =   390
      Width           =   555
   End
End
Attribute VB_Name = "frmREntitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jaddy 10/26/99 add a CheckBox and edit lblTitle(5).caption="Display"

Option Explicit

Private Sub chkAtt_Click(Value As Integer)
    If chkAtt Then chkCompTm = False: chkEntitle = False  'Jaddy 11/5/99
End Sub

Private Sub chkAtt_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkCompTm_Click(Value As Integer)
    If chkCompTm Then chkAtt.Value = False 'Jaddy 11/5/99
End Sub

Private Sub chkCompTm_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkEntitle_Click(Value As Integer)
    If chkEntitle Then chkAtt.Value = False 'Jaddy 11/5/99
End Sub

Private Sub chkEntitle_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbHours_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Entitlements Report Criteria", Me) Then Exit Sub
    Screen.MousePointer = HOURGLASS
     Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
     Call set_PrintState(True)
End If
Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString, , "INFO:HR"
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

'Private Sub cmdPrint_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
     Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
     Call set_PrintState(True)
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString, , "INFO:HR"
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ENTITLEMENTS", "VIEW")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

'Private Sub cmdView_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub



Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    Select Case intIdx%
    Case 0: strCd$ = "HREMP.ED_LOC"
    Case 1: strCd$ = "HREMP.ED_ORG"
    Case 2: strCd$ = "HREMP.ED_EMP"
    Case 3: strCd$ = "HREMP.ED_REGION"
    Case 4: strCd$ = "HREMP.ED_ADMINBY"
    Case 5: strCd$ = "HREMP.ED_SECTION"
    End Select
        CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
    End If
End If

If Len(CodeCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CodeCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CodeCri
    End If
    glbiOneWhere = True
End If
End Sub

Private Sub Cri_Div()
Dim DivCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level

If Len(clpDiv.Text) > 0 Then
    DivCri = "({HREMP.ED_DIV} = '" & clpDiv.Text & "')"
End If

If Len(DivCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = DivCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & DivCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_EE()
Dim EECri As String

If Len(elpEEID.Text) > 0 Then
    EECri = "{HREMP.ED_EMPNBR} IN [" & getEmpnbr(elpEEID.Text) & "] "
End If

If Len(EECri) >= 1 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_PT()
Dim EECri As String, OneSet%, x%

If Len(clpPT.Text) < 1 Then Exit Sub

EECri = "{HREMP.ED_PT}= '" & clpPT.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_FTDates()

Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%, x%
Dim FromDate, ToDate, SQLQ
Dim RsHRPARCO As New ADODB.Recordset

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HREMP.ED_EFDATE} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = Month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) and "
    
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = Month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " ({HREMP.ED_ETDATE} <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
End If


Cri_FTDatst:
If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = TempCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Function Cri_SetAll()
Dim x%, xNoFile, xNoWork
Dim WHSCCrpt As Boolean
Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""
WHSCCrpt = False
' call cri models set both glbiONeWhere and strSelCri

'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_PT
Call Cri_Shift
Call Cri_EE
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For x% = 0 To 5
    Call Cri_Code(x%)
Next x%

'Hemu - Muir's Cartage
If glbCompSerial = "S/N - 2340W" Then
    Call Cri_FTDates
End If
'Hemu

If chkAtt Then
    glbstrSelCri = glbstrSelCri & " AND {HRATTWRK.AD_WRKEMP}='" & glbUserID & "'"
    Call AttDetails
End If
    
If chkCompTm Then
    glbstrSelCri = glbstrSelCri & " AND {HRENTWRK.TE_WRKEMP}='" & glbUserID & "'"
    Call CompTime
End If
' danielk - 12/30/2002 - 7.0 Priority C changes
' Franks - Sep 16,02 for #2823
If glbCompSerial = "S/N - 2297W" Then 'OHTN
    If chkEntitle And cmbHours = "Days" Then
        Call ProjectYTD
    End If
End If
' danielk - 12/30/2002 - end
xNoFile = 7
xNoWork = 99
' report name
If glbCompSerial = "S/N - 2235W" Or glbCompSerial = "S/N - 2236W" Then   'laura 03/09/98
  If comGroup(0).Text <> "(none)" Then
      If chkAtt = True Then          'Jaddy 10/26/99
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsentat1.rpt"  'Jaddy 10/26/99
            xNoFile = 5          'Jaddy 10/26/99
      Else          'Jaddy 10/26/99
        If chkEntitle = True And chkCompTm = True Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsenthr5.rpt"
            xNoFile = 8
            xNoWork = 7
        End If
        If chkEntitle = True And chkCompTm = False Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsenthr2.rpt"
            xNoFile = 7
        End If
        If chkEntitle = False And chkCompTm = True Then
            '----------------------
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsentCm1.rpt"
            xNoFile = 8
            xNoWork = 7
        End If
        If chkEntitle = False And chkCompTm = False Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsenthrs.rpt"
            xNoFile = 5
        End If
    End If          'Jaddy 10/26/99
  Else
      If chkAtt = True Then          'Jaddy 10/26/99
         Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsentat2.rpt"  'Jaddy 10/26/99
         xNoFile = 5          'Jaddy 10/26/99
      Else          'Jaddy 10/26/99
        If chkEntitle = True And chkCompTm = True Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsenthr7.rpt"
            xNoFile = 8
            xNoWork = 7
        End If
        If chkEntitle = True And chkCompTm = False Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsenthr4.rpt"
            xNoFile = 7
        End If
        If chkEntitle = False And chkCompTm = True Then
            '----------------------
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsentCm3.rpt"
            xNoFile = 8
            xNoWork = 7
        End If
        If chkEntitle = False And chkCompTm = False Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsenthr1.rpt"
            xNoFile = 5
        End If
    End If      'Jaddy 10/26/99
  End If
Else
  If comGroup(0).Text <> "(none)" Then
    If cmbHours.ListIndex = 0 Then
        If chkAtt = True Then          'Jaddy 10/26/99
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentat1.rpt"  'Jaddy 10/26/99
            xNoFile = 5          'Jaddy 10/26/99
            xNoWork = 5
        Else          'Jaddy 10/26/99
            If chkEntitle = True And chkCompTm = True Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr5.rpt"
                xNoFile = 8
                xNoWork = 7
            End If
            If chkEntitle = True And chkCompTm = False Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr2.rpt"
                xNoFile = 7
            End If
            If chkEntitle = False And chkCompTm = True Then
                '----------------------
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentCm1.rpt"
                xNoFile = 9
                xNoWork = 7
            End If
            If chkEntitle = False And chkCompTm = False Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthrs.rpt"
                xNoFile = 5
            End If
        End If          'Jaddy 10/26/99
    Else
        If chkAtt = True Then          'Jaddy 10/26/99
            If Not glbWHSCC Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentat2.rpt"  'Jaddy 10/26/99
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "whentat2.rpt"
                WHSCCrpt = True
            End If
            xNoFile = 5          'Jaddy 10/26/99
            xNoWork = 5
        Else          'Jaddy 10/26/99
            If chkEntitle = True And chkCompTm = True Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr6.rpt"
                xNoFile = 8
                xNoWork = 7
            End If
            If chkEntitle = True And chkCompTm = False Then
                ' danielk - 12/30/2002 - 7.0 Priority C changes
                ' Franks - Sep 16,02 for #2823
                If glbCompSerial = "S/N - 2297W" Then 'OHTN
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2297r3.rpt"
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr3.rpt"
                End If
                ' danielk - 12/30/2002 - end
                xNoFile = 7
            End If
            If chkEntitle = False And chkCompTm = True Then
                '----------------------
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentCm2.rpt"
                xNoFile = 8
                xNoWork = 7
            End If
            If chkEntitle = False And chkCompTm = False Then
                If Not glbWHSCC Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentdys.rpt"
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "whentdys.rpt"
                    WHSCCrpt = True
                End If
                xNoFile = 5
            End If
        End If  'Jaddy 10/26/99
    End If
    
  Else
    If cmbHours.ListIndex = 0 Then
        If chkAtt = True Then          'Jaddy 10/26/99
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentat1.rpt"  'Jaddy 10/26/99
            xNoFile = 5          'Jaddy 10/26/99
            xNoWork = 5
        Else          'Jaddy 10/26/99
            If chkEntitle = True And chkCompTm = True Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr7.rpt"
                xNoFile = 8
                xNoWork = 7
            End If
            If chkEntitle = True And chkCompTm = False Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr4.rpt"
                xNoFile = 7
            End If
            If chkEntitle = False And chkCompTm = True Then
                '----------------------
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentCm3.rpt"
                xNoFile = 8
                xNoWork = 7
            End If
            If chkEntitle = False And chkCompTm = False Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr1.rpt"
                xNoFile = 5
            End If
        End If          'Jaddy 10/26/99
    Else
        If chkAtt = True Then          'Jaddy 10/26/99
            If Not glbWHSCC Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentat2.rpt"  'Jaddy 10/26/99
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "whentat2.rpt"
                WHSCCrpt = True
            End If
            xNoFile = 5          'Jaddy 10/26/99
            xNoWork = 5
        Else          'Jaddy 10/26/99
            xNoFile = 5
            If chkEntitle = True And chkCompTm = True Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentdy7.rpt"
                xNoFile = 8
                xNoWork = 7
            End If
            ' danielk - 12/30/2002 - 7.0 Priority C changes
            If chkEntitle = True And chkCompTm = False Then
                ' Franks - Sep 16,02 for #2823
                If glbCompSerial = "S/N - 2297W" Then 'OHTN
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2297Y4.rpt"
                    xNoFile = 7
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentdy4.rpt"
                End If
            End If
            ' danielk - 12/30/2002 - end
            If chkEntitle = False And chkCompTm = True Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentdy6.rpt"
                xNoFile = 8
                xNoWork = 7
            End If
            If chkEntitle = False And chkCompTm = False Then
                If Not glbWHSCC Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentdy1.rpt"
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "whentdy1.rpt"
                    WHSCCrpt = True
                End If
            End If
        End If 'Jaddy 10/26/99
    End If
  End If
End If

xNoFile = xNoFile + 1
'MsgBox Me.vbxCrystal.ReportFileName
' set to sorting/grouping criteria
x% = Cri_Sorts()   ' returns number of sections formated

'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
    If glbWHSCC And WHSCCrpt Then
        If Right(Me.vbxCrystal.ReportFileName, 12) = "whentdy1.rpt" Then
            vbxCrystal.SubreportToChange = "ASLdisp"
        Else
            vbxCrystal.SubreportToChange = "ASL_Disp"
        End If
        vbxCrystal.Connect = RptODBC_SQL
        vbxCrystal.SubreportToChange = "USB_Disp"
        vbxCrystal.Connect = RptODBC_SQL
        vbxCrystal.SubreportToChange = ""
    End If
Else
    If xNoWork < 99 Then
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To xNoFile - 1
            If xNoWork = x Then
              Me.vbxCrystal.DataFiles(x%) = glbIHRDBW
            Else
              Me.vbxCrystal.DataFiles(x%) = glbIHRDB
            End If
        Next x%
    Else
        Me.vbxCrystal.Connect = RptODBC_SQL
    End If

End If
' set security for database


' window title if appropriate
Me.vbxCrystal.WindowTitle = "Entitlements Report"

Cri_SetAll = True

Screen.MousePointer = DEFAULT
Exit Function


modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "ENTITLEMENT Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function
Private Sub Cri_Shift()
Dim EECri As String, OneSet%, x%

If Len(txtShift.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub
Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

Cri_Sorts = 0
grpField$ = getEGroup(comGroup(0).Text)
If comGroup(0).Text = "Employee Name" Then grpField$ = "{@EEFullName}"

grpField$ = Replace(grpField$, "HRTABL.", "HRTABL1.")
If comGroup(0) <> "(none)" Then
    ' <====
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup1 = '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    dscGroup$ = "descName1 = " & grpField$
    Me.vbxCrystal.Formulas(1) = dscGroup$
    
    '<====

    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
    grpCond$ = "GROUP" & CStr(2) & ";" & "{@EEFullName}" & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(1) = grpCond$

    Me.vbxCrystal.SectionFormat(0) = "GH1;T;T;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(1) = "GF1;F,F;F;X;X;X;X"
Else
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EEFullName}"
    End Select
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
    Me.vbxCrystal.SectionFormat(0) = "GH1;F;F;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(1) = "GF1;F,F;F;X;X;X;X"
    dscGroup$ = "descGroup1 = ''"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    dscGroup$ = "descName1 = ''"
    Me.vbxCrystal.Formulas(1) = dscGroup$
End If
'If chkAtt = False Then  'Jaddy 10/26/99 Jaddy 11/5/99
 Me.vbxCrystal.Formulas(2) = "DESCSICK =if not isnull({HREMP.ED_EFDATES}) or not isnull({HREMP.ED_ETDATES}) then totext({HREMP.ED_EFDATES}) +' TO ' + totext({HREMP.ED_ETDATES}) "
 Me.vbxCrystal.Formulas(3) = "DESCVAC = if not isnull({HREMP.ED_EFDATE}) or not isnull({HREMP.ED_ETDATE}) then totext({HREMP.ED_EFDATE}) +' TO ' + totext({HREMP.ED_ETDATE}) "
'End If  'Jaddy 10/26/99 Jaddy 11/5/99

Cri_Sorts = z% ' next section number to format

End Function

Private Function CriCheck()
Dim x%

CriCheck = False

If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    MsgBox lStr("If Division Entered - it must be known"), , "INFO:HR"
     clpDiv.SetFocus
    Exit Function
End If

If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known", , "INFO:HR"
     clpDept.SetFocus
    Exit Function
End If


For x% = 0 To 5
    If Not clpCode(x).ListChecker Then Exit Function

Next x%

If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    MsgBox lStr("Category code must be valid")
     clpPT.SetFocus
    Exit Function
End If
If Not elpEEID.ListChecker Then
    Exit Function
End If

CriCheck = True
End Function

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
                        
Screen.MousePointer = HOURGLASS
If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If
Call setRptCaption(Me)
Call comGrpLoad
If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
If glbLinamar Then clpCode(3).MaxLength = 8
If glbCompSerial = "S/N - 2235W" Then   'laura 03/09/98
    lblTitle(5).Visible = False
    cmbHours.Visible = False
ElseIf glbCompSerial = "S/N - 2236W" Then
    lblTitle(5).Visible = False
    cmbHours.Visible = False
Else
    lblTitle(5).Visible = True
    cmbHours.Visible = True
End If
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

'For Muir's Cartage (Ticket # 5626) - to be made across the board change in the next release
If glbCompSerial = "S/N - 2340W" Then
    lblFromTo.Visible = True
    dlpDateRange(0).Visible = True
    dlpDateRange(1).Visible = True
Else
    lblFromTo.Visible = False
    dlpDateRange(0).Visible = False
    dlpDateRange(1).Visible = False
End If

chkEntitle = False
Call INI_Controls(Me)
Screen.MousePointer = DEFAULT


End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub


Private Sub CompTime()
Dim SQLQ, SQLX, WSQLQ
On Error GoTo CompTm_Err

SQLQ = "DELETE FROM HRENTWRK " & in_SQL(glbIHRDBW) & " WHERE TE_WRKEMP='" & glbUserID & "'"

gdbAdoIhr001.Execute SQLQ

WSQLQ = WSQLQ & " WHERE " & glbSeleDeptUn
If clpDiv.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DIV='" & clpDiv.Text & "'"
If clpDept.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DEPTNO='" & clpDept.Text & "'"
If clpCode(0).Text <> "" Then WSQLQ = WSQLQ & "AND ED_LOC='" & clpCode(0).Text & "'"
If clpCode(1).Text <> "" Then WSQLQ = WSQLQ & "AND ED_ORG='" & clpCode(1).Text & "'"
If clpCode(2).Text <> "" Then WSQLQ = WSQLQ & "AND ED_EMP='" & clpCode(2).Text & "'"
If clpCode(3).Text <> "" Then WSQLQ = WSQLQ & "AND ED_REGION='" & clpCode(3).Text & "'"
If clpCode(4).Text <> "" Then WSQLQ = WSQLQ & "AND ED_ADMINBY='" & clpCode(4).Text & "'"
If clpCode(5).Text <> "" Then WSQLQ = WSQLQ & "AND ED_SECTION='" & clpCode(5).Text & "'"
If elpEEID.Text <> "" Then WSQLQ = WSQLQ & "AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ")"
If clpPT.Text <> "" Then WSQLQ = WSQLQ & "AND ED_PT='" & clpPT.Text & "'"

SQLQ = " SELECT '001' AS TE_COMPNO,AD_EMPNBR,'ADRE' AS TE_REASON_TABL,'CTOT' AS TE_REASON,"
If glbOracle Then
    SQLQ = SQLQ & " SUM(CASE WHEN substr(AD_REASON,1,2)='OT' THEN AD_HRS ELSE 0 END) AS TE_EARNHRS,"
    SQLQ = SQLQ & " SUM(CASE WHEN substr(AD_REASON,1,2)='CT' THEN AD_HRS ELSE 0 END) AS TE_USEDHRS"
ElseIf glbSQL Then
    SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='OT' THEN AD_HRS ELSE 0 END) AS TE_EARNHRS,"
    SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='CT' THEN AD_HRS ELSE 0 END) AS TE_USEDHRS"
Else
    SQLQ = SQLQ & " SUM(IIF(LEFT(AD_REASON,2)='OT',AD_HRS,0)) AS TE_EARNHRS, "
    SQLQ = SQLQ & " SUM(IIF(LEFT(AD_REASON,2)='CT',AD_HRS,0)) AS TE_USEDHRS "
End If

SQLQ = SQLQ & ",'" & glbUserID & "' AS TE_WRKEMP"
SQLQ = SQLQ & " FROM HR_ATTENDANCE WHERE AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ & ")"
SQLQ = SQLQ & " GROUP BY AD_COMPNO,AD_EMPNBR,AD_REASON_TABL "

SQLX = "INSERT INTO HRENTWRK "
SQLX = SQLX & "(TE_COMPNO,TE_EMPNBR,TE_REASON_TABL,TE_REASON,TE_EARNHRS,TE_USEDHRS,TE_WRKEMP)"
SQLX = SQLX & in_SQL(glbIHRDBW)
SQLX = SQLX & SQLQ
gdbAdoIhr001.Execute SQLX


SQLX = "INSERT INTO HRENTWRK (TE_COMPNO,TE_EMPNBR,TE_REASON_TABL,TE_REASON,TE_WRKEMP)"
SQLX = SQLX & in_SQL(glbIHRDBW)
SQLX = SQLX & "SELECT ED_COMPNO,ED_EMPNBR,'ADRE' AS AD_REASON_TABL,'SHOW' AS AD_REASON "
SQLX = SQLX & " ,'" & glbUserID & "' AS AD_WRKEMP"
SQLX = SQLX & " FROM HREMP " & WSQLQ
'Add by Franks Dec 4, 2001
'The previues code caused the problem: if both Show Hourly Entitlements and Show Compensatory Time checked,
'Compensatory Time Data not shown, 'SHOW' and Hourly Entitlement code on one can insert in the MDB
'If SHOW' and Hourly Entitlement code both inserted in the MDB, that will cause a problem
SQLX = SQLX & " AND ED_EMPNBR NOT IN (SELECT TE_EMPNBR FROM HRENTWRK "
SQLX = SQLX & in_SQL(glbIHRDBW)
SQLX = SQLX & " WHERE HRENTWRK.TE_WRKEMP = '" & glbUserID & "' "
SQLX = SQLX & ") "
'Add by Franks Dec 4, 2001
gdbAdoIhr001.Execute SQLX


Exit Sub

CompTm_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CompTime", "WORK File", "CREATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Sub
Private Sub AttDetails()

Dim SQLQ, SQLX, WSQLQ

On Error GoTo CompTm_Err

gdbAdoIhr001.Execute "DELETE FROM HRATTWRK " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"

WSQLQ = glbSeleDeptUn
If clpDiv.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DIV='" & clpDiv.Text & "'"
If clpDept.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DEPTNO='" & clpDept.Text & "'"
If clpCode(0).Text <> "" Then WSQLQ = WSQLQ & "AND ED_LOC='" & clpCode(0).Text & "'"
If clpCode(1).Text <> "" Then WSQLQ = WSQLQ & "AND ED_ORG='" & clpCode(1).Text & "'"
If clpCode(2).Text <> "" Then WSQLQ = WSQLQ & "AND ED_EMP='" & clpCode(2).Text & "'"
If clpCode(3).Text <> "" Then WSQLQ = WSQLQ & "AND ED_REGION='" & clpCode(3).Text & "'"
If clpCode(4).Text <> "" Then WSQLQ = WSQLQ & "AND ED_ADMINBY='" & clpCode(4).Text & "'"
If clpCode(5).Text <> "" Then WSQLQ = WSQLQ & "AND ED_SECTION='" & clpCode(5).Text & "'"
If elpEEID.Text <> "" Then WSQLQ = WSQLQ & "AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ")"
If clpPT.Text <> "" Then WSQLQ = WSQLQ & "AND ED_PT='" & clpPT.Text & "'"

SQLX = "INSERT INTO HRATTWRK (AD_COMPNO,AD_EMPNBR,AD_REASON_TABL,AD_REASON,AD_WRKEMP)"
SQLX = SQLX & in_SQL(glbIHRDBW)
SQLX = SQLX & "SELECT ED_COMPNO,ED_EMPNBR,'ADRE' AS AD_REASON_TABL,'SHOW' AS AD_REASON "
SQLX = SQLX & " ,'" & glbUserID & "' AS AD_WRKEMP"
SQLX = SQLX & " FROM HREMP WHERE " & WSQLQ

gdbAdoIhr001.Execute SQLX

SQLQ = "SELECT AD_COMPNO,AD_EMPNBR,'ADRE' AS AD_REASON_TABL,AD_REASON,AD_DOA,AD_HRS "
SQLQ = SQLQ & " ,'" & glbUserID & "' AS AD_WRKEMP"

If glbOracle Then
    SQLQ = SQLQ & " FROM HR_ATTENDANCE,HREMP WHERE HR_ATTENDANCE.AD_EMPNBR (+) = HREMP.ED_EMPNBR"
    SQLQ = SQLQ & " AND " & WSQLQ
Else
    SQLQ = SQLQ & " FROM HR_ATTENDANCE LEFT JOIN HREMP ON HR_ATTENDANCE.AD_EMPNBR = HREMP.ED_EMPNBR"
    SQLQ = SQLQ & " WHERE " & WSQLQ
End If
SQLQ = SQLQ & " AND ((HR_ATTENDANCE.AD_REASON like 'VAC%' "
SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_DOA>=HREMP.ED_EFDATE "
SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_DOA<=HREMP.ED_ETDATE) "
SQLQ = SQLQ & " OR  (HR_ATTENDANCE.AD_REASON like 'SIC%' "
SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_DOA>=HREMP.ED_EFDATES "
SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_DOA<=HREMP.ED_ETDATES))"


SQLX = "INSERT INTO HRATTWRK (AD_COMPNO,AD_EMPNBR,AD_REASON_TABL,AD_REASON,AD_DOA,AD_HRS,AD_WRKEMP) "
SQLX = SQLX & in_SQL(glbIHRDBW)
SQLQ = SQLX & SQLQ


gdbAdoIhr001.Execute SQLQ

Exit Sub

CompTm_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CompTime", "WORK File", "CREATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Sub



Private Sub comGrpLoad()

comGroup(0).AddItem lStr("Division")
comGroup(0).AddItem lStr("Department")
comGroup(0).AddItem lStr("Location")
comGroup(0).AddItem lStr("Union")
comGroup(0).AddItem lStr("Administered By")
comGroup(0).AddItem "Employee Name"
comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000

If glbLinamar Then ' Frank May 2,2001
    comGroup(0).AddItem "Employment Type"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem ("Home Line")
End If
If Not glbMulti Then comGroup(0).AddItem "Shift"
comGroup(0).AddItem "(none)"
comGroup(0).ListIndex = 0
comGroup(1).AddItem "Employee Name"
comGroup(1).ListIndex = 0
comGroup(1).Enabled = False

cmbHours.AddItem "Hours"
cmbHours.AddItem "Days"
cmbHours.ListIndex = 0
End Sub

Private Sub txtShift_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub ProjectYTD()
Dim rsVacEnt As New ADODB.Recordset
Dim rsVE As New ADODB.Recordset
Dim snapEntitle As New ADODB.Recordset
Dim rsWrk As New ADODB.Recordset
Dim SQLQ, SQLX, WSQLQ
Dim VacRule(16, 5), I, K, xOrder, nOrder
Dim xDiv, xDept, xORG, xEMP, xPT, xEDate, xGRPCD
Dim EmpNo As Long
Dim dblEntitle#, dblPrevEntitle#, strDivision$
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim xYTD, xYTD0, fglbWDate$, fglbWDate01$

    SQLQ = "DELETE FROM HRENTWRK " & IIf(Not glbSQL, " IN '" & glbIHRDBW & "' ", "")
    gdbAdoIhr001.Execute SQLQ

    Call Pause(3)
    
    SQLQ = "SELECT VE_DIV,VE_DEPT,VE_ORG,VE_EMP,VE_PT,VE_EDATE,VE_GRPCD FROM HRVACENT "
    SQLQ = SQLQ & "GROUP BY VE_DIV,VE_DEPT,VE_ORG,VE_EMP,VE_PT,VE_EDATE,VE_GRPCD"
    If rsVacEnt.State <> 0 Then rsVacEnt.Close
    rsVacEnt.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    Do While Not rsVacEnt.EOF
        xDiv = "": xDept = "": xORG = "": xEMP = "": xPT = "": xEDate = "": xGRPCD = ""
        If Not IsNull(rsVacEnt("VE_DIV")) Then
            xDiv = rsVacEnt("VE_DIV")
        End If
        If Not IsNull(rsVacEnt("VE_DEPT")) Then
            xDept = rsVacEnt("VE_DEPT")
        End If
        If Not IsNull(rsVacEnt("VE_ORG")) Then
            xORG = rsVacEnt("VE_ORG")
        End If
        If Not IsNull(rsVacEnt("VE_EMP")) Then
            xEMP = rsVacEnt("VE_EMP")
        End If
        If Not IsNull(rsVacEnt("VE_PT")) Then
            xPT = rsVacEnt("VE_PT")
        End If
        If Not IsNull(rsVacEnt("VE_EDATE")) Then
            xEDate = rsVacEnt("VE_EDATE")
        End If
        If Not IsNull(rsVacEnt("VE_GRPCD")) Then
            xGRPCD = rsVacEnt("VE_GRPCD")
        End If
        
        'Get data from HRVACENT
        SQLQ = "SELECT * FROM HRVACENT WHERE (1 =1) "
        If Len(xDiv) > 0 Then
            SQLQ = SQLQ & "AND VE_DIV = '" & xDiv & "' "
        End If
        If Len(xDept) > 0 Then
            SQLQ = SQLQ & "AND VE_DEPT = '" & xDept & "' "
        End If
        If Len(xORG) > 0 Then
            SQLQ = SQLQ & "AND VE_ORG = '" & xORG & "' "
        End If
        If Len(xEMP) > 0 Then
            SQLQ = SQLQ & "AND VE_EMP = '" & xEMP & "' "
        End If
        If Len(xPT) > 0 Then
            SQLQ = SQLQ & "AND VE_PT = '" & xPT & "' "
        End If
        If IsDate(xEDate) Then
        SQLQ = SQLQ & "AND VE_EDATE = " & IIf(glbSQL, "", "CVDATE") & "('" & Format(xEDate, "mmm dd,yyyy") & "')"
        End If
        If Len(xGRPCD) > 0 Then
            SQLQ = SQLQ & "AND VE_GRPCD = '" & xGRPCD & "' "
        End If
        SQLQ = SQLQ & "Order By VE_DIV,VE_DEPT,VE_ORG, VE_EDATE,VE_EMP,VE_PT,VE_ORDER "
        
        rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset

        nOrder = 0
        Do While Not rsVE.EOF
            xOrder = rsVE("VE_ORDER")
            nOrder = Format(Val(xOrder), "##0") - 1
            If Not (nOrder < 0 Or nOrder > 16) Then
                I = nOrder
                VacRule(I, 1) = nOrder
                VacRule(I, 2) = rsVE("VE_BMONTH")
                VacRule(I, 3) = rsVE("VE_EMONTH")
                VacRule(I, 4) = rsVE("VE_ENTITLE")
                VacRule(I, 5) = rsVE("VE_TYPE")
            End If
            rsVE.MoveNext
        Loop
        rsVE.Close
        K = nOrder
        
        'Get the employee recordset
        SQLQ = "SELECT HREMP.ED_EMPNBR, qry_JobCurrent.JB_GRPCD, HREMP.ED_VACPC, HREMP.ED_PVAC, HREMP.ED_VAC, HREMP.ED_PSICK, HREMP.ED_SICK, qry_JobCurrent.JH_DHRS, HREMP.ED_DIV, HREMP.ED_EMP, HREMP.ED_DEPTNO, HREMP.ED_PT, HREMP.ED_DOH, HREMP.ED_SENDTE, HREMP.ED_UNION, HREMP.ED_LTHIRE, HREMP.ED_USRDAT1, HREMP.ED_ORG, qry_JobCurrent.JH_FTENUM, qry_JobCurrent.JH_DHRS, HREMP.ED_SECTION, "
        SQLQ = SQLQ & "HREMP.ED_EFDATE, HREMP.ED_ETDATE "
        SQLQ = SQLQ & " FROM HREMP LEFT JOIN qry_JobCurrent ON HREMP.ED_EMPNBR = qry_JobCurrent.JH_EMPNBR "

        SQLQ = SQLQ & "WHERE (1=1) "
        If Len(xDept) > 0 Then SQLQ = SQLQ & " AND  ED_DEPTNO = '" & xDept & "' "
        If Len(xDiv) > 0 Then SQLQ = SQLQ & " AND ED_DIV = '" & xDiv & "' "
        If Len(xORG) > 0 Then SQLQ = SQLQ & " AND ED_ORG = '" & xORG & "' "
        If Len(xEMP) > 0 Then SQLQ = SQLQ & " AND ED_EMP = '" & xEMP & "' "
        If Len(xPT) > 0 Then SQLQ = SQLQ & " AND ED_PT = '" & xPT & "' "
        If Len(xGRPCD) > 0 Then SQLQ = SQLQ & " AND JB_GRPCD = '" & xGRPCD & "'; "
        
        If snapEntitle.State <> 0 Then snapEntitle.Close
        snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Do While Not snapEntitle.EOF

            EmpNo& = snapEntitle("ED_EMPNBR")

            If IsNull(snapEntitle("ED_EFDATE")) Then GoTo lblNextRec
            fglbWDate01$ = snapEntitle("ED_EFDATE")
            fglbWDate$ = DateAdd("D", -1, DateAdd("M", 1, fglbWDate01$))
            
            varStartDate = snapEntitle("ED_DOH")
            
            If Not IsNumeric(snapEntitle("JH_DHRS")) Then
                dblDHours# = 0
            Else
                dblDHours# = snapEntitle("JH_DHRS")
            End If
            
            If Not IsNumeric(snapEntitle("JH_FTENUM")) Then
                dblFTEHours# = 0
            Else
                dblFTEHours# = snapEntitle("JH_FTENUM")
            End If
            
            'dblServiceYears# = (DateDiff("d", varStartDate, CVDate(fglbWDate$)) / 365) * 12
            dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(fglbWDate$))
            intWhereFit& = -1   ' first record can be just less than
            xYTD = 0
            For I = 1 To 12
                For x% = 0 To K '16 xYTD
                    If dblServiceYears# >= VacRule(x%, 2) And dblServiceYears# <= VacRule(x%, 3) Then
                        intWhereFit& = x%
                        xYTD0 = VacRule(x%, 4)
                        If VacRule(x%, 5) = "D" Then           ' Entitlements entered in days
                            xYTD = xYTD + xYTD0
                        End If
                        If VacRule(x%, 5) = "H" Then
                            If dblDHours# > 0 Then
                                xYTD = xYTD + (xYTD0 / dblDHours#)
                            End If
                        End If
                        If VacRule(x%, 5) = "F" Then
                            If dblDHours# > 0 Then
                                xYTD = xYTD + (xYTD0 * dblFTEHours#) / dblDHours#
                            End If
                        End If
                    End If
                Next x%
                fglbWDate01$ = DateAdd("M", 1, fglbWDate01$)
                fglbWDate$ = DateAdd("D", -1, DateAdd("M", 1, fglbWDate01$))
                'dblServiceYears# = (DateDiff("d", varStartDate, CVDate(fglbWDate$)) / 365) * 12
                dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(fglbWDate$))
            Next I
            
            If xYTD > 0 Then
                SQLQ = "SELECT * FROM HRENTWRK WHERE TE_EMPNBR = " & EmpNo&
                If rsWrk.State <> 0 Then rsWrk.Close
                rsWrk.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
                If rsWrk.EOF Then rsWrk.AddNew
                rsWrk("TE_COMPNO") = "001"
                rsWrk("TE_EMPNBR") = EmpNo&
                rsWrk("TE_EARNHRS") = xYTD
                rsWrk.Update
                rsWrk.Close
            End If
        
lblNextRec:
            
            snapEntitle.MoveNext
        Loop
        
        rsVacEnt.MoveNext
    Loop
    rsVacEnt.Close

End Sub

Public Property Get ChangeAction() As UpdateStateEnum
ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = Reports
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = False
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property
Public Property Get Updateble() As Boolean
Updateble = False
End Property
Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Call set_Buttons
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

