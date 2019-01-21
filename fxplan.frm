VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmOPlan 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan Number Look up "
   ClientHeight    =   3915
   ClientLeft      =   1575
   ClientTop       =   1590
   ClientWidth     =   6630
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3915
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   6
      Top             =   3255
      Width           =   6630
      _Version        =   65536
      _ExtentX        =   11695
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
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   5280
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
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
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
         Height          =   375
         Left            =   90
         TabIndex        =   7
         Tag             =   "Select this Department"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   870
         TabIndex        =   8
         Tag             =   "Close and exit this screen"
         Top             =   150
         Width           =   795
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   1710
         TabIndex        =   9
         Tag             =   "Print Departmental Listing"
         Top             =   150
         Width           =   735
      End
      Begin Threed.SSOption optSort 
         Height          =   255
         Index           =   0
         Left            =   3570
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "Sort by Code"
         Top             =   210
         Width           =   855
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Code"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optSort 
         Height          =   255
         Index           =   1
         Left            =   4530
         TabIndex        =   11
         Tag             =   "Sort by Description"
         Top             =   210
         Width           =   855
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Descr."
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sort by"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2730
         TabIndex        =   12
         Top             =   210
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   285
      Left            =   5145
      TabIndex        =   5
      Tag             =   "Find specific record"
      Top             =   2775
      Width           =   960
   End
   Begin VB.TextBox txtFindDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1350
      MaxLength       =   30
      TabIndex        =   4
      Tag             =   "00-Search Plan Description"
      Top             =   2775
      Width           =   3165
   End
   Begin VB.TextBox txtFindKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   210
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "00-Search Plan Number"
      Top             =   2775
      Width           =   960
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      DataField       =   "PP_DESC"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1365
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "01-Plan Description"
      Top             =   2340
      Width           =   3135
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      DataField       =   "PP_PLAN"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   210
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "01-Plan Number"
      Top             =   2340
      Width           =   975
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxplan.frx":0000
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "fxplan.frx":0014
      TabIndex        =   0
      Tag             =   "Plan Number Listings"
      Top             =   60
      Width           =   6375
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6120
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowWidth     =   480
      WindowTitle     =   "Department Codes"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   2
      ReportSource    =   1
      WindowControls  =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "frmOPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_GotFocus()
Call SetPanHelp(ActiveControl) '19Aug99 js
End Sub

Private Sub cmdFind_Click()
Dim SQLQ As String

If Len(txtFindKey) > 0 Then
    SQLQ = "PP_PLAN = '" & txtFindKey.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
    Else
        txtFindKey = ""
    End If
    txtFindKey.SetFocus
    Exit Sub
End If

If Len(txtFindDesc) > 0 Then
    SQLQ = "PP_DESC >= '" & Trim(txtFindDesc.Text) & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    txtFindDesc.SetFocus
    If Data1.Recordset.EOF Then
        Data1.Refresh
    Else
        txtFindDesc = ""
    End If
    Exit Sub
End If

End Sub

Private Sub cmdFind_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPrint_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = "Positions Listing"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdSelect_Click()
glbPlan = Data1.Recordset("PP_PLAN")
glbPlanDesc = Data1.Recordset("PP_DESC")
glbSurvDate = Data1.Recordset("PP_SURVEYD")
If IsNull(Data1.Recordset("PP_DUEDATE")) Then
   glbDueDate = ""
Else
  glbDueDate = Data1.Recordset("PP_DUEDATE")
End If

Unload frmOPlan

End Sub

Private Sub cmdSelect_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRDEPTS", "SELECT")


End Sub



Private Sub Form_Load()
Dim SQLQ As String

On Error GoTo Job_Err


SQLQ = "SELECT * FROM HRPARCOP ORDER BY PP_DESC"
Data1.ConnectionString = glbAdoIHRDB


Me.Data1.RecordSource = SQLQ

Screen.MousePointer = HOURGLASS
Me.Data1.Refresh
Me.vbxTrueGrid.Refresh
Screen.MousePointer = DEFAULT
'Exit Sub
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    cmdSelect.Enabled = False
End If

Exit Sub
Job_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Error", "HRJobs", "Cancel")
Resume Next


End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub optSort_Click(Index As Integer, Value As Integer)
Dim SQLQ As String, srtSQL As String

Screen.MousePointer = HOURGLASS
txtFindDesc = ""
txtFindKey = ""
If optSort(0).Value = True Then
    srtSQL = "HRPARCOP.PP_PLAN"
    txtFindKey.SetFocus
End If
If optSort(1).Value = True Then
    srtSQL = "HRPARCOP.PP_DESC"
    txtFindDesc.SetFocus
End If
On Error GoTo Plan_Err2

'Data1.DatabaseName = glbIHRDB
Data1.ConnectionString = glbAdoIHRDB

SQLQ = "SELECT PP_PLAN, PP_DESC, PP_SURVEYD, PP_DUEDATE FROM HRPARCOP"
SQLQ = SQLQ & " ORDER BY " & srtSQL

Me.Data1.RecordSource = SQLQ

Me.Data1.Refresh
Me.vbxTrueGrid.Refresh
Screen.MousePointer = DEFAULT

Exit Sub

Plan_Err2:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Plan Error", "HRPARCOP", "Cancel")
Resume Next


End Sub

Private Sub optSort_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl) '19Aug99 js
End Sub

Private Sub txtCode_GotFocus()
Call SetPanHelp(ActiveControl) '19Aug99 js
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub txtDesc_GotFocus()
Call SetPanHelp(ActiveControl) '19Aug99 js
End Sub

Private Sub txtFindDesc_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub txtFindKey_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub vbxTrueGrid_DblClick()

glbPlan = Data1.Recordset("PP_PLAN")
glbPlanDesc = Data1.Recordset("PP_DESC")
glbSurvDate = Data1.Recordset("PP_SURVEYD")
If IsNull(Data1.Recordset("PP_DUEDATE")) Then
    glbDueDate = ""
Else
    glbDueDate = Data1.Recordset("PP_DUEDATE")
End If

Unload Me

End Sub

Private Sub vbxTrueGrid_GotFocus()
Call SetPanHelp(ActiveControl)
txtFindDesc.SetFocus      'Sept 21, Laura
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT * FROM HRPARCOP "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the enter key was struck
    KeyAscii = 0
    cmdClose.SetFocus
End If

End Sub




