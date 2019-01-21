VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEmpJOBS 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Positions Lookup "
   ClientHeight    =   3900
   ClientLeft      =   1395
   ClientTop       =   1500
   ClientWidth     =   6765
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
   ScaleHeight     =   3900
   ScaleWidth      =   6765
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   285
      Left            =   4740
      TabIndex        =   5
      Tag             =   "Find specific record"
      Top             =   2820
      Width           =   960
   End
   Begin VB.TextBox txtFindDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1380
      TabIndex        =   4
      Tag             =   "00-Search Description"
      Top             =   2820
      Width           =   3165
   End
   Begin VB.TextBox txtFindKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   180
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "00-Search Code"
      Top             =   2820
      Width           =   960
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      DataField       =   "JB_DESCR"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1380
      MaxLength       =   25
      TabIndex        =   2
      Tag             =   "01-Description of Code"
      Top             =   2340
      Width           =   3165
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      DataField       =   "JB_CODE"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   180
      MaxLength       =   6
      TabIndex        =   1
      Tag             =   "01-Position's Code"
      Top             =   2340
      Width           =   975
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxjobsem.frx":0000
      Height          =   2085
      Left            =   180
      OleObjectBlob   =   "fxjobsem.frx":0014
      TabIndex        =   0
      Tag             =   "Position Listings"
      Top             =   120
      Width           =   6375
   End
   Begin Threed.SSOption optSort 
      Height          =   225
      Index           =   0
      Left            =   3810
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4350
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "Code"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Left            =   4680
      TabIndex        =   8
      Top             =   4335
      Visible         =   0   'False
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
   Begin Threed.SSOption optSort 
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4335
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "Group"
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
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   10
      Top             =   3240
      Width           =   6765
      _Version        =   65536
      _ExtentX        =   11933
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
         Height          =   375
         Left            =   4800
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
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
         Caption         =   "Ado1"
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
         Left            =   120
         TabIndex        =   11
         Tag             =   "Select this Department"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   900
         TabIndex        =   12
         Tag             =   "Close and exit this screen"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Tag             =   "Print Departmental Listing"
         Top             =   120
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   4995
         Top             =   120
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sort by"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3120
      TabIndex        =   6
      Top             =   4335
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmEmpJOBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Dim SQLQ As String

If Len(txtFindKey) > 0 Then
    SQLQ = "JB_CODE = '" & txtFindKey.Text & "'"
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
    SQLQ = "JB_DESCR >= '" & Trim(txtFindDesc.Text) & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
    Else
        txtFindDesc = ""
    End If
    txtFindDesc.SetFocus
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
glbJob = Data1.Recordset("JB_CODE")
glbJobDesc = Data1.Recordset("JB_DESCR")
Unload frmEmpJOBS


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

Screen.MousePointer = HOURGLASS


SQLQ = ""
SQLQ = SQLQ & "SELECT HR_JOB_HISTORY.JH_EMPNBR, HRJOB.JB_CODE, HRJOB.JB_DESCR"
SQLQ = SQLQ & " FROM HR_JOB_HISTORY INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE"
SQLQ = SQLQ & " WHERE (((HR_JOB_HISTORY.JH_EMPNBR)=" & CStr(glbLEE_ID) & ")); "

Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = SQLQ
Data1.Refresh
Screen.MousePointer = DEFAULT
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

Private Sub Form_Unload(Cancel As Integer)
Set frmEmpJOBS = Nothing
End Sub

Private Sub optSort_Click(Index As Integer, Value As Integer)
Dim SQLQ As String, srtSQL As String

Screen.MousePointer = HOURGLASS
txtFindDesc = ""
txtFindKey = ""
If optSort(0).Value = True Then
    srtSQL = "HRJOB.JB_CODE"
    txtFindKey.SetFocus
End If
If optSort(1).Value = True Then
    srtSQL = "HRJOB.JB_DESCR"
    txtFindDesc.SetFocus
End If
If optSort(2).Value = True Then
    srtSQL = "JB_GRPCD, JB_DESCR"
    txtFindDesc.SetFocus
End If
On Error GoTo Job_Err2

'Data1.DatabaseName = glbIHRDB
Data1.ConnectionString = glbAdoIHRDB

SQLQ = "SELECT JB_CODE, JB_DESCR, JB_GRPCD FROM HRJOB"
SQLQ = SQLQ & " ORDER BY " & srtSQL

Me.Data1.RecordSource = SQLQ

Me.Data1.Refresh
Me.vbxTrueGrid.Refresh
Screen.MousePointer = DEFAULT

Exit Sub

Job_Err2:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Error", "HRJobs", "Cancel")
Resume Next


End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
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

glbJob = Data1.Recordset("JB_CODE")
glbJobDesc = Data1.Recordset("JB_DESCR")
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
        
        SQLQ = SQLQ & "SELECT HR_JOB_HISTORY.JH_EMPNBR, HRJOB.JB_CODE, HRJOB.JB_DESCR"
        SQLQ = SQLQ & " FROM HR_JOB_HISTORY INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE"
        SQLQ = SQLQ & " WHERE (((HR_JOB_HISTORY.JH_EMPNBR)=" & CStr(glbLEE_ID) & ")); "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the enter key was struck
    KeyAscii = 0
    cmdClose.SetFocus
End If

End Sub




