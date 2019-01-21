VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmJOBS 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Positions Lookup"
   ClientHeight    =   5550
   ClientLeft      =   1455
   ClientTop       =   1770
   ClientWidth     =   7740
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
   ScaleHeight     =   5550
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkHideInactive 
      Caption         =   "Hide Inactive Positions"
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
      Left            =   240
      TabIndex        =   15
      Top             =   4560
      Value           =   1  'Checked
      Width           =   2355
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   7
      Top             =   4890
      Width           =   7740
      _Version        =   65536
      _ExtentX        =   13652
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
         Left            =   4320
         Top             =   480
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
         Left            =   105
         TabIndex        =   8
         Tag             =   "Select this Department"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   945
         TabIndex        =   9
         Tag             =   "Close and exit this screen"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   1785
         TabIndex        =   10
         Tag             =   "Print Departmental Listing"
         Top             =   135
         Width           =   735
      End
      Begin Threed.SSOption optSort 
         Height          =   255
         Index           =   0
         Left            =   3585
         TabIndex        =   11
         TabStop         =   0   'False
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
            Size            =   24.27
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
         Left            =   4560
         TabIndex        =   12
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
            Size            =   24.27
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   210
         Width           =   855
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Group"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24.27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   7320
         Top             =   240
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
      Begin Threed.SSOption optSort 
         Height          =   255
         Index           =   3
         Left            =   6480
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   210
         Width           =   855
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Band"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24.27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sort by"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2805
         TabIndex        =   14
         Top             =   210
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   285
      Left            =   6180
      TabIndex        =   6
      Tag             =   "Find specific record"
      Top             =   4080
      Width           =   1200
   End
   Begin VB.TextBox txtFindDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2100
      TabIndex        =   5
      Tag             =   "00-Search Description"
      Top             =   4095
      Width           =   3885
   End
   Begin VB.TextBox txtFindKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      MaxLength       =   25
      TabIndex        =   4
      Tag             =   "00-Search Code"
      Top             =   4095
      Width           =   1695
   End
   Begin VB.TextBox txtGroup 
      Appearance      =   0  'Flat
      DataField       =   "JB_GRPCD"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   6180
      MaxLength       =   4
      TabIndex        =   3
      Top             =   3540
      Width           =   1215
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      DataField       =   "JB_DESCR"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2100
      MaxLength       =   25
      TabIndex        =   2
      Tag             =   "01-Description of Code"
      Top             =   3540
      Width           =   3915
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      DataField       =   "JB_CODE"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      MaxLength       =   25
      TabIndex        =   1
      Tag             =   "01-Position's Code"
      Top             =   3540
      Width           =   1695
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxjobs.frx":0000
      Height          =   3195
      Left            =   120
      OleObjectBlob   =   "fxjobs.frx":0014
      TabIndex        =   0
      Tag             =   "Position Listings"
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmJOBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fRS As ADODB.Recordset
 
Private Sub chkHideInactive_Click()
    Dim SQLQ  As String
    
    SQLQ = "SELECT * FROM HRJOB WHERE 1=1 "
'    If Len(fglbseleEMPCode) > 0 Then
'        SQLQ = SQLQ & " AND JB_CODE IN ('" & Replace(fglbseleEMPCode, ",", "','") & "')"
'    End If
'    If glbLinamar Then
'        SQLQ = SQLQ & " AND JB_LOCGROUP IN (SELECT LOCGROUP FROM HR_DIVISION WHERE " & glbSeleDiv & ")"
'        If glbTransDiv <> "ALL" And glbTransDiv <> "" Then
'            SQLQ = SQLQ & " AND JB_LOCGROUP IN (SELECT LOCGROUP FROM HR_DIVISION WHERE DIV='" & glbTransDiv & "')"
'        End If
'        SQLQ = SQLQ & " AND JB_STATUS<>'INA'"
'    Else
        If chkHideInactive Then
            SQLQ = SQLQ & " AND JB_STATUS<>'INAC'"
            If glbOracle Then 'Ticket #16416
                SQLQ = SQLQ & " AND UPPER(SUBSTR(JB_DESCR,1,2)) <> 'Z '"
            ElseIf glbSQL Then
                SQLQ = SQLQ & " AND UPPER(LEFT(JB_DESCR, 2)) <> 'Z '"
            Else
                SQLQ = SQLQ & " AND UCASE(LEFT(JB_DESCR, 2)) <> 'Z '"
            End If
        End If
'    End If

    If glbOracle Then
        SQLQ = SQLQ & " ORDER BY UPPER(JB_DESCR)"
    Else
        SQLQ = SQLQ & " ORDER BY JB_DESCR"
    End If

    Data1.RecordSource = SQLQ
    Data1.Refresh
    Set fRS = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Dim SQLQ As String

If Len(txtFindKey) > 0 Then
    If optSort(3).Value Then 'Ticket #20479 Franks 07/11/2011
        SQLQ = "JB_BAND like '" & txtFindKey.Text & "%'"
    Else
        SQLQ = "JB_CODE like '" & txtFindKey.Text & "%'"
    End If
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

'Private Sub cmdSelect_Click()
'glbJob = Data1.Recordset("JB_CODE")
'glbJobDesc = Data1.Recordset("JB_DESCR")
'Unload frmJOBS
'End Sub

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
glbOnTop = "FRMJOBS"
Data1.ConnectionString = glbAdoIHRDB

SQLQ = SQLQ & "SELECT * FROM HRJOB WHERE 1 = 1"
If chkHideInactive Then
    SQLQ = SQLQ & " AND JB_STATUS<>'INAC'"
    If glbOracle Then 'Ticket #16416
        SQLQ = SQLQ & " AND UPPER(SUBSTR(JB_DESCR,1,2)) <> 'Z '"
    ElseIf glbSQL Then
        SQLQ = SQLQ & " AND UPPER(LEFT(JB_DESCR, 2)) <> 'Z '"
    Else
        SQLQ = SQLQ & " AND UCASE(LEFT(JB_DESCR, 2)) <> 'Z '"
    End If
End If
SQLQ = SQLQ & "ORDER BY JB_DESCR"

Data1.RecordSource = SQLQ '"SELECT * FROM HRJOB ORDER BY JB_DESCR"
Data1.Refresh
Set fRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Screen.MousePointer = DEFAULT
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    cmdSelect.Enabled = False
End If
txtCode.Enabled = False
txtDesc.Enabled = False
txtGroup.Enabled = False

If Not glbWFC Then 'Ticket #20479 Franks 07/11/2011, this field is for WFC only
    vbxTrueGrid.Columns(3).Visible = False
    optSort(3).Visible = False
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
    srtSQL = "HRJOB.JB_CODE"
    txtFindKey.SetFocus
End If
If optSort(1).Value = True Then
    srtSQL = "HRJOB.JB_DESCR"
    txtFindDesc.SetFocus
End If
If optSort(2).Value = True Then
    srtSQL = "JB_GRPCD, JB_DESCR"
    txtFindKey.SetFocus
End If
If glbWFC Then 'Ticket #20479 Franks 07/11/2011, this is for WFC only
    If optSort(3).Value = True Then
        srtSQL = "JB_BAND, JB_DESCR"
        txtFindKey.SetFocus
    End If
End If

On Error GoTo Job_Err2

'Data1.Databasename= glbIHRDB
Data1.ConnectionString = glbAdoIHRDB


SQLQ = "SELECT JB_CODE, JB_DESCR, JB_GRPCD, JB_STATUS, JB_BAND FROM HRJOB WHERE 1 = 1"
If chkHideInactive Then
    SQLQ = SQLQ & " AND JB_STATUS<>'INAC'"
    If glbOracle Then 'Ticket #16416
        SQLQ = SQLQ & " AND UPPER(SUBSTR(JB_DESCR,1,2)) <> 'Z '"
    ElseIf glbSQL Then
        SQLQ = SQLQ & " AND UPPER(LEFT(JB_DESCR, 2)) <> 'Z '"
    Else
        SQLQ = SQLQ & " AND UCASE(LEFT(JB_DESCR, 2)) <> 'Z '"
    End If
End If
SQLQ = SQLQ & " ORDER BY " & srtSQL

Data1.RecordSource = SQLQ
Data1.Refresh
Set fRS = Data1.Recordset.Clone
vbxTrueGrid.Refresh
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
If Not (Data1.Recordset.EOF Or Data1.Recordset.EOF) Then
    frmFind = True
    Call cmdSelect_Click
End If
Unload Me

End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    'added by Bryan 15/Nov/05 Ticket#9766
    fRS.Bookmark = Bookmark
    If fRS("JB_STATUS") = "INAC" Or UCase(Left(fRS("JB_DESCR"), 2)) = "Z " Then
        RowStyle.ForeColor = vbRed
    End If
End Sub

Private Sub vbxTrueGrid_GotFocus()

Call SetPanHelp(ActiveControl)
'txtFindDesc.SetFocus      'Sept 21, Laura
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT * FROM HRJOB WHERE 1 = 1"
        If chkHideInactive Then
            SQLQ = SQLQ & " AND JB_STATUS<>'INAC'"
            If glbOracle Then 'Ticket #16416
                SQLQ = SQLQ & " AND UPPER(SUBSTR(JB_DESCR,1,2)) <> 'Z '"
            ElseIf glbSQL Then
                SQLQ = SQLQ & " AND UPPER(LEFT(JB_DESCR, 2)) <> 'Z '"
            Else
                SQLQ = SQLQ & " AND UCASE(LEFT(JB_DESCR, 2)) <> 'Z '"
            End If
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
        Set fRS = Data1.Recordset.Clone
        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the enter key was struck
    KeyAscii = 0
    cmdClose.SetFocus
End If

End Sub

Private Sub cmdSelect_Click()
Dim x As Integer
'If vbxTrueGrid.MultiSelect = 2 Then
If vbxTrueGrid.MultiSelect = 2 And vbxTrueGrid.SelBookmarks.count > 0 Then
    If vbxTrueGrid.SelBookmarks.count > 1000 Then
        MsgBox vbxTrueGrid.SelBookmarks.count & " Positions are selected." + Chr(10) + " Please make that less than 1000 Positions"
        Exit Sub
    End If
    glbPos = ""
    For x = 0 To vbxTrueGrid.SelBookmarks.count - 1
        vbxTrueGrid.Bookmark = vbxTrueGrid.SelBookmarks(x)
        glbPos = glbPos & Data1.Recordset!JB_CODE & ","
    Next
Else
    glbPos = Data1.Recordset("JB_CODE")
    glbPosDesc = Data1.Recordset("JB_DESCR")
End If
Unload frmJOBS
End Sub
