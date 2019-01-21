VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMJobMaster 
   Caption         =   "Job Lookup"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   ControlBox      =   0   'False
   LinkTopic       =   "Job Lookup"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkHideInactive 
      Caption         =   "Hide Inactive Positions"
      Height          =   315
      Left            =   7560
      TabIndex        =   15
      Top             =   4080
      Value           =   1  'Checked
      Width           =   2355
   End
   Begin VB.TextBox txtFindKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   3
      Tag             =   "00-Search Division"
      Top             =   4440
      Width           =   1560
   End
   Begin VB.TextBox txtFindDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      MaxLength       =   100
      TabIndex        =   4
      Tag             =   "00-Search Description"
      Top             =   4440
      Width           =   3525
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Tag             =   "Find specific record"
      Top             =   4395
      Width           =   720
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JB_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   8760
      MaxLength       =   25
      TabIndex        =   12
      Text            =   "LUser"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JB_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   8400
      MaxLength       =   25
      TabIndex        =   11
      Text            =   "LTime"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JB_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   8520
      MaxLength       =   25
      TabIndex        =   10
      Text            =   "Ldate"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox txtJobDescr 
      Appearance      =   0  'Flat
      DataField       =   "JB_JOBDESCR"
      Height          =   285
      Left            =   3000
      MaxLength       =   100
      TabIndex        =   2
      Tag             =   "01-Position Description"
      Top             =   4080
      Width           =   3495
   End
   Begin VB.TextBox txtJob 
      Appearance      =   0  'Flat
      DataField       =   "JB_JOBCODE"
      Height          =   285
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   1
      Tag             =   "01-Job Code (Unique)"
      Top             =   4080
      Width           =   1545
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmMJobMaster.frx":0000
      Height          =   3855
      Left            =   0
      OleObjectBlob   =   "frmMJobMaster.frx":0014
      TabIndex        =   0
      Top             =   120
      Width           =   9795
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   13
      Top             =   4980
      Width           =   9900
      _Version        =   65536
      _ExtentX        =   17462
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
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Tag             =   "Print Division Listing"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   855
         TabIndex        =   7
         Tag             =   "Close and exit this screen"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
         Height          =   375
         Left            =   15
         TabIndex        =   6
         Tag             =   "Select this Division"
         Top             =   105
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   3600
         Top             =   0
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7200
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Label lblDivSearch 
      Caption         =   "Search Job"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblJob 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Job Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   4110
      Width           =   675
   End
End
Attribute VB_Name = "frmMJobMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRSOld As String, glbEmptyNew  As Integer
Dim fglbNewRec% ' new record
Dim RSDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim Ctrl As Control 'Sam add July 2002 * Remove ADO
Dim FRS As ADODB.Recordset

Private Sub chkHideInactive_Click()
Dim SQLQ

SQLQ = "SELECT * FROM HRJOBMASTER WHERE (1=1) "
If chkHideInactive Then
    SQLQ = SQLQ & " AND JB_STATUS<>'INAC'"
    SQLQ = SQLQ & " AND UPPER(LEFT(JB_JOBDESCR, 2)) <> 'Z ' "
End If
SQLQ = SQLQ & "ORDER BY JB_JOBDESCR"
Data1.RecordSource = SQLQ ' "SELECT * FROM HRJOBMASTER ORDER BY JB_JOBDESCR "
Data1.Refresh

End Sub

Private Sub cmdClose_Click()
    'glbDiv = ""
    'glbDivDesc = ""
    fglbNewRec% = False
    
    Unload Me
End Sub


Private Sub cmdFind_Click()
Dim SQLQ As String

If Len(txtFindKey) > 0 Then
    'SQLQ = "JB_JOBCODE = '" & txtFindKey.Text & "'"
    SQLQ = "JB_JOBCODE like '" & txtFindKey.Text & "%'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
        
        Set FRS = Data1.Recordset.Clone
        vbxTrueGrid.FetchRowStyle = True
        
    Else
        txtFindKey = ""
    End If
    Exit Sub
End If

If Len(txtFindDesc) > 0 Then
    SQLQ = "JB_JOBDESCR >= '" & txtFindDesc.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
    
        Set FRS = Data1.Recordset.Clone
        vbxTrueGrid.FetchRowStyle = True
    
    Else
        txtFindDesc = ""
    End If
    Exit Sub
End If
End Sub

Private Sub cmdPrint_Click()
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

Me.vbxCrystal.ReportTitle = "All Table Codes"
Me.vbxCrystal.BoundReportHeading = Me.Caption
Me.vbxCrystal.WindowTitle = Me.Caption & " Report"
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdSelect_Click()
glbJobMaster = Data1.Recordset("JB_JOBCODE")
glbJobMasterDesc = Data1.Recordset("JB_JOBDESCR")
Unload Me

End Sub

Private Sub Form_Activate()
Dim SQLQ

SQLQ = "SELECT * FROM HRJOBMASTER WHERE (1=1) "
If chkHideInactive Then
    SQLQ = SQLQ & " AND JB_STATUS<>'INAC'"
    SQLQ = SQLQ & " AND UPPER(LEFT(JB_JOBDESCR, 2)) <> 'Z ' "
End If
SQLQ = SQLQ & "ORDER BY JB_JOBDESCR"
Data1.RecordSource = SQLQ ' "SELECT * FROM HRJOBMASTER ORDER BY JB_JOBDESCR "
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub Form_Load()
Dim SQLQ, I, ctylist, X

'glbOnTop = "FRMDIVISIONS"

Data1.ConnectionString = glbAdoIHRDB

'SQLQ = "SELECT * FROM HRJOBMASTER ORDER BY JB_JOBDESCR "
SQLQ = "SELECT * FROM HRJOBMASTER WHERE (1=1) "
If chkHideInactive Then
    SQLQ = SQLQ & " AND JB_STATUS<>'INAC'"
    SQLQ = SQLQ & " AND UPPER(LEFT(JB_JOBDESCR, 2)) <> 'Z ' "
End If
SQLQ = SQLQ & "ORDER BY JB_JOBDESCR"

Data1.RecordSource = SQLQ
Data1.LockType = adLockReadOnly
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Screen.MousePointer = HOURGLASS

'Me.vbxTrueGrid.Refresh

Screen.MousePointer = DEFAULT

Call modSTUPD(False)


Call WFCScreenSetup

Call INI_Controls(Me)

End Sub

Private Sub WFCScreenSetup() 'Ticket #25911 Franks 09/30/2014

    vbxTrueGrid.Columns(2).Caption = lStr("Job Group")
    'vbxTrueGrid.Columns(3).Caption = lStr("Job Level")
    vbxTrueGrid.Columns(3).Caption = lStr("Job Status")
    vbxTrueGrid.Columns(4).Caption = lStr("Job User Defined 1")
    vbxTrueGrid.Columns(5).Caption = lStr("Job User Defined 2")
    
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


cmdFind.Enabled = FT            '
txtJob.Enabled = TF
txtJobDescr.Enabled = TF


txtFindDesc.Enabled = FT        '
txtFindKey.Enabled = FT         '
cmdClose.Enabled = FT           '
cmdSelect.Enabled = FT          '
cmdPrint.Enabled = FT           '
        
If glbJobMasterInhSel% Then
    cmdSelect.Enabled = False
End If
End Sub

Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        Exit Sub
    End If
  
    SQLQ = "select * from HRJOBMASTER WHERE JB_JOBCODE = '" & Data1.Recordset!JB_JOBCODE & "' " '& " order by Division_Name"
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If RSDATA.EOF Or RSDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, RSDATA)
    
End Sub

Private Sub txtFindDesc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub txtFindKey_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub vbxTrueGrid_DblClick()
If cmdSelect.Enabled Then
    If Not Me.vbxTrueGrid.EditActive Then
        glbJobMaster = Data1.Recordset("JB_JOBCODE")
        glbJobMasterDesc = Data1.Recordset("JB_JOBDESCR")
        Unload Me
    Else
        MsgBox "Save/cancel changes first"
    End If
End If
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
    Dim SQLQ As String
           
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If

    'SQLQ = "SELECT * FROM HRJOBMASTER "
    SQLQ = "SELECT * FROM HRJOBMASTER WHERE (1=1) "
    If chkHideInactive Then
        SQLQ = SQLQ & " AND JB_STATUS<>'INAC'"
        SQLQ = SQLQ & " AND UPPER(LEFT(JB_JOBDESCR, 2)) <> 'Z ' "
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

    Data1.RecordSource = SQLQ
    Data1.Refresh

    Set FRS = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Display_Value
End Sub

Private Function chkJobMaster()
Dim Div As String, SQLQ As String, Msg$
Dim snapDivs As New ADODB.Recordset
Dim X
chkJobMaster = False
On Error GoTo chkJobMaster_Err

If Len(txtJob) < 1 Then
    MsgBox ("Job Code is a required field")
    txtJob.SetFocus
    Exit Function
End If

If Len(txtJobDescr) < 1 Then
    MsgBox lStr("Job Description is a required field")
    txtJobDescr.SetFocus
    Exit Function
End If

If fglbNewRec% Then
    SQLQ = "SELECT * FROM HRJOBMASTER "
    SQLQ = SQLQ & "WHERE JB_JOBCODE = '" & txtJob.Text & "'"
    
    If snapDivs.State <> 0 Then snapDivs.Close
    snapDivs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If snapDivs.BOF And snapDivs.EOF Then
        snapDivs.Close
    Else
        Msg$ = ("This Job Code already exists")
        MsgBox Msg$
        snapDivs.Close
        Exit Function
    End If
End If



chkJobMaster = True

Exit Function

chkJobMaster_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkJobMaster", "HRJOBMASTER", "Cancel")
Resume Next

End Function


