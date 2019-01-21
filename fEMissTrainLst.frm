VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEMissTrainLst 
   Appearance      =   0  'Flat
   Caption         =   "Missing Training Records"
   ClientHeight    =   5700
   ClientLeft      =   -150
   ClientTop       =   765
   ClientWidth     =   8985
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
   ForeColor       =   &H00000000&
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5700
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   4920
      Width           =   1215
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fEMissTrainLst.frx":0000
      Height          =   4005
      Left            =   90
      OleObjectBlob   =   "fEMissTrainLst.frx":0014
      TabIndex        =   0
      Tag             =   "Training Listing"
      Top             =   750
      Width           =   8715
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6480
      Top             =   5040
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
      LockType        =   1
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
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   5400
      Width           =   8985
      _Version        =   65536
      _ExtentX        =   15849
      _ExtentY        =   529
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6120
         Top             =   90
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
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
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8985
      _Version        =   65536
      _ExtentX        =   15849
      _ExtentY        =   873
      _StockProps     =   15
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
      Begin VB.Label lblEEProdLine 
         AutoSize        =   -1  'True
         Caption         =   "Product Line"
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
         Left            =   6360
         TabIndex        =   6
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   155
         Width           =   1005
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
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
         Left            =   1440
         TabIndex        =   3
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Employee Name"
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
         Left            =   3100
         TabIndex        =   2
         Top             =   135
         Width           =   1740
      End
   End
   Begin VB.CommandButton cmdAddContEdu 
      Caption         =   "Add to Continuing Education "
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "TR_EMPNBR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEMissTrainLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim rsDATA As New ADODB.Recordset
Dim rsGrid As ADODB.Recordset
Dim Ctrl As Control
Dim xCourseCode, xJob
Dim xRenewDt

Private Function chkMissTrainRec()
    Dim oCode As String, OCodeD As String
    
    chkMissTrainRec = False
    
    On Error GoTo chkMissTrainRec_Err
       
    chkMissTrainRec = True

Exit Function

chkMissTrainRec_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkMissTrainRec", "HR_TRAIN", "validation")
    Call RollBack
End Function

Sub cmdCancel_Click()
    Dim X
    On Error GoTo Can_Err
    
    Call Display_Value
    
    fglbNew = False
    
    Call SET_UP_MODE
    'Call ST_UPD_MODE(True)  ' reset screen's attributes
    
   
    fglbNew = False
    Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_TRAIN", "Cancel")
    Call RollBack
End Sub

Private Sub cmdAddContEdu_Click()
    'Call procedure to add this training record into the Continuing Education screen
    Call Add_to_Continuing_Education
    If EERetrieve() = False Then Exit Sub
    'Call Display_Value
End Sub

Sub cmdClose_Click()
    'Initialise
    xCourseCode = ""
    xJob = ""

    Unload Me
    If glbOnTop = "FRMEMISSTRAINLST" Then glbOnTop = ""
End Sub

Sub cmdDelete_Click()
    Dim a As Integer, Msg As String, X
    
    If Data1.Recordset.BOF And Data1.Recordset.EOF Then
        MsgBox "Nothing to Delete"
        Exit Sub
    End If
    
    On Error GoTo Del_Err
    
    Msg = "Are You Sure You Want To Delete "
    Msg = Msg & "This Record?"
    
    a% = MsgBox(Msg, 36, "Confirm Delete")
    If a% <> 6 Then Exit Sub
    
    'Friesens - Ticket #16189
    'If glbtermopen Then
    '  gdbAdoIhr001X.BeginTrans
    '  rsDATA.Delete
    '  gdbAdoIhr001X.CommitTrans
    '  Data1.Refresh
    'Else
        
    gdbAdoIhr001.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001.CommitTrans
    Data1.Refresh
    
    Set rsGrid = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True
      
    'End If
    
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        Call Display_Value
    End If
    
    fglbNew = False
    
    Call SET_UP_MODE
    
    'Call ST_UPD_MODE(True)

Exit Sub

Del_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_TRAIN", "Delete")
    Call RollBack
End Sub

Sub cmdNew_Click()
    Dim SQLQ As String
    
    On Error GoTo AddN_Err
    
    fglbNew = True
    
    Call SET_UP_MODE
      
    Call Set_Control("B", Me)
    
    rsDATA.AddNew
    
    'Friesens - Ticket #16189
    'If glbtermopen Then lblEEID = glbTERM_ID Else
    
Exit Sub
    
AddN_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_TRAIN", "Add")
    Call RollBack
End Sub

Sub cmdOK_Click()
    Dim X
    On Error GoTo Add_Err
    
    If Not chkMissTrainRec() Then Exit Sub
        
    Call UpdUStats(Me) ' update user's stats (who did it and when)
    
    Call Set_Control("U", Me, rsDATA)
    
    'Friesens - Ticket #16189
    'If glbtermopen Then
    '    rsDATA!TERM_SEQ = glbTERM_Seq
    '    gdbAdoIhr001X.BeginTrans
    '    rsDATA.Update
    '    gdbAdoIhr001X.CommitTrans
    'Else
        gdbAdoIhr001.BeginTrans
        rsDATA.Update
        gdbAdoIhr001.CommitTrans
    'End If
    Data1.Refresh

    Set rsGrid = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True

    'Call ST_UPD_MODE(True)

    fglbNew = False
    
    Call SET_UP_MODE
    
    Me.vbxTrueGrid.SetFocus
    
Exit Sub

Add_Err:
    If Err = 3022 Then
        MsgBox "Duplicate record existed - not entered"
        Err = 0
        Resume Next
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_TRAIN", "Update")
    Call RollBack
End Sub

Sub cmdPrint_Click()
    Dim RHeading As String
    'Ticket #20447 - Jerry asked to change to Training Plan for everyone except Friesens and
    'Chatham-Kent but Chatham-Kent are not using 7.9
    If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        RHeading = lblEEName & "'s Training List"
    Else
        RHeading = lblEEName & "'s Training Plan"
    End If
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
    Me.vbxCrystal.Destination = 1
    Me.vbxCrystal.Action = 1
End Sub

Sub cmdView_Click()
    Dim RHeading As String

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    'Ticket #20447 - Jerry asked to change to Training Plan for everyone except Friesens and
    'Chatham-Kent but Chatham-Kent are not using 7.9
    If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        RHeading = lblEEName & "'s Training List"
    Else
        RHeading = lblEEName & "'s Training Plan"
    End If
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.Action = 1
End Sub

Function EERetrieve()
    Dim SQLQ As String
    
    Screen.MousePointer = HOURGLASS
    
    EERetrieve = False
    
    On Error GoTo EERError
    
    'Friesens - Ticket #16189
    'If glbtermopen Then         'Lucy July 5, 2000
    '    SQLQ = "Select * from Term_TRADE"
    '    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    '    SQLQ = SQLQ & " ORDER BY TD_CODE"
    'Else
    If glbSQL Then
        SQLQ = "SELECT HR_TRAIN.*, (CASE TR_POS_TYPE WHEN 'C' THEN 'Current' WHEN 'T' THEN 'Temporary' WHEN 'P' THEN 'Previous' END) AS POSTYPE FROM HR_TRAIN"
        SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
        'SQLQ = SQLQ & " AND TR_RENEW <= " & Date_SQL(Format(Now, "mm/dd/yyyy"))
        SQLQ = SQLQ & " AND TR_CRSCODE NOT IN (SELECT ES_CRSCODE FROM HREDSEM WHERE ES_EMPNBR = " & glbLEE_ID & " AND ES_RENEW IS NULL AND ES_DATCOMP IS NULL AND ES_CRSCODE = TR_CRSCODE AND ES_JOB = TR_JOB)"
        SQLQ = SQLQ & " AND TR_JOB NOT IN (SELECT es_job FROM HREDSEM WHERE ES_EMPNBR = " & glbLEE_ID & " AND ES_RENEW IS NULL AND ES_DATCOMP IS NULL AND ES_CRSCODE = TR_CRSCODE AND (ES_JOB = TR_JOB OR ES_JOB IS NULL))"
        SQLQ = SQLQ & " ORDER BY TR_RENEW"
    ElseIf glbOracle Then
        SQLQ = "SELECT HR_TRAIN.* FROM HR_TRAIN"
        SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
        'SQLQ = SQLQ & " AND TR_RENEW <= " & Date_SQL(Format(Now, "mm/dd/yyyy"))
        SQLQ = SQLQ & " AND TR_CRSCODE NOT IN (SELECT ES_CRSCODE FROM HREDSEM WHERE ES_EMPNBR = " & glbLEE_ID & " AND ES_RENEW IS NULL AND ES_DATCOMP IS NULL AND ES_CRSCODE = TR_CRSCODE AND ES_JOB = TR_JOB)"
        SQLQ = SQLQ & " AND TR_JOB NOT IN (SELECT es_job FROM HREDSEM WHERE ES_EMPNBR = " & glbLEE_ID & " AND ES_RENEW IS NULL AND ES_DATCOMP IS NULL AND ES_CRSCODE = TR_CRSCODE AND (ES_JOB = TR_JOB OR ES_JOB IS NULL))"
        SQLQ = SQLQ & " ORDER BY TR_RENEW"
    End If
    'End If
    
    Data1.RecordSource = SQLQ
    Data1.Refresh
        
    Set rsGrid = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True
        
    EERetrieve = True
    Screen.MousePointer = DEFAULT
    
Exit Function
    
EERError:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TrainRetrieve", "HR_TRAIN", "SELECT")
    Call RollBack
End Function

Private Sub Form_Activate()
    glbOnTop = "FRMEMISSTRAINLST"
    Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEMISSTRAINLST"
End Sub

Private Sub Form_Load()
    Dim Answer, DefVal, Msg, Title
    Dim RFound As Integer
    
    glbOnTop = "FRMEMISSTRAINLST"
    
    'Friesens - Ticket #16189
    'If glbtermopen Then
    '    Data1.ConnectionString = glbAdoIHRAUDIT
    'Else
        Data1.ConnectionString = glbAdoIHRDB
    'End If
    
    Screen.MousePointer = DEFAULT
    
    'Initialise
    xCourseCode = ""
    xJob = ""
    
    'If Not glbtermopen Then
    '    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    '    If glbLEE_ID = 0 Then Unload Me: Exit Sub
    'Friesens - Ticket #16189
    'Else
    '    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    '    If glbTERM_ID = 0 Then Unload Me: Exit Sub
    'End If
    
    If EERetrieve() = False Then
        'MsgBox "Sorry, Employee can not be found"
        'Friesens - Ticket #16189
        'If glbtermopen Then frmTERMEMPL.Show 1 Else
        'frmEEFIND.Show 1
        Exit Sub
    'Else
    '    Me.Show
        'Friesens - Ticket #16189
        'If glbtermopen Then lblEEID = glbTERM_ID Else
    '    lblEEID = glbLEE_ID
    End If
    
    If Len(glbLEE_SName) < 1 Then Exit Sub
    
    Screen.MousePointer = HOURGLASS
        
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
        Me.Caption = "Missing Training Records - " & Left$(glbLEE_SName, 5)
        Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    
    lblEENum.Caption = ShowEmpnbr(lblEEID)
    
    Call Display_Value
    Call ST_UPD_MODE(True)             '
    Call INI_Controls(Me)
    
    vbxTrueGrid.Columns(4).Visible = False
    
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    Screen.MousePointer = DEFAULT
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
    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Call NextForm
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
    
    fUPMode = TF    ' update mode
    
    'Friesens - Ticket #16189
'    chkCompPaid.Enabled = TF
'    medDuesPaid.Enabled = TF
'    clpCode(1).Enabled = TF
'    dlpDate(0).Enabled = TF
'    dlpDate(1).Enabled = TF
    
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    End If
    
End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    rsGrid.Bookmark = Bookmark
    If Not IsNull(rsGrid("TR_COURSE_TAKEN")) And rsGrid("TR_COURSE_TAKEN") <> "" Then
        If CVDate(rsGrid("TR_RENEW")) < CVDate(Format(Now, "Short Date")) And CVDate(rsGrid("TR_COURSE_TAKEN")) < CVDate(Format(Now, "Short Date")) Then
            RowStyle.ForeColor = vbRed
        End If
    ElseIf CVDate(rsGrid("TR_RENEW")) < CVDate(Format(Now, "Short Date")) And (rsGrid("TR_COURSE_TAKEN") = "" Or IsNull(rsGrid("TR_COURSE_TAKEN"))) Then
        RowStyle.ForeColor = vbRed
    End If
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
    Dim SQLQ As String
       
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    'Friesens - Ticket #16189
    'If glbtermopen Then         'Lucy July 5, 2000
    '    SQLQ = "Select * from Term_TRADE"
    '    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    'Else
    If glbSQL Then
        SQLQ = "Select HR_TRAIN.*,(CASE TR_POS_TYPE WHEN 'C' THEN 'Current' WHEN 'T' THEN 'Temporary' WHEN 'P' THEN 'Previous' END) AS POSTYPE FROM HR_TRAIN"
        SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
    ElseIf glbOracle Then
        SQLQ = "Select HR_TRAIN.* FROM HR_TRAIN"
        SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
    End If
   ' End If
    If glbOracle And ColIndex <> 5 Then
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    ElseIf glbSQL Then
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    End If

    Data1.RecordSource = SQLQ
    Data1.Refresh
    
    Set rsGrid = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True
    
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then ' if the tab key was struck
        KeyAscii = 0
    End If
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim tdcode$
    Dim SQLQ As String
    
    On Error GoTo Tab1_Err
    'If Not Fnd_Match_Data2() Then Exit Sub 'MsgBox "No Records Found"
    
    ' ' set description for code
    'If Data1.Recordset.RecordCount <> 0 Then
    '    If Not IsNull(Data2.Recordset("TD_RENEWDT")) Then
    '        txtDate(1) = Data2.Recordset("TD_RENEWDT")
    '    Else
    '        txtDate(1) = ""
    '    End If
    'End If
        
    vbxTrueGrid.Columns(4).Visible = False
    If glbOracle Then
        vbxTrueGrid.Columns(5).Visible = False
    End If
        
    Call Display_Value
        
Exit Sub
    
Tab1_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_TRAIN", "Add")
    Call RollBack
End Sub

Private Function RollBack()
On Error GoTo rr
    Screen.MousePointer = DEFAULT
    
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
rr:
End Function

Sub Display_Value()
    Dim SQLQ
    
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        'Friesens - Ticket #16189
        'If glbtermopen Then
        '    rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        'Else
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        'End If
        Call SET_UP_MODE
        
        Exit Sub
    End If
          
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    'Friesens - Ticket #16189
    'If glbtermopen Then
    '    SQLQ = "Select * from Term_TRADE"
    '    SQLQ = SQLQ & " WHERE TERM_SEQ = " & Data1.Recordset!TERM_SEQ
    '    SQLQ = SQLQ & " ORDER BY TD_CODE"
    '    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    'Else
        SQLQ = "SELECT HR_TRAIN.*,(CASE TR_POS_TYPE WHEN 'C' THEN 'Current' WHEN 'T' THEN 'Temporary' WHEN 'P' THEN 'Previous' END) AS POSTYPE FROM HR_TRAIN"
        SQLQ = SQLQ & " WHERE TR_ID = " & Data1.Recordset!TR_ID
        'SQLQ = SQLQ & " AND TR_RENEW <= " & Date_SQL(Format(Now, "mm/dd/yyyy"))
        SQLQ = SQLQ & " ORDER BY TR_RENEW"
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    'End If

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    
    Call Set_Control("R", Me, rsDATA)
    Call SET_UP_MODE
    
    'Get the record values:
    xCourseCode = Data1.Recordset("TR_CRSCODE")
    xJob = Data1.Recordset("TR_JOB")
    xRenewDt = Data1.Recordset("TR_RENEW")
        
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
    RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
    UpdateRight = gSec_Upd_Training_List
End Property

Public Property Get Addable() As Boolean
    Addable = True
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
        TF = True
    End If
    
    Call set_Buttons(UpdateState)
    If Not UpdateRight Then TF = False
    Call ST_UPD_MODE(TF)
End Sub

Private Sub lblEEID_Change()
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
        frmEMissTrainLst.Caption = "Missing Training Records - " & Left$(glbLEE_SName, 5)
        frmEMissTrainLst.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
    
    lblEENum = ShowEmpnbr(lblEEID)
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = glbLEE_ProdLine
    Else
        lblEEProdLine = ""
    End If
End Sub

Private Sub Add_to_Continuing_Education()
    Dim rsContEdu As New ADODB.Recordset
    Dim rsCourseMst As New ADODB.Recordset
    Dim SQLQ As String
    Dim xRecAdd As Boolean
    
    'Initialise
    xRecAdd = False
    
    SQLQ = "SELECT * FROM HREDSEM "
    SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
    SQLQ = SQLQ & " AND ES_JOB = '" & xJob & "'"
    SQLQ = SQLQ & " AND (ES_RENEW IS NULL OR ES_RENEW = '')"
    SQLQ = SQLQ & " AND (ES_DATCOMP IS NULL OR ES_DATCOMP = '')"

    'SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(xRenewDt)
    'SQLQ = SQLQ & " AND ((ES_RENEW IS NULL OR ES_RENEW = '') OR (ES_RENEW = " & Date_SQL(Data1.Recordset("TR_RENEW")) & "))"
    'SQLQ = SQLQ & " ORDER BY TR_RENEW DESC"
    rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If rsContEdu.EOF Then
        'Add the Training record as a new record on the Continuing Education screen
        'leave, Start Date, Completed Date and Renewal Date blank
        rsContEdu.AddNew
        rsContEdu("ES_COMPNO") = "001"
        rsContEdu("ES_EMPNBR") = glbLEE_ID
        rsContEdu("ES_CRSCODE") = xCourseCode
        rsContEdu("ES_COURSE") = GetTABLDesc("ESCD", xCourseCode)
        rsContEdu("ES_JOB") = xJob

        'Retrieve rest of the data from Course Code Master screen
        SQLQ = "SELECT * FROM HR_COURSECODE_MASTER"
        SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & xCourseCode & "'"
        rsCourseMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsCourseMst.EOF Then
            rsContEdu("ES_CTYPE") = rsCourseMst("ES_CTYPE")
            rsContEdu("ES_COORDINATED") = rsCourseMst("ES_COORDINATED")
            rsContEdu("ES_COMPANYNAME") = rsCourseMst("ES_COMPANYNAME")
            rsContEdu("ES_TRAINNER") = rsCourseMst("ES_TRAINNER")
            rsContEdu("ES_HOURS") = rsCourseMst("ES_HOURS")
            rsContEdu("ES_TBEMP") = rsCourseMst("ES_TBEMP")
            rsContEdu("ES_EMPCUR") = rsCourseMst("ES_EMPCUR")
            rsContEdu("ES_OTHER") = rsCourseMst("ES_OTHER")
            rsContEdu("ES_OTCUR") = rsCourseMst("ES_OTCUR")
            rsContEdu("ES_TBCO") = rsCourseMst("ES_TBCO")
            rsContEdu("ES_EMPLOYCUR") = rsCourseMst("ES_EMPLOYCUR")
            rsContEdu("ES_ACCOM") = rsCourseMst("ES_ACCOM")
            rsContEdu("ES_ACOMCUR") = rsCourseMst("ES_ACOMCUR")
            rsContEdu("ES_LEARNING") = rsCourseMst("ES_LEARNING")
            rsContEdu("ES_LEARNINGCUR") = rsCourseMst("ES_LEARNINGCUR")
            rsContEdu("ES_TOTCUR") = rsCourseMst("ES_TOTCUR")
        End If
        rsCourseMst.Close
        Set rsCourseMst = Nothing
        
        rsContEdu("ES_LDATE") = Format(Now, "SHORT DATE")
        rsContEdu("ES_LTIME") = Time$
        rsContEdu("ES_LUSER") = glbUserID
        rsContEdu.Update
        
        xRecAdd = True
    End If
    rsContEdu.Close
    Set rsContEdu = Nothing
    
    If xRecAdd Then
        MsgBox "Missing training record added successfully. "
    End If
End Sub
