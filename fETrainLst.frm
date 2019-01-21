VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmETRAINLST 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Training List"
   ClientHeight    =   8160
   ClientLeft      =   -150
   ClientTop       =   765
   ClientWidth     =   10515
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
   ScaleHeight     =   8160
   ScaleWidth      =   10515
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fETrainLst.frx":0000
      Height          =   2085
      Left            =   120
      OleObjectBlob   =   "fETrainLst.frx":0014
      TabIndex        =   0
      Tag             =   "Training Listing"
      Top             =   750
      Width           =   10275
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6240
      Top             =   7560
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
      Height          =   660
      Left            =   0
      TabIndex        =   12
      Top             =   7500
      Width           =   10515
      _Version        =   65536
      _ExtentX        =   18547
      _ExtentY        =   1164
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
      Begin VB.CommandButton cmdReset 
         Appearance      =   0  'Flat
         Caption         =   "&Reset the Training List for this Employee"
         Height          =   525
         Left            =   240
         TabIndex        =   18
         Tag             =   "Reset the Training List for this employee"
         Top             =   0
         Visible         =   0   'False
         Width           =   2625
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   9810
         Top             =   210
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TR_LDATE"
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
      Left            =   6120
      MaxLength       =   25
      TabIndex        =   3
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TR_LTIME"
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
      Left            =   6960
      MaxLength       =   25
      TabIndex        =   4
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TR_LUSER"
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
      Left            =   7800
      MaxLength       =   25
      TabIndex        =   5
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10515
      _Version        =   65536
      _ExtentX        =   18547
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
         TabIndex        =   13
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   135
         Width           =   1740
      End
   End
   Begin INFOHR_Controls.DateLookup dlpRenewal 
      DataField       =   "TR_RENEW"
      Height          =   285
      Left            =   1860
      TabIndex        =   2
      Tag             =   "40-Date when course is to be renewed"
      Top             =   4050
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "TR_CRSCODE"
      Height          =   285
      Index           =   0
      Left            =   1860
      TabIndex        =   1
      Tag             =   "00-Course Code"
      Top             =   3600
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESCD"
      MaxLength       =   8
      Enabled         =   0   'False
   End
   Begin VB.Label lblCourseTakenDt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Taken"
      DataField       =   "TR_COURSE_TAKEN"
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
      Left            =   4200
      TabIndex        =   22
      Top             =   6480
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblPosType 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Pos"
      DataField       =   "TR_POS_TYPE"
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
      Left            =   3120
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblStartDt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      DataField       =   "TR_SDATE"
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
      Left            =   3960
      TabIndex        =   20
      Top             =   6120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblPosCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pos Code"
      DataField       =   "TR_JOB"
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
      Left            =   3120
      TabIndex        =   19
      Top             =   6120
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblPosDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "PosDesc"
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
      Left            =   2160
      TabIndex        =   17
      Top             =   3180
      Width           =   645
   End
   Begin VB.Label LabelPos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position / Start Date :"
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
      Left            =   240
      TabIndex        =   16
      Top             =   3180
      Width           =   1530
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Renewal Date"
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
      Index           =   18
      Left            =   240
      TabIndex        =   15
      Top             =   4095
      Width           =   1365
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Code"
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
      Index           =   21
      Left            =   240
      TabIndex        =   14
      Top             =   3645
      Width           =   1200
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
      Left            =   5520
      TabIndex        =   10
      Top             =   6120
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "TR_COMPNO"
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
      Left            =   4920
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmETRAINLST"
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
Dim oRenewalDt

Private Function chkETrainList()
    Dim oCode As String, OCodeD As String
    
    chkETrainList = False
    
    On Error GoTo chkETrainList_Err
    
    If fglbNew Then
        If Len(Trim(clpCode(0).Text)) = 0 Then
            MsgBox "Course code is a required field"
            clpCode(0).SetFocus
            Exit Function
        End If
        If Not clpCode(0).ListChecker Then
            clpCode(0).SetFocus
            Exit Function
        End If
    End If
    
    If Len(Trim(dlpRenewal.Text)) = 0 Then
        MsgBox "Renewal Date cannot be blank"
        dlpRenewal.SetFocus
        Exit Function
    ElseIf Not IsDate(dlpRenewal.Text) Then
        MsgBox "Invalid Renewal Date"
        dlpRenewal.SetFocus
        Exit Function
    End If
    
    'Friesens - Ticket #16189
'    If Len(clpCode(1).Text) < 1 Then
'        MsgBox "Association code is a required field"
'        clpCode(1).SetFocus
'        Exit Function
'    End If
'
'    If clpCode(1).Caption = "Unassigned" Then
'        MsgBox "Association code must be valid"
'        clpCode(1).SetFocus
'        Exit Function
'    End If
'
'    If chkCompPaid.Value = 1 Then
'        txtCompPaid.Text = "Y"
'    Else
'        txtCompPaid.Text = "N"
'    End If
'
'    If Len(dlpDate(0).Text) < 1 Then
'        MsgBox "Starting Date is Required Field"
'        dlpDate(0).SetFocus
'        Exit Function
'    Else
'        If Not IsDate(dlpDate(0).Text) Then
'            MsgBox "Starting Date is not a valid date."
'            dlpDate(0).SetFocus
'            Exit Function
'        End If
'    End If
'
'    If Len(dlpDate(1).Text) > 0 Then
'        If Not IsDate(dlpDate(1).Text) Then
'            MsgBox "Renewal Date is not a valid date."
'            dlpDate(1).Text = ""
'            dlpDate(1).SetFocus
'            Exit Function
'        End If
'    End If
'
'    If Len(Trim(medDuesPaid)) = 0 Then
'        medDuesPaid = 0
'    Else
'        If Not IsNumeric(medDuesPaid) Then
'            MsgBox "Dues Paid must be numeric"
'            medDuesPaid.SetFocus
'            Exit Function
'        End If
'    End If
    
    chkETrainList = True

Exit Function

chkETrainList_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkTrainList", "HR_TRAIN", "validation")
    Call RollBack
End Function

Sub cmdCancel_Click()
    Dim X
    On Error GoTo Can_Err
    
    rsDATA.CancelUpdate
    
    Call Display_Value
    
    fglbNew = False
    
    Call SET_UP_MODE
    'Call ST_UPD_MODE(True)  ' reset screen's attributes
    
    clpCode(0).Enabled = False
    lblTitle(21).FontBold = False
    lblTitle(18).FontBold = False
    
    fglbNew = False
    Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_TRAIN", "Cancel")
    Call RollBack
End Sub

Sub cmdClose_Click()
    Unload Me
    If glbOnTop = "FRMETRAINLST" Then glbOnTop = ""
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
    
    'Call procedure to delete from Follow Up record and update renewal date on
    'Continuing Education screen for this course if Position Code is found
    'If Not IsNull(Data1.Recordset("TR_JOB")) And Data1.Recordset("TR_JOB") <> "" Then
        'One of the position's required courses.
        'Delete follow up record and update renewal date on continuing education record
        Call Del_Upd_FollowUp_ContinuingEducation
    'End If
    
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
    
    clpCode(0).Enabled = True
    lblPosDesc.Caption = ""
    lblTitle(21).FontBold = True
    lblTitle(18).FontBold = True
    lblCNum.Caption = "001"
    
    Call Set_Control("B", Me)
    
    rsDATA.AddNew
    
    'Friesens - Ticket #16189
    'If glbtermopen Then lblEEID = glbTERM_ID Else
    lblEEID = glbLEE_ID
    
    lblCNum.Caption = "001"
    clpCode(0).SetFocus
    
Exit Sub
    
AddN_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_TRAIN", "Add")
    Call RollBack
End Sub

Sub cmdOK_Click()
    Dim X
    Dim xComments As String
    Dim SQLQ As String
    Dim rsFollowUp As New ADODB.Recordset
    
    On Error GoTo Add_Err
    
    If Not chkETrainList() Then Exit Sub
    
    'Check if Renewal Date has changed
    If Not fglbNew And oRenewalDt <> dlpRenewal.Text Then
        'Call procedure to update Follow Up record and Continuing Education record
        'only if this is one of the required courses
        'If Not IsNull(Data1.Recordset("TR_JOB")) And Data1.Recordset("TR_JOB") <> "" Then
            Call Update_Renewal_Dates
        'End If
    End If
    
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
        
        'Add a new Follow Up record if a new course is added
        If fglbNew Then
            'Add Follow Up record for this course
            If Len(Trim(lblPosCode.Caption)) = 0 Then
                rsDATA("TR_FOLLOWUP_ID") = Add_FollowUp_Record_For_Independant_Course(clpCode(0).Text, dlpRenewal.Text)
            Else
                rsDATA("TR_FOLLOWUP_ID") = Add_FollowUp_Record_For_Independant_Course(clpCode(0).Text, dlpRenewal.Text, lblPosCode.Caption)
            End If
        End If
        
        'Because somewhere in the code in other forms is not saving the Follow Up ID (a bug), I am temp. trying
        'to find the related follow record and updating HR_TRAIN with the Follow Up ID.
        If IsNull(rsDATA("TR_FOLLOWUP_ID")) Then
            xComments = "Course: " & rsDATA("TR_CRSCODE") & " "
            SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
            SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
            SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' AND EF_FDATE = " & Date_SQL(rsDATA("TR_RENEW"))
            rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsFollowUp.EOF Then
                rsDATA("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
            End If
            rsFollowUp.Close
            Set rsFollowUp = Nothing
        End If
        
        rsDATA.Update
        
        If fglbNew Then
            If IsDate(lblCourseTakenDt.Caption) Then
                Call Update_Continuing_Education_Renewal_Date(clpCode(0).Text, dlpRenewal.Text)
            End If
        End If
                
        gdbAdoIhr001.CommitTrans
        
        
    'End If
    Data1.Refresh

    Set rsGrid = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True

    'Call ST_UPD_MODE(True)

    fglbNew = False
    
    Call SET_UP_MODE
    
    clpCode(0).Enabled = False
    lblTitle(21).FontBold = False
    lblTitle(18).FontBold = False

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
            SQLQ = SQLQ & " ORDER BY TR_RENEW"
        ElseIf glbOracle Then
            SQLQ = "SELECT HR_TRAIN.* FROM HR_TRAIN"    '(CASE WHEN TR_POS_TYPE ='C' THEN 'Current' WHEN  TR_POS_TYPE ='T' THEN 'Temporary' WHEN  TR_POS_TYPE ='P' THEN 'Previous' END) AS POSTYPE
            SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
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

Private Sub clpCode_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
    If Index = 0 Then
        clpCode(Index).TransDiv = Get_Not_Required_Courses
    End If
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
    If Index = 0 And fglbNew And clpCode(0).ListChecker Then
        'Call procedure to check for Position code which requires this course
        Call Retrieve_Position_Requiring_this_Course(clpCode(0).Text)
    End If
End Sub

Private Sub Form_Activate()
    glbOnTop = "FRMETRAINLST"
    Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMETRAINLST"
End Sub

Private Sub Form_Load()
    Dim Answer, DefVal, Msg, Title
    Dim RFound As Integer
    
    glbOnTop = "FRMETRAINLST"
    
    'Friesens - Ticket #16189
    'If glbtermopen Then
    '    Data1.ConnectionString = glbAdoIHRAUDIT
    'Else
        Data1.ConnectionString = glbAdoIHRDB
    'End If
    
    Screen.MousePointer = DEFAULT
    
    If Not glbtermopen Then
        If glbLEE_ID = 0 Then frmEEFIND.Show 1
        If glbLEE_ID = 0 Then Unload Me: Exit Sub
    'Friesens - Ticket #16189
    'Else
    '    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    '    If glbTERM_ID = 0 Then Unload Me: Exit Sub
    End If
    
    If EERetrieve() = False Then
        MsgBox "Sorry, Employee can not be found"
        'Friesens - Ticket #16189
        'If glbtermopen Then frmTERMEMPL.Show 1 Else
        frmEEFIND.Show 1
    Else
        Me.Show
        'Friesens - Ticket #16189
        'If glbtermopen Then lblEEID = glbTERM_ID Else
        lblEEID = glbLEE_ID
    End If
    
    If Len(glbLEE_SName) < 1 Then Exit Sub
    
    Screen.MousePointer = HOURGLASS
    Me.vbxTrueGrid.SetFocus
    
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
        'Ticket #20447 - Jerry asked to change to Training Plan for everyone except Friesens and
        'Chatham-Kent but Chatham-Kent are not using 7.9
        If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
            Me.Caption = "Training List - " & Left$(glbLEE_SName, 5)
        Else
            Me.Caption = "Training Plan - " & Left$(glbLEE_SName, 5)
        End If
        Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    
    lblEENum.Caption = ShowEmpnbr(lblEEID)
    
    Call Display_Value
    Call ST_UPD_MODE(True)             '
    Call INI_Controls(Me)
    
    vbxTrueGrid.Columns(4).Visible = False
    
    'Ticket #20605 - Jerry said to hide Type of Position for all the clients except Friesens
    If glbCompSerial <> "S/N - 2279W" Then
        vbxTrueGrid.Columns(5).Visible = False
    End If
    
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
    ElseIf Not IsNull(rsGrid("TR_RENEW")) And rsGrid("TR_RENEW") <> "" Then
        If CVDate(rsGrid("TR_RENEW")) < CVDate(Format(Now, "Short Date")) And (rsGrid("TR_COURSE_TAKEN") = "" Or IsNull(rsGrid("TR_COURSE_TAKEN"))) Then
            RowStyle.ForeColor = vbRed
        End If
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
        SQLQ = "SELECT HR_TRAIN.*,(CASE TR_POS_TYPE WHEN 'C' THEN 'Current' WHEN 'T' THEN 'Temporary' WHEN 'P' THEN 'Previous' END) AS POSTYPE FROM HR_TRAIN"
        SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
    ElseIf glbOracle Then
        SQLQ = "SELECT HR_TRAIN.* FROM HR_TRAIN"    '(CASE WHEN TR_POS_TYPE ='C' THEN 'Current' WHEN  TR_POS_TYPE ='T' THEN 'Temporary' WHEN  TR_POS_TYPE ='P' THEN 'Previous' END) AS POSTYPE
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
        lblPosDesc.Caption = ""
        oRenewalDt = ""
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
        If glbSQL Then
            SQLQ = "SELECT HR_TRAIN.*, (CASE TR_POS_TYPE WHEN 'C' THEN 'Current' WHEN 'T' THEN 'Temporary' WHEN 'P' THEN 'Previous' END) AS POSTYPE FROM HR_TRAIN"
            SQLQ = SQLQ & " WHERE TR_ID = " & Data1.Recordset!TR_ID
            SQLQ = SQLQ & " ORDER BY TR_RENEW"
        ElseIf glbOracle Then
            SQLQ = "SELECT HR_TRAIN.* FROM HR_TRAIN" ', (CASE WHEN TR_POS_TYPE ='C' THEN 'Current' WHEN  TR_POS_TYPE ='T' THEN 'Temporary' WHEN  TR_POS_TYPE ='P' THEN 'Previous' END) AS POSTYPE FROM HR_TRAIN"
            SQLQ = SQLQ & " WHERE TR_ID = " & Data1.Recordset!TR_ID
            SQLQ = SQLQ & " ORDER BY TR_RENEW"
        End If
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    'End If

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    
    Call Set_Control("R", Me, rsDATA)
    Call SET_UP_MODE
    
    'Store the original Renewal Date
    oRenewalDt = dlpRenewal.Text
    
    'Display Position Description and Start Date
    lblPosDesc.Caption = GetJobData(Data1.Recordset("TR_JOB"), "JB_DESCR") & "   /   " & Data1.Recordset("TR_SDATE")
        
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
        'Ticket #20447 - Jerry asked to change to Training Plan for everyone except Friesens and
        'Chatham-Kent but Chatham-Kent are not using 7.9
        If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
            frmETRAINLST.Caption = "Training List - " & Left$(glbLEE_SName, 5)
        Else
            frmETRAINLST.Caption = "Training Plan - " & Left$(glbLEE_SName, 5)
        End If
        frmETRAINLST.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
    
    lblEENum = ShowEmpnbr(lblEEID)
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = glbLEE_ProdLine
    Else
        lblEEProdLine = ""
    End If
End Sub

Private Sub Update_Renewal_Dates()
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim SQLQ As String
    Dim xComments As String
    Dim xFollowUpID As Integer
    
    'Update Follow Up record - Effective Date
    'Update with Follow Up ID if the HR_TRAIN is missing one
    If IsNull(Data1.Recordset("TR_FOLLOWUP_ID")) Then
        xComments = "Course: " & Data1.Recordset("TR_CRSCODE") & " "
        SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
        SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' AND EF_FDATE = " & Date_SQL(Data1.Recordset("TR_RENEW"))
        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsFollowUp.EOF Then
            rsDATA("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
            xFollowUpID = rsFollowUp("EF_FOLLOWUP_ID")
        End If
        rsFollowUp.Close
        Set rsFollowUp = Nothing
    End If
    
    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & IIf(IsNull(Data1.Recordset("TR_FOLLOWUP_ID")), xFollowUpID, Data1.Recordset("TR_FOLLOWUP_ID"))
    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsFollowUp("EF_FDATE") = dlpRenewal.Text
    'rsFollowUp("EF_COMMENTS") = "Course: " & rsCourseCode("ES_CRSCODE") & " - " & GetTABLDesc("ESCD", rsCourseCode("ES_CRSCODE")) & " for Position: " & rsEmpJobs("TW_JOB")
    rsFollowUp("EF_LDATE") = Date
    rsFollowUp("EF_LUSER") = glbUserID
    rsFollowUp("EF_LTIME") = Time$
    rsFollowUp.Update
    
    rsFollowUp.Close
    Set rsFollowUp = Nothing

    'Update the Continuing Education record for this course and this employee
    'with Renewal Date and Job Code
    SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
    SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
    If Not IsNull(Data1.Recordset("TR_JOB")) And Data1.Recordset("TR_JOB") <> "" Then
        SQLQ = SQLQ & " AND ES_JOB = '" & Data1.Recordset("TR_JOB") & "'"
    Else
        SQLQ = SQLQ & " AND (ES_JOB IS NULL OR ES_JOB = '')"
    End If
    SQLQ = SQLQ & " AND ES_CRSCODE = '" & Data1.Recordset("TR_CRSCODE") & "'"
    SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDt)
    rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsContEdu.EOF Then
        rsContEdu("ES_RENEW") = dlpRenewal.Text
        'rsContEdu("ES_JOB") = rsEmpJobs("TW_JOB")
        rsContEdu("ES_LDATE") = Date
        rsContEdu("ES_LUSER") = glbUserID
        rsContEdu("ES_LTIME") = Time$
        rsContEdu.Update
    Else
        rsContEdu.Close
        Set rsContEdu = Nothing
        
        'Search for Continuing Education record without the Job Code, cause it's possible the course was previously
        'taken as independent course
        SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
        SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND (ES_JOB IS NULL OR ES_JOB = '')"
        SQLQ = SQLQ & " AND ES_CRSCODE = '" & Data1.Recordset("TR_CRSCODE") & "'"
        SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDt)
        rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsContEdu.EOF Then
            rsContEdu("ES_RENEW") = dlpRenewal.Text
            'rsContEdu("ES_JOB") = rsEmpJobs("TW_JOB")
            rsContEdu("ES_LDATE") = Date
            rsContEdu("ES_LUSER") = glbUserID
            rsContEdu("ES_LTIME") = Time$
            rsContEdu.Update
        End If
    End If
    rsContEdu.Close
    Set rsContEdu = Nothing

End Sub

Private Function Get_Not_Required_Courses()
Dim rsCourses As New ADODB.Recordset
Dim SQLQ As String
Dim xNonReqCourses As String

    xNonReqCourses = "'*'"
    SQLQ = "SELECT * FROM HR_COURSECODE_MASTER"
    SQLQ = SQLQ & " WHERE ES_UNIQUE_FOR_POS = 0"
    SQLQ = SQLQ & " AND ES_CRSCODE NOT IN (SELECT TR_CRSCODE FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & ")"
    'SQLQ = SQLQ & " AND ES_CRSCODE NOT IN (SELECT PC_CRSCODE FROM HR_JOB_COURSE WHERE PC_JOB IN (SELECT JH_JOB FROM QRY_CROSS_TRAINING_RPT))"
    rsCourses.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsCourses.EOF
        xNonReqCourses = xNonReqCourses & ",'" & rsCourses("ES_CRSCODE") & "'"
        rsCourses.MoveNext
    Loop
    Get_Not_Required_Courses = xNonReqCourses

End Function

Private Sub Del_Upd_FollowUp_ContinuingEducation()
    Dim rsContEdu As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim SQLQ As String
    Dim xComments As String
    Dim xFollowUpID As Long  ' As Integer
        
    'Clear the Renewal date for this course and for this employee from
    'Continuing Education screen
    SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_DATCOMP,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
    SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
    If Not IsNull(Data1.Recordset("TR_JOB")) And Data1.Recordset("TR_JOB") <> "" Then
        SQLQ = SQLQ & " AND ES_JOB = '" & Data1.Recordset("TR_JOB") & "'"
    End If
    SQLQ = SQLQ & " AND ES_CRSCODE = '" & Data1.Recordset("TR_CRSCODE") & "'"
    SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(Data1.Recordset("TR_RENEW"))
    SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(Data1.Recordset("TR_COURSE_TAKEN"))
    rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsContEdu.EOF Then
        rsContEdu("ES_RENEW") = Null
        rsContEdu("ES_LDATE") = Date
        rsContEdu("ES_LUSER") = glbUserID
        rsContEdu("ES_LTIME") = Time$
        rsContEdu.Update
        
        If Not IsNull(rsContEdu("ES_DATCOMP")) Then
            'If follow up id is null then find the id
            xFollowUpID = 0
            If IsNull(Data1.Recordset("TR_FOLLOWUP_ID")) Then
                xComments = "Course: " & Data1.Recordset("TR_CRSCODE") & " "
                SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(Data1.Recordset("TR_RENEW"))
                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsFollowUp.EOF Then
                    'Data1.Recordset("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    xFollowUpID = rsFollowUp("EF_FOLLOWUP_ID")
                End If
                rsFollowUp.Close
                Set rsFollowUp = Nothing
            Else
                xFollowUpID = Data1.Recordset("TR_FOLLOWUP_ID")
            End If
        
            'Since the course was completed - mark the Follow Up as
            'Completed instead of deleting it.
            SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & xFollowUpID  'Data1.Recordset("TR_FOLLOWUP_ID")
            gdbAdoIhr001.Execute SQLQ
        Else
            
            'If follow up id is null then find the id
            xFollowUpID = 0
            If IsNull(Data1.Recordset("TR_FOLLOWUP_ID")) Then
                xComments = "Course: " & Data1.Recordset("TR_CRSCODE") & " "
                SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(Data1.Recordset("TR_RENEW"))
                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsFollowUp.EOF Then
                    'Data1.Recordset("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    xFollowUpID = rsFollowUp("EF_FOLLOWUP_ID")
                End If
                rsFollowUp.Close
                Set rsFollowUp = Nothing
            Else
                xFollowUpID = Data1.Recordset("TR_FOLLOWUP_ID")
            End If
                        
            'If Not IsNull(Data1.Recordset("TR_FOLLOWUP_ID")) And Data1.Recordset("TR_FOLLOWUP_ID") <> "" Then
                'Delete the Follow Up record for this training record
                SQLQ = "DELETE FROM HR_FOLLOW_UP"
                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & xFollowUpID  'Data1.Recordset("TR_FOLLOWUP_ID")
                gdbAdoIhr001.Execute SQLQ
                
                'Clear the Follow Up ID in the Position record
                'if the course code is TRAIN
                If Data1.Recordset("TR_CRSCODE") = "TRAIN" Then
                    'Check in both the tables to see which record has the follow id - clear it
                    'Search HR_JOB_HISTORY table for this Position record
                    'and clear with Follow Up Id
                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & xFollowUpID 'Data1.Recordset("TR_FOLLOWUP_ID")
                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsTJob.EOF Then
                        rsTJob("JH_FOLLOWUP_ID") = Null
                        rsTJob.Update
                    End If
                    rsTJob.Close
                    Set rsTJob = Nothing
                                    
                    'Search HR_TEMP_WORK table for this Position record
                    'and clear with Follow Up Id
                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & xFollowUpID   'Data1.Recordset("TR_FOLLOWUP_ID")
                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsTJob.EOF Then
                        rsTJob("TW_FOLLOWUP_ID") = Null
                        rsTJob.Update
                    End If
                    rsTJob.Close
                    Set rsTJob = Nothing
                End If
            'End If
        End If
    Else
    
        'If follow up id is null then find the id
        xFollowUpID = 0
        If IsNull(Data1.Recordset("TR_FOLLOWUP_ID")) Then
            xComments = "Course: " & Data1.Recordset("TR_CRSCODE") & " "
            SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
            SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
            SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
            SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(Data1.Recordset("TR_RENEW"))
            rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsFollowUp.EOF Then
                'Data1.Recordset("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                xFollowUpID = rsFollowUp("EF_FOLLOWUP_ID")
            End If
            rsFollowUp.Close
            Set rsFollowUp = Nothing
        Else
            xFollowUpID = Data1.Recordset("TR_FOLLOWUP_ID")
        End If
    
        'Delete the Follow Up record for this training record
        'If Not IsNull(Data1.Recordset("TR_FOLLOWUP_ID")) And Data1.Recordset("TR_FOLLOWUP_ID") <> "" Then
            SQLQ = "DELETE FROM HR_FOLLOW_UP"
            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & xFollowUpID  'Data1.Recordset("TR_FOLLOWUP_ID")
            gdbAdoIhr001.Execute SQLQ
        
            'Clear the Follow Up ID in the Position record
            'if the course code is TRAIN
            If Data1.Recordset("TR_CRSCODE") = "TRAIN" Then
                'Check in both the tables to see which record has the follow id - clear it
                'Search HR_JOB_HISTORY table for this Position record
                'and clear with Follow Up Id
                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & xFollowUpID 'Data1.Recordset("TR_FOLLOWUP_ID")
                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsTJob.EOF Then
                    rsTJob("JH_FOLLOWUP_ID") = Null
                    rsTJob.Update
                End If
                rsTJob.Close
                Set rsTJob = Nothing
                                
                'Search HR_TEMP_WORK table for this Position record
                'and clear with Follow Up Id
                SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & xFollowUpID   'Data1.Recordset("TR_FOLLOWUP_ID")
                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsTJob.EOF Then
                    rsTJob("TW_FOLLOWUP_ID") = Null
                    rsTJob.Update
                End If
                rsTJob.Close
                Set rsTJob = Nothing
            End If
        'End If
    End If
    rsContEdu.Close
    Set rsContEdu = Nothing

End Sub

Private Function Add_FollowUp_Record_For_Independant_Course(xCourseCode, xRenewalDate, Optional xJob)
    Dim rsFollowUp As New ADODB.Recordset
    Dim SQLQ As String
    
    'Add a Follow Up record for this Training course
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE 1 = 2"
    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsFollowUp.AddNew
    rsFollowUp("EF_COMPNO") = "001"
    rsFollowUp("EF_EMPNBR") = glbLEE_ID
    rsFollowUp("EF_FDATE") = xRenewalDate
    rsFollowUp("EF_FREAS_TABL") = "FURE"
    'Ticket #24257 - Do not update Admin By for them only
    If glbCompSerial <> "S/N - 2262W" Then
        rsFollowUp("EF_ADMINBY_TABL") = "EDAB"
        rsFollowUp("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
    End If
    rsFollowUp("EF_FREAS") = "EDUC"
    If IsMissing(xJob) Then
        rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode)
    Else
        rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & xJob
    End If
    rsFollowUp("EF_LDATE") = Date
    rsFollowUp("EF_LTIME") = Time$
    rsFollowUp("EF_LUSER") = glbUserID
    rsFollowUp.Update
    
    Add_FollowUp_Record_For_Independant_Course = rsFollowUp("EF_FOLLOWUP_ID")
    
    rsFollowUp.Close
    Set rsFollowUp = Nothing
        
End Function

Private Sub Retrieve_Position_Requiring_this_Course(xCourse)
    Dim rsEmpJob As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim SQLQ As String
    Dim xCrsTaken
        
    'Check first if this Course was taken before in the Continuing Education screen
    SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_JOB, ES_DATCOMP,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
    SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
'    If flgUnqForPos Then
'        SQLQ = SQLQ & " AND ES_JOB = '" & xJob & "'"
'    End If
    SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourse & "'"
    SQLQ = SQLQ & " AND (ES_RENEW = '' OR ES_RENEW IS NULL)"
    SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
    SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
    rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsContEdu.EOF Then
        'Course Taken Before
        rsContEdu.MoveFirst
        xCrsTaken = rsContEdu("ES_DATCOMP")
        lblCourseTakenDt = rsContEdu("ES_DATCOMP")
    Else
        'Course not taken before
        xCrsTaken = ""
        
        '7.9 - Enhancement - Open the City of Chatham-Kent logic for all
        'Ticket #19816
        'Search for Cont Edu with Renewal Date
        'If glbCompSerial = "S/N - 2188W" Then
        If glbCompSerial <> "S/N - 2279W" Then
            'Renewal Date is not null
            rsContEdu.Close
            Set rsContEdu = Nothing
            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_JOB, ES_DATCOMP,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
            SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourse & "'"
            SQLQ = SQLQ & " AND (ES_RENEW IS NOT NULL)"
            SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
            SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsContEdu.EOF Then
                'Course Taken Before
                rsContEdu.MoveFirst
                xCrsTaken = rsContEdu("ES_DATCOMP")
                lblCourseTakenDt = rsContEdu("ES_DATCOMP")
                dlpRenewal.Text = rsContEdu("ES_RENEW")
            Else
                'Course not taken before
                xCrsTaken = ""
            End If
        End If
    End If
    rsContEdu.Close
    Set rsContEdu = Nothing
    
    'Get list of employee's Positions (Primary Current, Temp Current and Tracked Positions) requiring this course
    'independent of type of Renewal Periods
    SQLQ = "SELECT JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
    SQLQ = SQLQ & " AND JH_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & xCourse & "')"
    SQLQ = SQLQ & " UNION "
    SQLQ = SQLQ & " SELECT TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
    SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
    SQLQ = SQLQ & " AND TW_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & xCourse & "')"
    SQLQ = SQLQ & " ORDER BY TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmpJob.EOF Then
        'The first record's Position code gets associated with the independent Training List Course.
        'the order is Primary Current, Temp Current and then Previous depending on most recent end date
        rsEmpJob.MoveFirst
                
        lblPosDesc.Caption = GetJobData(rsEmpJob("TW_JOB"), "JB_DESCR") & "   /   " & rsEmpJob("TW_SDATE")
        
        'Assign to fields
        lblPosCode.Caption = rsEmpJob("TW_JOB")
        lblStartDt.Caption = rsEmpJob("TW_SDATE")
        If IIf(IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")), False, rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
            lblPosType.Caption = "P"
        Else
            If rsEmpJob("POS_TYPE") = "CURRENT" Then
                lblPosType.Caption = "C"
            Else
                lblPosType.Caption = "T"
            End If
        End If
        lblCourseTakenDt = xCrsTaken
    End If
    rsEmpJob.Close
    Set rsEmpJob = Nothing
End Sub

Private Sub Update_Continuing_Education_Renewal_Date(xCourse, xRenewalDt)
    Dim rsContEdu As New ADODB.Recordset
    Dim SQLQ As String
    
    SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_JOB, ES_DATCOMP,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
    SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourse & "'"
    SQLQ = SQLQ & " AND (ES_RENEW = '' OR ES_RENEW IS NULL)"
    SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
    SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
    rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsContEdu.EOF Then

        rsContEdu.MoveFirst
        
        rsContEdu("ES_RENEW") = CVDate(xRenewalDt)
        'rsContEdu("ES_JOB") = rsEmpJobs("TW_JOB")
        rsContEdu("ES_LDATE") = Date
        rsContEdu("ES_LUSER") = glbUserID
        rsContEdu("ES_LTIME") = Time$
        rsContEdu.Update
    
    Else
        '7.9 - Enhancement - Open this City of Chatham-Kent logic for all
        'Ticket #19816
        'Search for Cont Edu with Renewal Date
        'If glbCompSerial = "S/N - 2188W" Then
        If glbCompSerial <> "S/N - 2279W" Then
            'Renewal Date is not null
            rsContEdu.Close
            Set rsContEdu = Nothing
            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_JOB, ES_DATCOMP,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
            SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourse & "'"
            SQLQ = SQLQ & " AND (ES_RENEW IS NOT NULL)"
            SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
            SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsContEdu.EOF Then
                rsContEdu.MoveFirst
                
                rsContEdu("ES_RENEW") = CVDate(xRenewalDt)
                rsContEdu("ES_JOB") = lblPosCode.Caption
                rsContEdu("ES_LDATE") = Date
                rsContEdu("ES_LUSER") = glbUserID
                rsContEdu("ES_LTIME") = Time$
                rsContEdu.Update
            End If
        End If
    End If
    rsContEdu.Close
    Set rsContEdu = Nothing
    
End Sub

