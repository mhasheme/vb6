VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmPosCourse 
   Caption         =   "Courses for Position"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   10440
   WindowState     =   2  'Maximized
   Begin VB.Frame fraCType 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   360
      TabIndex        =   29
      Top             =   5400
      Width           =   3015
      Begin VB.OptionButton OptCType 
         Caption         =   "Status"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   32
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton OptCType 
         Caption         =   "Band"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton OptCType 
         Caption         =   "Position"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   30
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCopyReqCourses 
      Caption         =   "&Copy these Required Courses To..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Tag             =   "Copy the required courses to another position"
      Top             =   5880
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.TextBox txtCurDWMY 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "PC_CUR_PRD_DWMY"
      Height          =   285
      Left            =   4890
      TabIndex        =   25
      Top             =   4290
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ComboBox cmbCurDWMY 
      Height          =   315
      ItemData        =   "Fspcour.frx":0000
      Left            =   3990
      List            =   "Fspcour.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Tag             =   "40-Select Day, Week, Month or Year"
      Top             =   4275
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPrvDWMY 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "PC_PRV_PRD_DWMY"
      Height          =   285
      Left            =   4890
      TabIndex        =   24
      Top             =   4650
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ComboBox cmbPrvDWMY 
      Height          =   315
      ItemData        =   "Fspcour.frx":0038
      Left            =   3990
      List            =   "Fspcour.frx":0048
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Tag             =   "40-Select Day, Week, Month or Year"
      Top             =   4635
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtFlwuUpDWMY 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "PC_FLWUP_PRD_DWMY"
      Height          =   285
      Left            =   4890
      TabIndex        =   23
      Top             =   3930
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ComboBox cmbFlwUpDWMY 
      Height          =   315
      ItemData        =   "Fspcour.frx":0070
      Left            =   3990
      List            =   "Fspcour.frx":0080
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "40-Select Day, Week, Month or Year"
      Top             =   3915
      Visible         =   0   'False
      Width           =   975
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   6000
      TabIndex        =   2
      Tag             =   "00-Course Type Code"
      Top             =   2760
      Visible         =   0   'False
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "ESCD"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PC_CRSCODE"
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Tag             =   "01-Course Code"
      Top             =   2760
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESCD"
   End
   Begin Threed.SSCheck chkLegis 
      DataField       =   "PC_LEGISLATED"
      Height          =   255
      Left            =   1995
      TabIndex        =   4
      Tag             =   "40-Legistated"
      Top             =   3525
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   450
      _StockProps     =   78
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
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10440
      _Version        =   65536
      _ExtentX        =   18415
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
      Begin VB.Label lblPosDesc 
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
         Left            =   2880
         TabIndex        =   17
         Top             =   135
         Width           =   630
      End
      Begin VB.Label lblPosition 
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
         TabIndex        =   16
         Top             =   135
         Width           =   630
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
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
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   165
         Width           =   690
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "Fspcour.frx":00A8
      Height          =   1995
      Left            =   120
      OleObjectBlob   =   "Fspcour.frx":00BC
      TabIndex        =   0
      Tag             =   "Required Courses"
      Top             =   480
      Width           =   10125
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   5520
      Top             =   6120
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6000
      Top             =   6240
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
   Begin MSMask.MaskEdBox medPrvPosRenewal 
      DataField       =   "PC_RENEW_CRS_PRV"
      Height          =   285
      Left            =   3225
      TabIndex        =   9
      Tag             =   "20-Previous Position's Renewal Period"
      Top             =   4650
      Visible         =   0   'False
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medFlwUpEffective 
      DataField       =   "PC_RENEW_FOLLOWUP"
      Height          =   285
      Left            =   3225
      TabIndex        =   5
      Tag             =   "20-Follow Up Effective Date Period"
      Top             =   3930
      Visible         =   0   'False
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medCurPosRenewal 
      DataField       =   "PC_RENEW_CRS_CUR"
      Height          =   285
      Left            =   3240
      TabIndex        =   7
      Tag             =   "20-Current Position's Renewal Period"
      Top             =   4290
      Visible         =   0   'False
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      DataField       =   "PC_DEPTNO"
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   3125
      Visible         =   0   'False
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   34
      Top             =   6855
      Width           =   10440
      _Version        =   65536
      _ExtentX        =   18415
      _ExtentY        =   970
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
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
      Begin VB.CommandButton cmdResetTrainPlanAll 
         Caption         =   "Refresh Training Plan for ALL the Positions"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   515
         Left            =   3120
         TabIndex        =   13
         Tag             =   "Copy the required courses to another position"
         Top             =   0
         Width           =   2595
      End
      Begin VB.CommandButton cmdResetTrainPlan 
         Caption         =   "Refresh Training Plan for the SELECTED Position"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   515
         Left            =   360
         TabIndex        =   12
         Tag             =   "Copy the required courses to another position"
         Top             =   0
         Width           =   2595
      End
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   33
      Top             =   3170
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Pos. Renewal Period"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   360
      TabIndex        =   28
      Top             =   4335
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Pos. Renewal Period"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   13
      Left            =   360
      TabIndex        =   27
      Top             =   4695
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Follow Up Effective Date Period"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   14
      Left            =   360
      TabIndex        =   26
      Top             =   3975
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label lblID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1800
      TabIndex        =   22
      Top             =   6480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Legislated"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   21
      Top             =   3555
      Width           =   1200
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   20
      Top             =   2805
      Width           =   1200
   End
   Begin VB.Label lblPositions 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "POST"
      DataField       =   "PC_JOB"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1080
      TabIndex        =   19
      Top             =   6480
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CompNo"
      DataField       =   "PC_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   360
      TabIndex        =   18
      Top             =   6480
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "frmPosCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRecords%, fglbEditMode%
Dim RSDATA As New ADODB.Recordset
Dim fglbNew As Boolean
Dim CodeArray, CodeCount As Integer
Dim flgUniqforPos As Boolean
Dim xCurRenewal, xPrvRenewal, xFlwUpRenewal
Dim xCurDWMY, xPrvDWMY, xFlwUpDWMY
Dim oCurRen, oPrvRen, oFolRen
Dim oCurRenTyp, oPrvRenTyp, oFolRenTyp
Dim flgCopied

Public Sub cmdCancel_Click()
On Error GoTo Can_Err

clpCode(0).Visible = True
clpCode(1).Visible = False

'rsDATA.CancelUpdate

fglbNew = False

Call Display_Value
'
'Data1.Recordset.CancelUpdate
'If Not glbSQL And Not glbOracle Then Call Pause(0.5)
'Data1.Refresh

'Call ST_UPD_MODE(True)  ' reset screen's attributes


Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_JOB_COURSE", "Cancel")
End Sub

'Public Sub cmdClose_Click()
'glbUserUploadMode = SwitchForm: Unload Me
'End Sub

Public Sub cmdDelete_Click()
Dim a As Integer, Msg As String, INo&
Dim SQLQ As String

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    fglbRecords% = False
    Exit Sub
Else
    fglbRecords% = True
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete This Record?"
'Msg = Msg & Chr(10) & "This Record?  "

'7.9 - Enhancement - For all clients now
'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
    'Ticket #20447 - Jerry asked to change to Training Plan for everyone except Friesens and
    'Chatham-Kent but Chatham-Kent are not using 7.9
    If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        Msg = Msg & Chr(10) & Chr(10) & "Note: This will also delete the Training List course records of the employees with this Position."
    Else
        Msg = Msg & Chr(10) & Chr(10) & "Note: This will also delete the Training Plan course records of the employees with this Position."
    End If
'End If

a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

'gdbAdoIhr001.BeginTrans
'rsDATA.Delete
'gdbAdoIhr001.CommitTrans
'Ticket #12513
If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    SQLQ = "DELETE FROM HR_JOB_COURSE WHERE PC_JOB = '" & Data1.Recordset("PC_JOB") & "' "
    SQLQ = SQLQ & "AND PC_CRSCODE = '" & Data1.Recordset("PC_CRSCODE") & "' "
    
    'Ticket #25609 - Training Plan by Department
    If glbCompSerial <> "S/N - 2279W" Then
        If Len(Data1.Recordset("PC_DEPTNO")) > 0 Then
            SQLQ = SQLQ & "AND PC_DEPTNO = '" & Data1.Recordset("PC_DEPTNO") & "' "
        End If
    End If
    
    gdbAdoIhr001.Execute SQLQ
    
    '7.9 - Enhancement - For all clients now
    'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
    'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        'Call procedure to delete from employee's Training List as well
        Call Deleted_Training_List_Records
    'End If
Else
    Exit Sub
End If

If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

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
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRJOBCOURSE", "Delete")
End Sub

Public Sub cmdModify_Click()

On Error GoTo Edit_Err
Call SET_UP_MODE
'Call ST_UPD_MODE(True)

'Friesens - Ticket #16189
If glbCompSerial = "S/N - 2279W" Then
    'Keep existing values of the renewal periods
    oCurRen = medCurPosRenewal.Text
    oCurRenTyp = txtCurDWMY.Text
    oPrvRen = medPrvPosRenewal.Text
    oPrvRenTyp = txtPrvDWMY.Text
    oFolRen = medFlwUpEffective.Text
    oFolRenTyp = txtFlwuUpDWMY.Text
    
    clpCode(0).Enabled = False
    
'7.9 - Enhancement - For all clients now
Else 'If glbCompSerial = "S/N - 2188W" Then        'City of Chatham-Kent - Ticket #16794
    'Keep existing values of the renewal periods
    oCurRen = medCurPosRenewal.Text
    oCurRenTyp = txtCurDWMY.Text
    oFolRen = medFlwUpEffective.Text
    oFolRenTyp = txtFlwuUpDWMY.Text
    
    clpCode(0).Enabled = False

    'Ticket #25609 - Training Plan by Department
    clpDept.Enabled = True
End If

fglbEditMode% = True

Exit Sub

Edit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdModify", "HRJOBSKL", "Edit")
End Sub

Public Sub cmdNew_Click()
    Dim SQLQ As String
    
    fglbNew = True
    Call SET_UP_MODE
    'Call ST_UPD_MODE(True)
    
    On Error GoTo AddN_Err
    
    clpCode(1).Visible = True
    clpCode(0).Visible = False
    
    Call Set_Control("B", Me, RSDATA)
    'rsDATA.AddNew
    
    'Friesens - Ticket #16189
    If glbCompSerial = "S/N - 2279W" Then
        'lbltitle(12).FontBold = False
        'lbltitle(13).FontBold = False
        lbltitle(14).FontBold = False
        
        medCurPosRenewal.Enabled = False
        medPrvPosRenewal.Enabled = False
        medFlwUpEffective.Enabled = False
        cmbCurDWMY.Enabled = False
        cmbPrvDWMY.Enabled = False
        cmbFlwUpDWMY.Enabled = False
        
        cmdCopyReqCourses.Enabled = False
    '7.9 - Enhancement - For all clients now
    Else 'If glbCompSerial = "S/N - 2188W" Then   'City of Chatham-Kent - Ticket #16794
        lbltitle(14).FontBold = False
        
        medCurPosRenewal.Enabled = False
        medFlwUpEffective.Enabled = False
        cmbCurDWMY.Enabled = False
        cmbFlwUpDWMY.Enabled = False
        
        cmdCopyReqCourses.Enabled = False
        
        'Ticket #25609 - Training Plan by Department
        clpDept.Enabled = True
    End If
    
    fglbEditMode% = True
    lblCNum.Caption = "001"
    lblPositions.Caption = glbPos$
    
    
    'clpCode(0).Enabled = True
    'clpCode(0).SetFocus
    clpCode(1).Text = ""
    clpCode(1).SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRJOBSKL", "Add")
End Sub

Public Sub cmdOK_Click()
    Dim xRes As Integer
    
On Error GoTo OK_Err

If Not chkPosCourses() Then Exit Sub

clpCode(0).Visible = True
clpCode(1).Visible = False

If fglbNew Then 'Ticket #12513
    Call SaveMultiCode(glbPos$, chkLegis, clpDept.Text)
Else
    'Friesens - Ticket #16189
    If glbCompSerial = "S/N - 2279W" Then
        txtCurDWMY.Text = Left(cmbCurDWMY, 1)
        txtPrvDWMY.Text = Left(cmbPrvDWMY, 1)
        txtFlwuUpDWMY.Text = Left(cmbFlwUpDWMY, 1)
    '7.9 - Enhancement - For all clients now
    Else 'If glbCompSerial = "S/N - 2188W" Then        'City of Chatham-Kent - Ticket #16794
        txtCurDWMY.Text = Left(cmbCurDWMY, 1)
        txtFlwuUpDWMY.Text = Left(cmbFlwUpDWMY, 1)
    End If
    
    'Friesens - Ticket #16189
    If glbCompSerial = "S/N - 2279W" Then
        'Has the renewal periods changes?
        If oCurRen <> medCurPosRenewal.Text Or oCurRenTyp <> txtCurDWMY.Text Or _
            oPrvRen <> medPrvPosRenewal.Text Or oPrvRenTyp <> txtPrvDWMY.Text Or _
            oFolRen <> medFlwUpEffective.Text Or oFolRenTyp <> txtFlwuUpDWMY.Text Then
        
            xRes = MsgBox("Course Renewal Period(s) have changed. Employee's Course Renewal Date will be recomputed on Training List screen." & Chr(10) & Chr(10) & "Do you wish to proceed?", vbYesNo + vbExclamation, "info:HR - Course Renewal Periods")
            If xRes = vbNo Then GoTo Skip_Save
        End If
    '7.9 - Enhancement - For all clients now
    Else 'If glbCompSerial = "S/N - 2188W" Then        'City of Chatham-Kent - Ticket #16794
        'Has the renewal periods changes?
        If oCurRen <> medCurPosRenewal.Text Or oCurRenTyp <> txtCurDWMY.Text Or _
            oFolRen <> medFlwUpEffective.Text Or oFolRenTyp <> txtFlwuUpDWMY.Text Then
        
            'Ticket #20447 - Jerry asked to change to Training Plan for everyone except Friesens and
            'Chatham-Kent but Chatham-Kent are not using 7.9
            If glbCompSerial = "S/N - 2188W" Then
                xRes = MsgBox("Course Renewal Period(s) have changed. Employee's Course Renewal Date will be recomputed on Training List screen." & Chr(10) & Chr(10) & "Do you wish to proceed?", vbYesNo + vbExclamation, "info:HR - Course Renewal Periods")
            Else
                xRes = MsgBox("Course Renewal Period(s) have changed. Employee's Course Renewal Date will be recomputed on Training Plan screen." & Chr(10) & Chr(10) & "Do you wish to proceed?", vbYesNo + vbExclamation, "info:HR - Course Renewal Periods")
            End If
            If xRes = vbNo Then GoTo Skip_Save
        End If
    End If
    
    
    Call Set_Control("U", Me, RSDATA)
    gdbAdoIhr001.BeginTrans
    RSDATA.Update
    gdbAdoIhr001.CommitTrans
    
    'Friesens - Ticket #16189
    If glbCompSerial = "S/N - 2279W" Then
    
        'Has the renewal periods changes?
        If oCurRen <> medCurPosRenewal.Text Or oCurRenTyp <> txtCurDWMY.Text Or _
            oPrvRen <> medPrvPosRenewal.Text Or oPrvRenTyp <> txtPrvDWMY.Text Or _
            oFolRen <> medFlwUpEffective.Text Or oFolRenTyp <> txtFlwuUpDWMY.Text Then
    
            'Check if renewal periods are being added instead of being modified
            If (oCurRen = "") And (medCurPosRenewal.Text <> "") Then
                'Current Renewal Period added
                Call Add_Training_List_Rec_for_New_Renewal_Period(clpCode(0).Text, medCurPosRenewal.Text, txtCurDWMY.Text, medPrvPosRenewal.Text, txtPrvDWMY.Text, medFlwUpEffective.Text, txtFlwuUpDWMY.Text)
                
                'Check if Previous Renewal has been added as well
                If (oPrvRen = "") And (medPrvPosRenewal.Text <> "") Then
                    'Previous Renewal Period added
                    Call Add_Training_List_Rec_for_New_Prv_Renewal_Period(clpCode(0).Text, medCurPosRenewal.Text, txtCurDWMY.Text, medPrvPosRenewal.Text, txtPrvDWMY.Text, medFlwUpEffective.Text, txtFlwuUpDWMY.Text)
                End If
            ElseIf (oPrvRen = "") And (medPrvPosRenewal.Text <> "") Then
                'Previous Renewal Period added
                Call Add_Training_List_Rec_for_New_Prv_Renewal_Period(clpCode(0).Text, medCurPosRenewal.Text, txtCurDWMY.Text, medPrvPosRenewal.Text, txtPrvDWMY.Text, medFlwUpEffective.Text, txtFlwuUpDWMY.Text)
            End If
            
            'Changing from one value to another
            If (oCurRen <> "") Or (oPrvRen <> "") Then
                'Course renewal Period has changed - update Training List, Follow Up and Continuing Education records
                Call Course_Renewal_Period_Change(clpCode(0).Text, , medCurPosRenewal.Text, txtCurDWMY.Text, medPrvPosRenewal.Text, txtPrvDWMY.Text, medFlwUpEffective.Text, txtFlwuUpDWMY.Text)
            End If
        End If
    '7.9 - Enhancement - For all clients now
    Else 'If glbCompSerial = "S/N - 2188W" Then        'City of Chatham-Kent - Ticket #16794
        'Has the renewal periods changes?
        If oCurRen <> medCurPosRenewal.Text Or oCurRenTyp <> txtCurDWMY.Text Or _
            oFolRen <> medFlwUpEffective.Text Or oFolRenTyp <> txtFlwuUpDWMY.Text Then
    
            'Check if renewal periods are being added instead of being modified
            If (oCurRen = "") And (medCurPosRenewal.Text <> "") Then
                'Current Renewal Period added
                
                'Ticket #25609 - Training Plan by Department
                If Len(clpDept.Text) > 0 Then
                    Call Add_Training_List_Rec_for_New_Renewal_Period(clpCode(0).Text, medCurPosRenewal.Text, txtCurDWMY.Text, "", "", medFlwUpEffective.Text, txtFlwuUpDWMY.Text, clpDept.Text)
                Else
                    Call Add_Training_List_Rec_for_New_Renewal_Period(clpCode(0).Text, medCurPosRenewal.Text, txtCurDWMY.Text, "", "", medFlwUpEffective.Text, txtFlwuUpDWMY.Text)
                End If
            End If
            
            'Changing from one value to another
            If (oCurRen <> "") Then
                'Course renewal Period has changed - update Training List, Follow Up and Continuing Education records
                If Len(clpDept.Text) > 0 Then
                    'Ticket #25609 - Training Plan by Department
                    Call Course_Renewal_Period_Change(clpCode(0).Text, , medCurPosRenewal.Text, txtCurDWMY.Text, "", "", medFlwUpEffective.Text, txtFlwuUpDWMY.Text, clpDept.Text)
                Else
                    Call Course_Renewal_Period_Change(clpCode(0).Text, , medCurPosRenewal.Text, txtCurDWMY.Text, "", "", medFlwUpEffective.Text, txtFlwuUpDWMY.Text)
                End If
            End If
        End If
    End If
    
    'Data1.Recordset("PC_CRSCODE") = clpCode(0).Text & ""
    'Data1.Recordset("PC_LEGISLATED") = IIf(chkLegis, -1, 0)
    'Data1.Recordset.UpdateBatch
    'If Not glbSQL And Not glbOracle Then Call Pause(0.5)
End If

Skip_Save:
Data1.Refresh

fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(False)
fglbEditMode% = False

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRJOBSKL", "Update")

Unload Me

End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String

RHeading = Me.Caption
RHeading = RHeading & "-"
RHeading = RHeading & " " & lblPosDesc.Caption

Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Public Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = Me.Caption
RHeading = RHeading & "-"
RHeading = RHeading & " " & lblPosDesc.Caption

Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub clpCode_GotFocus(Index As Integer)
    '7.9 - Enhancement - For all clients now
    'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
    'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        If Index = 1 Then
            clpCode(Index).TransDiv = Get_Course_Code_Master_Codes
        End If
    'End If
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
    Dim xCourses
    Dim xCodeCount
    
    'Friesens - Ticket #16189
    If glbCompSerial = "S/N - 2279W" Then
        If fglbNew Then 'Ticket #12513
            xCourses = clpCode(1).Text
            xCourses = Replace(xCourses, ",,", ",")
            If Len(xCourses) > 0 Then
                If Right(xCourses, 1) = "," Then xCourses = Left(xCourses, Len(xCourses) - 1)
                xCodeCount = CountCharInString(xCourses, ",") + 1
                If xCodeCount = 1 Then
                    'Check if the course code is valid
                    If Course_Code_Valid(xCourses) Then
                        'Display the Renewal Periods if the course code is Unique for each Position
                        If Not Course_Unique_For_Position(xCourses) Then
                            'lbltitle(12).FontBold = False
                            'lbltitle(13).FontBold = False
                            lbltitle(14).FontBold = False
                        
                            medCurPosRenewal.Text = IIf(IsNull(xCurRenewal), "", xCurRenewal)
                            medPrvPosRenewal.Text = IIf(IsNull(xPrvRenewal), "", xPrvRenewal)
                            'Ticket #19816
                            If IsNull(xFlwUpRenewal) Then
                                MsgBox "Please setup the 'Follow Up Effective Date Period' on the Course Code Master screen.", vbExclamation, "Course Code Master Setup missing"
                                Exit Sub
                            Else
                                medFlwUpEffective.Text = xFlwUpRenewal
                            End If
                            txtCurDWMY.Text = IIf(IsNull(xCurDWMY), "", xCurDWMY)
                            txtPrvDWMY.Text = IIf(IsNull(xPrvDWMY), "", xPrvDWMY)
                            txtFlwuUpDWMY.Text = xFlwUpDWMY
                            
                            medCurPosRenewal.Enabled = False
                            medPrvPosRenewal.Enabled = False
                            medFlwUpEffective.Enabled = False
                            cmbCurDWMY.Enabled = False
                            cmbPrvDWMY.Enabled = False
                            cmbFlwUpDWMY.Enabled = False
                        Else
                            'lbltitle(12).FontBold = True
                            'lbltitle(13).FontBold = True
                            lbltitle(14).FontBold = True
                        
                            medCurPosRenewal.Enabled = True
                            medPrvPosRenewal.Enabled = True
                            medFlwUpEffective.Enabled = True
                            cmbCurDWMY.Enabled = True
                            cmbPrvDWMY.Enabled = True
                            cmbFlwUpDWMY.Enabled = True
                        
                            medCurPosRenewal.Text = ""
                            medPrvPosRenewal.Text = ""
                            medFlwUpEffective.Text = ""
                            txtCurDWMY.Text = ""
                            txtPrvDWMY.Text = ""
                            txtFlwuUpDWMY.Text = ""
                        End If
                    Else
                        medCurPosRenewal.Text = ""
                        medPrvPosRenewal.Text = ""
                        medFlwUpEffective.Text = ""
                        
                        medCurPosRenewal.Enabled = False
                        medPrvPosRenewal.Enabled = False
                        medFlwUpEffective.Enabled = False
                        
                        cmbCurDWMY.ListIndex = -1
                        cmbPrvDWMY.ListIndex = -1
                        cmbFlwUpDWMY.ListIndex = -1
                        
                        cmbCurDWMY.Enabled = False
                        cmbPrvDWMY.Enabled = False
                        cmbFlwUpDWMY.Enabled = False
                    End If
                End If
            End If
        End If
    End If
    
    '7.9 - Enhancement - For all clients now
    'City of Chatham-Kent - Ticket #16794
    If glbCompSerial <> "S/N - 2279W" Then   'glbCompSerial = "S/N - 2188W" Then
        If fglbNew Then
            xCourses = clpCode(1).Text
            xCourses = Replace(xCourses, ",,", ",")
            If Len(xCourses) > 0 Then
                If Right(xCourses, 1) = "," Then xCourses = Left(xCourses, Len(xCourses) - 1)
                xCodeCount = CountCharInString(xCourses, ",") + 1
                If xCodeCount = 1 Then
                    'Check if the course code is valid
                    If Course_Code_Valid(xCourses) Then
                        'Display the Renewal Periods if the course code is Unique for each Position
                        If Not Course_Unique_For_Position(xCourses) Then
                            lbltitle(14).FontBold = False
                        
                            medCurPosRenewal.Text = IIf(IsNull(xCurRenewal), "", xCurRenewal)
                            'Ticket #19816
                            If IsNull(xFlwUpRenewal) Then
                                MsgBox "Please setup the 'Follow Up Effective Date Period' on the Course Code Master screen.", vbExclamation, "Course Code Master Setup missing"
                                Exit Sub
                            Else
                                medFlwUpEffective.Text = xFlwUpRenewal
                            End If
                            txtCurDWMY.Text = IIf(IsNull(xCurDWMY), "", xCurDWMY)
                            txtFlwuUpDWMY.Text = xFlwUpDWMY
                            
                            medCurPosRenewal.Enabled = False
                            medFlwUpEffective.Enabled = False
                            cmbCurDWMY.Enabled = False
                            cmbFlwUpDWMY.Enabled = False
                        Else
                            lbltitle(14).FontBold = True
                        
                            medCurPosRenewal.Enabled = True
                            medFlwUpEffective.Enabled = True
                            cmbCurDWMY.Enabled = True
                            cmbFlwUpDWMY.Enabled = True
                        
                            medCurPosRenewal.Text = ""
                            medFlwUpEffective.Text = ""
                            txtCurDWMY.Text = ""
                            txtFlwuUpDWMY.Text = ""
                        End If
                    Else
                        medCurPosRenewal.Text = ""
                        medFlwUpEffective.Text = ""
                        
                        medCurPosRenewal.Enabled = False
                        medFlwUpEffective.Enabled = False
                        
                        cmbCurDWMY.ListIndex = -1
                        cmbFlwUpDWMY.ListIndex = -1
                        
                        cmbCurDWMY.Enabled = False
                        cmbFlwUpDWMY.Enabled = False
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub cmbCurDWMY_Lostfocus()
    txtCurDWMY.Text = Left(cmbCurDWMY, 1)
End Sub

Private Sub cmbFlwUpDWMY_Lostfocus()
    txtFlwuUpDWMY.Text = Left(cmbFlwUpDWMY, 1)
End Sub

Private Sub cmbPrvDWMY_Lostfocus()
    txtPrvDWMY.Text = Left(cmbPrvDWMY, 1)
End Sub

Private Sub cmdCopyReqCourses_Click()
    Dim xOrgPos, xCopyToPos, xMultipleCodes, xExtPosCode
    Dim Msg As String
    Dim answ
        
    Msg = "Are you sure you want to copy these required courses to other Positions?"
    
    '7.9 - Enhancement - For all clients now
    'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
    'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
        'Ticket #20447 - Jerry asked to change to Training Plan for everyone except Friesens and
        'Chatham-Kent but Chatham-Kent are not using 7.9
        If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
            Msg = Msg & Chr(10) & Chr(10) & "Note: This option will also add these courses to employee's Training List screen."
        Else
            Msg = Msg & Chr(10) & Chr(10) & "Note: This option will also add these courses to employee's Training Plan screen."
        End If
    'End If
    answ = MsgBox(Msg, 36, "Confirm Delete")
    
    If answ <> 6 Then Exit Sub
    
    
    'Store the current Position
    xOrgPos = glbPos
    If OptCType(0).Value Then 'Position
        frmJOBS.vbxTrueGrid.MultiSelect = 2
        frmJOBS.Show 1
        frmJOBS.vbxTrueGrid.MultiSelect = 1
    End If
    'Ticket #21670 Franks 03/06/2012 - begin
    If OptCType(1).Value Then 'Band
        Call Get_Code_Normal("WFBD", "Band Codes", "")
        'get the position list based on glbCode
        glbPos = getLocJobCodes(glbCode, "BAND")
    End If
    If OptCType(2).Value Then 'Status
        Call Get_Code_Normal("JBST", "Position Status Codes", "")
        'get the position list based on glbCode
        glbPos = getLocJobCodes(glbCode, "PStatus")
    End If
    'Ticket #21670 Franks 03/06/2012 - end
    
    'Restore back to the orginal Position and copy the 'Copy To Position' to a variable
    xCopyToPos = glbPos
    glbPos = xOrgPos
    
    'Call procedure to copy the required courses to a 'Copy to Position'
    Screen.MousePointer = HOURGLASS
    flgCopied = False
    
    'Check if Multiple Position code selected
    If InStr(1, xCopyToPos, ",") > 0 Then
        'Multiple codes selected
        xMultipleCodes = xCopyToPos
        
        Do While Len(xMultipleCodes) <> 0
            'Extract each code and call function to copy the required courses
            If InStr(1, xMultipleCodes, ",") > 0 Then
                xExtPosCode = Mid(xMultipleCodes, 1, InStr(1, xMultipleCodes, ",") - 1)
            Else
                xExtPosCode = xMultipleCodes
            End If
            
            'Call function to To Copy the Courses
            If glbPos <> xExtPosCode Then
                Call Copy_Required_Course_To(xExtPosCode)
            End If
            
            'Trim the code already copied out
            If InStr(1, xMultipleCodes, ",") > 0 Then
                xMultipleCodes = Mid(xMultipleCodes, InStr(1, xMultipleCodes, ",") + 1)
            Else
                xMultipleCodes = ""
            End If
        Loop
    Else
        If glbPos <> xCopyToPos Then
            Call Copy_Required_Course_To(xCopyToPos)
        End If
    End If
    Screen.MousePointer = DEFAULT
    
    If flgCopied Then
        MsgBox "Required courses from Position Code '" & glbPos & "' copied successfully to '" & xCopyToPos & "'.", vbOKOnly, "Required Courses Copy"
    Else
        If glbPos <> xCopyToPos Then
            MsgBox "No required courses from Position Code '" & glbPos & "' were copied to '" & xCopyToPos & "'.", vbOKOnly, "Required Courses Copy"
        End If
    End If
End Sub

Private Sub cmdResetTrainPlan_Click()
    Dim Answer
    Answer = MsgBox("Are you sure you want to refresh the Training Plan for the SELECTED Position?", 36, "Refresh Training Plan")
                
    If Answer <> 6 Then Exit Sub

    MDIMain.panHelp(0).Caption = "Please wait...reseting Position's Training Plans"
    Screen.MousePointer = HOURGLASS
    
    Call Refresh_Training_Plan(glbPos$)
    
    MDIMain.panHelp(0).Caption = " "
    Screen.MousePointer = DEFAULT
End Sub

Private Sub cmdResetTrainPlanAll_Click()
    Dim Answer
    Answer = MsgBox("Are you sure you want to refresh the Training Plan for the ALL Positions?", 36, "Refresh Training Plan")
                
    If Answer <> 6 Then Exit Sub

    MDIMain.panHelp(0).Caption = "Please wait...reseting ALL Training Plans"
    Screen.MousePointer = HOURGLASS
    
    Call Refresh_Training_Plan
    
    MDIMain.panHelp(0).Caption = " "
    Screen.MousePointer = DEFAULT
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE

glbOnTop = "FRMPOSCOURSE"

Me.cmdModify_Click
End Sub

Private Sub Form_Load()
Dim RFound As Integer
'On Error GoTo Error_Load
glbOnTop = "FRMPOSCOURSE"

Screen.MousePointer = HOURGLASS
lblPosition.Caption = glbPos$
'lblPositions.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$
Data1.ConnectionString = glbAdoIHRDB

'If Not EERetrieve(glbPos$) Then
'    Exit Sub        '  modGet it sets fglbRecords
'End If

'MDIMain.lstPanel.Visible = False
'MDIMain.lstView.Visible = False

'Call ST_UPD_MODE(False)
'Call INI_Controls(Me)
'Screen.MousePointer = DEFAULT
'Exit Sub

'Error_Load:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err

'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form load Error", "Courses", "Select")
Dim SQLQ
On Error GoTo FLErr


Me.Height = 6000
Me.Width = 7000

'Data1.ConnectionString = glbAdoIHRDB

If glbWFC Then 'Ticket #25911 Franks 10/21/2014
    If glbPos = "" Then frmJOBSWFC.Show 1
Else
    If glbPos = "" Then frmJOBS.Show 1
End If
If glbPos = "" Then glbUserUploadMode = UploadFormWithoutCheck: Unload Me: Exit Sub

'Ticket #25609 - Training Plan by Department
If glbCompSerial <> "S/N - 2279W" Then
    clpDept.Visible = True
    lbltitle(3).Visible = True
    Call setCaption(lbltitle(3))
End If

'Friesens - Ticket #16189
If glbCompSerial = "S/N - 2279W" Then
    lbltitle(12).Visible = True
    lbltitle(13).Visible = True
    lbltitle(14).Visible = True
    medCurPosRenewal.Visible = True
    medPrvPosRenewal.Visible = True
    medFlwUpEffective.Visible = True
    cmbCurDWMY.Visible = True
    cmbPrvDWMY.Visible = True
    cmbFlwUpDWMY.Visible = True
    cmdCopyReqCourses.Visible = True
'7.9 - Enhancement - For all clients now
Else 'If glbCompSerial = "S/N - 2188W" Then        'City of Chatham-Kent - Ticket #16794
    lbltitle(12).Visible = True
    lbltitle(13).Visible = False
    lbltitle(14).Visible = True
    medCurPosRenewal.Visible = True
    medPrvPosRenewal.Visible = False
    medFlwUpEffective.Visible = True
    cmbCurDWMY.Visible = True
    cmbPrvDWMY.Visible = False
    cmbFlwUpDWMY.Visible = True
    cmdCopyReqCourses.Visible = True
End If

If EERetrieve() = False Then
    MsgBox "Sorry, Position can not be found"
    If glbWFC Then 'Ticket #25911 Franks 10/21/2014
        frmJOBSWFC.Show 1
    Else
        frmJOBS.Show 1
    End If
Else
    Me.Show
    'lblID = glbPos
End If

''Hemu - 05/29/2003 Begin - Ticket # 4204
'If glbCompSerial = "S/N - 2161W" Then
'    clpCode(0).TextBoxWidth = 1200
'    clpCode(0).MaxLength = 8
'Else
'    clpCode(0).TextBoxWidth = 870
'    clpCode(0).MaxLength = 4
'End If
''Hemu - 05/29/2003 End
'Ticket #15688 Frank, all customer should have this
clpCode(0).TextBoxWidth = 1200
clpCode(0).MaxLength = 8

Screen.MousePointer = DEFAULT

Call Display_Value
Call INI_Controls(Me)

clpCode(1).Left = clpCode(0).Left

If glbWFC Then 'Ticket #21670 Franks 03/06/2012
    OptCType(1).Visible = True
End If

Screen.MousePointer = DEFAULT

Exit Sub

FLErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form load Error", "Required Courses", "Select")
Call RollBack

End Sub

Public Function EERetrieve()   '(StrPos$)
Dim SQLQ$

EERetrieve = False
Screen.MousePointer = HOURGLASS

On Error GoTo EERetrieveErr


' out or left join query not updateable - so do straight.
''SQLQ$ = "SELECT * FROM HR_JOB_COURSE "
''SQLQ$ = SQLQ$ & "WHERE PC_JOB = '" & glbPos & "' "            'StrPos$ & "' "
''SQLQ$ = SQLQ$ & "ORDER BY PC_JOB"
'Ticket #12513 Frank 04/02/2007
If glbOracle Then
    SQLQ$ = "SELECT *, HRTABL.TB_NAME, HRTABL.TB_KEY, HRTABL.TB_DESC FROM HR_JOB_COURSE,HRTABL "
    SQLQ$ = SQLQ$ & "WHERE PC_JOB = '" & glbPos & "' "            'StrPos$ & "' "
    SQLQ$ = SQLQ$ & "AND HR_JOB_COURSE.PC_CRSCODE_TABL = HRTABL.TB_NAME  AND HR_JOB_COURSE.PC_CRSCODE = HRTABL.TB_KEY "
    SQLQ$ = SQLQ$ & "ORDER BY PC_JOB"
Else
    SQLQ$ = "SELECT *, HRTABL.TB_NAME, HRTABL.TB_KEY, HRTABL.TB_DESC FROM HR_JOB_COURSE "
    SQLQ$ = SQLQ$ & "LEFT JOIN HRTABL ON HR_JOB_COURSE.PC_CRSCODE_TABL = HRTABL.TB_NAME  AND HR_JOB_COURSE.PC_CRSCODE = HRTABL.TB_KEY "
    SQLQ$ = SQLQ$ & "WHERE PC_JOB = '" & glbPos & "' "            'StrPos$ & "' "
    SQLQ$ = SQLQ$ & "ORDER BY PC_JOB"
End If
Data1.RecordSource = SQLQ$
Data1.Refresh

lblPosition.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    fglbRecords% = False
'    cmdModify.Enabled = False       'Laura jan 06, 1998
Else
    fglbRecords% = True
End If
EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERetrieveErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Pos Skills", "HR_JOB_COURSE", "SELECT")
Call RollBack
End Function


Private Sub ST_UPD_MODE(YN)
Dim TF As Boolean, FT As Boolean

    If YN Then
        TF = True
        FT = False
    Else
        TF = False
        FT = True
    End If
    
    glbOHSEdit% = TF
    
    'fUPMode = TF    ' update mode
    'cmdOK.Enabled = TF
    'cmdCancel.Enabled = TF
    
    'cmdClose.Enabled = FT
    'cmdModify.Enabled = FT
    'cmdNew.Enabled = FT
    'cmdDelete.Enabled = FT
    'cmdPrint.Enabled = FT
    
    clpCode(0).Enabled = TF
    chkLegis.Enabled = TF


    'Friesens - Ticket #16189
    If glbCompSerial = "S/N - 2279W" Then
        If Data1.Recordset.EOF Or Data1.Recordset.EOF Then
            'lbltitle(12).FontBold = False
            'lbltitle(13).FontBold = False
            lbltitle(14).FontBold = False
                
            medCurPosRenewal.Enabled = False
            medPrvPosRenewal.Enabled = False
            medFlwUpEffective.Enabled = False
            cmbCurDWMY.Enabled = False
            cmbPrvDWMY.Enabled = False
            cmbFlwUpDWMY.Enabled = False
            
            cmdCopyReqCourses.Enabled = False
        End If
    '7.9 - Enhancement - For all clients now
    Else 'If glbCompSerial = "S/N - 2188W" Then        'City of Chatham-Kent - Ticket #16794
        If Data1.Recordset.EOF Or Data1.Recordset.EOF Then
            lbltitle(14).FontBold = False
                
            medCurPosRenewal.Enabled = False
            medFlwUpEffective.Enabled = False
            cmbCurDWMY.Enabled = False
            cmbFlwUpDWMY.Enabled = False
            
            cmdCopyReqCourses.Enabled = False
            
            'Ticket #25609 - Training Plan by Department
            clpDept.Enabled = TF
        End If
    End If
    'If Data1.Recordset.EOF Or Data1.Recordset.EOF Then
    '    cmdDelete.Enabled = False
    '    cmdModify.Enabled = False
    'End If

    If Not Data1.Recordset.EOF Then
        If gSec_Upd_Training_List Then
            cmdResetTrainPlan.Enabled = True
            cmdResetTrainPlanAll.Enabled = True
        End If
    Else
        cmdResetTrainPlan.Enabled = False
        cmdResetTrainPlanAll.Enabled = False
    End If
End Sub


Private Function chkPosCourses()
Dim SQLQ As String, Msg As String, dd#, PID&, Expr#, Skill$
Dim I As Integer
Dim xDept As String

chkPosCourses = False

On Error GoTo chkPosCourses_Err

If fglbNew Then 'Ticket #12513
    Skill$ = clpCode(1).Text
    Skill$ = Replace(Skill$, ",,", ",")
    If Len(Skill$) < 1 Then
        MsgBox "Course code is a required field"
        clpCode(1).SetFocus
        Exit Function
    End If
    
    'Ticket #25609 - Training Plan by Department
    xDept = ""
    If glbCompSerial <> "S/N - 2279W" Then
        If clpDept.Caption = "Unassigned" And Len(clpDept.Text) > 0 Then
            MsgBox lStr("Department Code must be valid")
            clpDept.SetFocus
            Exit Function
        Else
            xDept = clpDept.Text
        End If
    End If
    
    CodeCount = CountCharInString(Skill$, ",") + 1
    CodeArray = Split(Skill$, ",")
    
    'Check code validity - Begin
    For I = 0 To CodeCount - 1
        If InvalidCode(CodeArray(I)) Then
            MsgBox "Course code '" & CodeArray(I) & "' is not valid"
            clpCode(1).SetFocus
            Exit Function
        End If
        If MultiDupCode(glbPos$, CodeArray(I), xDept) Then
            If Len(xDept) > 0 Then
                MsgBox "Course Code with Department must be unique"
            Else
                MsgBox "Course code '" & CodeArray(I) & "' is not unique"
            End If
            clpCode(1).SetFocus
            Exit Function
        End If
        
        '7.9 - Enhancement - For all clients now
        'Friesens - Ticket #16189 or City of Chatham-Kent - Ticket #16794
        'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
            flgUniqforPos = Course_Unique_For_Position(CodeArray(I))
            If flgUniqforPos <> 0 Then
                'If chkRenewalPeriod(I - (CodeCount - 1)) = False Then
                If chkRenewalPeriod(CodeCount - 1) = False Then
                    'If I <> (CodeCount - 1) Then
                    If (CodeCount) > 1 Then
                        MsgBox "Course code '" & CodeArray(I) & "' is marked as 'Unique for each Position'. Renewal Period is required for this course." & vbCrLf & " If multiple courses are being entered then enter 'Unique for each Position' course separately."
                        clpCode(1).SetFocus
                    End If
                    Exit Function
                End If
            End If
        'End If
    Next
    'Check code validity - End
    
    PID& = CLng(Val(lblID))

Else
    If Len(clpCode(0).Text) < 1 Then
        MsgBox "Course code is a required field"
        clpCode(0).SetFocus
        Exit Function
    End If
    
    If clpCode(0).Caption = "Unassigned" Then
        MsgBox "Course code must be valid"
        clpCode(0).SetFocus
        Exit Function
    End If
    Skill$ = clpCode(0).Text
    
    PID& = CLng(Val(lblID))
    
    'Ticket #25609 - Training Plan by Department
    xDept = ""
    If glbCompSerial <> "S/N - 2279W" Then
        If clpDept.Caption = "Unassigned" And Len(clpDept.Text) > 0 Then
            MsgBox lStr("Department Code must be valid")
            clpDept.SetFocus
            Exit Function
        Else
            xDept = clpDept.Text
        End If
    End If
    
    If modISDupSkill(glbPos$, Skill$, PID&, xDept) Then
        If Len(xDept) > 0 Then
            MsgBox "Course Code with Department must be unique"
        Else
            MsgBox "Course Code must be unique"
        End If
        clpCode(0).SetFocus
        Exit Function
    End If
End If


'Friesens - Ticket #16189
If glbCompSerial = "S/N - 2279W" And flgUniqforPos <> 0 Then
    'If Len(Trim(medCurPosRenewal.Text)) = 0 Then
    '    MsgBox "Current Position Renewal Period cannot be blank"
    '    medCurPosRenewal.SetFocus
    '    Exit Function
    'End If
    If Len(Trim(medCurPosRenewal.Text)) > 0 Then
        If Not IsNumeric(medCurPosRenewal.Text) Then
            MsgBox "Current Position Renewal Period is not numeric"
            medCurPosRenewal.SetFocus
            Exit Function
        End If
        If cmbCurDWMY.ListIndex = -1 Then
            MsgBox "Select Day(s)/Month(s)/Week(s)/Year(s) for Current Position Renewal Period"
            cmbCurDWMY.SetFocus
            Exit Function
        End If
    Else
        cmbCurDWMY.ListIndex = -1
    End If
    
    'If Len(Trim(medPrvPosRenewal.Text)) = 0 Then
    '    MsgBox "Previous Position Renewal Period cannot be blank"
    '    medPrvPosRenewal.SetFocus
    '    Exit Function
    'End If
    If Len(Trim(medPrvPosRenewal.Text)) > 0 Then
        If Not IsNumeric(medPrvPosRenewal.Text) Then
            MsgBox "Previous Position Renewal Period is not numeric"
            medPrvPosRenewal.SetFocus
            Exit Function
        End If
        If cmbPrvDWMY.ListIndex = -1 Then
            MsgBox "Select Day(s)/Month(s)/Week(s)/Year(s) for Previous Position Renewal Period"
            cmbPrvDWMY.SetFocus
            Exit Function
        End If
    Else
         cmbPrvDWMY.ListIndex = -1
    End If
    
    If Len(Trim(medFlwUpEffective.Text)) = 0 Then
        MsgBox "Follow Up Effective Date Period cannot be blank"
        medFlwUpEffective.SetFocus
        Exit Function
    End If
    If Not IsNumeric(medFlwUpEffective.Text) Then
        MsgBox "Follow Up Effective Date Period is not numeric"
        medFlwUpEffective.SetFocus
        Exit Function
    End If
    If cmbFlwUpDWMY.ListIndex = -1 Then
        MsgBox "Select Day(s)/Month(s)/Week(s)/Year(s) for Follow Up Effective Date Period"
        cmbFlwUpDWMY.SetFocus
        Exit Function
    End If
End If

'7.9 - Enhancement - For all clients now
'City of Chatham-Kent - Ticket #16794
'If glbCompSerial = "S/N - 2188W" And flgUniqforPos <> 0 Then
If glbCompSerial <> "S/N - 2279W" And flgUniqforPos <> 0 Then
    If Len(Trim(medCurPosRenewal.Text)) > 0 Then
        If Not IsNumeric(medCurPosRenewal.Text) Then
            MsgBox "Current Position Renewal Period is not numeric"
            medCurPosRenewal.SetFocus
            Exit Function
        End If
        If cmbCurDWMY.ListIndex = -1 Then
            MsgBox "Select Day(s)/Month(s)/Week(s)/Year(s) for Current Position Renewal Period"
            cmbCurDWMY.SetFocus
            Exit Function
        End If
    Else
        cmbCurDWMY.ListIndex = -1
    End If
        
    If Len(Trim(medFlwUpEffective.Text)) = 0 Then
        MsgBox "Follow Up Effective Date Period cannot be blank"
        medFlwUpEffective.SetFocus
        Exit Function
    End If
    If Not IsNumeric(medFlwUpEffective.Text) Then
        MsgBox "Follow Up Effective Date Period is not numeric"
        medFlwUpEffective.SetFocus
        Exit Function
    End If
    If cmbFlwUpDWMY.ListIndex = -1 Then
        MsgBox "Select Day(s)/Month(s)/Week(s)/Year(s) for Follow Up Effective Date Period"
        cmbFlwUpDWMY.SetFocus
        Exit Function
    End If
End If

chkPosCourses = True

Exit Function

chkPosCourses_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HRJOBCOURSE", "edit/Add")

End Function

Private Function modISDupSkill(Pos$, Skill$, ID&, xDept)
Dim SQLQ$
Dim snapSkill As New ADODB.Recordset

modISDupSkill = True

On Error GoTo modISDupSkill_Err

Screen.MousePointer = HOURGLASS

SQLQ$ = "SELECT * FROM HR_JOB_COURSE "
SQLQ$ = SQLQ$ & "Where " '(JS_COMPNO = '001' "
SQLQ$ = SQLQ$ & " PC_JOB = '" & Pos$ & "' "

'Ticket #25609 - Training Plan by Department
If Len(xDept) > 0 Then
    SQLQ$ = SQLQ$ & "AND PC_DEPTNO = '" & xDept & "' "
End If

If fglbNew Then
    SQLQ$ = SQLQ$ & "AND PC_CRSCODE IN ('" & Replace(Skill$, ",", "','") & "') "
Else
    SQLQ$ = SQLQ$ & "AND PC_CRSCODE = '" & Skill$ & "' "
    SQLQ$ = SQLQ$ & "AND PC_ID <> " & ID& & " "
End If

snapSkill.Open SQLQ$, gdbAdoIhr001, adOpenStatic

If snapSkill.BOF And snapSkill.EOF Then
    modISDupSkill = False
End If

Screen.MousePointer = DEFAULT
snapSkill.Close

Exit Function

modISDupSkill_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Code Snap", "TABL", "SELECT")
Call RollBack
End Function

Public Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
        RSDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Call SET_UP_MODE
        Me.cmdModify_Click
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HR_JOB_COURSE "
    SQLQ = SQLQ & "WHERE PC_ID = " & Data1.Recordset!PC_ID
    SQLQ = SQLQ & " ORDER BY PC_JOB,PC_CRSCODE"

    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

    If RSDATA.EOF Or RSDATA.BOF Then Exit Sub
    lblID = RSDATA!PC_ID
    Call Set_Control("R", Me, RSDATA)
    
    'Check if this course is Unique for each Position
    'Friesens - Ticket #16189
    If glbCompSerial = "S/N - 2279W" Then
        cmdCopyReqCourses.Enabled = True
        
        flgUniqforPos = False
        
        flgUniqforPos = Course_Unique_For_Position(RSDATA!PC_CRSCODE)
        If flgUniqforPos <> 0 Then  'Make Period entry mandatory
            'lbltitle(12).FontBold = True
            'lbltitle(13).FontBold = True
            lbltitle(14).FontBold = True
            
            medCurPosRenewal.Enabled = True
            medPrvPosRenewal.Enabled = True
            medFlwUpEffective.Enabled = True
            cmbCurDWMY.Enabled = True
            cmbPrvDWMY.Enabled = True
            cmbFlwUpDWMY.Enabled = True
        Else
            'lbltitle(12).FontBold = False
            'lbltitle(13).FontBold = False
            lbltitle(14).FontBold = False
        
            medCurPosRenewal.Enabled = False
            medPrvPosRenewal.Enabled = False
            medFlwUpEffective.Enabled = False
            cmbCurDWMY.Enabled = False
            cmbPrvDWMY.Enabled = False
            cmbFlwUpDWMY.Enabled = False
        End If
    '7.9 - Enhancement - For all clients now
    Else 'If glbCompSerial = "S/N - 2188W" Then        'City of Chatham-Kent - Ticket #16794
        cmdCopyReqCourses.Enabled = True
        
        flgUniqforPos = False
        
        flgUniqforPos = Course_Unique_For_Position(RSDATA!PC_CRSCODE)
        If flgUniqforPos <> 0 Then  'Make Period entry mandatory
            lbltitle(14).FontBold = True
            
            medCurPosRenewal.Enabled = True
            medFlwUpEffective.Enabled = True
            cmbCurDWMY.Enabled = True
            cmbFlwUpDWMY.Enabled = True
        Else
            lbltitle(14).FontBold = False
        
            medCurPosRenewal.Enabled = False
            medFlwUpEffective.Enabled = False
            cmbCurDWMY.Enabled = False
            cmbFlwUpDWMY.Enabled = False
        End If
    End If
    
Call SET_UP_MODE
Me.cmdModify_Click
End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub txtCurDWMY_Change()
    cmbCurDWMY.ListIndex = -1
    Select Case txtCurDWMY
    Case "D"
        cmbCurDWMY.ListIndex = 0
    Case "W"
        cmbCurDWMY.ListIndex = 1
    Case "M"
        cmbCurDWMY.ListIndex = 2
    Case "Y"
        cmbCurDWMY.ListIndex = 3
    End Select
End Sub

Private Sub txtFlwuUpDWMY_Change()
    cmbFlwUpDWMY.ListIndex = -1
    Select Case txtFlwuUpDWMY
    Case "D"
        cmbFlwUpDWMY.ListIndex = 0
    Case "W"
        cmbFlwUpDWMY.ListIndex = 1
    Case "M"
        cmbFlwUpDWMY.ListIndex = 2
    Case "Y"
        cmbFlwUpDWMY.ListIndex = 3
    End Select
End Sub

Private Sub txtPrvDWMY_Change()
    cmbPrvDWMY.ListIndex = -1
    Select Case txtPrvDWMY
    Case "D"
        cmbPrvDWMY.ListIndex = 0
    Case "W"
        cmbPrvDWMY.ListIndex = 1
    Case "M"
        cmbPrvDWMY.ListIndex = 2
    Case "Y"
        cmbPrvDWMY.ListIndex = 3
    End Select
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
    Dim SQLQ As String
    
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    ' out or left join query not updateable - so do straight. - CAUSES ERROR WHEN SORTED ON TB_DESC
    'SQLQ$ = "SELECT * FROM HR_JOB_COURSE "
    'SQLQ$ = SQLQ$ & "WHERE PC_JOB = '" & glbPos & "' "
    'SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    
    If glbOracle Then
        SQLQ$ = "SELECT *, HRTABL.TB_NAME, HRTABL.TB_KEY, HRTABL.TB_DESC FROM HR_JOB_COURSE,HRTABL "
        SQLQ$ = SQLQ$ & "WHERE PC_JOB = '" & glbPos & "' "
        SQLQ$ = SQLQ$ & "AND HR_JOB_COURSE.PC_CRSCODE_TABL = HRTABL.TB_NAME  AND HR_JOB_COURSE.PC_CRSCODE = HRTABL.TB_KEY "
        'SQLQ$ = SQLQ$ & "ORDER BY PC_JOB"
        SQLQ$ = SQLQ$ & "ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    Else
        SQLQ$ = "SELECT *, HRTABL.TB_NAME, HRTABL.TB_KEY, HRTABL.TB_DESC FROM HR_JOB_COURSE "
        SQLQ$ = SQLQ$ & "LEFT JOIN HRTABL ON HR_JOB_COURSE.PC_CRSCODE_TABL = HRTABL.TB_NAME  AND HR_JOB_COURSE.PC_CRSCODE = HRTABL.TB_KEY "
        SQLQ$ = SQLQ$ & "WHERE PC_JOB = '" & glbPos & "' "
        'SQLQ$ = SQLQ$ & "ORDER BY PC_JOB"
        SQLQ$ = SQLQ$ & "ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    End If
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Display_Value
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub

Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)

End Sub

Private Sub lblPositions_Change()
lblPosDesc.Caption = glbPosDesc$
Me.Caption = "Required Courses - " & lblPosition
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

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
    Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Function CountCharInString(xText, xLetter)
Dim I As Integer, J As Integer
    J = 0
    For I = 1 To Len(xText)
        If Mid$(xText, I, 1) = xLetter Then J = J + 1
    Next
    CountCharInString = J
End Function

Private Function InvalidCode(xCode) As Boolean
Dim rsVCode As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ESCD' "
    SQLQ = SQLQ & "AND TB_KEY = '" & xCode & "' "
    rsVCode.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsVCode.EOF Then
        InvalidCode = True
    Else
        InvalidCode = False
    End If
    rsVCode.Close
End Function

Private Function MultiDupCode(Pos$, xCode, xDept) As Boolean
Dim rsVCode As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_JOB = '" & Pos$ & "' "
    SQLQ = SQLQ & "AND PC_CRSCODE = '" & xCode & "' "
    'Ticket #25609 - Training Plan by Department
    If Len(xDept) > 0 Then
        SQLQ = SQLQ & "AND PC_DEPTNO = '" & xDept & "' "
    End If
    rsVCode.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsVCode.EOF Then
        MultiDupCode = False
    Else
        MultiDupCode = True
    End If
    rsVCode.Close

End Function

Private Sub SaveMultiCode(Pos$, Legis$, xDept)
Dim rsMCRSCode As New ADODB.Recordset
Dim I As Integer
Dim SQLQ As String

    SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE 1=2 "
    rsMCRSCode.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    For I = 0 To CodeCount - 1
        rsMCRSCode.AddNew
        rsMCRSCode("PC_COMPNO") = "001"
        rsMCRSCode("PC_JOB") = Pos$
        rsMCRSCode("PC_CRSCODE") = CodeArray(I)
        If Legis$ Then
            rsMCRSCode("PC_LEGISLATED") = -1
        Else
            rsMCRSCode("PC_LEGISLATED") = 0
        End If
                  
        'Friesens - Ticket #16189
        If glbCompSerial = "S/N - 2279W" Then
            'if Course exists in Course Code master then do the following
            If Course_Exists(CodeArray(I)) Then
                xCurRenewal = ""
                xPrvRenewal = ""
                xFlwUpRenewal = ""
                xCurDWMY = ""
                xPrvDWMY = ""
                xFlwUpDWMY = ""
                'If not unique for each position then get the values from Course Code Master
                If Not Course_Unique_For_Position(CodeArray(I)) Then
                    If xCurRenewal <> "" Then
                        rsMCRSCode("PC_RENEW_CRS_CUR") = xCurRenewal
                    End If
                    If xPrvRenewal <> "" Then
                        rsMCRSCode("PC_RENEW_CRS_PRV") = xPrvRenewal
                    End If
                    rsMCRSCode("PC_RENEW_FOLLOWUP") = xFlwUpRenewal
                    If xCurRenewal <> "" Then
                        rsMCRSCode("PC_CUR_PRD_DWMY") = xCurDWMY
                    End If
                    If xPrvRenewal <> "" Then
                        rsMCRSCode("PC_PRV_PRD_DWMY") = xPrvDWMY
                    End If
                    rsMCRSCode("PC_FLWUP_PRD_DWMY") = xFlwUpDWMY
                Else
                    If medCurPosRenewal.Text <> "" Then
                        rsMCRSCode("PC_RENEW_CRS_CUR") = medCurPosRenewal.Text
                    End If
                    If medPrvPosRenewal.Text <> "" Then
                        rsMCRSCode("PC_RENEW_CRS_PRV") = medPrvPosRenewal.Text
                    End If
                    rsMCRSCode("PC_RENEW_FOLLOWUP") = medFlwUpEffective.Text
                    If medCurPosRenewal.Text <> "" Then
                        rsMCRSCode("PC_CUR_PRD_DWMY") = Left(cmbCurDWMY, 1)
                    End If
                    If medPrvPosRenewal.Text <> "" Then
                        rsMCRSCode("PC_PRV_PRD_DWMY") = Left(cmbPrvDWMY, 1)
                    End If
                    rsMCRSCode("PC_FLWUP_PRD_DWMY") = Left(cmbFlwUpDWMY, 1)
                End If
                
                'Add the new course in the employee's Training List if the employee has this position as current
                'or tracked for course renewals
                Call Add_New_Required_Courses_to_TrainingList(CodeArray(I))
            End If
        '7.9 - Enhancement - For all clients now
        Else 'If glbCompSerial = "S/N - 2188W" Then        'City of Chatham-Kent - Ticket #16794
            'Ticket #25609 - Training Plan by Department
            If Len(xDept) > 0 Then
                rsMCRSCode("PC_DEPTNO") = xDept
            End If

            'if Course exists in Course Code master then do the following
            If Course_Exists(CodeArray(I)) Then
                xCurRenewal = ""
                xFlwUpRenewal = ""
                xCurDWMY = ""
                xFlwUpDWMY = ""
                'If not unique for each position then get the values from Course Code Master
                If Not Course_Unique_For_Position(CodeArray(I)) Then
                    If xCurRenewal <> "" Then
                        rsMCRSCode("PC_RENEW_CRS_CUR") = xCurRenewal
                    End If
                    rsMCRSCode("PC_RENEW_FOLLOWUP") = xFlwUpRenewal
                    If xCurRenewal <> "" Then
                        rsMCRSCode("PC_CUR_PRD_DWMY") = xCurDWMY
                    End If
                    rsMCRSCode("PC_FLWUP_PRD_DWMY") = xFlwUpDWMY
                Else
                    If medCurPosRenewal.Text <> "" Then
                        rsMCRSCode("PC_RENEW_CRS_CUR") = medCurPosRenewal.Text
                    End If
                    rsMCRSCode("PC_RENEW_FOLLOWUP") = medFlwUpEffective.Text
                    If medCurPosRenewal.Text <> "" Then
                        rsMCRSCode("PC_CUR_PRD_DWMY") = Left(cmbCurDWMY, 1)
                    End If
                    rsMCRSCode("PC_FLWUP_PRD_DWMY") = Left(cmbFlwUpDWMY, 1)
                End If
                
                'Add the new course in the employee's Training List if the employee has this position as current
                'or tracked for course renewals
                If Len(xDept) > 0 Then
                    'Ticket #25609 - Training Plan by Department
                    Call Add_New_Required_Courses_to_TrainingList(CodeArray(I), , , , , , , , xDept)
                Else
                    Call Add_New_Required_Courses_to_TrainingList(CodeArray(I))
                End If
            End If
        End If
        rsMCRSCode.Update
    Next
    rsMCRSCode.Close
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelatePOS
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_ReqCourses
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
ElseIf RSDATA.EOF Then
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

Private Function Course_Unique_For_Position(xCourseCode)
    Dim rsCourseCodeMst As New ADODB.Recordset
    Dim SQLQ As String
    
    Course_Unique_For_Position = False
    
    SQLQ = "SELECT ES_UNIQUE_FOR_POS,ES_RENEW_CRS_CUR,ES_RENEW_CRS_PRV,ES_RENEW_FOLLOWUP,ES_CUR_PRD_DWMY,ES_PRV_PRD_DWMY,ES_FLWUP_PRD_DWMY FROM HR_COURSECODE_MASTER"
    SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & xCourseCode & "'"
    rsCourseCodeMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsCourseCodeMst.EOF Then
        If rsCourseCodeMst("ES_UNIQUE_FOR_POS") <> 0 Then
            Course_Unique_For_Position = True
        Else
            xCurRenewal = rsCourseCodeMst("ES_RENEW_CRS_CUR")
            
            '7.9 - Enhancement - For all clients now
            'Not City of Chatham-Kent - Ticket #16794
            'If glbCompSerial <> "S/N - 2188W" Then
            If glbCompSerial = "S/N - 2279W" Then
                xPrvRenewal = rsCourseCodeMst("ES_RENEW_CRS_PRV")
                xPrvDWMY = rsCourseCodeMst("ES_PRV_PRD_DWMY")
            End If
            
            xFlwUpRenewal = rsCourseCodeMst("ES_RENEW_FOLLOWUP")
            
            xCurDWMY = rsCourseCodeMst("ES_CUR_PRD_DWMY")
            xFlwUpDWMY = rsCourseCodeMst("ES_FLWUP_PRD_DWMY")

            Course_Unique_For_Position = False
        End If
    End If
    rsCourseCodeMst.Close
    Set rsCourseCodeMst = Nothing
End Function

Private Function Course_Exists(xCourseCode)
    Dim rsCourseCodeMst As New ADODB.Recordset
    Dim SQLQ As String
    
    Course_Exists = False
    
    SQLQ = "SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER"
    SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & xCourseCode & "'"
    rsCourseCodeMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsCourseCodeMst.EOF Then
        Course_Exists = True
    End If
    rsCourseCodeMst.Close
    Set rsCourseCodeMst = Nothing

End Function

Private Function chkRenewalPeriod(xArrCount)

    chkRenewalPeriod = False

    If Len(Trim(medCurPosRenewal.Text)) > 0 Then
        If Not IsNumeric(medCurPosRenewal.Text) Then
            If xArrCount = 0 Then
                MsgBox "Current Position Renewal Period is not numeric"
                medCurPosRenewal.SetFocus
            End If
            Exit Function
        End If
        If cmbCurDWMY.ListIndex = -1 Then
            If xArrCount = 0 Then
                MsgBox "Select Day(s)/Month(s)/Week(s)/Year(s) for Current Position Renewal Period"
                cmbCurDWMY.SetFocus
            End If
            Exit Function
        End If
    Else
        cmbCurDWMY.ListIndex = -1
    End If
    
    '7.9 - Enhancement - For all clients now
    'Not City of Chatham-Kent - Ticket #16794
    'If glbCompSerial <> "S/N - 2188W" Then
    If glbCompSerial = "S/N - 2279W" Then
        If Len(Trim(medPrvPosRenewal.Text)) > 0 Then
            If Not IsNumeric(medPrvPosRenewal.Text) Then
                If xArrCount = 0 Then
                    MsgBox "Previous Position Renewal Period is not numeric"
                    medPrvPosRenewal.SetFocus
                End If
                Exit Function
            End If
            If cmbPrvDWMY.ListIndex = -1 Then
                If xArrCount = 0 Then
                    MsgBox "Select Day(s)/Month(s)/Week(s)/Year(s) for Previous Position Renewal Period"
                    cmbPrvDWMY.SetFocus
                End If
                Exit Function
            End If
        Else
            cmbPrvDWMY.ListIndex = -1
        End If
    End If
    
    If Len(Trim(medFlwUpEffective.Text)) = 0 Then
        If xArrCount = 0 Then
            MsgBox "Follow Up Effective Date Period cannot be blank"
            medFlwUpEffective.SetFocus
        End If
        Exit Function
    End If
    If Not IsNumeric(medFlwUpEffective.Text) Then
        If xArrCount = 0 Then
            MsgBox "Current Position Renewal Period is not numeric"
            medFlwUpEffective.SetFocus
        End If
        Exit Function
    End If
    If cmbFlwUpDWMY.ListIndex = -1 Then
        If xArrCount = 0 Then
            MsgBox "Select Day(s)/Month(s)/Week(s)/Year(s) for Follow Up Effective Date Period"
            cmbFlwUpDWMY.SetFocus
        End If
        Exit Function
    End If
    
    chkRenewalPeriod = True

End Function

Private Sub Copy_Required_Course_To(xCopyToPos)
    Dim rsSReqCourses As New ADODB.Recordset
    Dim rsDReqCourses As New ADODB.Recordset
    Dim SQLQ As String
    
    flgCopied = False
    
    If Len(xCopyToPos) = 0 Then Exit Sub
    
    'Retrieve required courses of the source position.
    SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE PC_JOB = '" & glbPos & "'"
    rsSReqCourses.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    If Not rsSReqCourses.EOF Then
        rsSReqCourses.MoveFirst
        
        Do While Not rsSReqCourses.EOF
            'Copy to destination position after checking if the position does not already exists
            SQLQ = "SELECT * FROM HR_JOB_COURSE "
            SQLQ = SQLQ & "WHERE PC_JOB = '" & xCopyToPos & "'"
            SQLQ = SQLQ & "AND PC_CRSCODE = '" & rsSReqCourses("PC_CRSCODE") & "'"
            
            'Ticket #25609 - Training Plan by Department
            If glbCompSerial <> "S/N - 2279W" Then
                If Len(rsSReqCourses("PC_DEPTNO")) > 0 Then
                    SQLQ = SQLQ & "AND PC_DEPTNO = '" & rsSReqCourses("PC_DEPTNO") & "'"
                End If
            End If
            
            rsDReqCourses.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
            If rsDReqCourses.EOF Then
                rsDReqCourses.AddNew
                rsDReqCourses("PC_COMPNO") = rsSReqCourses("PC_COMPNO")
                rsDReqCourses("PC_JOB") = xCopyToPos
                rsDReqCourses("PC_CRSCODE") = rsSReqCourses("PC_CRSCODE")
            
                'Ticket #25609 - Training Plan by Department
                If glbCompSerial <> "S/N - 2279W" Then
                    rsDReqCourses("PC_DEPTNO") = rsSReqCourses("PC_DEPTNO")
                End If
                
                rsDReqCourses("PC_LEGISLATED") = rsSReqCourses("PC_LEGISLATED")
                rsDReqCourses("PC_RENEW_CRS_CUR") = rsSReqCourses("PC_RENEW_CRS_CUR")
                
                '7.9 - Enhancement - For all clients now
                'Not City of Chatham-Kent - Ticket #16794
                'If glbCompSerial <> "S/N - 2188W" Then
                If glbCompSerial = "S/N - 2279W" Then
                    rsDReqCourses("PC_RENEW_CRS_PRV") = rsSReqCourses("PC_RENEW_CRS_PRV")
                    rsDReqCourses("PC_PRV_PRD_DWMY") = rsSReqCourses("PC_PRV_PRD_DWMY")
                End If
                
                rsDReqCourses("PC_RENEW_FOLLOWUP") = rsSReqCourses("PC_RENEW_FOLLOWUP")
                
                rsDReqCourses("PC_CUR_PRD_DWMY") = rsSReqCourses("PC_CUR_PRD_DWMY")
                rsDReqCourses("PC_FLWUP_PRD_DWMY") = rsSReqCourses("PC_FLWUP_PRD_DWMY")
                
                rsDReqCourses.Update
                
                flgCopied = True
                
                'Add the copied course in the employee's Training List if the employee has this position as current
                'or tracked for course renewals
                'City of Chatham-Kent - Ticket #16794
                '7.9 - Enhancement - For all clients now
                'If glbCompSerial = "S/N - 2188W" Then
                If glbCompSerial <> "S/N - 2279W" Then
                    Call Add_New_Required_Courses_to_TrainingList(rsSReqCourses("PC_CRSCODE"), xCopyToPos, rsSReqCourses("PC_RENEW_CRS_CUR"), rsSReqCourses("PC_CUR_PRD_DWMY"), "", "", rsSReqCourses("PC_RENEW_FOLLOWUP"), rsSReqCourses("PC_FLWUP_PRD_DWMY"))
                Else
                    Call Add_New_Required_Courses_to_TrainingList(rsSReqCourses("PC_CRSCODE"), xCopyToPos, rsSReqCourses("PC_RENEW_CRS_CUR"), rsSReqCourses("PC_CUR_PRD_DWMY"), rsSReqCourses("PC_RENEW_CRS_PRV"), rsSReqCourses("PC_PRV_PRD_DWMY"), rsSReqCourses("PC_RENEW_FOLLOWUP"), rsSReqCourses("PC_FLWUP_PRD_DWMY"))
                End If
                
            End If
            rsDReqCourses.Close
            Set rsDReqCourses = Nothing
            
            rsSReqCourses.MoveNext
        Loop
    End If
    rsSReqCourses.Close
End Sub

Private Sub Deleted_Training_List_Records()
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim SQLQ As String
    Dim xComments As String
    Dim xFollowUpID As Integer
    
    'If a required course is deleted then the corresponding Training List records with this course
    'should be deleted as well. The follow up record should be marked Completed if the Course has been taken.
    'The Course Renewal Date on the Continuing Education screen be cleared.
    'If TRAIN course then delete the Follow Up ref in HR_JOB_HISTORY and HR_TEMP_WORK
    
    
    'Retrieve all training list records with this course and position being deleted
    SQLQ = "SELECT * FROM HR_TRAIN "
    SQLQ = SQLQ & " WHERE TR_JOB = '" & glbPos & "' AND TR_CRSCODE = '" & clpCode(0).Text & "'"
    rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRTrain.EOF Then
        'Records found in Training List with thsi Job and Course
        rsHRTrain.MoveFirst
        
        Do While Not rsHRTrain.EOF
            'Clear the Renewal date for this course and for this employee from
            'Continuing Education screen
            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_DATCOMP,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
            SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
            SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsHRTrain("TR_CRSCODE") & "'"
            SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(rsHRTrain("TR_RENEW"))
            SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
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
                    If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                        xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                        SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                        SQLQ = SQLQ & " WHERE EF_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                        SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(rsHRTrain("TR_RENEW"))
                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsFollowUp.EOF Then
                            'Data1.Recordset("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                            xFollowUpID = rsFollowUp("EF_FOLLOWUP_ID")
                        End If
                        rsFollowUp.Close
                        Set rsFollowUp = Nothing
                    Else
                        xFollowUpID = rsHRTrain("TR_FOLLOWUP_ID")
                    End If
                
                    'Since the course was completed - mark the Follow Up as
                    'Completed instead of deleting it.
                    SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & xFollowUpID  'rsHRTrain("TR_FOLLOWUP_ID")
                    gdbAdoIhr001.Execute SQLQ
                Else
                
                    'If follow up id is null then find the id
                    xFollowUpID = 0
                    If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                        xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                        SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                        SQLQ = SQLQ & " WHERE EF_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                        SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(rsHRTrain("TR_RENEW"))
                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsFollowUp.EOF Then
                            'Data1.Recordset("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                            xFollowUpID = rsFollowUp("EF_FOLLOWUP_ID")
                        End If
                        rsFollowUp.Close
                        Set rsFollowUp = Nothing
                    Else
                        xFollowUpID = rsHRTrain("TR_FOLLOWUP_ID")
                    End If
                
                    'Delete the Follow Up record for this training record
                    'as no Course completion record found
                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & xFollowUpID  'rsHRTrain("TR_FOLLOWUP_ID")
                    gdbAdoIhr001.Execute SQLQ
                
                    'Clear the Follow Up ID in the Temp/Cross Training Position record
                    'if the course code is TRAIN
                    If rsHRTrain("TR_CRSCODE") = "TRAIN" Then
                        'Search HR_JOB_HISTORY table for this Position record
                        'and update with Follow Up Id
                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & xFollowUpID 'rsHRTrain("TR_FOLLOWUP_ID")
                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsTJob.EOF Then
                            rsTJob("JH_FOLLOWUP_ID") = Null
                            rsTJob.Update
                        End If
                        rsTJob.Close
                        Set rsTJob = Nothing
                        
                        'Search HR_TEMP_WORK table for this Position record
                        'and update with Follow Up Id
                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & xFollowUpID   'rsHRTrain("TR_FOLLOWUP_ID")
                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsTJob.EOF Then
                            rsTJob("TW_FOLLOWUP_ID") = Null
                            rsTJob.Update
                        End If
                        rsTJob.Close
                        Set rsTJob = Nothing
                    End If
                End If
            Else
                'If follow up id is null then find the id
                xFollowUpID = 0
                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                    SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(rsHRTrain("TR_RENEW"))
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        'Data1.Recordset("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                        xFollowUpID = rsFollowUp("EF_FOLLOWUP_ID")
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                Else
                    xFollowUpID = rsHRTrain("TR_FOLLOWUP_ID")
                End If
                
                'Delete the Follow Up record for this training record
                'as no Course completion record found
                SQLQ = "DELETE FROM HR_FOLLOW_UP"
                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & xFollowUpID  'rsHRTrain("TR_FOLLOWUP_ID")
                gdbAdoIhr001.Execute SQLQ
            
                'Clear the Follow Up Id in the Temp/Cross Training Position record
                'if the course code is TRAIN
                If rsHRTrain("TR_CRSCODE") = "TRAIN" Then
                    'Search HR_JOB_HISTORY table for this Position record
                    'and update with Follow Up Id
                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & xFollowUpID 'rsHRTrain("TR_FOLLOWUP_ID")
                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsTJob.EOF Then
                        rsTJob("JH_FOLLOWUP_ID") = Null
                        rsTJob.Update
                    End If
                    rsTJob.Close
                    Set rsTJob = Nothing
                    
                    'Search HR_TEMP_WORK table for this Position record
                    'and update with Follow Up Id
                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & xFollowUpID   'rsHRTrain("TR_FOLLOWUP_ID")
                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsTJob.EOF Then
                        rsTJob("TW_FOLLOWUP_ID") = Null
                        rsTJob.Update
                    End If
                    rsTJob.Close
                    Set rsTJob = Nothing
                End If
            End If
            rsContEdu.Close
            Set rsContEdu = Nothing
            
            'Delete this Training List record as the course is deleted from this position
            'SQLQ = "DELETE FROM HR_TRAIN"
            'SQLQ = SQLQ & " WHERE TR_EMPNBR = " & rsHRTrain("TR_EMPNBR")
            'SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
            'SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsHRTrain("TR_CRSCODE") & "'"
            'gdbAdoIhr001.Execute SQLQ
            rsHRTrain.Delete
            
            rsHRTrain.MoveNext
        Loop
    End If
    rsHRTrain.Close
    Set rsHRTrain = Nothing
    
    'Call procedure to update Training List records of the employees requiring this course through
    'other Current or Tracked Positions of these employee.
    Call Update_Other_EmpPositions_Require_This_Course(clpCode(0).Text)
    
End Sub

Private Sub Add_New_Required_Courses_to_TrainingList(xCourseCode, Optional xJobCode, Optional xCurRen, Optional xCurRenTyp, Optional xPrvRen, Optional xPrvRenTyp, Optional xFolRen, Optional xFolRenTyp, Optional xDept)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsEmpJob As New ADODB.Recordset
    Dim rsCourseMst As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim SQLQ As String
    Dim flgUnqForPos, flgNoPrvRnwl, flgNoCurRnwl, flgCrsTakenBefore As Boolean
    Dim xDWMY As String
    Dim CurRen, PrvRen, FolRen
    Dim CurRenTyp, PrvRenTyp, FolRenTyp
    Dim oRenewalDate As Date
    Dim flgRenewalPeriod As Boolean
    Dim oJob As String
    Dim xPrvEndDate
    Dim flgIfCond_AddTrainRec As Boolean
    Dim xComments As String

    If Not IsMissing(xJobCode) Then
        CurRen = xCurRen
        PrvRen = xPrvRen
        FolRen = xFolRen
        
        CurRenTyp = xCurRenTyp
        PrvRenTyp = xPrvRenTyp
        FolRenTyp = xFolRenTyp
    Else
        CurRen = medCurPosRenewal.Text
        '7.9 - Enhancement - For all clients now
        'Not City of Chatham-Kent - Ticket #16794
        'If glbCompSerial <> "S/N - 2188W" Then
        If glbCompSerial = "S/N - 2279W" Then
            PrvRen = medPrvPosRenewal.Text
            PrvRenTyp = txtPrvDWMY.Text
        End If
        FolRen = medFlwUpEffective.Text
        
        CurRenTyp = txtCurDWMY.Text
        FolRenTyp = txtFlwuUpDWMY.Text
    End If

    'If a new required course is added to a Position then that course should be added to the Training List records of
    'all the employees having this Position as either Current or marked to Track Course Renewal.
    
    flgUnqForPos = False
    

    'Check if this new required course is Unique for each Position.
    'If so, then it will have to be added in the Training List even
    'though the Course code already exists for this employee for another positions
    'Course should be existing in Course Code Master first
    SQLQ = "SELECT ES_CRSCODE,ES_UNIQUE_FOR_POS,ES_RENEW_CRS_CUR,ES_CUR_PRD_DWMY, ES_RENEW_CRS_PRV,ES_PRV_PRD_DWMY, ES_RENEW_FOLLOWUP, ES_FLWUP_PRD_DWMY FROM HR_COURSECODE_MASTER"
    SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & xCourseCode & "'"
    rsCourseMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsCourseMst.EOF Then
        flgUnqForPos = IIf(IsNull(rsCourseMst("ES_UNIQUE_FOR_POS")), False, rsCourseMst("ES_UNIQUE_FOR_POS"))
        
        If flgUnqForPos = 0 Then    'Not flgUnqForPos
            'Ticket #20518 - If Follow Up renewal is 99years then skip adding to training list
            If rsCourseMst("ES_RENEW_FOLLOWUP") = 99 And rsCourseMst("ES_FLWUP_PRD_DWMY") = "Y" Then Exit Sub
            
            If Not IsMissing(xJobCode) Then
                CurRen = xCurRen
                PrvRen = xPrvRen
                FolRen = xFolRen
                
                CurRenTyp = xCurRenTyp
                PrvRenTyp = xPrvRenTyp
                FolRenTyp = xFolRenTyp
            Else
                CurRen = rsCourseMst("ES_RENEW_CRS_CUR") 'medCurPosRenewal.Text
                '7.9 - Enhancement - For all clients now
                'Not City of Chatham-Kent - Ticket #16794
                'If glbCompSerial <> "S/N - 2188W" Then
                If glbCompSerial = "S/N - 2279W" Then
                    PrvRen = rsCourseMst("ES_RENEW_CRS_PRV") 'medPrvPosRenewal.Text
                    PrvRenTyp = rsCourseMst("ES_PRV_PRD_DWMY") 'txtPrvDWMY.Text
                End If
                FolRen = rsCourseMst("ES_RENEW_FOLLOWUP") 'medFlwUpEffective.Text
                
                CurRenTyp = rsCourseMst("ES_CUR_PRD_DWMY") 'txtCurDWMY.Text
                FolRenTyp = rsCourseMst("ES_FLWUP_PRD_DWMY") 'txtFlwuUpDWMY.Text
            End If
        End If
    Else
        'Course not defined in the Course Code Master - skip this process
        'Not a valid course
        rsCourseMst.Close
        Set rsCourseMst = Nothing

        Exit Sub
    End If

    'Get list of employees with this Position as Current or marked to Track for Course Renewal in
    'HR_JOB_HISTORY and HR_TEMP_WORK tables
    SQLQ = "SELECT 'C' AS JOBTYPE, JH_ID AS TW_ID, JH_EMPNBR AS TW_EMPNBR, JH_JOB AS TW_JOB, JH_SDATE AS TW_SDATE, JH_CURRENT AS TW_CURRENT, JH_ENDDATE AS TW_ENDDATE, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL FROM HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
    If IsMissing(xJobCode) Then
        SQLQ = SQLQ & " AND JH_JOB = '" & glbPos & "'"
    Else
        SQLQ = SQLQ & " AND JH_JOB = '" & xJobCode & "'"
    End If
    
    'Ticket #25609 - Training Plan by Department
    If Not IsMissing(xDept) Then
        SQLQ = SQLQ & " AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_DEPTNO = '" & xDept & "')"
    End If
    
    SQLQ = SQLQ & " UNION "
    SQLQ = SQLQ & " SELECT 'T' AS JOBTYPE, TW_ID, TW_EMPNBR, TW_JOB, TW_SDATE, TW_CURRENT, TW_ENDDATE, TW_TRK_CRS_RENEWAL FROM HR_TEMP_WORK "
    SQLQ = SQLQ & " WHERE ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
    If IsMissing(xJobCode) Then
        SQLQ = SQLQ & " AND TW_JOB = '" & glbPos & "'"
    Else
        SQLQ = SQLQ & " AND TW_JOB = '" & xJobCode & "'"
    End If
    
    'Ticket #25609 - Training Plan by Department
    If Not IsMissing(xDept) Then
        SQLQ = SQLQ & " AND TW_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_DEPTNO = '" & xDept & "')"
    End If
    
    SQLQ = SQLQ & " ORDER BY TW_EMPNBR,TW_TRK_CRS_RENEWAL ASC,JOBTYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmpJob.EOF Then
        rsEmpJob.MoveFirst

        Do While Not rsEmpJob.EOF
            flgNoPrvRnwl = False
            flgNoCurRnwl = False

            'Add the Required Courses in the Training List
            'if it does not already exists for this employee or Unique for each Position
            SQLQ = "SELECT * FROM HR_TRAIN"
            SQLQ = SQLQ & " WHERE TR_EMPNBR = " & rsEmpJob("TW_EMPNBR")
            SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourseCode & "'"
            If flgUnqForPos <> 0 Then
                If IsMissing(xJobCode) Then
                    SQLQ = SQLQ & " AND TR_JOB = '" & glbPos & "'"
                Else
                    SQLQ = SQLQ & " AND TR_JOB = '" & xJobCode & "'"
                End If

                'If rsEmpJob("TW_CURRENT") And rsEmpJob("JOBTYPE") = "C" Then
                '    SQLQ = SQLQ & " AND TR_POS_TYPE = 'C'"
                'ElseIf rsEmpJob("TW_CURRENT") And rsEmpJob("JOBTYPE") = "T" Then
                '    SQLQ = SQLQ & " AND TR_POS_TYPE = 'T'"
                'ElseIf rsEmpJob("TW_TRK_CRS_RENEWAL") Then
                '    SQLQ = SQLQ & " AND TR_POS_TYPE = 'P'"
                'End If
            End If
            rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsHRTrain.EOF Then
                'TRAINING RECORD DOES NOT EXISTS - ADD NEW ONE
                
                'Check first if this Course was taken before in the Continuing Education screen
                flgCrsTakenBefore = False
                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_JOB,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsEmpJob("TW_EMPNBR")
                'SQLQ = SQLQ & " AND ES_JOB = '" & xJob & "'"
                If flgUnqForPos <> 0 Then
                    If IsMissing(xJobCode) Then
                        SQLQ = SQLQ & " AND ES_JOB = '" & glbPos & "'"
                    Else
                        SQLQ = SQLQ & " AND ES_JOB = '" & xJobCode & "'"
                    End If
                End If
                SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                SQLQ = SQLQ & " AND (ES_RENEW = '' OR ES_RENEW IS NULL)"
                SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
                SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsContEdu.EOF Then
                    'Course Taken Before
                    rsContEdu.MoveFirst
                    flgCrsTakenBefore = True
                Else
                    'Course not taken before
                    flgCrsTakenBefore = False
                    
                    'Ticket #19816
                    'Search for Cont Edu with Renewal Date
                    'Search for Cont Edu without Job - for Unique for Position Courses
                    '7.9 - Enhancement - For all clients now
                    'If glbCompSerial = "S/N - 2188W" Then
                    If glbCompSerial <> "S/N - 2279W" Then
                        'Renewal Date is not null
                        rsContEdu.Close
                        Set rsContEdu = Nothing
                        SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_JOB,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
                        SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsEmpJob("TW_EMPNBR")
                        If flgUnqForPos <> 0 Then
                            If IsMissing(xJobCode) Then
                                SQLQ = SQLQ & " AND ES_JOB = '" & glbPos & "'"
                            Else
                                SQLQ = SQLQ & " AND ES_JOB = '" & xJobCode & "'"
                            End If
                        End If
                        SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                        SQLQ = SQLQ & " AND (ES_RENEW IS NOT NULL)"
                        SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
                        SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
                        rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsContEdu.EOF Then
                            'Course Taken Before
                            rsContEdu.MoveFirst
                            flgCrsTakenBefore = True
                        Else
                            'Course not taken before
                            flgCrsTakenBefore = False
                            
                            'Search for Cont Edu without Job - for Unique for Position Courses
                            'Renewal Date null and without Job
                            rsContEdu.Close
                            Set rsContEdu = Nothing
                            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_JOB,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
                            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsEmpJob("TW_EMPNBR")
                            SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                            SQLQ = SQLQ & " AND (ES_RENEW = '' OR ES_RENEW IS NULL)"
                            SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
                            SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
                            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsContEdu.EOF Then
                                'Course Taken Before
                                rsContEdu.MoveFirst
                                flgCrsTakenBefore = True
                            Else
                                'Course not taken before
                                flgCrsTakenBefore = False
                                
                                'Search for Cont Edu without Job - for Unique for Position Courses
                                'Renewal Date not null and without Job
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_JOB,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsEmpJob("TW_EMPNBR")
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                                SQLQ = SQLQ & " AND (ES_RENEW IS NOT NULL)"
                                SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
                                SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'Course Taken Before
                                    rsContEdu.MoveFirst
                                    flgCrsTakenBefore = True
                                Else
                                    'Course not taken before
                                    flgCrsTakenBefore = False
                                End If
                            End If
                        End If
                    End If
                End If
                                
                'If the course is being added for the Previous Position and this course
                'does not have previous renewal period then do not add this course
                'If (rsEmpJob("JOBTYPE") = "C" Or rsEmpJob("JOBTYPE") = "T") Or (rsEmpJob("TW_TRK_CRS_RENEWAL") And (Not IsNull(PrvRen)) And PrvRen <> 0) Then
                
                
                'If Course was taken and it's Position is Current then
                'make sure Current Renewal Period is there otherwise do not add the course
                'If the course is being added for the Previous Position and this course
                'does not have previous renewal period then do not add this course
                
                'Ticket #19816
                'Since Chatham-Kent do not use Prv Renewal, Temp. Position and Track Course Renewal,
                'I have to determine if to add Training Rec value differently using IF condition. The values
                'returned from Prv Renewal, Temp. Position and Track Course Renewal is either null or empty and
                'so the IF condition does not return correct result to add Training List or Not.
                flgIfCond_AddTrainRec = False
                '7.9 - Enhancement - For all clients now
                'If glbCompSerial = "S/N - 2188W" Then
                If glbCompSerial <> "S/N - 2279W" Then
                    If (flgCrsTakenBefore = True And (rsEmpJob("JOBTYPE") = "C" Or rsEmpJob("JOBTYPE") = "T") And rsEmpJob("TW_CURRENT") And (CurRen <> "") And CurRen <> 0) Or _
                        (flgCrsTakenBefore = False And (rsEmpJob("JOBTYPE") = "C" Or rsEmpJob("JOBTYPE") = "T")) Then
                        
                        flgIfCond_AddTrainRec = True
                    Else
                        flgIfCond_AddTrainRec = False
                    End If
                Else
                    If (flgCrsTakenBefore = True And (rsEmpJob("JOBTYPE") = "C" Or rsEmpJob("JOBTYPE") = "T") And rsEmpJob("TW_CURRENT") And (Not rsEmpJob("TW_TRK_CRS_RENEWAL")) And (CurRen <> "") And CurRen <> 0) Or _
                        (flgCrsTakenBefore = False And (rsEmpJob("JOBTYPE") = "C" Or rsEmpJob("JOBTYPE") = "T")) Or (flgCrsTakenBefore = True And rsEmpJob("TW_TRK_CRS_RENEWAL") And (PrvRen <> "") And PrvRen <> 0) Or _
                        (flgCrsTakenBefore = False And rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                        
                        flgIfCond_AddTrainRec = True
                    Else
                        flgIfCond_AddTrainRec = False
                    End If
                End If
                'If (flgCrsTakenBefore = True And (rsEmpJob("JOBTYPE") = "C" Or rsEmpJob("JOBTYPE") = "T") And rsEmpJob("TW_CURRENT") And (Not rsEmpJob("TW_TRK_CRS_RENEWAL")) And (CurRen <> "") And CurRen <> 0) Or _
                '    (flgCrsTakenBefore = False And (rsEmpJob("JOBTYPE") = "C" Or rsEmpJob("JOBTYPE") = "T")) Or (flgCrsTakenBefore = True And rsEmpJob("TW_TRK_CRS_RENEWAL") And (PrvRen <> "") And PrvRen <> 0) Or _
                '    (flgCrsTakenBefore = False And rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                If flgIfCond_AddTrainRec Then
                    'Add Training Record
                    rsHRTrain.AddNew
                    rsHRTrain("TR_COMPNO") = "001"
                    rsHRTrain("TR_EMPNBR") = rsEmpJob("TW_EMPNBR")
                    rsHRTrain("TR_CRSCODE") = xCourseCode
                    
                    If flgCrsTakenBefore = False Then   'Course not taken before
                        If CurRen <> "" And CurRen <> 0 Then
                            'Current Course Renewal found
                            Select Case FolRenTyp
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            If rsEmpJob("JOBTYPE") = "C" Or rsEmpJob("JOBTYPE") = "T" Or IIf(IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")), False, rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, FolRen, CVDate(rsEmpJob("TW_SDATE")))
                            
                            'For courses not taken and are now Previous, the renewal date is based
                            'on Follow Up Renewal Period and not Previous Renewal Period - above
                            'ElseIf rsEmpJob("JOBTYPE") = "P" Then
                            '    Select Case PrvRenTyp
                            '        Case "D"
                            '            xDWMY = "d"
                            '        Case "W"
                            '            xDWMY = "ww"
                            '        Case "M"
                            '            xDWMY = "m"
                            '        Case "Y"
                            '            xDWMY = "yyyy"
                            '    End Select
                            '    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, PrvRen, CVDate(rsEmpJob("TW_SDATE")))
                            End If
                        Else    'No Current Course Renewal Period
                            If ((rsEmpJob("JOBTYPE") = "C" Or rsEmpJob("JOBTYPE") = "T") And rsEmpJob("TW_CURRENT")) Or _
                                IIf(IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")), False, rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                                Select Case FolRenTyp
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                If Not IsNull(FolRenTyp) Then
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, FolRen, CVDate(rsEmpJob("TW_SDATE")))
                                End If
                            'ElseIf rsEmpJob("TW_TRK_CRS_RENEWAL") Then
                            '    'For courses not taken and are now Previous, the renewal date is based
                            '    'on Follow Up Renewal Period and not Previous Renewal Period.
                            '    'If there is no current renewal then it's based on End Date only and
                            '    'Prev Renewal Period - when course taken.
                            '    'Compute Renewal with Position End Date because there is no Current Renewal Period defined
                            '    Select Case PrvRenTyp
                            '        Case "D"
                            '            xDWMY = "d"
                            '        Case "W"
                            '            xDWMY = "ww"
                            '        Case "M"
                            '            xDWMY = "m"
                            '        Case "Y"
                            '            xDWMY = "yyyy"
                            '    End Select
                            '    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, PrvRen, CVDate(rsEmpJob("TW_ENDDATE")))
                            End If
                        End If
                    Else    'Course Has Been Taken Before
                        'Course has been taken before, compute Renewal Date based on Course Taken Date
                        If (rsEmpJob("JOBTYPE") = "C" Or rsEmpJob("JOBTYPE") = "T") And rsEmpJob("TW_CURRENT") Then
                            Select Case CurRenTyp
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            'Ticket #19816
                            '7.9 - Enhancement - For all clients now
                            'If glbCompSerial = "S/N - 2188W" Then
                            If glbCompSerial <> "S/N - 2279W" Then
                                If IsDate(rsContEdu("ES_RENEW")) Then
                                    rsHRTrain("TR_RENEW") = rsContEdu("ES_RENEW")   'If they have already entered the Renewal Date then follow that.
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, CurRen, CVDate(rsContEdu("ES_DATCOMP")))
                                End If
                            Else
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, CurRen, CVDate(rsContEdu("ES_DATCOMP")))
                            End If
                            rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                        ElseIf IIf(IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")), False, rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                            Select Case PrvRenTyp
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            If CurRen <> "" And CurRen <> 0 Then
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, PrvRen, CVDate(rsContEdu("ES_DATCOMP")))
                            Else
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, PrvRen, CVDate(rsEmpJob("TW_ENDDATE")))
                            End If
                            rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                        End If
                        
                        'Update Continuing Education with new Renewal Date
                        If IsMissing(xJobCode) Then
                            rsContEdu("ES_JOB") = glbPos
                        Else
                            rsContEdu("ES_JOB") = xJobCode
                        End If
                        rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                        rsContEdu("ES_LDATE") = Date
                        rsContEdu("ES_LUSER") = glbUserID
                        rsContEdu("ES_LTIME") = Time$
                        rsContEdu.Update
                    End If
                    
                    If IsMissing(xJobCode) Then
                        rsHRTrain("TR_JOB") = glbPos
                    Else
                        rsHRTrain("TR_JOB") = xJobCode
                    End If
                    rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                    If rsEmpJob("TW_CURRENT") And rsEmpJob("JOBTYPE") = "C" Then
                        rsHRTrain("TR_POS_TYPE") = "C"
                    ElseIf rsEmpJob("TW_CURRENT") And rsEmpJob("JOBTYPE") = "T" Then
                        rsHRTrain("TR_POS_TYPE") = "T"
                    ElseIf IIf(IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")), False, rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                        rsHRTrain("TR_POS_TYPE") = "P"
                    End If
                    'rsHRTrain("TR_COURSE_TAKEN")   - Remains BLANK
                    rsHRTrain("TR_LDATE") = Date
                    rsHRTrain("TR_LTIME") = Time$
                    rsHRTrain("TR_LUSER") = glbUserID

                    'Add a Follow Up record for this Training course
                    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE 1 = 2"
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    rsFollowUp.AddNew
                    rsFollowUp("EF_COMPNO") = "001"
                    rsFollowUp("EF_EMPNBR") = rsEmpJob("TW_EMPNBR")
                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                    rsFollowUp("EF_FREAS_TABL") = "FURE"
                    'Ticket #24257 - Do not update Admin By for them only
                    If glbCompSerial <> "S/N - 2262W" Then
                        rsFollowUp("EF_ADMINBY_TABL") = "EDAB"
                        rsFollowUp("EF_ADMINBY") = GetEmpData(rsEmpJob("TW_EMPNBR"), "ED_ADMINBY", Null)
                    End If
                    rsFollowUp("EF_FREAS") = "EDUC"
                    If IsMissing(xJobCode) Then
                        rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & glbPos
                    Else
                        rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & xJobCode
                    End If
                    rsFollowUp("EF_LDATE") = Date
                    rsFollowUp("EF_LTIME") = Time$
                    rsFollowUp("EF_LUSER") = glbUserID
                    rsFollowUp.Update

                    rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    rsHRTrain.Update

                    rsFollowUp.Close
                    Set rsFollowUp = Nothing

                    'Update Position record with Follow Up ID
                    'if the course code is TRAIN
                    If xCourseCode = "TRAIN" Then
                        'Search HR_JOB_HISTORY/HR_TEMP_WORK table for this Position record
                        'and update with Follow Up Id
                        If rsEmpJob("JOBTYPE") = "C" Then
                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & rsEmpJob("TW_ID")
                        Else
                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & rsEmpJob("TW_ID")
                        End If
                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsTJob.EOF Then
                            If rsEmpJob("JOBTYPE") = "C" Then
                                rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                            Else
                                rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                            End If
                            rsTJob.Update
                        End If
                        rsTJob.Close
                        Set rsTJob = Nothing
                    End If
                Else
                    'Ticket #19816
                    '7.9 - Enhancement - For all clients now
                    'If glbCompSerial = "S/N - 2188W" And flgCrsTakenBefore = True And (IsNull(CurRen) Or (CurRen = "") Or CurRen = 0) Then
                    If glbCompSerial <> "S/N - 2279W" And flgCrsTakenBefore = True And (IsNull(CurRen) Or (CurRen = "") Or CurRen = 0) Then
                        'Update Continuing Education with Job
                        If IsNull(rsContEdu("ES_JOB")) Then
                            If IsMissing(xJobCode) Then
                                rsContEdu("ES_JOB") = glbPos
                            Else
                                rsContEdu("ES_JOB") = xJobCode
                            End If
                            'rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                            rsContEdu("ES_LDATE") = Date
                            rsContEdu("ES_LUSER") = glbUserID
                            rsContEdu("ES_LTIME") = Time$
                            rsContEdu.Update
                        End If
                    End If
                End If
                
                rsContEdu.Close
                Set rsContEdu = Nothing
            Else
                'TRAINING LIST RECORD EXISTS
                'Check if the course has been taken
                If Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                    'COURSE TAKEN
                    oRenewalDate = rsHRTrain("TR_RENEW")
                    flgRenewalPeriod = True
                    
                    
                    'Check if it's Independant course, Current, Temporary or Previous record.
                    If IsNull(rsHRTrain("TR_JOB")) Then
                        'INDEPENDANT COURSE
                        'Compute Renewal Date based on the Type of the Position
                        If rsEmpJob("TW_CURRENT") And (rsEmpJob("JOBTYPE") = "C" Or rsEmpJob("JOBTYPE") = "T") Then
                            'Current Position or Temporary Position
                            'Primary/Temporary Current Position - See if Current Renewal Period found
                            If Not IsNull(CurRen) And CurRen <> 0 Then
                                'Current Renewal Period found
                                'Calculate Renewal Date based on the Renewal Period and Course Taken Date
                                Select Case CurRenTyp
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                'Ticket #16858 - Do not change the Renewal Date
                                'rsHRTrain("TR_RENEW") = DateAdd(xDWMY, CurRen, CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                            Else
                                'No Current Renewal Period
                                flgRenewalPeriod = False
                                
                                'Delete Training List record, and update Follow Up and Continuing Education records
                                GoTo Delete_Training_Record
                            End If
                        
                        ElseIf IIf(IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")), False, rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                            'Previous Position - See if Previous Renewal Period found
                            If Not IsNull(PrvRen) And PrvRen <> 0 Then
                                'Previous Renewal Period found
                                'Calculate Renewal Date based on the Renewal Period and Course Taken Date
                                Select Case PrvRenTyp
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                'Ticket #16858 - Do not change the Renewal Date
                                'rsHRTrain("TR_RENEW") = DateAdd(xDWMY, PrvRen, CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                            Else
                                'No Previous Renewal Period
                                flgRenewalPeriod = False
                                
                                'Delete Training List record, and update Follow Up and Continuing Education records
                                GoTo Delete_Training_Record
                            End If
                        End If
                        
                        'Simply update rest of the fields with Position information
                        If IsMissing(xJobCode) Then
                            rsHRTrain("TR_JOB") = glbPos
                        Else
                            rsHRTrain("TR_JOB") = xJobCode
                        End If
                        rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                        If rsEmpJob("TW_CURRENT") And rsEmpJob("JOBTYPE") = "C" Then
                            rsHRTrain("TR_POS_TYPE") = "C"
                        ElseIf rsEmpJob("TW_CURRENT") And rsEmpJob("JOBTYPE") = "T" Then
                            rsHRTrain("TR_POS_TYPE") = "T"
                        ElseIf IIf(IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")), False, rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                            rsHRTrain("TR_POS_TYPE") = "P"
                        End If
                        'rsHRTrain("TR_COURSE_TAKEN")   - Remains BLANK
                        rsHRTrain("TR_LDATE") = Date
                        rsHRTrain("TR_LTIME") = Time$
                        rsHRTrain("TR_LUSER") = glbUserID
                        rsHRTrain.Update
                        
                        'Update Follow Up record
                        If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                            SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsFollowUp.EOF Then
                                If IsMissing(xJobCode) Then
                                    rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & glbPos
                                Else
                                    rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & xJobCode
                                End If
                                'Ticket #16858 - Do not change the Renewal Date
                                'rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")  'new renewal date
                                rsFollowUp("EF_LDATE") = Date
                                rsFollowUp("EF_LUSER") = glbUserID
                                rsFollowUp("EF_LTIME") = Time$
                                rsFollowUp.Update
                            End If
                            rsFollowUp.Close
                            Set rsFollowUp = Nothing
                        End If
                        
                        'Update Continuing Education record
                        SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                        SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                        SQLQ = SQLQ & " AND ES_JOB IS NULL"
                        SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                        SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDate)   'Retrieve record with old renewal date to update with new date
                        SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
                        rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsContEdu.EOF Then
                            If IsMissing(xJobCode) Then
                                rsContEdu("ES_JOB") = glbPos     'Job Code
                            Else
                                rsContEdu("ES_JOB") = xJobCode   'Job Code
                            End If
                            'Ticket #16858 - Do not change the Renewal Date
                            'rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                            rsContEdu("ES_LDATE") = Date
                            rsContEdu("ES_LUSER") = glbUserID
                            rsContEdu("ES_LTIME") = Time$
                            rsContEdu.Update
                        End If
                        rsContEdu.Close
                        Set rsContEdu = Nothing
                    Else
                        oJob = rsHRTrain("TR_JOB")
                        'COURSE WITH JOB INFORMATION - CURRENT JOB
                        If rsEmpJob("TW_CURRENT") And rsEmpJob("JOBTYPE") = "C" Then
                            'Course being added for Current Position
                            'What Type of Position is assigned to this Training List record
                            If rsHRTrain("TR_POS_TYPE") = "C" Then
                                'Current Position Type
                                'This cannot happen, the required course just being added for Current Position
                                'so this course should not have been existed from before.
                            ElseIf rsHRTrain("TR_POS_TYPE") = "T" Then
                                'Temporary Position Type
                                'This will compute the same date because the Course has been taken and
                                'it's using the same renewal period = Current Renewal Period
                                'Though the Job information will change
                                
                                'Simply update rest of the fields with this Position information
                                If IsMissing(xJobCode) Then
                                    rsHRTrain("TR_JOB") = glbPos
                                Else
                                    rsHRTrain("TR_JOB") = xJobCode
                                End If
                                rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                                rsHRTrain("TR_POS_TYPE") = "C"
                                'rsHRTrain("TR_COURSE_TAKEN")   - does not change
                                rsHRTrain("TR_LDATE") = Date
                                rsHRTrain("TR_LTIME") = Time$
                                rsHRTrain("TR_LUSER") = glbUserID
                                rsHRTrain.Update
                                
                                'Update Follow Up record
                                If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        If IsMissing(xJobCode) Then
                                            rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & glbPos
                                        Else
                                            rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & xJobCode
                                        End If
                                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")  'new renewal date
                                        rsFollowUp("EF_LDATE") = Date
                                        rsFollowUp("EF_LUSER") = glbUserID
                                        rsFollowUp("EF_LTIME") = Time$
                                        rsFollowUp.Update
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Update Continuing Education record
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                                SQLQ = SQLQ & " AND ES_JOB = '" & oJob & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDate)   'Retrieve record with old renewal date to update with new date
                                SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    If IsMissing(xJobCode) Then
                                        rsContEdu("ES_JOB") = glbPos     'Job Code
                                    Else
                                        rsContEdu("ES_JOB") = xJobCode   'Job Code
                                    End If
                                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                            
                            ElseIf rsHRTrain("TR_POS_TYPE") = "P" Then
                                'Previous Position Type
                                If Not IsNull(CurRen) And CurRen <> 0 Then
                                    'Current Renewal Period found
                                    'Calculate Renewal Date based on the Renewal Period and Course Taken Date
                                    Select Case CurRenTyp
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, CurRen, CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                    
                                    'Simply update rest of the fields with this Position information
                                    If IsMissing(xJobCode) Then
                                        rsHRTrain("TR_JOB") = glbPos
                                    Else
                                        rsHRTrain("TR_JOB") = xJobCode
                                    End If
                                    rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                                    rsHRTrain("TR_POS_TYPE") = "C"
                                    'rsHRTrain("TR_COURSE_TAKEN")   - does not change
                                    rsHRTrain("TR_LDATE") = Date
                                    rsHRTrain("TR_LTIME") = Time$
                                    rsHRTrain("TR_LUSER") = glbUserID
                                    rsHRTrain.Update
                                    
                                    'Update Follow Up record
                                    If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                        SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsFollowUp.EOF Then
                                            If IsMissing(xJobCode) Then
                                                rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & glbPos
                                            Else
                                                rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & xJobCode
                                            End If
                                            rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")  'new renewal date
                                            rsFollowUp("EF_LDATE") = Date
                                            rsFollowUp("EF_LUSER") = glbUserID
                                            rsFollowUp("EF_LTIME") = Time$
                                            rsFollowUp.Update
                                        End If
                                        rsFollowUp.Close
                                        Set rsFollowUp = Nothing
                                    End If
                                    
                                    'Update Continuing Education record
                                    SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                    SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                                    SQLQ = SQLQ & " AND ES_JOB = '" & oJob & "'"
                                    SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                                    SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDate)   'Retrieve record with old renewal date to update with new date
                                    SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
                                    rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsContEdu.EOF Then
                                        If IsMissing(xJobCode) Then
                                            rsContEdu("ES_JOB") = glbPos     'Job Code
                                        Else
                                            rsContEdu("ES_JOB") = xJobCode   'Job Code
                                        End If
                                        rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                                        rsContEdu("ES_LDATE") = Date
                                        rsContEdu("ES_LUSER") = glbUserID
                                        rsContEdu("ES_LTIME") = Time$
                                        rsContEdu.Update
                                    End If
                                    rsContEdu.Close
                                    Set rsContEdu = Nothing
                                    
                                    
                                End If
                            End If
                        ElseIf rsEmpJob("TW_CURRENT") And rsEmpJob("JOBTYPE") = "T" Then
                            'COURSE WITH JOB INFORMATION - TEMPORARY JOB
                            'Course being added for Temporary Position
                            'What Type of Position is assigned to this Training List record
                            If rsHRTrain("TR_POS_TYPE") = "C" Then
                                'Current Position Type
                                'Do not do anything - Current Position takes the precedence
                            ElseIf rsHRTrain("TR_POS_TYPE") = "T" Then
                                'Temporary Position Type
                                'This cannot happen, the required course just being added for Temp. Position
                                'so this course should not have been existed from before.
                            ElseIf rsHRTrain("TR_POS_TYPE") = "P" Then
                                'Previous Position Type
                                If Not IsNull(CurRen) And CurRen <> 0 Then
                                    'Current Renewal Period found
                                    'Calculate Renewal Date based on the Renewal Period and Course Taken Date
                                    Select Case CurRenTyp
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, CurRen, CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                    
                                    'Simply update rest of the fields with this Position information
                                    If IsMissing(xJobCode) Then
                                        rsHRTrain("TR_JOB") = glbPos
                                    Else
                                        rsHRTrain("TR_JOB") = xJobCode
                                    End If
                                    rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                                    rsHRTrain("TR_POS_TYPE") = "T"
                                    'rsHRTrain("TR_COURSE_TAKEN")   - does not change
                                    rsHRTrain("TR_LDATE") = Date
                                    rsHRTrain("TR_LTIME") = Time$
                                    rsHRTrain("TR_LUSER") = glbUserID
                                    rsHRTrain.Update
                                    
                                    'Update Follow Up record
                                    If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                        SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsFollowUp.EOF Then
                                            If IsMissing(xJobCode) Then
                                                rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & glbPos
                                            Else
                                                rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & xJobCode
                                            End If
                                            rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")  'new renewal date
                                            rsFollowUp("EF_LDATE") = Date
                                            rsFollowUp("EF_LUSER") = glbUserID
                                            rsFollowUp("EF_LTIME") = Time$
                                            rsFollowUp.Update
                                        End If
                                        rsFollowUp.Close
                                        Set rsFollowUp = Nothing
                                    End If
                                    
                                    'Update Continuing Education record
                                    SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                    SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                                    SQLQ = SQLQ & " AND ES_JOB = '" & oJob & "'"
                                    SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                                    SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDate)   'Retrieve record with old renewal date to update with new date
                                    SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
                                    rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsContEdu.EOF Then
                                        If IsMissing(xJobCode) Then
                                            rsContEdu("ES_JOB") = glbPos     'Job Code
                                        Else
                                            rsContEdu("ES_JOB") = xJobCode   'Job Code
                                        End If
                                        rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                                        rsContEdu("ES_LDATE") = Date
                                        rsContEdu("ES_LUSER") = glbUserID
                                        rsContEdu("ES_LTIME") = Time$
                                        rsContEdu.Update
                                    End If
                                    rsContEdu.Close
                                    Set rsContEdu = Nothing
                                End If
                            End If
                        ElseIf IIf(IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")), False, rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                            'COURSE WITH JOB INFORMATION - PREVIOUS TO Previous
                            'Course being added for Previous Position
                            xPrvEndDate = Get_Position_End_Date(rsHRTrain("TR_JOB"), rsHRTrain("TR_SDATE"))
                            If Not IsDate(xPrvEndDate) Then xPrvEndDate = rsHRTrain("TR_SDATE")
                            'If CVDate(rsHRTrain("TR_SDATE")) < CVDate(rsEmpJob("TW_SDATE")) Then
                            If CVDate(xPrvEndDate) < CVDate(IIf(Not IsDate(xPrvEndDate), rsEmpJob("TW_SDATE"), rsEmpJob("TW_ENDDATE"))) Then
                                'Training List has older Position Start Date so update with new Position info.
                                If Not IsNull(PrvRen) And PrvRen <> 0 Then
                                    'Previous Renewal Period found
                                    'Calculate Renewal Date based on the Renewal Period and Course Taken Date
                                    Select Case PrvRenTyp
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, PrvRen, CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                    
                                    'Simply update rest of the fields with this Position information
                                    If IsMissing(xJobCode) Then
                                        rsHRTrain("TR_JOB") = glbPos
                                    Else
                                        rsHRTrain("TR_JOB") = xJobCode
                                    End If
                                    rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                                    rsHRTrain("TR_POS_TYPE") = "T"
                                    'rsHRTrain("TR_COURSE_TAKEN")   - does not change
                                    rsHRTrain("TR_LDATE") = Date
                                    rsHRTrain("TR_LTIME") = Time$
                                    rsHRTrain("TR_LUSER") = glbUserID
                                    rsHRTrain.Update
                                    
                                    'Update Follow Up record
                                    If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                        SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsFollowUp.EOF Then
                                            If IsMissing(xJobCode) Then
                                                rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & glbPos
                                            Else
                                                rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & xJobCode
                                            End If
                                            rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")  'new renewal date
                                            rsFollowUp("EF_LDATE") = Date
                                            rsFollowUp("EF_LUSER") = glbUserID
                                            rsFollowUp("EF_LTIME") = Time$
                                            rsFollowUp.Update
                                        End If
                                        rsFollowUp.Close
                                        Set rsFollowUp = Nothing
                                    End If
                                    
                                    'Update Continuing Education record
                                    SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                    SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                                    SQLQ = SQLQ & " AND ES_JOB = '" & oJob & "'"
                                    SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                                    SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDate)   'Retrieve record with old renewal date to update with new date
                                    SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
                                    rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsContEdu.EOF Then
                                        If IsMissing(xJobCode) Then
                                            rsContEdu("ES_JOB") = glbPos     'Job Code
                                        Else
                                            rsContEdu("ES_JOB") = xJobCode   'Job Code
                                        End If
                                        rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                                        rsContEdu("ES_LDATE") = Date
                                        rsContEdu("ES_LUSER") = glbUserID
                                        rsContEdu("ES_LTIME") = Time$
                                        rsContEdu.Update
                                    End If
                                    rsContEdu.Close
                                    Set rsContEdu = Nothing
                                    
                                End If
                            End If
                        End If
                    End If
                                        
Delete_Training_Record:
                    If flgRenewalPeriod = False Then
                        'Retrieve Continuing Education record
                        SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_DATCOMP,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                        SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                        If Not IsNull(rsHRTrain("TR_JOB")) Then
                            SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                        End If
                        SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                        SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDate)   'Retrieve record with old renewal date to update with new date
                        SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
                            
                        rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsContEdu.EOF Then
                            rsContEdu("ES_RENEW") = Null
                            rsContEdu("ES_LDATE") = Date
                            rsContEdu("ES_LUSER") = glbUserID
                            rsContEdu("ES_LTIME") = Time$
                            rsContEdu.Update
                            
                            If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                                'Since the course was completed - mark the Follow Up as
                                'Completed instead of deleting it.
                                
                                'If follow up id is null then find the id
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                    SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                                    If Not IsNull(rsHRTrain("TR_JOB")) Then
                                        SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    End If
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsHRTrain("TR_CRSCODE") & "')"
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                gdbAdoIhr001.Execute SQLQ
                            Else
                                'Delete the Follow Up record for this training record
                                'as no Course completion record found
                                
                                'If follow up id is null then find the id
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                    SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                                    If Not IsNull(rsHRTrain("TR_JOB")) Then
                                        SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    End If
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsHRTrain("TR_CRSCODE") & "')"
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                
                                SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                gdbAdoIhr001.Execute SQLQ
                               
                                'Clear the Follow Up Id on the Position record
                                'if the course code is TRAIN
                                If xCourseCode = "TRAIN" Then
                                    'Search HR_JOB_HISTORY and HR_TEMP_WORK table for this Position record
                                    'and update with Follow Up Id
                                    If rsEmpJob("JOBTYPE") = "C" Then
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    ElseIf rsEmpJob("JOBTYPE") = "T" Then
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    End If
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        If rsEmpJob("JOBTYPE") = "C" Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                        ElseIf rsEmpJob("JOBTYPE") = "T" Then
                                            rsTJob("TW_FOLLOWUP_ID") = Null
                                        End If
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            End If
                        Else
                            'Delete the Follow Up record for this training record
                            'as no Course record found
                            
                            'If follow up id is null then find the id
                            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                                SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                                If Not IsNull(rsHRTrain("TR_JOB")) Then
                                    SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                End If
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsHRTrain("TR_CRSCODE") & "')"
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                            
                            SQLQ = "DELETE FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            gdbAdoIhr001.Execute SQLQ
                            
                            'Clear the Follow Up ID in the Position record
                            'if the course code is TRAIN
                            If xCourseCode = "TRAIN" Then
                                'Search HR_JOB_HISTORY and HR_TEMP_WORK table for this Position record
                                'and update with Follow Up Id
                                If rsEmpJob("JOBTYPE") = "C" Then
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                ElseIf rsEmpJob("JOBTYPE") = "T" Then
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                End If
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    If rsEmpJob("JOBTYPE") = "C" Then
                                        rsTJob("JH_FOLLOWUP_ID") = Null
                                    ElseIf rsEmpJob("JOBTYPE") = "T" Then
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                    End If
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                            End If
                        End If
                        rsContEdu.Close
                        Set rsContEdu = Nothing
                            
                        'Delete this Training List record as the course
                        rsHRTrain.Delete
                    End If
                    
                Else
                    'COURSE NOT TAKEN
                    'Check if Independant Course
                    If IsNull(rsHRTrain("TR_JOB")) Then
                        'INDEPENDANT COURSE
                        'Compute the Renewal Date
                        Select Case FolRenTyp
                            Case "D"
                                xDWMY = "d"
                            Case "W"
                                xDWMY = "ww"
                            Case "M"
                                xDWMY = "m"
                            Case "Y"
                                xDWMY = "yyyy"
                        End Select
                        'Ticket #16858 - Do not replace the Renewal Period
                        'rsHRTrain("TR_RENEW") = DateAdd(xDWMY, FolRen, CVDate(rsEmpJob("TW_SDATE")))
                        
                        
                        'Simply update rest of the fields with Position information
                        If IsMissing(xJobCode) Then
                            rsHRTrain("TR_JOB") = glbPos
                        Else
                            rsHRTrain("TR_JOB") = xJobCode
                        End If
                        rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                        If rsEmpJob("TW_CURRENT") And rsEmpJob("JOBTYPE") = "C" Then
                            rsHRTrain("TR_POS_TYPE") = "C"
                        ElseIf rsEmpJob("TW_CURRENT") And rsEmpJob("JOBTYPE") = "T" Then
                            rsHRTrain("TR_POS_TYPE") = "T"
                        ElseIf IIf(IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")), False, rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                            rsHRTrain("TR_POS_TYPE") = "P"
                        End If
                        'rsHRTrain("TR_COURSE_TAKEN")   - Remains BLANK
                        rsHRTrain("TR_LDATE") = Date
                        rsHRTrain("TR_LTIME") = Time$
                        rsHRTrain("TR_LUSER") = glbUserID
                        rsHRTrain.Update
                        
                        'Update Follow Up record
                        If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                            SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsFollowUp.EOF Then
                                rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")  'new renewal date
                                If IsMissing(xJobCode) Then
                                    rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & glbPos
                                Else
                                    rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & xJobCode
                                End If
                                rsFollowUp("EF_LDATE") = Date
                                rsFollowUp("EF_LUSER") = glbUserID
                                rsFollowUp("EF_LTIME") = Time$
                                rsFollowUp.Update
                            End If
                            rsFollowUp.Close
                            Set rsFollowUp = Nothing
                        End If
                    Else
                        'COURSES WITH JOB INFORMATION
                        'Compute the Renewal Date
                        Select Case FolRenTyp
                            Case "D"
                                xDWMY = "d"
                            Case "W"
                                xDWMY = "ww"
                            Case "M"
                                xDWMY = "m"
                            Case "Y"
                                xDWMY = "yyyy"
                        End Select
                        
                        If rsEmpJob("TW_CURRENT") And rsEmpJob("JOBTYPE") = "C" Then
                            If rsHRTrain("TR_POS_TYPE") = "C" Then
                                'This will not happen
                            ElseIf rsHRTrain("TR_POS_TYPE") = "T" Or rsHRTrain("TR_POS_TYPE") = "P" Then
                                'Change to Current Job as it takes the precedence.
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, FolRen, CVDate(rsHRTrain("TR_SDATE")))
                                
                                'Simply update rest of the fields with Position information
                                If IsMissing(xJobCode) Then
                                    rsHRTrain("TR_JOB") = glbPos
                                Else
                                    rsHRTrain("TR_JOB") = xJobCode
                                End If
                                rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                                rsHRTrain("TR_POS_TYPE") = "C"
                                'rsHRTrain("TR_COURSE_TAKEN")   - Remains BLANK
                                rsHRTrain("TR_LDATE") = Date
                                rsHRTrain("TR_LTIME") = Time$
                                rsHRTrain("TR_LUSER") = glbUserID
                                rsHRTrain.Update
                                
                                'Update Follow Up record
                                If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")  'new renewal date
                                        If IsMissing(xJobCode) Then
                                            rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & glbPos
                                        Else
                                            rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & xJobCode
                                        End If
                                        rsFollowUp("EF_LDATE") = Date
                                        rsFollowUp("EF_LUSER") = glbUserID
                                        rsFollowUp("EF_LTIME") = Time$
                                        rsFollowUp.Update
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                            End If
                        ElseIf rsEmpJob("TW_CURRENT") And rsEmpJob("JOBTYPE") = "T" Then
                            If rsHRTrain("TR_POS_TYPE") = "C" Then
                                'Do not do anything, Current Job takes the precedence
                            ElseIf rsHRTrain("TR_POS_TYPE") = "T" Then
                                'This will not happen
                            ElseIf rsHRTrain("TR_POS_TYPE") = "P" Then
                                'Change to Temp Job as it takes the precedence.
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, FolRen, CVDate(rsHRTrain("TR_SDATE")))
                                
                                'Simply update rest of the fields with Position information
                                If IsMissing(xJobCode) Then
                                    rsHRTrain("TR_JOB") = glbPos
                                Else
                                    rsHRTrain("TR_JOB") = xJobCode
                                End If
                                rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                                rsHRTrain("TR_POS_TYPE") = "T"
                                'rsHRTrain("TR_COURSE_TAKEN")   - Remains BLANK
                                rsHRTrain("TR_LDATE") = Date
                                rsHRTrain("TR_LTIME") = Time$
                                rsHRTrain("TR_LUSER") = glbUserID
                                rsHRTrain.Update
                                
                                'Update Follow Up record
                                If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")  'new renewal date
                                        If IsMissing(xJobCode) Then
                                            rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & glbPos
                                        Else
                                            rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & xJobCode
                                        End If
                                        rsFollowUp("EF_LDATE") = Date
                                        rsFollowUp("EF_LUSER") = glbUserID
                                        rsFollowUp("EF_LTIME") = Time$
                                        rsFollowUp.Update
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                            End If
                        ElseIf IIf(IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")), False, rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                            If rsHRTrain("TR_POS_TYPE") = "C" Then
                                'Do not do anything, Current Job takes the precedence
                            ElseIf rsHRTrain("TR_POS_TYPE") = "T" Then
                                'Do not do anything, Temp. Job takes the precedence
                            ElseIf rsHRTrain("TR_POS_TYPE") = "P" Then
                                'Most recent position takes the precedence
                                xPrvEndDate = Get_Position_End_Date(rsHRTrain("TR_JOB"), rsHRTrain("TR_SDATE"))
                                If Not IsDate(xPrvEndDate) Then xPrvEndDate = rsHRTrain("TR_SDATE")
                                'If CVDate(rsHRTrain("TR_SDATE")) < CVDate(rsEmpJob("TW_SDATE")) Then
                                If CVDate(xPrvEndDate) < CVDate(IIf(Not IsDate(xPrvEndDate), rsEmpJob("TW_SDATE"), rsEmpJob("TW_ENDDATE"))) Then
                                    'Training List has older Position Start Date so update with new Position info.
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, FolRen, CVDate(rsHRTrain("TR_SDATE")))
                                    
                                    'Simply update rest of the fields with Position information
                                    If IsMissing(xJobCode) Then
                                        rsHRTrain("TR_JOB") = glbPos
                                    Else
                                        rsHRTrain("TR_JOB") = xJobCode
                                    End If
                                    rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                                    rsHRTrain("TR_POS_TYPE") = "P"
                                    'rsHRTrain("TR_COURSE_TAKEN")   - Remains BLANK
                                    rsHRTrain("TR_LDATE") = Date
                                    rsHRTrain("TR_LTIME") = Time$
                                    rsHRTrain("TR_LUSER") = glbUserID
                                    rsHRTrain.Update
                                    
                                    'Update Follow Up record
                                    If Not IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                        SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsFollowUp.EOF Then
                                            rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")  'new renewal date
                                            If IsMissing(xJobCode) Then
                                                rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & glbPos
                                            Else
                                                rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & xJobCode
                                            End If
                                            rsFollowUp("EF_LDATE") = Date
                                            rsFollowUp("EF_LUSER") = glbUserID
                                            rsFollowUp("EF_LTIME") = Time$
                                            rsFollowUp.Update
                                        End If
                                        rsFollowUp.Close
                                        Set rsFollowUp = Nothing
                                    End If
                                    
                                Else
                                    'Do not do anything
                                End If
                            End If
                        End If
                        
                    End If
                End If
            End If
                        
            rsHRTrain.Close
            Set rsHRTrain = Nothing
            
            rsEmpJob.MoveNext
        Loop

    End If
    rsEmpJob.Close
    Set rsEmpJob = Nothing

    rsCourseMst.Close
    Set rsCourseMst = Nothing

End Sub

Private Sub Course_Renewal_Period_Change(xCourseCode, Optional xJobCode, Optional xCurRen, Optional xCurRenTyp, Optional xPrvRen, Optional xPrvRenTyp, Optional xFolRen, Optional xFolRenTyp, Optional xDept)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsEmpJob As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim SQLQ As String
    Dim flgUnqForPos, flgNoPrvRnwl, flgNoCurRnwl As Boolean
    Dim xDWMY As String
    Dim oRenewalDate As Date
    Dim flgRenewalPeriod As Boolean
    Dim xComments As String

    'Since this course's renewal period has changed, this must be a Unique for each Position course, retrieve
    'corresponding Training List and recompute the Renewal Date, and then update in the Continuing Education
    'and Follow Up records.
    
       
    'Get list of employees with this Position as Current or marked to Track for Course Renewal in
    'HR_JOB_HISTORY and HR_TEMP_WORK tables
    SQLQ = "SELECT 'C' AS JOBTYPE, JH_ID AS TW_ID, JH_EMPNBR AS TW_EMPNBR, JH_JOB AS TW_JOB, JH_SDATE AS TW_SDATE, JH_CURRENT AS TW_CURRENT, JH_ENDDATE AS TW_ENDDATE, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL FROM HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
    If IsMissing(xJobCode) Then
        SQLQ = SQLQ & " AND JH_JOB = '" & glbPos & "'"
    Else
        SQLQ = SQLQ & " AND JH_JOB = '" & xJobCode & "'"
    End If
    
    'Ticket #25609 - Training Plan by Department
    If Not IsMissing(xDept) Then
        SQLQ = SQLQ & " AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_DEPTNO = '" & xDept & "')"
    End If
    
    SQLQ = SQLQ & " UNION "
    SQLQ = SQLQ & " SELECT 'T' AS JOBTYPE, TW_ID, TW_EMPNBR, TW_JOB, TW_SDATE, TW_CURRENT, TW_ENDDATE, TW_TRK_CRS_RENEWAL FROM HR_TEMP_WORK "
    SQLQ = SQLQ & " WHERE ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
    If IsMissing(xJobCode) Then
        SQLQ = SQLQ & " AND TW_JOB = '" & glbPos & "'"
    Else
        SQLQ = SQLQ & " AND TW_JOB = '" & xJobCode & "'"
    End If
    
    'Ticket #25609 - Training Plan by Department
    If Not IsMissing(xDept) Then
        SQLQ = SQLQ & " AND TW_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_DEPTNO = '" & xDept & "')"
    End If
    
    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmpJob.EOF Then
        rsEmpJob.MoveFirst
        
        Do While Not rsEmpJob.EOF
            flgRenewalPeriod = True
            
            'Retrieve Training List record for this Job and Course
            SQLQ = "SELECT * FROM HR_TRAIN"
            SQLQ = SQLQ & " WHERE TR_EMPNBR = " & rsEmpJob("TW_EMPNBR")
            SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourseCode & "'"
            'If flgUnqForPos Then
                If IsMissing(xJobCode) Then
                    SQLQ = SQLQ & " AND TR_JOB = '" & glbPos & "'"
                Else
                    SQLQ = SQLQ & " AND TR_JOB = '" & xJobCode & "'"
                End If

                If rsEmpJob("TW_CURRENT") And rsEmpJob("JOBTYPE") = "C" Then
                    SQLQ = SQLQ & " AND TR_POS_TYPE = 'C'"
                ElseIf rsEmpJob("TW_CURRENT") And rsEmpJob("JOBTYPE") = "T" Then
                    SQLQ = SQLQ & " AND TR_POS_TYPE = 'T'"
                ElseIf IIf(IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")), False, rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                    SQLQ = SQLQ & " AND TR_POS_TYPE = 'P'"
                End If
            'End If
            rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsHRTrain.EOF Then
                'Training List record found
                oRenewalDate = rsHRTrain("TR_RENEW")
                flgRenewalPeriod = True
                
                'Course Taken?
                If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                    'Course Not Taken - Renewal Date based on Follow Up Period
                    Select Case xFolRenTyp
                        Case "D"
                            xDWMY = "d"
                        Case "W"
                            xDWMY = "ww"
                        Case "M"
                            xDWMY = "m"
                        Case "Y"
                            xDWMY = "yyyy"
                    End Select
                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xFolRen, CVDate(rsHRTrain("TR_SDATE")))
                Else
                    'Course Taken - Renewal Date based on the Renewal Period
                    'Check what type of Position it is and see if Renewal Period found for that
                    If rsEmpJob("TW_CURRENT") And (rsEmpJob("JOBTYPE") = "C" Or rsEmpJob("JOBTYPE") = "T") Then
                        'Primary/Temporary Current Position - See if Current Renewal Period found
                        If Not IsNull(xCurRen) And xCurRen <> 0 And xCurRen <> "" Then
                            'Current Renewal Period found
                            'Calculate Renewal Date based on the Renewal Period and Course Taken Date
                            Select Case xCurRenTyp
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xCurRen, CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                        Else
                            'No Current Renewal Period
                            flgRenewalPeriod = False
                            
                            'Delete Training List record, and update Follow Up and Continuing Education records
                            GoTo Delete_Training_Record
                        End If
                    ElseIf IIf(IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")), False, rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                         'Previous Position - See if Previous Renewal Period found
                        If Not IsNull(xPrvRen) And xPrvRen <> 0 And xPrvRen <> "" Then
                            'Previous Renewal Period found
                            'Calculate Renewal Date based on the Renewal Period and Course Taken Date
                            Select Case xPrvRenTyp
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xPrvRen, CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                        Else
                            'No Previous Renewal Period
                            flgRenewalPeriod = False
                            
                            'Delete Training List record, and update Follow Up and Continuing Education records
                            GoTo Delete_Training_Record
                        End If
                    End If
                End If
                rsHRTrain("TR_LDATE") = Date
                rsHRTrain("TR_LUSER") = glbUserID
                rsHRTrain("TR_LTIME") = Time$
                
                
                'If follow up id is null then find the id
                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                    SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & rsHRTrain("TR_EMPNBR") & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsHRTrain("TR_CRSCODE") & "')"
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                End If
                                
                rsHRTrain.Update
                
                'Update Continuing Education record
                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDate)   'Retrieve record with old renewal date to update with new date
                SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsContEdu.EOF Then
                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                    rsContEdu("ES_LDATE") = Date
                    rsContEdu("ES_LUSER") = glbUserID
                    rsContEdu("ES_LTIME") = Time$
                    rsContEdu.Update
                End If
                rsContEdu.Close
                Set rsContEdu = Nothing
                
                'Update Follow Up record
                SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsFollowUp.EOF Then
                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")  'new renewal date
                    rsFollowUp("EF_LDATE") = Date
                    rsFollowUp("EF_LUSER") = glbUserID
                    rsFollowUp("EF_LTIME") = Time$
                    rsFollowUp.Update
                End If
                rsFollowUp.Close
                Set rsFollowUp = Nothing
                
            
Delete_Training_Record:
                If flgRenewalPeriod = False Then
                    'Retrieve Continuing Education record
                    SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_DATCOMP,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                    SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                    SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                    SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                    SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDate)   'Retrieve record with old renewal date to update with new date
                    SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
                        
                    rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsContEdu.EOF Then
                        rsContEdu("ES_RENEW") = Null
                        rsContEdu("ES_LDATE") = Date
                        rsContEdu("ES_LUSER") = glbUserID
                        rsContEdu("ES_LTIME") = Time$
                        rsContEdu.Update
                        
                        If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                            'Since the course was completed - mark the Follow Up as
                            'Completed instead of deleting it.
                            SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            gdbAdoIhr001.Execute SQLQ
                        Else
                            'Delete the Follow Up record for this training record
                            'as no Course completion record found
                            SQLQ = "DELETE FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            gdbAdoIhr001.Execute SQLQ
                           
                            'Clear the Follow Up Id on the Position record
                            'if the course code is TRAIN
                            If xCourseCode = "TRAIN" Then
                                'Search HR_JOB_HISTORY and HR_TEMP_WORK table for this Position record
                                'and update with Follow Up Id
                                If rsEmpJob("JOBTYPE") = "C" Then
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                ElseIf rsEmpJob("JOBTYPE") = "T" Then
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                End If
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    If rsEmpJob("JOBTYPE") = "C" Then
                                        rsTJob("JH_FOLLOWUP_ID") = Null
                                    ElseIf rsEmpJob("JOBTYPE") = "T" Then
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                    End If
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                            End If
                        End If
                    Else
                        'Delete the Follow Up record for this training record
                        'as no Course record found
                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                        gdbAdoIhr001.Execute SQLQ
                        
                        'Clear the Follow Up ID in the Position record
                        'if the course code is TRAIN
                        If xCourseCode = "TRAIN" Then
                            'Search HR_JOB_HISTORY and HR_TEMP_WORK table for this Position record
                            'and update with Follow Up Id
                            If rsEmpJob("JOBTYPE") = "C" Then
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            ElseIf rsEmpJob("JOBTYPE") = "T" Then
                                SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            End If
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                If rsEmpJob("JOBTYPE") = "C" Then
                                    rsTJob("JH_FOLLOWUP_ID") = Null
                                ElseIf rsEmpJob("JOBTYPE") = "T" Then
                                    rsTJob("TW_FOLLOWUP_ID") = Null
                                End If
                                rsTJob.Update
                            End If
                            rsTJob.Close
                            Set rsTJob = Nothing
                        End If
                    End If
                    rsContEdu.Close
                    Set rsContEdu = Nothing
                        
                    'Delete this Training List record as the course
                    rsHRTrain.Delete
                End If
            
            End If
            rsHRTrain.Close
            Set rsHRTrain = Nothing
            
            rsEmpJob.MoveNext
        Loop
    End If
    rsEmpJob.Close
    Set rsEmpJob = Nothing
    
End Sub

Private Function Course_Code_Valid(xCourseCode)
    Dim rsCourseCodeMst As New ADODB.Recordset
    Dim SQLQ As String
    
    Course_Code_Valid = False
    
    SQLQ = "SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER"
    SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & xCourseCode & "'"
    rsCourseCodeMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsCourseCodeMst.EOF Then
        Course_Code_Valid = True
    Else
        Course_Code_Valid = False
    End If
    rsCourseCodeMst.Close
    Set rsCourseCodeMst = Nothing
    
End Function

Private Function Get_Course_Code_Master_Codes()
Dim rsCourses As New ADODB.Recordset
Dim SQLQ As String
Dim xCourses As String

    xCourses = "'*'"
    SQLQ = "SELECT * FROM HR_COURSECODE_MASTER"
    'SQLQ = SQLQ & " WHERE ES_CRSCODE NOT IN (SELECT TR_CRSCODE FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & ")"
    'SQLQ = SQLQ & " AND ES_CRSCODE NOT IN (SELECT PC_CRSCODE FROM HR_JOB_COURSE WHERE PC_JOB IN (SELECT JH_JOB FROM QRY_CROSS_TRAINING_RPT))"
    rsCourses.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsCourses.EOF
        xCourses = xCourses & ",'" & rsCourses("ES_CRSCODE") & "'"
        rsCourses.MoveNext
    Loop
    Get_Course_Code_Master_Codes = xCourses

End Function

Private Sub Add_Training_List_Rec_for_New_Renewal_Period(xCourseCode, Optional xCurRen, Optional xCurRenTyp, Optional xPrvRen, Optional xPrvRenTyp, Optional xFolRen, Optional xFolRenTyp, Optional xDept)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsPosCourse As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsEmpJob As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim SQLQ, xDWMY, oJob As String
    Dim oRenewalDate As Date
    Dim flgChanged, flgCrsTakenBefore As Boolean
    
    'Renewal Period added to this course which was not existing before. Retrieve all the Jobs requiring this
    'course from the Required Courses table and then check which employee has this Job as Current or Tracked.
    'Job list should be ordered as Current, Temporary and Previous (Start Date Descending)
    'For all those jobs, check in the Training List based on the Type of Job - Current/Temp/Previous matching
    'the type of Renewal Period just added, if a Training List exists.
    'if CURRENT RENEWAL PERIOD added:
        'If the Course Taken is Blank then:
        
        '- If Type of Position is Current and the employee Position is Current - Skip to next record
        '- If Type of Position is Current and the employee Position is Temporary - Skip to next record
        '- If Type of Position is Current and the employee Position is Previous - Skip to next record
        'This is because Current Position takes precedence and Follow Up period has been used.
        
        '- If Type of Position is Temporary and the employee Position is Current
            '- change the Training List record Job and Type of Position to this Current Job. And renewal Period
            'based on the Current Job Position Start Date and Follow Up Period
        '- If Type of Position is Temporary and the employee Position is Temporary - Skip to next record
        '- If Type of Position is Temporary and the employee Position is Previous - Skip to next record
        
    
    'Retrieve employees with Job marked as Current only as Current Renewal has changed
    SQLQ = "SELECT 'C' AS JOBTYPE, JH_ID AS TW_ID, JH_EMPNBR AS TW_EMPNBR, JH_JOB AS TW_JOB, JH_SDATE AS TW_SDATE, JH_CURRENT AS TW_CURRENT, JH_ENDDATE AS TW_ENDDATE, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL FROM HR_JOB_HISTORY "
    'SQLQ = SQLQ & " WHERE ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
    SQLQ = SQLQ & " WHERE (JH_CURRENT <> 0)"
    SQLQ = SQLQ & " AND JH_JOB = '" & glbPos & "'"
    
    'Ticket #25609 - Training Plan by Department
    If Not IsMissing(xDept) Then
        SQLQ = SQLQ & " AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_DEPTNO = '" & xDept & "')"
    End If
    
    SQLQ = SQLQ & " UNION "
    SQLQ = SQLQ & " SELECT 'T' AS JOBTYPE, TW_ID, TW_EMPNBR, TW_JOB, TW_SDATE, TW_CURRENT, TW_ENDDATE, TW_TRK_CRS_RENEWAL FROM HR_TEMP_WORK "
    'SQLQ = SQLQ & " WHERE ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
    SQLQ = SQLQ & " WHERE (TW_CURRENT <> 0)"
    SQLQ = SQLQ & " AND TW_JOB = '" & glbPos & "'"
    
    'Ticket #25609 - Training Plan by Department
    If Not IsMissing(xDept) Then
        SQLQ = SQLQ & " AND TW_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_DEPTNO = '" & xDept & "')"
    End If
    
    SQLQ = SQLQ & " ORDER BY TW_EMPNBR, JOBTYPE ASC"
    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmpJob.EOF Then
        rsEmpJob.MoveFirst
        
        Do While Not rsEmpJob.EOF
            'Check in the Training List if this course exists
            SQLQ = "SELECT * FROM HR_TRAIN"
            SQLQ = SQLQ & " WHERE TR_EMPNBR = " & rsEmpJob("TW_EMPNBR")
            SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourseCode & "'"
            SQLQ = SQLQ & " AND TR_JOB = '" & glbPos & "'"
            rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsHRTrain.EOF Then
                'Retain the original date
                oRenewalDate = rsHRTrain("TR_RENEW")
                oJob = rsHRTrain("TR_JOB")
                flgChanged = False
                
                'For PRIMACY or TEMPORARY Current type of Jobs
                If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                    'Course Not Taken
                    If rsHRTrain("TR_POS_TYPE") <> "C" Then
                        'Course had not been taken and it's not a Current Type Training List record,
                        'reset the Renewal Date
                        Select Case xFolRenTyp
                            Case "D"
                                xDWMY = "d"
                            Case "W"
                                xDWMY = "ww"
                            Case "M"
                                xDWMY = "m"
                            Case "Y"
                                xDWMY = "yyyy"
                        End Select
                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xFolRen, CVDate(rsEmpJob("TW_SDATE")))
                        flgChanged = True
                    End If
                Else
                    'Course Taken
                    If rsHRTrain("TR_POS_TYPE") <> "C" Then
                        'Course Taken by another Job Type - Current takes the precedence
                        'Recompute the Renewal Date for Current Job
                        If xCurRen <> "" Then
                            Select Case xCurRenTyp
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xCurRen, CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                            flgChanged = True
                        Else
                            flgChanged = False
                        End If
                    End If
                End If
                
                If flgChanged = True Then
                    'Change took place - update the rest of the fields and table
                    'rsHRTrain("TR_JOB") = rsEmpJob("TW_JOB")
                    rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                    If (rsEmpJob("JOBTYPE") = "C") Then
                        rsHRTrain("TR_POS_TYPE") = "C"
                    ElseIf (rsEmpJob("JOBTYPE") = "T") Then
                        rsHRTrain("TR_POS_TYPE") = "T"
                    End If
                    rsHRTrain("TR_LDATE") = Date
                    rsHRTrain("TR_LUSER") = glbUserID
                    rsHRTrain("TR_LTIME") = Time$
                    rsHRTrain.Update
                    
                    'Update Continuing Education record
                    If Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                        'if Course Taken
                        SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                        SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                        SQLQ = SQLQ & " AND ES_JOB = '" & glbPos & "'"
                        SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                        SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDate)   'Retrieve record with old renewal date to update with new date
                        SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
                        
                        rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsContEdu.EOF Then
                            rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                            'rsContEdu("ES_JOB") = rsEmpJob("TW_JOB")
                            rsContEdu("ES_LDATE") = Date
                            rsContEdu("ES_LUSER") = glbUserID
                            rsContEdu("ES_LTIME") = Time$
                            rsContEdu.Update
                        End If
                        rsContEdu.Close
                        Set rsContEdu = Nothing
                    End If
                    
                    'Update Follow Up record
                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")  'new renewal date
                        'rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & rsEmpJob("TW_JOB")
                        rsFollowUp("EF_LDATE") = Date
                        rsFollowUp("EF_LUSER") = glbUserID
                        rsFollowUp("EF_LTIME") = Time$
                        rsFollowUp.Update
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                
                    'Update Position record with Follow Up ID
                    'if the course code is TRAIN
                    If xCourseCode = "TRAIN" Then
                        'Clear the Follow Up ID from the older job record
                        If (rsEmpJob("JOBTYPE") = "C") Then
                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                        Else
                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                        End If
                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsTJob.EOF Then
                            If (rsEmpJob("JOBTYPE") = "C") Then
                                rsTJob("JH_FOLLOWUP_ID") = Null
                            Else
                                rsTJob("TW_FOLLOWUP_ID") = Null
                            End If
                            rsTJob.Update
                        End If
                        rsTJob.Close
                        Set rsTJob = Nothing
                        
                        'Search HR_JOB_HISTORY or HR_TEMP_WORK table for this Position record
                        'and update with Follow Up Id
                        If (rsEmpJob("JOBTYPE") = "C") Then
                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & rsEmpJob("TW_ID")
                        Else
                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & rsEmpJob("TW_ID")
                        End If
                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsTJob.EOF Then
                            If (rsEmpJob("JOBTYPE") = "C") Then
                                rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                            Else
                                rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                            End If
                            rsTJob.Update
                        End If
                        rsTJob.Close
                        Set rsTJob = Nothing
                    End If
                End If
            Else
                'No Training List records found for this Job
                'Add Training List record
                flgCrsTakenBefore = False
                
                rsHRTrain.AddNew
                rsHRTrain("TR_COMPNO") = "001"
                rsHRTrain("TR_EMPNBR") = rsEmpJob("TW_EMPNBR")
                rsHRTrain("TR_CRSCODE") = xCourseCode
                
                'Check first if this Course was taken before in the Continuing Education screen
                flgCrsTakenBefore = False
                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_JOB, ES_DATCOMP,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsEmpJob("TW_EMPNBR")
                SQLQ = SQLQ & " AND ES_JOB = '" & glbPos & "'"
                SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                SQLQ = SQLQ & " AND (ES_RENEW = '' OR ES_RENEW IS NULL)"
                SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
                SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsContEdu.EOF Then
                    'Course Taken Before
                    rsContEdu.MoveFirst
                    flgCrsTakenBefore = True
                Else
                    'Course not taken before
                    flgCrsTakenBefore = False
                End If
                
                If flgCrsTakenBefore = True Then
                    'Course Taken Before
                    'Compute the Renewal Date based on last Course Taken Date and Current Renewal Period
                    If xCurRen <> "" Then
                        Select Case xCurRenTyp
                            Case "D"
                                xDWMY = "d"
                            Case "W"
                                xDWMY = "ww"
                            Case "M"
                                xDWMY = "m"
                            Case "Y"
                                xDWMY = "yyyy"
                        End Select
                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xCurRen, CVDate(rsContEdu("ES_DATCOMP")))
                        rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                        
                        'Update Continuing Education record as well
                        rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                        'rsContEdu("ES_JOB") = rsEmpJob("TW_JOB")
                        rsContEdu("ES_LDATE") = Date
                        rsContEdu("ES_LUSER") = glbUserID
                        rsContEdu("ES_LTIME") = Time$
                        rsContEdu.Update
                    End If
                Else
                    'Course Not Taken
                    'Compute Renewal Date based on Follow Up Renewal Period
                    Select Case xFolRenTyp
                        Case "D"
                            xDWMY = "d"
                        Case "W"
                            xDWMY = "ww"
                        Case "M"
                            xDWMY = "m"
                        Case "Y"
                            xDWMY = "yyyy"
                    End Select
                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xFolRen, CVDate(rsEmpJob("TW_SDATE")))
                End If
                
                rsHRTrain("TR_JOB") = rsEmpJob("TW_JOB")
                rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                If (rsEmpJob("JOBTYPE") = "C") Then
                    rsHRTrain("TR_POS_TYPE") = "C"
                ElseIf (rsEmpJob("JOBTYPE") = "T") Then
                    rsHRTrain("TR_POS_TYPE") = "T"
                End If
                rsHRTrain("TR_LDATE") = Date
                rsHRTrain("TR_LTIME") = Time$
                rsHRTrain("TR_LUSER") = glbUserID
                
                'Add a Follow Up record for this Training course
                SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE 1 = 2"
                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                rsFollowUp.AddNew
                rsFollowUp("EF_COMPNO") = "001"
                rsFollowUp("EF_EMPNBR") = rsEmpJob("TW_EMPNBR")
                rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                rsFollowUp("EF_FREAS_TABL") = "FURE"
                'Ticket #24257 - Do not update Admin By for them only
                If glbCompSerial <> "S/N - 2262W" Then
                    rsFollowUp("EF_ADMINBY_TABL") = "EDAB"
                    rsFollowUp("EF_ADMINBY") = GetEmpData(rsEmpJob("TW_EMPNBR"), "ED_ADMINBY", Null)
                End If
                rsFollowUp("EF_FREAS") = "EDUC"
                rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & rsEmpJob("TW_JOB")
                rsFollowUp("EF_LDATE") = Date
                rsFollowUp("EF_LTIME") = Time$
                rsFollowUp("EF_LUSER") = glbUserID
                rsFollowUp.Update
                
                rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                rsHRTrain.Update
                
                rsFollowUp.Close
                Set rsFollowUp = Nothing
            
                'Update Position record with Follow Up ID
                'if the course code is TRAIN
                If xCourseCode = "TRAIN" Then
                    'Search HR_JOB_HISTORY or HR_TEMP_WORK table for this Position record
                    'and update with Follow Up Id
                    If (rsEmpJob("JOBTYPE") = "C") Then
                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & rsEmpJob("TW_ID")
                    Else
                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & rsEmpJob("TW_ID")
                    End If
                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsTJob.EOF Then
                        If (rsEmpJob("JOBTYPE") = "C") Then
                            rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                        Else
                            rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                        End If
                        rsTJob.Update
                    End If
                    rsTJob.Close
                    Set rsTJob = Nothing
                End If
                
                rsContEdu.Close
                Set rsContEdu = Nothing
            End If
            rsHRTrain.Close
            Set rsHRTrain = Nothing
            
            rsEmpJob.MoveNext
        Loop
    End If
    rsEmpJob.Close
    Set rsEmpJob = Nothing
            
End Sub

Private Sub Add_Training_List_Rec_for_New_Prv_Renewal_Period(xCourseCode, Optional xCurRen, Optional xCurRenTyp, Optional xPrvRen, Optional xPrvRenTyp, Optional xFolRen, Optional xFolRenTyp)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsPosCourse As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsEmpJob As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim SQLQ, xDWMY, oJob As String
    Dim oRenewalDate, lstEndDate As Date
    Dim flgChanged, flgCrsTakenBefore As Boolean
    Dim lstEmpNo As Integer
    
    'Renewal Period added to this course which was not existing before. Retrieve all the Jobs requiring this
    'course from the Required Courses table and then check which employee has this Job as Current or Tracked.
    'Job list should be ordered as Current, Temporary and Previous (Start Date Descending)
    'For all those jobs, check in the Training List based on the Type of Job - Current/Temp/Previous matching
    'the type of Renewal Period just added, if a Training List exists.
    'if PREVIOUS RENEWAL PERIOD added:
        'If the Course Taken is Blank then:
                
    
    'Retrieve employees with Job marked as Current only as Current Renewal has changed
    SQLQ = "SELECT 'C' AS JOBTYPE, JH_ID AS TW_ID, JH_EMPNBR AS TW_EMPNBR, JH_JOB AS TW_JOB, JH_SDATE AS TW_SDATE, JH_CURRENT AS TW_CURRENT, JH_ENDDATE AS TW_ENDDATE, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL FROM HR_JOB_HISTORY "
    'SQLQ = SQLQ & " WHERE ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
    SQLQ = SQLQ & " WHERE (JH_TRK_CRS_RENEWAL <> 0)"
    SQLQ = SQLQ & " AND JH_JOB = '" & glbPos & "'"
    SQLQ = SQLQ & " UNION "
    SQLQ = SQLQ & " SELECT 'T' AS JOBTYPE, TW_ID, TW_EMPNBR, TW_JOB, TW_SDATE, TW_CURRENT, TW_ENDDATE, TW_TRK_CRS_RENEWAL FROM HR_TEMP_WORK "
    'SQLQ = SQLQ & " WHERE ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
    SQLQ = SQLQ & " WHERE (TW_TRK_CRS_RENEWAL <> 0)"
    SQLQ = SQLQ & " AND TW_JOB = '" & glbPos & "'"
    SQLQ = SQLQ & " ORDER BY TW_EMPNBR, TW_ENDDATE DESC"
    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmpJob.EOF Then
        rsEmpJob.MoveFirst
        
        Do While Not rsEmpJob.EOF
            'Check in the Training List if this course exists
            SQLQ = "SELECT * FROM HR_TRAIN"
            SQLQ = SQLQ & " WHERE TR_EMPNBR = " & rsEmpJob("TW_EMPNBR")
            SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourseCode & "'"
            SQLQ = SQLQ & " AND TR_JOB = '" & glbPos & "'"
            rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsHRTrain.EOF Then
                'Retain the original date
                oRenewalDate = rsHRTrain("TR_RENEW")
                oJob = rsHRTrain("TR_JOB")
                flgChanged = False
                
                'Last record
                If (lstEmpNo <> rsEmpJob("TW_EMPNBR")) Or (lstEmpNo <> rsEmpJob("TW_EMPNBR") And lstEndDate <> rsEmpJob("TW_ENDDATE")) Then
                    lstEmpNo = rsEmpJob("TW_EMPNBR")
                    lstEndDate = rsEmpJob("TW_ENDDATE")
                End If
                
                'For Previous type of Jobs
                If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                    'Course Not Taken
                    If rsHRTrain("TR_POS_TYPE") <> "C" And rsHRTrain("TR_POS_TYPE") <> "T" Then
                        If (lstEmpNo <> rsEmpJob("TW_EMPNBR")) Or (lstEmpNo <> rsEmpJob("TW_EMPNBR") And lstEndDate <> rsEmpJob("TW_ENDDATE")) Then
                            'Course had not been taken and it's not a Current/Temp Type Training List record,
                            'reset the Renewal Date
                            Select Case xFolRenTyp
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xFolRen, CVDate(rsEmpJob("TW_SDATE")))
                            flgChanged = True
                        Else
                            flgChanged = False
                        End If
                    End If
                Else
                    'Course Taken
                    If rsHRTrain("TR_POS_TYPE") <> "C" And rsHRTrain("TR_POS_TYPE") <> "T" Then
                        If (lstEmpNo <> rsEmpJob("TW_EMPNBR")) Or (lstEmpNo <> rsEmpJob("TW_EMPNBR") And lstEndDate <> rsEmpJob("TW_ENDDATE")) Then
                            'Course Taken by another Job Type - Current/Temp takes the precedence
                            'Recompute the Renewal Date for Previous Job
                            If xPrvRen <> "" Then
                                Select Case xPrvRenTyp
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xPrvRen, CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                flgChanged = True
                            Else
                                flgChanged = False
                            End If
                        Else
                            flgChanged = False
                        End If
                    End If
                End If
                
                If flgChanged = True Then
                    'Change took place - update the rest of the fields and table
                    'rsHRTrain("TR_JOB") = rsEmpJob("TW_JOB")
                    rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                    rsHRTrain("TR_POS_TYPE") = "P"
                    rsHRTrain("TR_LDATE") = Date
                    rsHRTrain("TR_LUSER") = glbUserID
                    rsHRTrain("TR_LTIME") = Time$
                    rsHRTrain.Update
                    
                    'Update Continuing Education record
                    If Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                        'if Course Taken
                        SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                        SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsHRTrain("TR_EMPNBR")
                        SQLQ = SQLQ & " AND ES_JOB = '" & glbPos & "'"
                        SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                        SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(oRenewalDate)   'Retrieve record with old renewal date to update with new date
                        SQLQ = SQLQ & " AND ES_DATCOMP = " & Date_SQL(rsHRTrain("TR_COURSE_TAKEN"))
                        
                        rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsContEdu.EOF Then
                            rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                            'rsContEdu("ES_JOB") = rsEmpJob("TW_JOB")
                            rsContEdu("ES_LDATE") = Date
                            rsContEdu("ES_LUSER") = glbUserID
                            rsContEdu("ES_LTIME") = Time$
                            rsContEdu.Update
                        End If
                        rsContEdu.Close
                        Set rsContEdu = Nothing
                    End If
                    
                    'Update Follow Up record
                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsFollowUp.EOF Then
                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")  'new renewal date
                        'rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & rsEmpJob("TW_JOB")
                        rsFollowUp("EF_LDATE") = Date
                        rsFollowUp("EF_LUSER") = glbUserID
                        rsFollowUp("EF_LTIME") = Time$
                        rsFollowUp.Update
                    End If
                    rsFollowUp.Close
                    Set rsFollowUp = Nothing
                
                    'Update Position record with Follow Up ID
                    'if the course code is TRAIN
                    If xCourseCode = "TRAIN" Then
                        'Clear the Follow Up ID from the older job record
                        If (rsEmpJob("JOBTYPE") = "C") Then
                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                        Else
                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                        End If
                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsTJob.EOF Then
                            If (rsEmpJob("JOBTYPE") = "C") Then
                                rsTJob("JH_FOLLOWUP_ID") = Null
                            Else
                                rsTJob("TW_FOLLOWUP_ID") = Null
                            End If
                            rsTJob.Update
                        End If
                        rsTJob.Close
                        Set rsTJob = Nothing
                        
                        'Search HR_JOB_HISTORY or HR_TEMP_WORK table for this Position record
                        'and update with Follow Up Id
                        If (rsEmpJob("JOBTYPE") = "C") Then
                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & rsEmpJob("TW_ID")
                        Else
                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & rsEmpJob("TW_ID")
                        End If
                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsTJob.EOF Then
                            If (rsEmpJob("JOBTYPE") = "C") Then
                                rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                            Else
                                rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                            End If
                            rsTJob.Update
                        End If
                        rsTJob.Close
                        Set rsTJob = Nothing
                    End If
                End If
            Else
                'No Training List records found for this Job
                'Add Training List record
                flgCrsTakenBefore = False
                
                rsHRTrain.AddNew
                rsHRTrain("TR_COMPNO") = "001"
                rsHRTrain("TR_EMPNBR") = rsEmpJob("TW_EMPNBR")
                rsHRTrain("TR_CRSCODE") = xCourseCode
                
                'Check first if this Course was taken before in the Continuing Education screen
                flgCrsTakenBefore = False
                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_JOB, ES_DATCOMP,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsEmpJob("TW_EMPNBR")
                SQLQ = SQLQ & " AND ES_JOB = '" & glbPos & "'"
                SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourseCode & "'"
                SQLQ = SQLQ & " AND (ES_RENEW = '' OR ES_RENEW IS NULL)"
                SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
                SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsContEdu.EOF Then
                    'Course Taken Before
                    rsContEdu.MoveFirst
                    flgCrsTakenBefore = True
                Else
                    'Course not taken before
                    flgCrsTakenBefore = False
                End If
                
                If flgCrsTakenBefore = True Then
                    'Course Taken Before
                    'Compute the Renewal Date based on last Course Taken Date and Current Renewal Period
                    If xPrvRen <> "" Then
                        Select Case xPrvRenTyp
                            Case "D"
                                xDWMY = "d"
                            Case "W"
                                xDWMY = "ww"
                            Case "M"
                                xDWMY = "m"
                            Case "Y"
                                xDWMY = "yyyy"
                        End Select
                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xPrvRen, CVDate(rsContEdu("ES_DATCOMP")))
                        rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                        
                        'Update Continuing Education record as well
                        rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                        'rsContEdu("ES_JOB") = rsEmpJob("TW_JOB")
                        rsContEdu("ES_LDATE") = Date
                        rsContEdu("ES_LUSER") = glbUserID
                        rsContEdu("ES_LTIME") = Time$
                        rsContEdu.Update
                    End If
                Else
                    'Course Not Taken
                    'Compute Renewal Date based on Follow Up Renewal Period
                    Select Case xFolRenTyp
                        Case "D"
                            xDWMY = "d"
                        Case "W"
                            xDWMY = "ww"
                        Case "M"
                            xDWMY = "m"
                        Case "Y"
                            xDWMY = "yyyy"
                    End Select
                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xFolRen, CVDate(rsEmpJob("TW_SDATE")))
                End If
                
                rsHRTrain("TR_JOB") = rsEmpJob("TW_JOB")
                rsHRTrain("TR_SDATE") = rsEmpJob("TW_SDATE")
                rsHRTrain("TR_POS_TYPE") = "P"
                rsHRTrain("TR_LDATE") = Date
                rsHRTrain("TR_LTIME") = Time$
                rsHRTrain("TR_LUSER") = glbUserID
                
                'Add a Follow Up record for this Training course
                SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE 1 = 2"
                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                rsFollowUp.AddNew
                rsFollowUp("EF_COMPNO") = "001"
                rsFollowUp("EF_EMPNBR") = rsEmpJob("TW_EMPNBR")
                rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                rsFollowUp("EF_FREAS_TABL") = "FURE"
                'Ticket #24257 - Do not update Admin By for them only
                If glbCompSerial <> "S/N - 2262W" Then
                    rsFollowUp("EF_ADMINBY_TABL") = "EDAB"
                    rsFollowUp("EF_ADMINBY") = GetEmpData(rsEmpJob("TW_EMPNBR"), "ED_ADMINBY", Null)
                End If
                rsFollowUp("EF_FREAS") = "EDUC"
                rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & rsEmpJob("TW_JOB")
                rsFollowUp("EF_LDATE") = Date
                rsFollowUp("EF_LTIME") = Time$
                rsFollowUp("EF_LUSER") = glbUserID
                rsFollowUp.Update
                
                rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                rsHRTrain.Update
                
                rsFollowUp.Close
                Set rsFollowUp = Nothing
            
                'Update Position record with Follow Up ID
                'if the course code is TRAIN
                If xCourseCode = "TRAIN" Then
                    'Search HR_JOB_HISTORY or HR_TEMP_WORK table for this Position record
                    'and update with Follow Up Id
                    If (rsEmpJob("JOBTYPE") = "C") Then
                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & rsEmpJob("TW_ID")
                    Else
                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & rsEmpJob("TW_ID")
                    End If
                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsTJob.EOF Then
                        If (rsEmpJob("JOBTYPE") = "C") Then
                            rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                        Else
                            rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                        End If
                        rsTJob.Update
                    End If
                    rsTJob.Close
                    Set rsTJob = Nothing
                End If
                
                rsContEdu.Close
                Set rsContEdu = Nothing
            End If
            rsHRTrain.Close
            Set rsHRTrain = Nothing
            
            rsEmpJob.MoveNext
        Loop
    End If
    rsEmpJob.Close
    Set rsEmpJob = Nothing
            
End Sub

Private Function Get_Position_End_Date(xJob, xStartDate)
    Dim rsEmpJob As New ADODB.Recordset
    Dim SQLQ As String
    
    Get_Position_End_Date = ""
    
    SQLQ = "SELECT JH_ID, JH_EMPNBR, JH_SDATE, JH_ENDDATE FROM HR_JOB_HISTORY"
    SQLQ = SQLQ & " WHERE JH_JOB = '" & xJob & "'"
    SQLQ = SQLQ & " AND JH_SDATE = " & Date_SQL(xStartDate)
    SQLQ = SQLQ & " AND JH_TRK_CRS_RENEWAL<>0"
    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmpJob.EOF Then
        Get_Position_End_Date = rsEmpJob("JH_ENDDATE")
    Else
        rsEmpJob.Close
        Set rsEmpJob = Nothing
        SQLQ = "SELECT TW_ID, TW_EMPNBR, TW_SDATE, TW_ENDDATE FROM HR_TEMP_WORK"
        SQLQ = SQLQ & " WHERE TW_JOB = '" & xJob & "'"
        SQLQ = SQLQ & " AND TW_SDATE = " & Date_SQL(xStartDate)
        SQLQ = SQLQ & " AND TW_TRK_CRS_RENEWAL<>0"
        rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmpJob.EOF Then
            Get_Position_End_Date = rsEmpJob("TW_ENDDATE")
        Else
            Get_Position_End_Date = ""
        End If
    End If
    rsEmpJob.Close
    Set rsEmpJob = Nothing
End Function

Private Sub Update_Other_EmpPositions_Require_This_Course(xCourse)
    Dim rsEmpJobs As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsCourseMst As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim SQLQ As String
    Dim xEduRec, xCurRenPrd, xPrvRenPrd, xFlwRenPrd As Integer
    Dim xCurRenTyp, xPrvRenTyp, xFlwRenTyp, xDWMY As String
    Dim flgCrsTakenBefore, flgUnqForPos As Boolean
    

    'Check if the course is unique for each position
    SQLQ = "SELECT ES_CRSCODE,ES_UNIQUE_FOR_POS FROM HR_COURSECODE_MASTER"
    SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & xCourse & "'"
    rsCourseMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsCourseMst.EOF Then
        'Course found
        If Not IsNull(rsCourseMst("ES_UNIQUE_FOR_POS")) Then
            flgUnqForPos = rsCourseMst("ES_UNIQUE_FOR_POS")
        Else
            MsgBox "Please setup the 'Follow Up Effective Date Period' on the Course Code Master screen.", vbExclamation, "Course Code Master Setup missing"
            Exit Sub
        End If

    Else
        flgUnqForPos = False
    End If
    rsCourseMst.Close
    Set rsCourseMst = Nothing
        
    
    'Check which Current or Tracked Position required this Course
    'Get list of Current/Temporary and Tracked Positions of this employee who has this job as Current or Tracked
    SQLQ = "SELECT JH_ID AS TW_ID, JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
    SQLQ = SQLQ & " AND JH_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & xCourse & "')"
    SQLQ = SQLQ & " AND (JH_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_JOB = '" & glbPos & "' AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0)))"
    SQLQ = SQLQ & " OR JH_EMPNBR IN (SELECT TW_EMPNBR FROM HR_TEMP_WORK WHERE TW_JOB = '" & glbPos & "' AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))))"

    'if Unique for each Position record that means each Current/Tracked Job requiring this
    'course will have it's own Training List record - of course depending on the Renewal Period
    'Retrieve only the job assigned to the deleted Course record.
    If flgUnqForPos <> 0 Then
        SQLQ = SQLQ & " AND JH_JOB = '" & glbPos & "'"
    End If
    
    SQLQ = SQLQ & " UNION "
    SQLQ = SQLQ & " SELECT TW_ID, TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
    SQLQ = SQLQ & " WHERE ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
    SQLQ = SQLQ & " AND TW_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & xCourse & "')"
    SQLQ = SQLQ & " AND (TW_EMPNBR IN (SELECT TW_EMPNBR FROM HR_TEMP_WORK WHERE TW_JOB = '" & glbPos & "' AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0)))"
    SQLQ = SQLQ & " OR TW_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_JOB = '" & glbPos & "' AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))))"
    
    'if Unique for each Position record that means each Current/Tracked Job requiring this
    'course will have it's own Training List record - of course depending on the Renewal Period
    'Retrieve only the job assigned to the deleted Course record.
    If flgUnqForPos <> 0 Then
        SQLQ = SQLQ & " AND TW_JOB = '" & glbPos & "'"
    End If
    
    SQLQ = SQLQ & " ORDER BY TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
    rsEmpJobs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmpJobs.EOF Then
        rsEmpJobs.MoveFirst
    
        Do While Not rsEmpJobs.EOF
            'Get the renewal periods of the course
            SQLQ = "SELECT PC_CRSCODE,PC_RENEW_CRS_CUR,PC_CUR_PRD_DWMY,PC_RENEW_CRS_PRV,PC_PRV_PRD_DWMY,PC_RENEW_FOLLOWUP,PC_FLWUP_PRD_DWMY FROM HR_JOB_COURSE "
            SQLQ = SQLQ & " WHERE PC_JOB = '" & rsEmpJobs("TW_JOB") & "'"
            SQLQ = SQLQ & " AND PC_CRSCODE = '" & xCourse & "'"
            rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsReqCourse.EOF Then
                'Course found
                xCurRenPrd = rsReqCourse("PC_RENEW_CRS_CUR")
                xCurRenTyp = rsReqCourse("PC_CUR_PRD_DWMY")
                
                '7.9 - Enhancement - For all clients now
                'Not City of Chatham-Kent - Ticket #16794
                'If glbCompSerial <> "S/N - 2188W" Then
                If glbCompSerial = "S/N - 2279W" Then
                    xPrvRenPrd = rsReqCourse("PC_RENEW_CRS_PRV")
                    xPrvRenTyp = rsReqCourse("PC_PRV_PRD_DWMY")
                End If
                
                xFlwRenPrd = rsReqCourse("PC_RENEW_FOLLOWUP")
                xFlwRenTyp = rsReqCourse("PC_FLWUP_PRD_DWMY")
            End If
            rsReqCourse.Close
            Set rsReqCourse = Nothing
            
            'if Unique for each Position Course check if the Training List existing for this Job
            'already exists - then skip to next Employee Position requiring this course
            'If flgUnqForPos Then
                SQLQ = "SELECT * FROM HR_TRAIN"
                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & rsEmpJobs("TW_EMPNBR")
                If flgUnqForPos <> 0 Then
                    SQLQ = SQLQ & " AND TR_JOB = '" & rsEmpJobs("TW_JOB") & "'"
                End If
                SQLQ = SQLQ & " AND TR_CRSCODE = '" & xCourse & "'"
                rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsHRTrain.EOF Then
                    'Skip to next Employee Job because for this Job, the training list record already
                    'exist for this unique for each position course.
                    GoTo next_EmpPosition
                Else
                    'Continue with the rest of the process
                End If
                rsHRTrain.Close
                Set rsHRTrain = Nothing
            'End If
            
            'Course Taken before?
            If rsContEdu.State <> 0 Then
                rsContEdu.Close
                Set rsContEdu = Nothing
            End If
            SQLQ = "SELECT ES_EMPNBR,ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & rsEmpJobs("TW_EMPNBR")
            SQLQ = SQLQ & " AND ES_CRSCODE = '" & xCourse & "'"
            If flgUnqForPos <> 0 Then
                'Unique for each position course then check if the course was taken for the right position
                SQLQ = SQLQ & " AND ES_JOB = '" & rsEmpJobs("TW_JOB") & "'"
            End If
            SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsContEdu.EOF Then
                'Course Taken before
                rsContEdu.MoveFirst
                flgCrsTakenBefore = True
            Else
                flgCrsTakenBefore = False
            End If
        
            
            SQLQ = "SELECT * FROM HR_TRAIN WHERE 1=2"
            rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            
            'Course Taken before?
            If flgCrsTakenBefore = True Then
                'Course Taken before
                'Compute the Renewal Period for this course and add a Training List and Follow Up record
                If (rsEmpJobs("POS_TYPE") = "CURRENT" Or rsEmpJobs("POS_TYPE") = "TEMPORARY") And rsEmpJobs("TW_CURRENT") Then
                    'Primary Current/Temporary Position
                    'Based on Current Renewal Period if found
                    If IsNull(xCurRenPrd) Or xCurRenPrd = 0 Or xCurRenPrd = "" Then
                        'No Renewal Period found, clear last course taken record's Renewal Date
                        'There won't be Training List record, because there was no Renewal Date on the
                        'deleted Course record.
                        rsContEdu("ES_RENEW") = Null
                        rsContEdu("ES_LDATE") = Date
                        rsContEdu("ES_LUSER") = glbUserID
                        rsContEdu("ES_LTIME") = Time$
                        rsContEdu.Update
                        
                        rsContEdu.Close
                        Set rsContEdu = Nothing
                    
                        'If flgUnqForPos Then
                        '    'Go to next position
                        '    GoTo next_EmpPosition
                        'Else
                            'Exit loop - only the first position gets this course
                            'Exit Do
                            GoTo next_EmpPosition
                       ' End If
                    Else
                        'Compute renewal date
                        Select Case xCurRenTyp
                            Case "D"
                                xDWMY = "d"
                            Case "W"
                                xDWMY = "ww"
                            Case "M"
                                xDWMY = "m"
                            Case "Y"
                                xDWMY = "yyyy"
                        End Select
                        'Add a new Training List record with Renewal Date based on Current Renewal Period and
                        'Course Taken Date
                        rsHRTrain.AddNew
                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xCurRenPrd, CVDate(rsContEdu("ES_DATCOMP")))
                        rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                        
                        'Update Continuing Education record as well
                        rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                        rsContEdu("ES_JOB") = rsEmpJobs("TW_JOB")
                        rsContEdu("ES_LDATE") = Date
                        rsContEdu("ES_LUSER") = glbUserID
                        rsContEdu("ES_LTIME") = Time$
                        rsContEdu.Update
                        
                        rsContEdu.Close
                        Set rsContEdu = Nothing
                    End If
                Else
                    'Previous position - Course Taken Before
                    'Based on Previous Renewal period if found
                    If IsNull(xPrvRenPrd) Or xPrvRenPrd = 0 Or xPrvRenPrd = "" Then
                        'No Renewal Period found, clear last course taken record's Renewal Date
                        'There won't be Training List record, because there was no Renewal Date on the
                        'deleted Course record.
                        rsContEdu("ES_RENEW") = Null
                        rsContEdu("ES_LDATE") = Date
                        rsContEdu("ES_LUSER") = glbUserID
                        rsContEdu("ES_LTIME") = Time$
                        rsContEdu.Update
                        
                        rsContEdu.Close
                        Set rsContEdu = Nothing
                        
                        'Exit Do
                        GoTo next_EmpPosition
                    Else
                        'Compute renewal date
                        Select Case xPrvRenTyp
                            Case "D"
                                xDWMY = "d"
                            Case "W"
                                xDWMY = "ww"
                            Case "M"
                                xDWMY = "m"
                            Case "Y"
                                xDWMY = "yyyy"
                        End Select
                        'Add a new Training List record with Renewal Date based on Prev Renewal Period
                        'Course Taken Date
                        rsHRTrain.AddNew
                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xPrvRenPrd, CVDate(rsContEdu("ES_DATCOMP")))
                        rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                        
                        'Update Continuing Education record as well
                        rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")   'new renewal date
                        rsContEdu("ES_JOB") = rsEmpJobs("TW_JOB")
                        rsContEdu("ES_LDATE") = Date
                        rsContEdu("ES_LUSER") = glbUserID
                        rsContEdu("ES_LTIME") = Time$
                        rsContEdu.Update
                        
                        rsContEdu.Close
                        Set rsContEdu = Nothing
                    End If
                End If
            Else
                'Course not taken before
                'Compute renewal date based on Follow Up Period
                Select Case xFlwRenTyp
                    Case "D"
                        xDWMY = "d"
                    Case "W"
                        xDWMY = "ww"
                    Case "M"
                        xDWMY = "m"
                    Case "Y"
                        xDWMY = "yyyy"
                End Select
                'Add a new Training List record with Renewal Date based on Follow Up Period
                rsHRTrain.AddNew
                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, xFlwRenPrd, CVDate(rsEmpJobs("TW_SDATE")))
            End If
            
            rsHRTrain("TR_COMPNO") = "001"
            rsHRTrain("TR_EMPNBR") = rsEmpJobs("TW_EMPNBR")
            rsHRTrain("TR_CRSCODE") = xCourse
            
            rsHRTrain("TR_JOB") = rsEmpJobs("TW_JOB")
            rsHRTrain("TR_SDATE") = rsEmpJobs("TW_SDATE")
            If (rsEmpJobs("POS_TYPE") = "CURRENT") And rsEmpJobs("TW_CURRENT") Then
                rsHRTrain("TR_POS_TYPE") = "C"
            ElseIf (rsEmpJobs("POS_TYPE") = "TEMPORARY") And rsEmpJobs("TW_CURRENT") Then
                rsHRTrain("TR_POS_TYPE") = "T"
            Else
                rsHRTrain("TR_POS_TYPE") = "P"
            End If
            rsHRTrain("TR_LDATE") = Date
            rsHRTrain("TR_LTIME") = Time$
            rsHRTrain("TR_LUSER") = glbUserID
            
            'Add a Follow Up record for this Training course
            SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE 1 = 2"
            rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            rsFollowUp.AddNew
            rsFollowUp("EF_COMPNO") = "001"
            rsFollowUp("EF_EMPNBR") = rsEmpJobs("TW_EMPNBR")
            rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
            rsFollowUp("EF_FREAS_TABL") = "FURE"
            'Ticket #24257 - Do not update Admin By for them only
            If glbCompSerial <> "S/N - 2262W" Then
                rsFollowUp("EF_ADMINBY_TABL") = "EDAB"
                rsFollowUp("EF_ADMINBY") = GetEmpData(rsEmpJobs("TW_EMPNBR"), "ED_ADMINBY", Null)
            End If
            rsFollowUp("EF_FREAS") = "EDUC"
            rsFollowUp("EF_COMMENTS") = "Course: " & xCourse & " - " & GetTABLDesc("ESCD", xCourse) & " for Position: " & rsEmpJobs("TW_JOB")
            rsFollowUp("EF_LDATE") = Date
            rsFollowUp("EF_LTIME") = Time$
            rsFollowUp("EF_LUSER") = glbUserID
            rsFollowUp.Update
            
            rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
            rsHRTrain.Update
            
            rsFollowUp.Close
            Set rsFollowUp = Nothing
        
            'Update Position record with Follow Up ID
            'if the course code is TRAIN
            If xCourse = "TRAIN" Then
                'Search HR_JOB_HISTORY or HR_TEMP_WORK table for this Position record
                'and update with Follow Up Id
                If (rsEmpJobs("POS_TYPE") = "CURRENT") Then
                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & rsEmpJobs("TW_ID")
                Else
                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & rsEmpJobs("TW_ID")
                End If
                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsTJob.EOF Then
                    If (rsEmpJobs("POS_TYPE") = "CURRENT") Then
                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                    Else
                        rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                    End If
                    rsTJob.Update
                End If
                rsTJob.Close
                Set rsTJob = Nothing
            End If
            
'            rsHRTrain.Close
'            Set rsHRTrain = Nothing
'
'            rsContEdu.Close
'            Set rsContEdu = Nothing
            
            'If flgUnqForPos Then
            '    'Go to next position
            '    GoTo next_EmpPosition
            'Else
                'Exit loop - only the first position gets this course
                'Exit Do
            'End If
            
            
next_EmpPosition:
            rsHRTrain.Close
            Set rsHRTrain = Nothing
            

            rsEmpJobs.MoveNext
                        
        Loop
    End If
    rsEmpJobs.Close
    Set rsEmpJobs = Nothing
        
End Sub

Private Function getLocJobCodes(xCodeList, xType)
Dim SQLQ As String
Dim rsTemp As New ADODB.Recordset
Dim retval As String
    retval = ""
    If Len(xCodeList) > 0 Then
        xCodeList = Replace(xCodeList, ",", "','")
        If xType = "PStatus" Then
            SQLQ = "SELECT * FROM HRJOB WHERE JB_STATUS IN ('" & xCodeList & "') "
        End If
        If xType = "BAND" Then
            SQLQ = "SELECT * FROM HRJOB WHERE JB_BAND IN ('" & xCodeList & "') "
        End If
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsTemp.EOF
            retval = retval & rsTemp("JB_CODE") & ","
            rsTemp.MoveNext
        Loop
        rsTemp.Close
    End If
    getLocJobCodes = retval
End Function

Private Sub Refresh_Training_Plan(Optional xPosSelected)
    Dim rsEmpJobs As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim SQLQ As String

    On Error GoTo Refresh_Training_Plan_Err



'        SQLQ = "DELETE FROM HR_TRAIN"
'        SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
''        SQLQ = SQLQ & " AND TR_JOB = '" & PosCode & "'"
''        If Not IsMissing(xPosType) Then
''            SQLQ = SQLQ & " AND TR_POS_TYPE = '" & xPosType & "'"
''        End If
'        SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0)"
'        gdbAdoIhr001.Execute SQLQ
'
'        'Delete this Training List record as the course is not required by other positions
'        SQLQ = "DELETE FROM HR_TRAIN"
'        SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
''        SQLQ = SQLQ & " AND TR_JOB = '" & PosCode & "'"
'        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
'        gdbAdoIhr001.Execute SQLQ

    'Clear all Training Plan first for all or for the specific Job
    If IsMissing(xPosSelected) Then
        SQLQ = "DELETE FROM HR_TRAIN WHERE TR_JOB IS NOT NULL OR TR_JOB <> ''"
    Else
        SQLQ = "DELETE FROM HR_TRAIN WHERE TR_JOB ='" & xPosSelected & "'"
    End If
    gdbAdoIhr001.Execute SQLQ


    'Retrieve the Required Courses for the selected Position or for ALL
    If IsMissing(xPosSelected) Then
        SQLQ = "SELECT * FROM HR_JOB_COURSE"
    Else
        SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE PC_JOB = '" & xPosSelected & "'"
    End If
        
    'Ticket #25609 - Training Plan by Department
    'Only courses matching employee's Department if the Course has Department Code assigned
    'SQLQ = SQLQ & " AND ((PC_DEPTNO IS NULL) OR (PC_DEPTNO = '" & GetEmpData(glbLEE_ID, "ED_DEPTNO") & "'))"
    
    rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsReqCourse.EOF Then
        rsReqCourse.MoveFirst
        
        Do While Not rsReqCourse.EOF
            SQLQ = "SELECT JH_ID AS TW_ID, JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
            'SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
            SQLQ = SQLQ & " WHERE ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
            'and not the position currently selected
            'SQLQ = SQLQ & " AND (JH_ID <> " & RSDATA!JH_ID & ")"
            SQLQ = SQLQ & " AND JH_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
            'Retrieve the Required Courses for the selected Position or for ALL
            If Not IsMissing(xPosSelected) Then
                SQLQ = SQLQ & " AND JH_JOB IN ('" & xPosSelected & "')"
            End If
            SQLQ = SQLQ & " UNION "
            SQLQ = SQLQ & " SELECT TW_ID, TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
            'SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
            SQLQ = SQLQ & " WHERE ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
            SQLQ = SQLQ & " AND TW_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
            'Retrieve the Required Courses for the selected Position or for ALL
            If Not IsMissing(xPosSelected) Then
                SQLQ = SQLQ & " AND TW_JOB IN ('" & xPosSelected & "')"
            End If
            SQLQ = SQLQ & " ORDER BY TW_EMPNBR, TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
            rsEmpJobs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsEmpJobs.EOF Then
                rsEmpJobs.MoveFirst
                
                Do While Not rsEmpJobs.EOF
                    If (rsEmpJobs("POS_TYPE") = "CURRENT" Or rsEmpJobs("POS_TYPE") = "TEMPORARY") And rsEmpJobs("TW_CURRENT") Then
                        Call Update_Employee_Job_Training_List(rsEmpJobs("TW_EMPNBR"), rsEmpJobs("TW_JOB"), IIf(rsEmpJobs("POS_TYPE") = "CURRENT", "Current", "Temporary"), rsEmpJobs("TW_SDATE"), , rsEmpJobs("TW_ID"))
                    Else
                        Call Update_Employee_Job_Training_List(rsEmpJobs("TW_EMPNBR"), rsEmpJobs("TW_JOB"), "Previous", rsEmpJobs("TW_SDATE"), rsEmpJobs("TW_ENDDATE"), rsEmpJobs("TW_ID"))
                    End If
                    
                    rsEmpJobs.MoveNext
                Loop
            End If
            rsEmpJobs.Close
            Set rsEmpJobs = Nothing
        
            rsReqCourse.MoveNext
        Loop
    End If
    rsReqCourse.Close
    Set rsReqCourse = Nothing

Exit Sub

Refresh_Training_Plan_Err:
If Err = 3018 Then
    Err = 0
    Resume Next
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
If Len(SQLQ) = 0 Then
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Refresh_Training_Plan", "HR_JOB_COURSE", "Refresh Training Plan")
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, SQLQ, "HR_JOB_COURSE", "Refresh Training Plan")
End If
Call RollBack '26July99 js
End Sub

Private Sub Update_Employee_Job_Training_List(xEmpnbr, xJob, xPosType, Optional xStartEndDate, Optional xEndDate, Optional xJobID)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsCourseMst As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim rsEmpJob As New ADODB.Recordset
    Dim xDWMY, xorgPosType, xorgJob As String
    Dim SQLQ  As String
    Dim flgUnqForPos, flgNoPrvRnwl, flgNoCurRnwl, flgCrsTakenBefore, flgProcCalled As Boolean
    Dim xPrvEndDate
    Dim xComments As String
    
    '''On Error GoTo Employee_Job_Training_Err

    'Note: If tracking is for the Previous Job then any courses for this job which does not have
    'Previous Renewal defined should be removed for this position or
    'If tracking is for Current Job then any courses for this job which does not have
    'Current Renewal defined should be removed for this position
    
    'if this procedure is called from another procedure and not an event
    If IsMissing(xStartEndDate) Then
        flgProcCalled = False
        xorgPosType = xPosType
        xorgJob = xJob
        xStartEndDate = ""
        xEndDate = ""
    Else
        flgProcCalled = True
    End If
    
    'Get the list of Required Courses for the Job
    SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE PC_JOB = '" & xJob & "'"
    
    'Ticket #25609 - Training Plan by Department
    'Only courses matching employee's Department if the Course has Department Code assigned
    SQLQ = SQLQ & " AND ((PC_DEPTNO IS NULL) OR (PC_DEPTNO = '" & GetEmpData(xEmpnbr, "ED_DEPTNO") & "'))"
    
    rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsReqCourse.EOF Then
        rsReqCourse.MoveFirst
        
        Do While Not rsReqCourse.EOF
            'Ticket #25609 - Training Plan by Department
            'Check if the Course has Department assigned. If so then check if the Department of the Course matches
            'employee's Department
            'If Not IsNull(rsReqCourse("PC_DEPTNO")) And rsReqCourse("PC_DEPTNO") <> "" Then
            '    If rsReqCourse("PC_DEPTNO") <> GetEmpData(xEmpnbr, "ED_DEPTNO") Then
            '        'Skip this course as Employee does not belong to the department this Course is setup for
            '        GoTo Next_Required_Course
            '    End If
            'End If
        
            'Check if this required course is Unique for each Position.
            'If so, then it will have to be added in the Training List even
            'though the Course code already exists for this employee for other positions
            flgUnqForPos = False
            flgNoPrvRnwl = False
            flgNoCurRnwl = False
            SQLQ = "SELECT ES_CRSCODE,ES_UNIQUE_FOR_POS,ES_RENEW_CRS_CUR,ES_CUR_PRD_DWMY, ES_RENEW_CRS_PRV,ES_PRV_PRD_DWMY, ES_RENEW_FOLLOWUP, ES_FLWUP_PRD_DWMY FROM HR_COURSECODE_MASTER"
            SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
            rsCourseMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsCourseMst.EOF Then
                flgUnqForPos = IIf(IsNull(rsCourseMst("ES_UNIQUE_FOR_POS")), False, rsCourseMst("ES_UNIQUE_FOR_POS"))
                
                If flgUnqForPos = False And rsCourseMst("ES_RENEW_FOLLOWUP") = 99 And rsCourseMst("ES_FLWUP_PRD_DWMY") = "Y" Then
                    'Skip this course
                    GoTo Next_Required_Course
                End If
            Else
                'Course not defined in the Course Code Master - skip this course
                GoTo Next_Required_Course
            End If
            'rsCourseMst.Close
            'Set rsCourseMst = Nothing
            
            'Follow Up Effective Date Period is mandatory. Check if it exists otherwise the logic below will give an error.
            If IsNull(rsReqCourse("PC_RENEW_FOLLOWUP")) Or rsReqCourse("PC_RENEW_FOLLOWUP") = "" Then
                'Follow Up Effective Date renewal Period missing
                GoTo Next_Required_Course
            End If
                        
            'Add the Required Courses in the Training List
            'if it does not already exists for this employee or Unique for each Position
            SQLQ = "SELECT * FROM HR_TRAIN"
            SQLQ = SQLQ & " WHERE TR_EMPNBR = " & xEmpnbr
            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
            If flgUnqForPos <> 0 Then
                SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
                'If xPosType = "Previous" And chkTrackCrsRenewal And chkCurrent(0) Then
                '    SQLQ = SQLQ & " AND TR_POS_TYPE = 'C'"
                'Else
                '    If chkTrackCrsRenewal And chkCurrent(0) Then
                '        SQLQ = SQLQ & " AND TR_POS_TYPE = 'P'"
                '    Else
                '        SQLQ = SQLQ & " AND TR_POS_TYPE = '" & Left(xPosType, 1) & "'"
                '    End If
                'End If
            End If
            rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsHRTrain.EOF Then
                'TRAINING RECORD DOES NOT EXISTS - ADD NEW ONE
                
                'Check first if this Course was taken before in the Continuing Education screen
                flgCrsTakenBefore = False
                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_JOB, ES_DATCOMP,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
                If flgUnqForPos <> 0 Then
                    SQLQ = SQLQ & " AND ES_JOB = '" & xJob & "'"
                End If
                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                SQLQ = SQLQ & " AND (ES_RENEW = '' OR ES_RENEW IS NULL)"
                SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
                SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsContEdu.EOF Then
                    'Course Taken Before
                    rsContEdu.MoveFirst
                    flgCrsTakenBefore = True
                Else
                    'Course not taken before
                    flgCrsTakenBefore = False
                End If
                
                
'                'May be Training List accidently deleted or messed up
'                'if the Course is Previous and procedure not called from another procdure then
'                'check if this course is required by another Primary or Temporary Current or Previous Position if so then
'                'change the xJob to that Position and Start & Date Date to that Position Start Date & End Date
'                If flgProcCalled = False And xPosType = "Previous" Then
'                    'Check if Primary Current or Previous or Temp Current or other Previous required this Course
'                    SQLQ = "SELECT JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
'                    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & xEmpnbr & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
'                    SQLQ = SQLQ & " AND JH_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                    If flgUnqForPos <> 0 Then
'                        SQLQ = SQLQ & " AND JH_JOB = '" & xJob & "'"
'                    End If
'                    SQLQ = SQLQ & " AND (JH_ID <> " & RSDATA!JH_ID & ")"
'                    SQLQ = SQLQ & " UNION "
'                    SQLQ = SQLQ & " SELECT TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
'                    SQLQ = SQLQ & " WHERE TW_EMPNBR = " & xEmpnbr & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
'                    SQLQ = SQLQ & " AND TW_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                    If flgUnqForPos <> 0 Then
'                        SQLQ = SQLQ & " AND TW_JOB = '" & xJob & "'"
'                    End If
'                    SQLQ = SQLQ & " ORDER BY TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
'                    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                    If Not rsEmpJob.EOF Then
'                        'The first record gets it
'                        'the order is Primary Current, Temp Current and then Previous depending on most recent end date
'                        rsEmpJob.MoveFirst
'                        If Not IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
'                            If rsEmpJob("TW_TRK_CRS_RENEWAL") Then
'                                If CVDate(rsEmpJob("TW_ENDDATE")) > CVDate(dlpENDDATE.Text) Then
'                                    'Previous Position requires this course
'                                    xJob = rsEmpJob("TW_JOB")
'                                    xStartEndDate = rsEmpJob("TW_SDATE")
'                                    xEndDate = rsEmpJob("TW_ENDDATE")
'                                    xPosType = "Previous"
'                                End If
'                            Else
'                                If rsEmpJob("POS_TYPE") = "CURRENT" Then
'                                    If xJob <> rsEmpJob("TW_JOB") Then    'If Current becoming Previous
'                                        xPosType = "Current"
'                                        xJob = rsEmpJob("TW_JOB")
'                                        xStartEndDate = rsEmpJob("TW_SDATE")
'                                    End If
'                                Else
'                                    xPosType = "Temporary"
'                                    xJob = rsEmpJob("TW_JOB")
'                                    xStartEndDate = rsEmpJob("TW_SDATE")
'                                End If
'                            End If
'                        Else
'                            If rsEmpJob("POS_TYPE") = "CURRENT" Then
'                                If xJob <> rsEmpJob("TW_JOB") Then    'If Current becoming Previous
'                                    xPosType = "Current"
'                                    xJob = rsEmpJob("TW_JOB")
'                                    xStartEndDate = rsEmpJob("TW_SDATE")
'                                End If
'                            Else
'                                xPosType = "Temporary"
'                                xJob = rsEmpJob("TW_JOB")
'                                xStartEndDate = rsEmpJob("TW_SDATE")
'                            End If
'                        End If
'                    Else
'                        xStartEndDate = ""  'Ticket #22951
'                    End If
'                    rsEmpJob.Close
'                    Set rsEmpJob = Nothing
'                Else
'                    'if Current then do not do anything as Current record takes precedence
'                End If
                
                'If the course is being added for the Previous Position and this course
                'does not have previous renewal period then do not add this course
                'If xPosType = "Current" Or (xPosType = "Previous" And (Not IsNull(rsReqCourse("PC_RENEW_CRS_PRV"))) And rsReqCourse("PC_RENEW_CRS_PRV") <> 0) Then
                
                'If Course was taken and it's Position is Current then
                'make sure Current Renewal Period is there otherwise do not add the course
                'If the course is being added for the Previous Position and this course
                'does not have previous renewal period then do not add this course
                'Changed
                If (flgCrsTakenBefore = True And (xPosType = "Current" Or xPosType = "Temporary") And (Not IsNull(rsReqCourse("PC_RENEW_CRS_CUR"))) And rsReqCourse("PC_RENEW_CRS_CUR") <> 0) Or _
                    (flgCrsTakenBefore = False And (xPosType = "Current" Or xPosType = "Temporary")) Or (flgCrsTakenBefore = True And xPosType = "Previous" And (Not IsNull(rsReqCourse("PC_RENEW_CRS_PRV"))) And rsReqCourse("PC_RENEW_CRS_PRV") <> 0) Or _
                    (flgCrsTakenBefore = False And xPosType = "Previous") Then
                    
                    'Add Training Record
                    rsHRTrain.AddNew
                    rsHRTrain("TR_COMPNO") = "001"
                    rsHRTrain("TR_EMPNBR") = xEmpnbr
                    rsHRTrain("TR_CRSCODE") = rsReqCourse("PC_CRSCODE")
                    
                    If flgCrsTakenBefore = False Then
                        If Not IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) And rsReqCourse("PC_RENEW_CRS_CUR") <> 0 Then
                            'Current Course Renewal found
                            Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            If xPosType = "Current" Or xPosType = "Temporary" Or xPosType = "Previous" Then
'                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
'                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
'                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
'                                End If
                                
                            'For courses not taken and are now Previous, the renewal date is based
                            'on Follow Up Renewal Period and not Previous Renewal Period - above
                            'ElseIf xPosType = "Previous" Then
                            '    Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                            '        Case "D"
                            '            xDWMY = "d"
                            '        Case "W"
                            '            xDWMY = "ww"
                            '        Case "M"
                            '            xDWMY = "m"
                            '        Case "Y"
                            '            xDWMY = "yyyy"
                            '    End Select
                            '    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(dlpStartDate.Text))
                            End If
                        Else    'No Current Course Renewal Period
                            If xPosType = "Current" Or xPosType = "Temporary" Or xPosType = "Previous" Then
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
'                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
'                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
'                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
'                                End If
                            'ElseIf xPosType = "Previous" Then
                            '    'For courses not taken and are now Previous, the renewal date is based
                            '    'on Follow Up Renewal Period and not Previous Renewal Period.
                            '    'If there is no current renewal then it's based on End Date only and
                            '    'Prev Renewal Period - for courses taken.
                            '    'Compute Renewal with Position End Date because there is no Current Renewal Period defined
                            '    Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                            '        Case "D"
                            '            xDWMY = "d"
                            '        Case "W"
                            '            xDWMY = "ww"
                            '        Case "M"
                            '            xDWMY = "m"
                            '        Case "Y"
                            '            xDWMY = "yyyy"
                            '    End Select
                            '    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(dlpENDDATE.Text))
                            End If
                        End If
                    Else    'Course Has Been Taken Before
                        'Course has been taken before, compute Renewal Date based on Course Taken Date
                        If xPosType = "Current" Or xPosType = "Temporary" Then
                            Select Case rsReqCourse("PC_CUR_PRD_DWMY")
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_CUR"), CVDate(rsContEdu("ES_DATCOMP")))
                            rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                        ElseIf xPosType = "Previous" Then
                            Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            If Not IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) And rsReqCourse("PC_RENEW_CRS_CUR") <> 0 Then
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(rsContEdu("ES_DATCOMP")))
                            Else
'                                If IsMissing(xEndDate) Or xEndDate = "" Or IsNull(xEndDate) Then
'                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(dlpENDDATE.Text))
'                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(xEndDate))
'                                End If
                            End If
                            rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                        End If
                        
                        'Update Continuing Education with new Renewal Date
                        rsContEdu("ES_JOB") = xJob
                        rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                        rsContEdu("ES_LDATE") = Date
                        rsContEdu("ES_LUSER") = glbUserID
                        rsContEdu("ES_LTIME") = Time$
                        rsContEdu.Update
                    End If
                    
                    rsHRTrain("TR_JOB") = xJob
'                    If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
'                        rsHRTrain("TR_SDATE") = dlpStartDate.Text
'                    Else
                        rsHRTrain("TR_SDATE") = xStartEndDate
'                    End If
                    If xPosType = "Current" Then
                        rsHRTrain("TR_POS_TYPE") = "C"
                    ElseIf xPosType = "Temporary" Then
                        rsHRTrain("TR_POS_TYPE") = "T"
                    ElseIf xPosType = "Previous" Then
                        rsHRTrain("TR_POS_TYPE") = "P"
                    End If
                    'rsHRTrain("TR_COURSE_TAKEN")   - Remains BLANK
                    rsHRTrain("TR_LDATE") = Date
                    rsHRTrain("TR_LTIME") = Time$
                    rsHRTrain("TR_LUSER") = glbUserID
                    
                    'Add a Follow Up record for this Training course
'                    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE 1 = 2"
'                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                    rsFollowUp.AddNew
'                    rsFollowUp("EF_COMPNO") = "001"
'                    rsFollowUp("EF_EMPNBR") = xEmpnbr
'                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
'                    rsFollowUp("EF_FREAS_TABL") = "FURE"
'                    'Ticket #24257 - Do not update Admin By for them only
'                    If glbCompSerial <> "S/N - 2262W" Then
'                        rsFollowUp("EF_ADMINBY_TABL") = "EDAB"
'                        rsFollowUp("EF_ADMINBY") = GetEmpData(xEmpnbr, "ED_ADMINBY", Null)
'                    End If
'                    rsFollowUp("EF_FREAS") = "EDUC"
'                    rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
'                    rsFollowUp("EF_LDATE") = Date
'                    rsFollowUp("EF_LTIME") = Time$
'                    rsFollowUp("EF_LUSER") = glbUserID
'                    rsFollowUp.Update
                    
                    'Ticket #24300
                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(xEmpnbr, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                    
                    rsHRTrain.Update
                    
                    'rsFollowUp.Close
                    'Set rsFollowUp = Nothing
                
                    'Update Position record with Follow Up ID
                    'if the course code is TRAIN
                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                        'Search HR_JOB_HISTORY table for this Position record
                        'and update with Follow Up Id
                        If xPosType = "Temporary" Then
                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & xJobID     'Data1.Recordset("JH_ID")
                        Else
                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & xJobID     'Data1.Recordset("JH_ID")
                        End If
                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsTJob.EOF Then
                            If xPosType = "Temporary" Then
                                rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                            Else
                                rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                            End If
                            rsTJob.Update
                        End If
                        rsTJob.Close
                        Set rsTJob = Nothing
                    End If
                End If
                rsContEdu.Close
                Set rsContEdu = Nothing
                
                If flgProcCalled = False Then
                    xPosType = xorgPosType
                    xJob = xorgJob
                End If
            Else
                'TRAINING RECORD FOUND
                
'                'May be Training List accidently deleted or messed up
'                'if the Course is Previous and procedure not called from another procdure then
'                'check if this course is required by another Primary or Temporary Current or Previous Position if so then
'                'change the xJob to that Position and Start & Date Date to that Position Start Date & End Date
'                If flgProcCalled = False And xPosType = "Previous" Then
'                    'Check if Primary Current or Previous or Temp Current or other Previous required this Course
'                    SQLQ = "SELECT JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
'                    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & xEmpnbr & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
'                    SQLQ = SQLQ & " AND JH_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                    If flgUnqForPos <> 0 Then
'                        SQLQ = SQLQ & " AND JH_JOB = '" & xJob & "'"
'                    End If
'                    SQLQ = SQLQ & " AND (JH_ID <> " & RSDATA!JH_ID & ")"
'                    SQLQ = SQLQ & " UNION "
'                    SQLQ = SQLQ & " SELECT TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
'                    SQLQ = SQLQ & " WHERE TW_EMPNBR = " & xEmpnbr & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
'                    SQLQ = SQLQ & " AND TW_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                    If flgUnqForPos <> 0 Then
'                        SQLQ = SQLQ & " AND TW_JOB = '" & xJob & "'"
'                    End If
'                    SQLQ = SQLQ & " ORDER BY TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
'                    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                    If Not rsEmpJob.EOF Then
'                        'The first record gets it
'                        'the order is Primary Current, Temp Current and then Previous depending on most recent end date
'                        rsEmpJob.MoveFirst
'                        If Not IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
'                            If rsEmpJob("TW_TRK_CRS_RENEWAL") Then
'                                If CVDate(rsEmpJob("TW_ENDDATE")) > CVDate(dlpENDDATE.Text) Then
'                                    'Previous Position requires this course
'                                    xJob = rsEmpJob("TW_JOB")
'                                    xStartEndDate = rsEmpJob("TW_SDATE")
'                                    xEndDate = rsEmpJob("TW_ENDDATE")
'                                    xPosType = "Previous"
'                                End If
'                            Else
'                                If rsEmpJob("POS_TYPE") = "CURRENT" Then
'                                    If xJob <> rsEmpJob("TW_JOB") Then    'If Current becoming Previous
'                                        xPosType = "Current"
'                                        xJob = rsEmpJob("TW_JOB")
'                                        xStartEndDate = rsEmpJob("TW_SDATE")
'                                    End If
'                                Else
'                                    xPosType = "Temporary"
'                                    xJob = rsEmpJob("TW_JOB")
'                                    xStartEndDate = rsEmpJob("TW_SDATE")
'                                End If
'                            End If
'                        Else
'                            If rsEmpJob("POS_TYPE") = "CURRENT" Then
'                                If xJob <> rsEmpJob("TW_JOB") Then    'If Current becoming Previous
'                                    xPosType = "Current"
'                                    xJob = rsEmpJob("TW_JOB")
'                                    xStartEndDate = rsEmpJob("TW_SDATE")
'                                End If
'                            Else
'                                xPosType = "Temporary"
'                                xJob = rsEmpJob("TW_JOB")
'                                xStartEndDate = rsEmpJob("TW_SDATE")
'                            End If
'                        End If
'                    Else
'                        xStartEndDate = ""  'Ticket #22951
'                    End If
'                    rsEmpJob.Close
'                    Set rsEmpJob = Nothing
'                Else
'                    'if Current then do not do anything as Current record takes precedence
'                End If
                
                
                
                'Training record for this course already exists so update the Renewal Date
                'Check which Type of Position is assigned to this course
                If rsHRTrain("TR_POS_TYPE") = "C" Then
                    'Currently the course is holding Primary Current Position Code
                    'Check which type of position requires this course
                    If xPosType = "Current" Then
                        'These courses are for new Current Primary Position so recalculate the
                        'Renewal Dates - based on Position Start Date or last Course Taken date
                        'See which Position Start Date is most recent
'                        If CVDate(rsHRTrain("TR_SDATE")) < CVDate(IIf(IsMissing(xStartEndDate) Or xStartEndDate = "", dlpStartDate.Text, xStartEndDate)) Then
                        If CVDate(rsHRTrain("TR_SDATE")) < CVDate(xStartEndDate) Then
                            'Training List has older Position Start Date so update with new Position info.
                            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
'                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
'                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
'                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
'                                End If
                            Else
                                'Check if Current Renewal period is defined
                                If IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0 Then
                                    'No Current Renewal Period defined so delete this job from this current position.
                                    'It should not be in the training list for any current job
                                    flgNoCurRnwl = True
                                Else
                                    Select Case rsReqCourse("PC_CUR_PRD_DWMY")
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_CUR"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                End If
                            End If
                            If flgNoCurRnwl = False Then
                                'Update Continuing Education record as well
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_JOB") = xJob
                                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                            
                                rsHRTrain("TR_JOB") = xJob
'                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
'                                    rsHRTrain("TR_SDATE") = dlpStartDate.Text
'                                Else
                                    rsHRTrain("TR_SDATE") = xStartEndDate
'                                End If
                                rsHRTrain("TR_POS_TYPE") = "C"   'Current Primary
                                ''If Renewal date is greater than today's date then clear the Course Taken Date
                                'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                                '    rsHRTrain("TR_COURSE_TAKEN") = Null
                                'End If
                                rsHRTrain("TR_LDATE") = Date
                                rsHRTrain("TR_LUSER") = glbUserID
                                rsHRTrain("TR_LTIME") = Time$
                                
                                'If follow up id is null then find the id
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & xEmpnbr
                                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                    SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    'Add a Follow Up record for this Training course
                                    'Ticket #24300
                                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(xEmpnbr, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                    
                                    rsHRTrain.Update
                                Else
                                    rsHRTrain.Update
                                
                                    'Update Follow Up record - Effective Date
                                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                        rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                        rsFollowUp("EF_LDATE") = Date
                                        rsFollowUp("EF_LUSER") = glbUserID
                                        rsFollowUp("EF_LTIME") = Time$
                                        rsFollowUp.Update
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Update Position record with Follow Up ID
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and update with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & xJobID   'Data1.Recordset("JH_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                
                                    'Clear the Follow Up Id on the other current position rec in the Temp Position table
                                    'Search HR_TEMP_WORK table for this Position record
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            Else
                                'CURRENT - Current
                                'No Current renewal found for this course
                                
                                'Clear the Renewal date for this course and for this employee from
                                'Continuing Education screen
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                
                                    If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                                        'Since the course was completed - mark the Follow Up as
                                        'Completed instead of deleting it.
                                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    Else
                                        'Delete the Follow Up record for this training record
                                        'as no Course completion record found
                                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                        'Ticket #26211 Franks 11/04/2014
                                        'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    
                                        'Clear the Follow Up Id on the Position record
                                        'if the course code is TRAIN
                                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                            'Search HR_JOB_HISTORY table for this Position record
                                            'and update with Follow Up Id
                                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                            If Not rsTJob.EOF Then
                                                rsTJob("JH_FOLLOWUP_ID") = Null
                                                rsTJob.Update
                                            End If
                                            rsTJob.Close
                                            Set rsTJob = Nothing
                                        End If
                                    End If
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    'Ticket #26211 Franks 11/04/2014
                                    'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                
                                    'Clear the Follow Up ID in the Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                
                                'Delete this Training List record as the course is not required by other positions
                                SQLQ = "DELETE FROM HR_TRAIN"
                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & xEmpnbr
                                SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                gdbAdoIhr001.Execute SQLQ
                            End If
                        Else
                            'Do not do anything because Training List has most recent Position Start Date
                        End If
                    ElseIf xPosType = "Previous" Then
                        'CURRENT - Previous
                        'Current Job becoming Previous
                        'Previous Primary Position is being tracked but Current Primary Position has this course
                        'Check if the Position in HR_TRAIN is same this Position
'                        If (rsHRTrain("TR_JOB") <> xJob) Or (rsHRTrain("TR_JOB") = xJob And CVDate(rsHRTrain("TR_SDATE")) <> CVDate(dlpStartDate.Text) And CVDate(rsHRTrain("TR_SDATE")) <> CVDate(IIf(IsMissing(xStartEndDate) Or xStartEndDate = "", dlpStartDate.Text, xStartEndDate))) Then
                        If (rsHRTrain("TR_JOB") <> xJob) Or (rsHRTrain("TR_JOB") = xJob And CVDate(rsHRTrain("TR_SDATE")) <> CVDate(xStartEndDate)) Then
                            'Do not do anything because Current takes the priority
                        Else
                            'Renewal Date based on last Course Taken date if present
                            'otherwise Follow Up Effective Date Period
                            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
'                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
'                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
'                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
'                                End If
                            Else
                                'Change the renewal dates if Previous renewal is defined
                                If IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) Or rsReqCourse("PC_RENEW_CRS_PRV") = 0 Then
                                    'No Previous Renewal Period defined so delete this job from this previous position.
                                    'It should not be in the training list for any previous job
                                    flgNoPrvRnwl = True
                                Else
                                    Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                End If
                            End If
                            If flgNoPrvRnwl = False Then
                                'Update Continuing Education record as well
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_JOB") = xJob
                                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                            
                                'Previous Renewal period available
                                rsHRTrain("TR_JOB") = xJob
'                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
'                                    rsHRTrain("TR_SDATE") = dlpStartDate.Text
'                                Else
                                    rsHRTrain("TR_SDATE") = xStartEndDate
'                                End If
                                rsHRTrain("TR_POS_TYPE") = "P"   'Previous Primary
                                ''If Renewal date is greater than today's date then clear the Course Taken Date
                                'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                                '    rsHRTrain("TR_COURSE_TAKEN") = Null
                                'End If
                                rsHRTrain("TR_LDATE") = Date
                                rsHRTrain("TR_LUSER") = glbUserID
                                rsHRTrain("TR_LTIME") = Time$
                                
                                'If follow up id is null then find the id
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & xEmpnbr
                                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                    SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Ticket #24300
                                'rsHRTrain.Update
                                
                                'Ticket #24300
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    'Add a Follow Up record for this Training course
                                    'Ticket #24300
                                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(xEmpnbr, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                    rsHRTrain.Update
                                Else
                                    rsHRTrain.Update
                                    'Update Follow Up record - Effective Date
                                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                        rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                        rsFollowUp("EF_LDATE") = Date
                                        rsFollowUp("EF_LUSER") = glbUserID
                                        rsFollowUp("EF_LTIME") = Time$
                                        rsFollowUp.Update
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                            
                                'Update Position record with Follow Up ID
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and update with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & xJobID   'Data1.Recordset("JH_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            Else
                                'CURRENT - Previous
                                'No Previous renewal found for this course
                                
                                'Clear the Renewal date for this course and for this employee from
                                'Continuing Education screen
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                
                                    If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                                        'Since the course was completed - mark the Follow Up as
                                        'Completed instead of deleting it.
                                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    Else
                                        'Delete the Follow Up record for this training record
                                        'as no Course completion record found
                                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                        'Ticket #26211 Franks 11/04/2014
                                        'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    
                                        'Clear the Follow Up ID in the Position record
                                        'if the course code is TRAIN
                                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                            'Search HR_JOB_HISTORY table for this Position record
                                            'and update with Follow Up Id
                                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & xJobID   'Data1.Recordset("JH_ID")
                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                            If Not rsTJob.EOF Then
                                                rsTJob("JH_FOLLOWUP_ID") = Null
                                                rsTJob.Update
                                            End If
                                            rsTJob.Close
                                            Set rsTJob = Nothing
                                        End If
                                    End If
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    'Ticket #26211 Franks 11/04/2014
                                    'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                
                                    'Clear the Follow Up ID in the Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & xJobID   'Data1.Recordset("JH_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                
                                'Delete this Training List record as the course is not required by other positions
                                SQLQ = "DELETE FROM HR_TRAIN"
                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & xEmpnbr
                                SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                gdbAdoIhr001.Execute SQLQ
                            End If
                        End If
                    End If
                ElseIf rsHRTrain("TR_POS_TYPE") = "T" Then
                    'Currently the Temporary Current Position is holding this course
                    'Check which type of position requires this course now
                    If xPosType = "Current" Then
                        'These courses are for new Current Primary Position so recalculate the
                        'Renewal Dates - based on Position Start Date or last Course Taken date
                        'See which Position Start Date is most recent
                        'If CVDate(rsHRTrain("TR_SDATE")) <= CVDate(IIf(IsMissing(xStartEndDate) Or xStartEndDate = "", dlpStartDate.Text, xStartEndDate)) Then
                            'Training List has older Position Start Date so update with new Position info.
                            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
'                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
'                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
'                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
'                                End If
                            Else
                                'Check if Current Renewal period is defined
                                If IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0 Then
                                    'No Current Renewal Period defined so delete this job from this current position.
                                    'It should not be in the training list for any current job
                                    flgNoCurRnwl = True
                                Else
                                    Select Case rsReqCourse("PC_CUR_PRD_DWMY")
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_CUR"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                End If
                            End If
                            If flgNoCurRnwl = False Then
                                'Update Continuing Education record as well
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_JOB") = xJob
                                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                                        
                                rsHRTrain("TR_JOB") = xJob
'                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
'                                    rsHRTrain("TR_SDATE") = dlpStartDate.Text
'                                Else
                                    rsHRTrain("TR_SDATE") = xStartEndDate
'                                End If
                                rsHRTrain("TR_POS_TYPE") = "C"   'Current Primary
                                ''If Renewal date is greater than today's date then clear the Course Taken Date
                                'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                                '    rsHRTrain("TR_COURSE_TAKEN") = Null
                                'End If
                                rsHRTrain("TR_LDATE") = Date
                                rsHRTrain("TR_LUSER") = glbUserID
                                rsHRTrain("TR_LTIME") = Time$
                                
                                'If follow up id is null then find the id
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & xEmpnbr
                                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                    SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Ticket #24300
                                'rsHRTrain.Update
                                
                                'Ticket #24300
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    'Add a Follow Up record for this Training course
                                    'Ticket #24300
                                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(xEmpnbr, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                    rsHRTrain.Update
                                Else
                                    rsHRTrain.Update
                                    
                                    'Update Follow Up record - Effective Date
                                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                        rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                        rsFollowUp("EF_LDATE") = Date
                                        rsFollowUp("EF_LUSER") = glbUserID
                                        rsFollowUp("EF_LTIME") = Time$
                                        rsFollowUp.Update
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Update Position record with Follow Up ID
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and update with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & xJobID   'Data1.Recordset("JH_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                    
                                    'Clear the Follow Up Id on the position in the Temp/Cross Training Position table
                                    'Search HR_TEMP_WORK table for this Position record
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            Else
                                'TEMPORARY - Current
                                'No Current renewal found for this course
                                
                                'Clear the Renewal date for this course and for this employee from
                                'Continuing Education screen
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                
                                    If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                                        'Since the course was completed - mark the Follow Up as
                                        'Completed instead of deleting it.
                                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    Else
                                        'Delete the Follow Up record for this training record
                                        'as no Course completion record found
                                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                        'Ticket #26211 Franks 11/04/2014
                                        'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                        
                                        'Clear the Follow Up ID in the Position record
                                        'if the course code is TRAIN
                                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                            'Search HR_TEMP_WORK table for this Position record
                                            'and clear with Follow Up Id
                                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                            If Not rsTJob.EOF Then
                                                rsTJob("TW_FOLLOWUP_ID") = Null
                                                rsTJob.Update
                                            End If
                                            rsTJob.Close
                                            Set rsTJob = Nothing
                                        End If
                                    End If
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    'Ticket #26211 Franks 11/04/2014
                                    'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                
                                    'Clear the Follow Up ID in the Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_TEMP_WORK table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("TW_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                
                                'Delete this Training List record as the course is not required by other positions
                                SQLQ = "DELETE FROM HR_TRAIN"
                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & xEmpnbr
                                SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                gdbAdoIhr001.Execute SQLQ
                            End If
                        'Else
                            'Do not do anything because Training List has most recent Position Start Date
                        'End If
                    ElseIf xPosType = "Previous" Then
                        'TEMPORARY - Previous
                        'Do not do anything because Training List record is of the Current
                        'Temporary/Cross Training Position
                    
'                        'Previous Primary Position is being tracked but Temp. Current Position is holding this course
'                        'Check if the Position in HR_TRAIN is same this Position
'                        If rsHRTrain("TR_JOB") <> xJob Then
'                            'Do not do anything because Current takes the  priority
'                        Else
'                            'Change the renewal dates if Previous renewal is defined
'                            If IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) Or rsReqCourse("PC_RENEW_CRS_PRV") = 0 Then
'                                'No Previous Renewal Period defined so delete this job from this previous position.
'                                'It should not be in the training list for any previous job
'                                flgNoPrvRnwl = True
'                            Else
'                                'Renewal Date based on last Course Taken date if present
'                                'otherwise Follow Up Effective Date Period
'                                If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
'                                    Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
'                                        Case "D"
'                                            xDWMY = "d"
'                                        Case "W"
'                                            xDWMY = "ww"
'                                        Case "M"
'                                            xDWMY = "m"
'                                        Case "Y"
'                                            xDWMY = "yyyy"
'                                    End Select
'                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
'                                Else
'                                    Select Case rsReqCourse("PC_PRV_PRD_DWMY")
'                                        Case "D"
'                                            xDWMY = "d"
'                                        Case "W"
'                                            xDWMY = "ww"
'                                        Case "M"
'                                            xDWMY = "m"
'                                        Case "Y"
'                                            xDWMY = "yyyy"
'                                    End Select
'                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
'                                End If
'                            End If
'                            If flgNoPrvRnwl = False Then
'                                'Previous Renewal period available
'                                rsHRTrain("TR_JOB") = xJob
'                                rsHRTrain("TR_SDATE") = dlpStartDate.Text
'                                rsHRTrain("TR_POS_TYPE") = "P"   'Previous Primary
'                                ''If Renewal date is greater than today's date then clear the Course Taken Date
'                                'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
'                                '    rsHRTrain("TR_COURSE_TAKEN") = Null
'                                'End If
'                                rsHRTrain("TR_LDATE") = Date
'                                rsHRTrain("TR_LUSER") = glbUserID
'                                rsHRTrain("TR_LTIME") = Time$
'                                rsHRTrain.Update
'
'                                'Update Follow Up record - Effective Date
'                                SQLQ = "SELECT * FROM HR_FOLLOW_UP"
'                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
'                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                If Not rsFollowUp.EOF Then
'                                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
'                                    rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
'                                    rsFollowUp("EF_LDATE") = Date
'                                    rsFollowUp("EF_LUSER") = glbUserID
'                                    rsFollowUp("EF_LTIME") = Time$
'                                    rsFollowUp.Update
'                                End If
'                                rsFollowUp.Close
'                                Set rsFollowUp = Nothing
'
'                                'Update Position record with Follow Up ID
'                                'if the course code is TRAIN
'                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
'                                    'Search HR_JOB_HISTORY table for this Position record
'                                    'and update with Follow Up Id
'                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("TW_ID")
'                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                    If Not rsTJob.EOF Then
'                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
'                                        rsTJob.Update
'                                    End If
'                                    rsTJob.Close
'                                    Set rsTJob = Nothing
'
'                                    'Clear the Follow Up Id on the position in the Temp/Cross Training Position table
'                                    'Search HR_TEMP_WORK table for this Position record
'                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
'                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                    If Not rsTJob.EOF Then
'                                        rsTJob("TW_FOLLOWUP_ID") = Null
'                                        rsTJob.Update
'                                    End If
'                                    rsTJob.Close
'                                    Set rsTJob = Nothing
'                                End If
'                            Else
'                                'No Previous renewal found for this course
'
'                                'Clear the Renewal date for this course and for this employee from
'                                'Continuing Education screen
'                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
'                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
'                                SQLQ = SQLQ & " AND ES_JOB = '" & xJob & "'"
'                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
'                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & xJob & "'"
'                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & xJob & "'"
'                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                If Not rsContEdu.EOF Then
'                                    rsContEdu("ES_RENEW") = Null
'                                    rsContEdu("ES_LDATE") = Date
'                                    rsContEdu("ES_LUSER") = glbUserID
'                                    rsContEdu("ES_LTIME") = Time$
'                                    rsContEdu.Update
'
'                                    If Not IsNull(rsContEdu("ES_DATCOMP")) Then
'                                        'Since the course was completed - mark the Follow Up as
'                                        'Completed instead of deleting it.
'                                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP"))
'                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & xJob & "'"
'                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                                        gdbAdoIhr001.Execute SQLQ
'                                    Else
'                                        'Delete the Follow Up record for this training record
'                                        'as no Course completion record found
'                                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
'                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & xJob & "'"
'                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                                        gdbAdoIhr001.Execute SQLQ
'
'                                        'Clear the Follow Up ID in the Position record
'                                        'if the course code is TRAIN
'                                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
'                                            'Search HR_TEMP_WORK table for this Position record
'                                            'and clear with Follow Up Id
'                                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
'                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                            If Not rsTJob.EOF Then
'                                                rsTJob("TW_FOLLOWUP_ID") = Null
'                                                rsTJob.Update
'                                            End If
'                                            rsTJob.Close
'                                            Set rsTJob = Nothing
'                                        End If
'                                    End If
'                                Else
'                                    'Delete the Follow Up record for this training record
'                                    'as no Course completion record found
'                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
'                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & xJob & "'"
'                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                                    gdbAdoIhr001.Execute SQLQ
'
'                                    'Clear the Follow Up ID in the Position record
'                                    'if the course code is TRAIN
'                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
'                                        'Search HR_TEMP_WORK table for this Position record
'                                        'and clear with Follow Up Id
'                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
'                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                        If Not rsTJob.EOF Then
'                                            rsTJob("TW_FOLLOWUP_ID") = Null
'                                            rsTJob.Update
'                                        End If
'                                        rsTJob.Close
'                                        Set rsTJob = Nothing
'                                    End If
'                                End If
'                                rsContEdu.Close
'                                Set rsContEdu = Nothing
'
'                                'Delete this Training List record as the course is not required by other positions
'                                SQLQ = "DELETE FROM HR_TRAIN"
'                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & xEmpnbr
'                                SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
'                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
'                                gdbAdoIhr001.Execute SQLQ
'                            End If
'                        End If
                    End If
                ElseIf rsHRTrain("TR_POS_TYPE") = "P" Then
                    'Previous Primary or Temporary position is holding this course
                    If xPosType = "Current" Then
                        'This course is required by new Current Primary Position so recalculate the
                        'Renewal Dates
                        If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                            'Check if Current Renewal period is defined
                            If IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0 Then
                                'No Current Renewal Period defined so delete this job from this current position.
                                'It should not be in the training list for any current job
                                flgNoCurRnwl = True
                                
                                'Check if the Course was taken before ever
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
                                'SQLQ = SQLQ & " AND ES_JOB = '" & xJob & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND (ES_RENEW = '' OR ES_RENEW IS NULL)"
                                SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
                                SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'Course Taken Before
                                    flgNoCurRnwl = True
                                Else
                                    'Course not taken before
                                    flgNoCurRnwl = False
                                    
                                    Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
'                                    If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
'                                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
'                                    Else
                                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
'                                    End If
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                            Else
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
'                               If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
'                                   rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
'                               Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
'                               End If
                            End If
                        Else
                            'Check if Current Renewal period is defined
                            If IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0 Then
                                'No Current Renewal Period defined so delete this job from this current position.
                                'It should not be in the training list for any current job
                                flgNoCurRnwl = True
                            Else
                                Select Case rsReqCourse("PC_CUR_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_CUR"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                            End If
                        End If
                        If flgNoCurRnwl = False Then
                            'Update Continuing Education record as well
                            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
                            SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsContEdu.EOF Then
                                'rsContEdu("ES_RENEW") = Null
                                rsContEdu("ES_JOB") = xJob
                                rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                rsContEdu("ES_LDATE") = Date
                                rsContEdu("ES_LUSER") = glbUserID
                                rsContEdu("ES_LTIME") = Time$
                                rsContEdu.Update
                            End If
                            rsContEdu.Close
                            Set rsContEdu = Nothing
                            
                            rsHRTrain("TR_JOB") = xJob
'                            If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
'                                rsHRTrain("TR_SDATE") = dlpStartDate.Text
'                            Else
                                rsHRTrain("TR_SDATE") = xStartEndDate
'                            End If
                            rsHRTrain("TR_POS_TYPE") = "C"   'Current Primary
                            ''If Renewal date is greater than today's date then clear the Course Taken Date
                            'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                            '    rsHRTrain("TR_COURSE_TAKEN") = Null
                            'End If
                            rsHRTrain("TR_LDATE") = Date
                            rsHRTrain("TR_LUSER") = glbUserID
                            rsHRTrain("TR_LTIME") = Time$
                            
                            'If follow up id is null then find the id
                            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_EMPNBR = " & xEmpnbr
                                SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                            
                            'Ticket #24300
                            'rsHRTrain.Update
                            
                            'Ticket #24300
                            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                'Add a Follow Up record for this Training course
                                'Ticket #24300
                                'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(xEmpnbr, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                rsHRTrain.Update
                            Else
                                rsHRTrain.Update
                                
                                'Update Follow Up record - Effective Date
                                SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                    rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                    rsFollowUp("EF_LDATE") = Date
                                    rsFollowUp("EF_LUSER") = glbUserID
                                    rsFollowUp("EF_LTIME") = Time$
                                    rsFollowUp.Update
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                        
                            'Update Position record with Follow Up ID
                            'if the course code is TRAIN
                            If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                'Clear the Follow Up from the Previous Job in Primary/Temp Position
                                'Search HR_JOB_HISTORY table for this Position record
                                'and clear with Follow Up Id
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("JH_FOLLOWUP_ID") = Null
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                                
                                'Since Previous in HR_TRAIN can be Primary or Temp Position
                                'Search HR_TEMP_WORK table for this Position record
                                'and clear with Follow Up Id
                                SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("TW_FOLLOWUP_ID") = Null
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                                
                                'Search HR_JOB_HISTORY table for this Position record
                                'and update with Follow Up Id
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & xJobID   'Data1.Recordset("JH_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                            End If
                        Else
                            'No Current renewal found for this course - Correct logic - confirmed with email -March 09, 2009 1:18 PM
                                                        
                            'Clear the Renewal date for this course and for this employee from
                            'Continuing Education screen
                            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
                            SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            'SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & xJob & "'"
                            'SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            If Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            End If
                            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsContEdu.EOF Then
                                rsContEdu("ES_RENEW") = Null
                                rsContEdu("ES_LDATE") = Date
                                rsContEdu("ES_LUSER") = glbUserID
                                rsContEdu("ES_LTIME") = Time$
                                rsContEdu.Update
                                
                                If Not IsNull(rsContEdu("ES_DATCOMP")) And Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                    'Since the course was completed - mark the Follow Up as
                                    'Completed instead of deleting it.
                                    SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    'Ticket #26211 Franks 11/04/2014
                                    'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                
                                    'Clear the Follow Up ID in the Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                        
                                        'Since Previous in HR_TRAIN can be Primary or Temp Position
                                        'Search HR_TEMP_WORK table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("TW_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                            Else
                                'Delete the Follow Up record for this training record
                                'as no Course completion record found
                                SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                'Ticket #26211 Franks 11/04/2014
                                'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                gdbAdoIhr001.Execute SQLQ
                                
                                'Since Previous in HR_TRAIN can be Primary or Temp Position
                                'Clear the Follow Up ID in the Temp/Cross Training Position record
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                    
                                    'Since Previous in HR_TRAIN can be Primary or Temp Position
                                    'Search HR_TEMP_WORK table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            End If
                            rsContEdu.Close
                            Set rsContEdu = Nothing
                            
                            'Delete this Training List record as the course is not required by other positions
                            SQLQ = "DELETE FROM HR_TRAIN"
                            SQLQ = SQLQ & " WHERE TR_EMPNBR = " & xEmpnbr
                            SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            gdbAdoIhr001.Execute SQLQ
                        End If
                    ElseIf xPosType = "Previous" Then
                        'PREVIOUS - Previous
                        'Track for the most recent previous position requiring this course
                        'These courses are for new Previous Primary Position so recalculate the
                        'Renewal Dates
                        xPrvEndDate = Get_Position_End_Date(rsHRTrain("TR_JOB"), rsHRTrain("TR_SDATE"))
                        If Not IsDate(xPrvEndDate) Then xPrvEndDate = rsHRTrain("TR_SDATE")
                        
                        'If CVDate(rsHRTrain("TR_SDATE")) < CVDate(IIf(IsMissing(xStartEndDate), dlpStartDate.Text, xStartEndDate)) Then
'                        If (dlpENDDATE.Text = "") And (IsNull(xEndDate) Or xEndDate = "" Or IsMissing(xEndDate)) Then
                        If (IsNull(xEndDate) Or xEndDate = "" Or IsMissing(xEndDate)) Then
                            'Do nothing
                        Else
'                        If CVDate(xPrvEndDate) < CVDate(IIf(IsMissing(xEndDate) Or xEndDate = "" Or IsNull(xEndDate), dlpENDDATE.Text, xEndDate)) Then
                        If CVDate(xPrvEndDate) < CVDate(xEndDate) Then
                            'Training List has older Position Start Date so update with new Position info.
                            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
'                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
'                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
'                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
'                                End If
                            Else
                                'Check if Previous Renewal period is defined
                                If IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) Or rsReqCourse("PC_RENEW_CRS_PRV") = 0 Then
                                    'No Previous Renewal Period defined so delete this job from this previous position.
                                    'It should not be in the training list for any previous job
                                    flgNoPrvRnwl = True
                                Else
                                    Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                End If
                            End If
                            If flgNoPrvRnwl = False Then
                                'Update Continuing Education record as well
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_JOB") = xJob
                                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                            
                                'Previous Renewal period available
                                rsHRTrain("TR_JOB") = xJob
'                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
'                                    rsHRTrain("TR_SDATE") = dlpStartDate.Text
'                                Else
                                    rsHRTrain("TR_SDATE") = xStartEndDate
'                                End If
                                rsHRTrain("TR_POS_TYPE") = "P"   'Previous Primary
                                ''If Renewal date is greater than today's date then clear the Course Taken Date
                                'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                                '    rsHRTrain("TR_COURSE_TAKEN") = Null
                                'End If
                                rsHRTrain("TR_LDATE") = Date
                                rsHRTrain("TR_LUSER") = glbUserID
                                rsHRTrain("TR_LTIME") = Time$
                                
                                'If follow up id is null then find the id
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & xEmpnbr
                                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                    SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Ticket #24300
                                'rsHRTrain.Update
                                
                                'Ticket #24300
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    'Add a Follow Up record for this Training course
                                    'Ticket #24300
                                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(xEmpnbr, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                    rsHRTrain.Update
                                Else
                                    rsHRTrain.Update
                                    
                                    'Update Follow Up record - Effective Date
                                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                        rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                        rsFollowUp("EF_LDATE") = Date
                                        rsFollowUp("EF_LUSER") = glbUserID
                                        rsFollowUp("EF_LTIME") = Time$
                                        rsFollowUp.Update
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Update Position record with Follow Up ID
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Clear the Follow Up from the Previous Job in Primary/Temp Position
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                    
                                    'Since Previous in HR_TRAIN can be Primary or Temp Position
                                    'Search HR_TEMP_WORK table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                    
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and update with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & xJobID   'Data1.Recordset("JH_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            Else
                                'No Previous renewal found for this course
                                
                                'Clear the Renewal date for this course and for this employee from
                                'Continuing Education screen
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                    
                                    If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                                        'Since the course was completed - mark the Follow Up as
                                        'Completed instead of deleting it.
                                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    Else
                                        'Delete the Follow Up record for this training record
                                        'as no Course completion record found
                                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                        'Ticket #26211 Franks 11/04/2014
                                        'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    
                                        'Clear the Follow Up ID in the Position record
                                        'if the course code is TRAIN
                                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                            'Search HR_JOB_HISTORY table for this Position record
                                            'and clear with Follow Up Id
                                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                            If Not rsTJob.EOF Then
                                                rsTJob("JH_FOLLOWUP_ID") = Null
                                                rsTJob.Update
                                            End If
                                            rsTJob.Close
                                            Set rsTJob = Nothing
                                            
                                            'Since Previous in HR_TRAIN can be Primary or Temp Position
                                            'Search HR_TEMP_WORK table for this Position record
                                            'and clear with Follow Up Id
                                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                            If Not rsTJob.EOF Then
                                                rsTJob("TW_FOLLOWUP_ID") = Null
                                                rsTJob.Update
                                            End If
                                            rsTJob.Close
                                            Set rsTJob = Nothing
                                        End If
                                    End If
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                                                
                                    'Since Previous in HR_TRAIN can be Primary or Temp Position
                                    'Clear the Follow Up ID in the Temp/Cross Training Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                        
                                        'Since Previous in HR_TRAIN can be Primary or Temp Position
                                        'Search HR_TEMP_WORK table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("TW_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                
                                'Delete this Training List record as the course is not required by other positions
                                SQLQ = "DELETE FROM HR_TRAIN"
                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & xEmpnbr
                                SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                gdbAdoIhr001.Execute SQLQ
                            End If
                        
                        Else
                            'Do not do anything because Training List has most recent Position Start Date
                        End If
                        End If
                    End If
                ElseIf IsNull(rsHRTrain("TR_POS_TYPE")) Or rsHRTrain("TR_POS_TYPE") = "" Then
                    'Check if the course was taken before. If taken then use the normal Training List logic based
                    'on the renewal date if the course should continue to exist or not
                    If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                        'COURSE NEVER TAKEN BEFORE
                        'It's an independent course and never taken before, update with this Position's information
                        'even though there is no renewal period for the type of position this is
                        rsHRTrain("TR_JOB") = xJob
'                        If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
'                            rsHRTrain("TR_SDATE") = dlpStartDate.Text
'                        Else
                            rsHRTrain("TR_SDATE") = xStartEndDate
'                        End If
                        If xPosType = "Current" Then
                            rsHRTrain("TR_POS_TYPE") = "C"
                        ElseIf xPosType = "Temporary" Then
                            rsHRTrain("TR_POS_TYPE") = "T"
                        ElseIf xPosType = "Previous" Then
                            rsHRTrain("TR_POS_TYPE") = "P"
                        End If
    
                        'Do not overwrite the Renewal Date entered for this independent course
                        'rsHRTrain("TR_RENEW")) =
                        rsHRTrain("TR_LDATE") = Date
                        rsHRTrain("TR_LUSER") = glbUserID
                        rsHRTrain("TR_LTIME") = Time$
                        
                        'If follow up id is null then find the id
                        If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                            xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                            SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_EMPNBR = " & xEmpnbr
                            SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                            SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(rsHRTrain("TR_RENEW"))
                            rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsFollowUp.EOF Then
                                rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                            End If
                            rsFollowUp.Close
                            Set rsFollowUp = Nothing
                        End If
                        
                        'Ticket #24300
                        'rsHRTrain.Update
                        
                        'Ticket #24300
                        If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                            'Add a Follow Up record for this Training course
                            'Ticket #24300
                            'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                            rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(xEmpnbr, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                            rsHRTrain.Update
                        Else
                            rsHRTrain.Update
                        
                            'Update Follow Up record - Comments with Position
                            SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsFollowUp.EOF Then
                                'No change to renewal date
                                'rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                rsFollowUp("EF_LDATE") = Date
                                rsFollowUp("EF_LUSER") = glbUserID
                                rsFollowUp("EF_LTIME") = Time$
                                rsFollowUp.Update
                            End If
                            rsFollowUp.Close
                            Set rsFollowUp = Nothing
                        End If
                    
                        'Update Position record with Follow Up ID
                        'if the course code is TRAIN
                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                            'Clear the Follow Up from the Previous Job in Primary/Temp Position
                            'Search HR_JOB_HISTORY table for this Position record
                            'and clear with Follow Up Id
                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                rsTJob("JH_FOLLOWUP_ID") = Null
                                rsTJob.Update
                            End If
                            rsTJob.Close
                            Set rsTJob = Nothing
                            
                            'Since Previous in HR_TRAIN can be Primary or Temp Position
                            'Search HR_TEMP_WORK table for this Position record
                            'and clear with Follow Up Id
                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                rsTJob("TW_FOLLOWUP_ID") = Null
                                rsTJob.Update
                            End If
                            rsTJob.Close
                            Set rsTJob = Nothing
                            
                            'Search HR_JOB_HISTORY table for this Position record
                            'and update with Follow Up Id
                            If xPosType = "Temporary" Then
                                SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & xJobID   'Data1.Recordset("JH_ID")
                            Else
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & xJobID   'Data1.Recordset("JH_ID")
                            End If
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                If xPosType = "Temporary" Then
                                    rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                Else
                                    rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                End If
                                rsTJob.Update
                            End If
                            rsTJob.Close
                            Set rsTJob = Nothing
                        End If
                    Else
                        'COURSE TAKEN BEFORE
                        'Which kind of Position is this
                        If xPosType = "Current" Or xPosType = "Temporary" Then
                            'Check if Current Renewal period is defined
                            If IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0 Then
                                'No Current Renewal Period defined so delete this job from this current position.
                                'It should not be in the training list for any current job
                                flgNoCurRnwl = True
                            Else
                                Select Case rsReqCourse("PC_CUR_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_CUR"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                flgNoCurRnwl = False
                            End If
                        ElseIf xPosType = "Previous" Then
                            'Check if Previous Renewal period is defined
                            If IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) Or rsReqCourse("PC_RENEW_CRS_PRV") = 0 Then
                                'No Previous Renewal Period defined so delete this job from this previous position.
                                'It should not be in the training list for any previous job
                                flgNoCurRnwl = True
                            Else
                                Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                flgNoCurRnwl = False
                            End If
                        End If
                        
                        If flgNoCurRnwl = False Then
                            'Renewal Period Found - updated existing records
                            'Update Continuing Education record as well
                            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
                            SQLQ = SQLQ & " AND (ES_JOB = '' OR ES_JOB IS NULL)"    'No Job - Independent course
                            SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsContEdu.EOF Then
                                'rsContEdu("ES_RENEW") = Null
                                rsContEdu("ES_JOB") = xJob
                                rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                rsContEdu("ES_LDATE") = Date
                                rsContEdu("ES_LUSER") = glbUserID
                                rsContEdu("ES_LTIME") = Time$
                                rsContEdu.Update
                            End If
                            rsContEdu.Close
                            Set rsContEdu = Nothing
                            
                            'Renewal Period available
                            rsHRTrain("TR_JOB") = xJob
'                            If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
'                                rsHRTrain("TR_SDATE") = dlpStartDate.Text
'                            Else
                                rsHRTrain("TR_SDATE") = xStartEndDate
'                            End If
                            If xPosType = "Current" Then
                                rsHRTrain("TR_POS_TYPE") = "C"
                            ElseIf xPosType = "Temporary" Then
                                rsHRTrain("TR_POS_TYPE") = "T"
                            ElseIf xPosType = "Previous" Then
                                rsHRTrain("TR_POS_TYPE") = "P"
                            End If
                            
                            rsHRTrain("TR_LDATE") = Date
                            rsHRTrain("TR_LUSER") = glbUserID
                            rsHRTrain("TR_LTIME") = Time$
                            
                            'If follow up id is null then find the id
                            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_EMPNBR = " & xEmpnbr
                                SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                            
                            'Ticket #24300
                            'rsHRTrain.Update
                            
                            'Ticket #24300
                            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                'Add a Follow Up record for this Training course
                                'Ticket #24300
                                'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(xEmpnbr, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                rsHRTrain.Update
                            Else
                                rsHRTrain.Update
                            
                                'Update Follow Up record - Effective Date
                                SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                    rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                    rsFollowUp("EF_LDATE") = Date
                                    rsFollowUp("EF_LUSER") = glbUserID
                                    rsFollowUp("EF_LTIME") = Time$
                                    rsFollowUp.Update
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                        
                            'Update Position record with Follow Up ID
                            'if the course code is TRAIN
                            If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                'Clear the Follow Up from the Previous Job in Primary/Temp Position
                                'Search HR_JOB_HISTORY table for this Position record
                                'and clear with Follow Up Id
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("JH_FOLLOWUP_ID") = Null
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                                
                                'Since Previous in HR_TRAIN can be Primary or Temp Position
                                'Search HR_TEMP_WORK table for this Position record
                                'and clear with Follow Up Id
                                SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("TW_FOLLOWUP_ID") = Null
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                                
                                'Search HR_JOB_HISTORY table for this Position record
                                'and update with Follow Up Id
                                If xPosType = "Temporary" Then
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_ID = " & xJobID   'Data1.Recordset("JH_ID")
                                Else
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & xJobID   'Data1.Recordset("JH_ID")
                                End If
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    If xPosType = "Temporary" Then
                                        rsTJob("TW_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                    Else
                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                    End If
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                            End If
                        Else
                            'No Renewal Period found for this course
                                                        
                            'Clear the Renewal date for this course and for this employee from
                            'Continuing Education screen
                            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & xEmpnbr
                            SQLQ = SQLQ & " AND (ES_JOB = '' OR ES_JOB IS NULL)"    'Independent course
                            SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            'SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND TR_JOB = '" & xJob & "'"
                            'SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            If Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            End If
                            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsContEdu.EOF Then
                                rsContEdu("ES_RENEW") = Null
                                rsContEdu("ES_LDATE") = Date
                                rsContEdu("ES_LUSER") = glbUserID
                                rsContEdu("ES_LTIME") = Time$
                                rsContEdu.Update
                                
                                If Not IsNull(rsContEdu("ES_DATCOMP")) And Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                    'Since the course was completed - mark the Follow Up as
                                    'Completed instead of deleting it.
                                    SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    'Ticket #26211 Franks 11/04/2014
                                    'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                
                                    'Clear the Follow Up ID in the Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                        
                                        'Since Previous in HR_TRAIN can be Primary or Temp Position
                                        'Search HR_TEMP_WORK table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("TW_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                            Else
                                'Delete the Follow Up record for this training record
                                'as no Course completion record found
                                SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                'Ticket #26211 Franks 11/04/2014
                                'SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & xEmpnbr & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                gdbAdoIhr001.Execute SQLQ
                                
                                'Since Previous in HR_TRAIN can be Primary or Temp Position
                                'Clear the Follow Up ID in the Temp/Cross Training Position record
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                    
                                    'Since Previous in HR_TRAIN can be Primary or Temp Position
                                    'Search HR_TEMP_WORK table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            End If
                            rsContEdu.Close
                            Set rsContEdu = Nothing
                            
                            'Delete this Training List record as the course is not required by other positions
                            SQLQ = "DELETE FROM HR_TRAIN"
                            SQLQ = SQLQ & " WHERE TR_EMPNBR = " & xEmpnbr
                            SQLQ = SQLQ & " AND (TR_JOB = '' OR TR_JOB IS NULL)"    'Independent course
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            gdbAdoIhr001.Execute SQLQ
                        End If
                        
                    End If
                End If
                
                If flgProcCalled = False Then
                    xPosType = xorgPosType
                    xJob = xorgJob
                End If
                
            End If
            rsHRTrain.Close
            Set rsHRTrain = Nothing
            
Next_Required_Course:
            rsCourseMst.Close
            Set rsCourseMst = Nothing

            rsReqCourse.MoveNext
        Loop
    End If
    rsReqCourse.Close
    Set rsReqCourse = Nothing

Exit Sub

Employee_Job_Training_Err:
If Err = 3018 Then
    Err = 0
    Resume Next
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
If Len(SQLQ) = 0 Then
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update_Employee_Job_Training_List", "HR_JOB_COURSE", "Update_Emp_Job_Training_List")
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, SQLQ, "HR_JOB_COURSE", "Update_Emp_Job_Training_List")
End If
Call RollBack '26July99 js
End Sub


