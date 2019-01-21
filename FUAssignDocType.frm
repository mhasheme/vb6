VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmUAssignDocType 
   AutoRedraw      =   -1  'True
   Caption         =   "Attach Document Type"
   ClientHeight    =   11490
   ClientLeft      =   15
   ClientTop       =   1020
   ClientWidth     =   13590
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   ScaleHeight     =   11490
   ScaleWidth      =   13590
   WindowState     =   2  'Maximized
   Begin VB.OptionButton optHSF7WrittenOfr 
      Caption         =   "Health && Safety - Written Modified Offer (Form 7)"
      Height          =   375
      Left            =   360
      TabIndex        =   31
      ToolTipText     =   "Return to Work - Modified Work Offer"
      Top             =   9720
      Width           =   4100
   End
   Begin VB.OptionButton optHSF7Concerns 
      Caption         =   "Health && Safety - Concerned about Claims Written Submission (Form 7)"
      Height          =   375
      Left            =   360
      TabIndex        =   30
      ToolTipText     =   "Concerned about the Claims - Written Submission"
      Top             =   9360
      Width           =   5295
   End
   Begin VB.OptionButton optHSIncidents 
      Caption         =   "Health && Safety - Incidents"
      Height          =   375
      Left            =   360
      TabIndex        =   29
      Top             =   9000
      Width           =   4100
   End
   Begin VB.OptionButton optPosSkills 
      Caption         =   "Position Skills"
      Height          =   375
      Left            =   360
      TabIndex        =   26
      Top             =   8400
      Width           =   4100
   End
   Begin VB.OptionButton optTermination 
      Caption         =   "Termination"
      Height          =   375
      Left            =   360
      TabIndex        =   25
      Top             =   8040
      Width           =   4100
   End
   Begin VB.CheckBox chkReplace 
      Alignment       =   1  'Right Justify
      Caption         =   "Replace Existing Document Type Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   180
      TabIndex        =   2
      Top             =   1080
      Value           =   1  'Checked
      Width           =   4395
   End
   Begin VB.TextBox txtDocDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "00-Dcoument Description"
      Top             =   720
      Width           =   5295
   End
   Begin VB.ListBox lstEmpFlagsList 
      Height          =   1425
      ItemData        =   "FUAssignDocType.frx":0000
      Left            =   1875
      List            =   "FUAssignDocType.frx":0002
      TabIndex        =   6
      Top             =   3240
      Width           =   2655
   End
   Begin VB.OptionButton optComments 
      Caption         =   "Comments"
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   7680
      Width           =   4100
   End
   Begin VB.OptionButton optCounselling 
      Caption         =   "Counseling"
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   7320
      Width           =   4100
   End
   Begin VB.OptionButton optFormalEdu 
      Caption         =   "Formal Education"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   6960
      Width           =   4100
   End
   Begin VB.OptionButton optContEdu 
      Caption         =   "Continuing Education"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   6600
      Width           =   4100
   End
   Begin VB.OptionButton optAssociation 
      Caption         =   "Associations"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   6240
      Width           =   4100
   End
   Begin VB.OptionButton optAttendance 
      Caption         =   "Attendance"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   5880
      Width           =   4100
   End
   Begin VB.OptionButton optDollarEnt 
      Caption         =   "Dollar Entitlements"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   5520
      Width           =   4100
   End
   Begin VB.OptionButton optPerfReview 
      Caption         =   "Performance"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   5160
      Width           =   4100
   End
   Begin VB.OptionButton optJobOffer 
      Caption         =   "Job Offer (Position screen)"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   4800
      Width           =   4100
   End
   Begin VB.OptionButton optEmpFlags 
      Caption         =   "Employee Flags"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Width           =   4100
   End
   Begin VB.OptionButton optOtherInfo 
      Caption         =   "Other Information (Other Information screen)"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   4100
   End
   Begin VB.OptionButton optResume 
      Caption         =   "Resume (Status/Dates screen)"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   4100
   End
   Begin VB.OptionButton optImportPhoto 
      Caption         =   "Import Photo into info:HR database"
      Height          =   375
      Left            =   9600
      TabIndex        =   45
      Top             =   10680
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Frame frImpPhoto 
      Caption         =   "Attachment File"
      Height          =   855
      Left            =   9600
      TabIndex        =   44
      Top             =   9720
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CheckBox chkFile 
         Caption         =   "File Names are equal to Employee Numbers"
         Height          =   315
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3555
      End
   End
   Begin VB.CheckBox chkImpWord 
      Caption         =   "Import Resume File"
      Height          =   315
      Left            =   9720
      TabIndex        =   33
      Top             =   12720
      Visible         =   0   'False
      Width           =   2475
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      TabIndex        =   38
      Top             =   11010
      Width           =   13590
      _Version        =   65536
      _ExtentX        =   23971
      _ExtentY        =   847
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
         Left            =   8520
         Top             =   120
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
      End
   End
   Begin VB.Frame frmFile 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   9540
      TabIndex        =   39
      Top             =   11100
      Visible         =   0   'False
      Width           =   7695
      Begin INFOHR_Controls.EmployeeLookup elpEEID 
         Height          =   315
         Left            =   1740
         TabIndex        =   36
         Top             =   90
         Width           =   5000
         _ExtentX        =   8811
         _ExtentY        =   556
         RefreshDescriptionWhen=   2
      End
      Begin VB.TextBox txtFileName 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         MaxLength       =   32
         TabIndex        =   37
         Tag             =   "00-File Name (Do not Enter Extension TXT)"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Number"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   150
         Width           =   1290
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Do Not Enter the file Extension (Must be 'JPG')."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3720
         TabIndex        =   41
         Top             =   480
         Width           =   4260
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Import File Name"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   525
         Width           =   1620
      End
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID_Del 
      Height          =   285
      Left            =   12480
      TabIndex        =   34
      Tag             =   "10-Enter Employee Number"
      Top             =   9840
      Visible         =   0   'False
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   6355
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   12480
      TabIndex        =   35
      Tag             =   "00-Section"
      Top             =   10185
      Visible         =   0   'False
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   2205
      TabIndex        =   0
      Tag             =   "01-Document Type Code "
      Top             =   360
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "DOCT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   6240
      TabIndex        =   12
      Tag             =   "01-Entitlement - Code"
      Top             =   5565
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOL"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   6240
      TabIndex        =   14
      Tag             =   "01-Attendance Reason"
      Top             =   5925
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ADRE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   6240
      TabIndex        =   16
      Tag             =   "01-Association/Membership- Code"
      Top             =   6285
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "TDCD"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   6240
      TabIndex        =   18
      Tag             =   "00-Course Code"
      Top             =   6645
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESCD"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   6240
      TabIndex        =   20
      Tag             =   "01-School - Code"
      Top             =   7005
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EUSC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   7
      Left            =   6240
      TabIndex        =   22
      Tag             =   "01-Counselling Type- Code"
      Top             =   7365
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "CETY"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   8
      Left            =   6240
      TabIndex        =   24
      Tag             =   "01-Comment Type- Code"
      Top             =   7725
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECOM"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   9
      Left            =   6240
      TabIndex        =   27
      Tag             =   "01-Skill- Code"
      Top             =   8445
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSK"
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Left            =   6240
      TabIndex        =   28
      Tag             =   "01-Position code"
      Top             =   8760
      Visible         =   0   'False
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   5
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   10
      Left            =   6240
      TabIndex        =   8
      Tag             =   "01-Reason for change in position - Code"
      Top             =   4845
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDRC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   11
      Left            =   6240
      TabIndex        =   10
      Tag             =   "00-Performance Rating - Code "
      Top             =   5205
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDPC"
   End
   Begin VB.Label lblJob 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      Height          =   195
      Left            =   4710
      TabIndex        =   53
      Top             =   8805
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"FUAssignDocType.frx":0004
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   795
      Left            =   5400
      TabIndex        =   52
      Top             =   1890
      Width           =   7140
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4710
      TabIndex        =   51
      Top             =   5250
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Document Description"
      Height          =   195
      Left            =   180
      TabIndex        =   50
      Top             =   765
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Document Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   49
      Top             =   405
      Width           =   1350
   End
   Begin VB.Label lblEmpFlagList 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select one Employee Flag to update Document Type for:"
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   630
      TabIndex        =   48
      Top             =   3240
      Width           =   1185
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Update Document Type Info. on:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   47
      Top             =   1920
      Width           =   2805
   End
   Begin VB.Label lblSec 
      AutoSize        =   -1  'True
      Caption         =   "Section"
      Height          =   195
      Left            =   9660
      TabIndex        =   46
      Top             =   12390
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   9660
      TabIndex        =   43
      Top             =   12045
      Visible         =   0   'False
      Width           =   1290
   End
End
Attribute VB_Name = "frmUAssignDocType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FPath, UPDTCNT
Dim ImportFile As String
Dim xDeleteable
Dim flgWrongDocTypeCode As Boolean

'Private Sub chkExportPhotos_Click()
'
'    lblPath.Visible = False
'    Drive1.Visible = False
'    Dir1.Visible = False
'
'    'Ticket #26315 Franks 11/26/2014 - Jerry asked to make this function generic in 8.1
'    'If glbWFC Then 'Ticket #26308 Franks 11/21/2014
'        Call WFCPhotoScreen(False)
'    'End If
'
'    Call SET_UP_MODE
'End Sub

Private Sub WFCPhotoScreen(xFlag) 'Ticket #26308 Franks 11/21/2014
    elpEEID_Del.Enabled = xFlag 'True
    lblEENum(1).Enabled = xFlag 'True
    lblSec.Visible = xFlag
    clpCode(0).Visible = xFlag
    lblSec.Caption = lStr("Section")
End Sub

Private Sub chkFile_Click()
    frmFile.Visible = 1 - chkFile.Value
    xDeleteable = 1 - chkFile.Value
    
    Call SET_UP_MODE
    
    'cmdDelete.Visible = 1 - chkFile.Value
    If chkFile Then
'        chkDelete.Value = vbUnchecked
    End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_GotFocus()
 Call SetPanHelp(Me.ActiveControl)
End Sub

Public Sub cmdDelete_Click()
Dim SQLQ As String, X
Dim Title$, Msg$, DgDef As Variant, Response%
Dim xESQLQ As String

Title = "Employee Photo Delete"


On Error GoTo Mod_Err

'If elpEEID.Caption = "Unassinged" Then
'    MsgBox "Employee Number is not valid."
'    elpEEID.SetFocus
'    Exit Sub
'End If
    
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
'Msg$ = "Are you sure you want to Delete " & elpEEID.Caption & "'s Photo?"
'8.0 - Ticket #22682 - Remove photos from HR_PHOTO table because we are moving to folder now
'If chkDeleteAll.Value = vbChecked Then
'    Msg$ = "Are you sure you want to DELETE ALL Employees Photos from info:HR database?"
'Else
    Msg$ = "Are you sure you want to Delete Employee's Document from info:HR database?"
'End If
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then Exit Sub

'If chkDeleteAll.Value = vbChecked Then
'    MDIMain.panHelp(0).Caption = "Deleting Employees Photos from info:HR database, please wait....."
'Else
    MDIMain.panHelp(0).Caption = "Deleting Employee's Document from info:HR database, please wait....." '
'End If

Screen.MousePointer = HOURGLASS

'8.0 - Ticket #22682 - Remove photos from HR_PHOTO table because we are moving to folder now
'If chkDeleteAll.Value = vbChecked Then
'    'As per Department security
'    xESQLQ = glbSeleDeptUn
'    gdbAdoIhr001.Execute "DELETE FROM HR_PHOTO WHERE PT_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & xESQLQ & ")"
'Else
    gdbAdoIhr001.Execute "DELETE FROM HR_PHOTO WHERE PT_EMPNBR IN (" & getEmpnbr(elpEEID_Del) & ")"
'End If

Screen.MousePointer = DEFAULT

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

'8.0 - Ticket #22682 - Remove photos from HR_PHOTO table because we are moving to folder now
'If chkDeleteAll.Value = vbChecked Then
'    MsgBox "ALL Employees Document DELETED from info:HR database successfully."
'Else
    MsgBox "Employee's Document Deleted from info:HR database successfully."
'End If

Exit Sub

Mod_Err:
If Err = 53 Then Resume Next

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDelete", "Attachment", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Sub cmdModify_Click()
Dim SQLQ As String, X
Dim Title$, Msg$, DgDef As Variant, Response%

'If chkImpWord Then
'    Title = "Employee Resume Import"
'    'If Not gSec_Import_Attendance Then
'    '    MsgBox "You Do Not Have Authority For This Transacaction"
'    '    Exit Sub
'    'End If
'
'    On Error GoTo Mod_Err
'
'    If Not chkPhoto() Then Exit Sub
'
'    DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
'    Msg$ = "Are you sure you want to Import Resume?"
'    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
'    If Response% = IDNO Then Exit Sub
'
'    Screen.MousePointer = HOURGLASS
'
'    ChDir FPath
'    If Not modUpdateSelectionResume() Then GoTo bpMod
'
'    MDIMain.panHelp(0).FloodPercent = 100
'
'    Close
'    '-----------------------------------------------------
'
'    Screen.MousePointer = DEFAULT
'    MDIMain.panHelp(0).FloodType = 0
'    MDIMain.panHelp(1).Caption = " Update Completed"
'    MDIMain.panHelp(2).Caption = ""
'    If UPDTCNT = 0 Then
'        Msg$ = "No Photo Imported "
'    Else
'        Msg$ = Str(UPDTCNT)
'        If UPDTCNT = 1 Then Msg$ = Msg$ & " Record " Else Msg$ = Msg$ & " Records "
'        Msg$ = Msg$ & "Imported Successfully "
'    End If
'    DgDef = MB_ICONINFORMATION
'    MsgBox Msg$, DgDef, Title
'Else
        Title = "Update Document Information"
        'If Not gSec_Import_Attendance Then
        '    MsgBox "You Do Not Have Authority For This Transacaction"
        '    Exit Sub
        'End If
    
        On Error GoTo Mod_Err
        
        If Not chkImpAttachment() Then Exit Sub
        
        DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
        Msg$ = "Are you sure you want to update Employee's Document with this Document Type Information?"
        Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
        If Response% = IDNO Then Exit Sub
    
        Screen.MousePointer = HOURGLASS
    
'        ChDir FPath
        glbUPDTCNT = 0
'        If Not Import_Attachment_Files() Then GoTo bpMod
        
        'Update Document Type Informnation
        '??????? Different approach to updating the Employee Flags document information as there could be multiple Employee Flags selected.
        
        glbDocName = ""
        If optResume Then
            glbDocName = "Resume"
        ElseIf optOtherInfo Then
            glbDocName = "OtherInfo"
        ElseIf optEmpFlags Then
            glbDocName = "EmployeeFlag"
            glbEmpFlagNo = Get_SelectedEmployeeFlag
            If glbEmpFlagNo <> -1 Then
            Else
                glbDocName = ""
            End If
        ElseIf optJobOffer Then
            glbDocName = "Offer"
        ElseIf optPerfReview Then
            glbDocName = "Performance"
        ElseIf optDollarEnt Then
            glbDocName = "DollarEnt"
        ElseIf optAttendance Then
            glbDocName = "Attendance"
        ElseIf optAssociation Then
            glbDocName = "Associations"
        ElseIf optContEdu Then
            glbDocName = "EdSem"
        ElseIf optFormalEdu Then
            glbDocName = "FormalEdu"
        ElseIf optCounselling Then
            glbDocName = "Counsel"
        ElseIf optComments Then
            glbDocName = "Comments"
        ElseIf optTermination Then
            glbDocName = "Termination"
        ElseIf optPosSkills Then
            glbDocName = "PositionSkill"
        ElseIf optHSIncidents Then
            glbDocName = "INCIDENT"
        ElseIf optHSF7Concerns Then
            glbDocName = "INJURYWF7"
        ElseIf optHSF7WrittenOfr Then
            glbDocName = "INJURYWF7_WRITTENOFR"
        End If
        
        If Len(glbDocName) > 0 Then
            Call UpdateDocumentTypeInfo(glbDocName, clpCode(1).Text, txtDocDesc.Text)
        Else
            MsgBox "No Document found to update Document Type Information with."
        End If
    
        MDIMain.panHelp(0).FloodPercent = 100
    
        '????
        Close
        '-----------------------------------------------------
    
        Screen.MousePointer = DEFAULT
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " Update Completed"
        MDIMain.panHelp(2).Caption = ""
        
        'If UPDTCNT = 0 Then
        If glbUPDTCNT = 0 Then
            Msg$ = "No Document found to update with Document Type Information."
        Else
            'Msg$ = Str(UPDTCNT)
            Msg$ = Str(glbUPDTCNT)
            'If UPDTCNT = 1 Then Msg$ = Msg$ & " Document " Else Msg$ = Msg$ & " Documents "
            If glbUPDTCNT = 1 Then Msg$ = Msg$ & " Document's " Else Msg$ = Msg$ & " Documents "
            Msg$ = Msg$ & "Document Type Information got Updated Successfully."
        End If
        DgDef = MB_ICONINFORMATION
        MsgBox Msg$, DgDef, Title
'End If

bpMod:

Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Exit Sub

Mod_Err:
If Err = 53 Then Resume Next

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Document Attachment", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
    If Len(clpCode(1).Text) > 0 Then
        Dim rs As New ADODB.Recordset
        Dim strSQL As String
        Dim xWrongPos, xPos, I
        Dim xList, xShowCell, xCell
        Dim xTemplate As String
        
        'If Not clpCode(1).ListChecker Then Exit Sub
        If clpCode(1).Caption = "Unassigned" Then Exit Sub
        
        '????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
        xTemplate = ""
        xTemplate = Get_Template(glbUserID)
        
        
        xList = clpCode(1).Text
        xWrongPos = 0
        xPos = 0
        Do While Len(xList) <> 0
            xWrongPos = xWrongPos + xPos
            xPos = InStr(xList, ",")
            If xPos = 0 Then
                xShowCell = xList
                xList = ""
            Else
                xShowCell = Left(xList, xPos - 1)
                xList = Mid(xList, xPos + 1)
            End If
            xCell = xShowCell
            
            If xTemplate = "" Or xTemplate = "TEMPLATE" Then
                strSQL = "SELECT ACCESSABLE FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(glbUserID, "'", "''") & "'"
            Else
                '????Ticket #24808 -  Retrieve template's security profile
                strSQL = "SELECT ACCESSABLE FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
            End If
            strSQL = strSQL & " AND CODENAME = '" & xCell & "' AND TB_NAME='DOCT'"
            rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
            If rs.EOF = False And rs.BOF = False Then
                If rs("ACCESSABLE") = 0 Then
                    flgWrongDocTypeCode = True
                    MsgBox "You do not have Authorization for '" & xCell & "' Document Type code", vbInformation + vbOKOnly, "Authorization Failure"
                    SendKeys "{HOME}"
                    For I = 1 To xWrongPos
                        SendKeys "{Right}"
                    Next
                    Exit Sub
                End If
            Else
                flgWrongDocTypeCode = True
                MsgBox "You do not have Authorization for '" & xCell & "' Document Type code", vbInformation + vbOKOnly, "Authorization Failure"
                SendKeys "{HOME}"
                For I = 1 To xWrongPos
                    SendKeys "{Right}"
                Next
                Exit Sub
            End If
            rs.Close
            Set rs = Nothing
        Loop
    End If

End Sub

'Private Sub Export_Photos()
'    Dim xPath As String
'    Dim Response%
'
'    'Make sure Export existing photos is selected
'    If chkExportPhotos.Value = vbChecked Then
'        'Get the user to enter the path to export the Photos to and
'        'Update that path to Company Pref. 'EMPLOYEEPHOTOPATH'
'        xPath = UCase(Dir1.Path) & UCase(IIf(Right(Dir1.Path, 1) = "\", "", "\"))
'
'        'Verify the export folder
'        'Ticket #26315 Franks 11/26/2014 - Jerry asked to make this function generic in 8.1
'        'If glbWFC Then 'Ticket #26308 Franks 11/21/2014
'            Response% = MsgBox("Are you sure you want to export these Employees Photos from info:HR database to '" & xPath & "' folder?", vbQuestion + vbYesNo, "Confirm Employees Photos export folder")
'        'Else
'        '    Response% = MsgBox("Are you sure you want to export all Employees Photos from info:HR database to '" & xPath & "' folder?", vbQuestion + vbYesNo, "Confirm Employees Photos export folder")
'        'End If
'        If Response% = vbNo Then Exit Sub
'
'        'Export photos
'        Screen.MousePointer = HOURGLASS
'
'        MDIMain.panHelp(0).Caption = "Exporting Photos, please wait...."
'
'        Call Export_Photos_FromDB(xPath)
'
'        MDIMain.panHelp(0).Caption = ""
'
'        Screen.MousePointer = DEFAULT
'    Else
'        MsgBox "To 'Export/Delete Photo from info:HR database', one of the 'Export/Delete Photo' checkboxes should be checked.", vbExclamation
'    End If
'End Sub

'Private Sub Export_Photos_FromDB(xAppPath)
'    Dim AppPath
'    Dim rsPhoto As New ADODB.Recordset
'    Dim byteChunk() As Byte
'
'    Dim FileNumber As Integer
'    Dim TempFile As String
'    Dim TempDir As String * 255
'
'    Dim rsPrefer As New ADODB.Recordset
'    Dim SQLQ As String
'    Dim xESQLQ As String
'
'    'Path user selected to export the Photos into
'    AppPath = xAppPath
'
'    'As per Department security
'    xESQLQ = glbSeleDeptUn
'
'    'Retrieve Photos of each employee
'    SQLQ = "SELECT * FROM HR_PHOTO WHERE PT_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & xESQLQ & ")"
'    'Ticket #26315 Franks 11/26/2014 - Jerry asked to make this function generic in 8.1
'    'If glbWFC Then 'Ticket #26308 Franks 11/21/2014
'        If Len(elpEEID_Del.Text) > 0 Then
'            SQLQ = SQLQ & " AND PT_EMPNBR IN (" & getEmpnbr(elpEEID_Del) & ") "
'        End If
'        If clpCode(0).Visible And Len(clpCode(0).Text) > 0 Then
'            SQLQ = SQLQ & " AND PT_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
'        End If
'    'End If
'    rsPhoto.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
'    If rsPhoto.EOF Then Exit Sub
'    Do While Not rsPhoto.EOF
'        'Set File Name using the Employee #
'        TempFile = AppPath & rsPhoto("PT_EMPNBR") & ".jpg"
'
'        'If file already exists, delete it
'        If (Dir(TempFile)) <> "" Then Kill TempFile
'
'        FileNumber = FreeFile
'        Open TempFile For Binary Access Write As FileNumber
'
'        ReDim byteChunk(rsPhoto("PT_PHOTO").ActualSize)
'        byteChunk() = rsPhoto("PT_PHOTO").GetChunk(rsPhoto("PT_PHOTO").ActualSize)
'        Put FileNumber, , byteChunk()
'
'        Close FileNumber
'
'        rsPhoto.MoveNext
'    Loop
'    rsPhoto.Close
'    Set rsPhoto = Nothing
'
'    'Update Company Pref with Employee Photo path
'    If glbWFC Then 'Ticket #26308 Franks 11/21/2014
'        Screen.MousePointer = DEFAULT
'        MsgBox "   Finished!   "
'    Else
'        SQLQ = "SELECT * FROM HRPREFERENCE WHERE HP_FUN_NAME = 'EMPLOYEEPHOTOPATH'"
'        rsPrefer.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'        If Not rsPrefer.EOF Then
'            rsPrefer("HP_EMAIL") = AppPath
'            rsPrefer.Update
'        End If
'        rsPrefer.Close
'        Set rsPrefer = Nothing
'
'        MDIMain.panHelp(0).Caption = "Employees Photos export completed."
'
'        Screen.MousePointer = DEFAULT
'
'        MsgBox "Employees Photos exported from info:HR database successfully." & vbCrLf & vbCrLf & "To view Employee's Photo in info:HR, please turn it ON from the 'Company Preference' screen under the 'Setup' menu in info:HR.", vbInformation, "Turn-ON Employee Photo view option"
'    End If
'End Sub

'Private Sub optImportPhoto_Click()
'    If optImportPhoto Then
'        frImpPhoto.Visible = True
''        frmDelExpPhotos.Visible = False
'
'        frmFile.Top = 2940
'        frmFile.Visible = False
'        lblPath.Caption = "Import From Path"
'
'        lblPath.Visible = True
'        Drive1.Visible = True
'        Dir1.Visible = True
'        File1.Visible = True
'
'        lblEENum(1).Visible = False
'        elpEEID_Del.Visible = False
'
'        'Feb 11th 2014: From the Meeting with Jerry and Mostafa, since Mostafa said ESS/Timesheet cannot read
'        'a folder thatdoes not reside on the web server which needs to use Network Service account (local account)
'        'to access any local folders from web module, we decided to put this message saying clients with web
'        'modules should not move the photos out of info:HR database.
''        lblMsg.Visible = False
'
'    Else
'        'Feb 11th 2014: From the Meeting with Jerry and Mostafa, since Mostafa said ESS/Timesheet cannot read
'        'a folder thatdoes not reside on the web server which needs to use Network Service account (local account)
'        'to access any local folders from web module, we decided to put this message saying clients with web
'        'modules should not move the photos out of info:HR database.
''        lblMsg.Visible = True
'
'        frImpPhoto.Visible = False
''        frmDelExpPhotos.Visible = True
'
''        chkDelete.Value = vbUnchecked
''        chkDeleteAll.Value = vbUnchecked
''        chkExportPhotos.Value = vbUnchecked
'
''        chkDelete.Enabled = True
''        chkDeleteAll.Enabled = True
''        chkExportPhotos.Enabled = True
'
'        lblPath.Caption = "Export to Path"
'        lblPath.Visible = False
'        File1.Pattern = "*.jpg"
'        Drive1.Visible = False
'        Dir1.Visible = False
'        File1.Visible = False
'
'        frmFile.Visible = False
'
'        lblEENum(1).Visible = True
'        elpEEID_Del.Visible = True
'    End If
'    Call SET_UP_MODE
'End Sub

'Private Sub chkImpWord_Click()
'If chkImpWord Then
'    optImportPhoto.Enabled = False
''    optExpDelPhoto.Enabled = False
'
'    File1.Pattern = "*.*"
'
''    chkDelete.Value = False
''    chkDeleteAll.Value = False
''    chkExportPhotos.Value = False
'
''    chkDelete.Visible = False
''    chkDeleteAll.Visible = False
''    chkExportPhotos.Visible = False
''
'    frImpPhoto.Caption = "Import Resume"
'    chkReplace.Caption = "Replace Existing Resume"
'    chkReplace.Enabled = False
'
'    frmFile.Visible = True
'    txtFileName.Text = ""
'    txtFileName.Enabled = False
'    Label2.Visible = False
'    frmFile.Top = 2940 ' lblEENum(1).Top
'    lblPath.Visible = True
'    lblPath.Caption = "Import From Path"
'    Drive1.Visible = True
'    Dir1.Visible = True
'    File1.Visible = True
'
'    lblEENum(1).Visible = False
'    elpEEID_Del.Visible = False
'Else
'    optImportPhoto.Enabled = True
''    optExpDelPhoto.Enabled = True
'    frImpPhoto.Caption = "Import Photo"
'    chkReplace.Enabled = True
'
'    If gsEMPLOYEEPHOTO Then
'        optImportPhoto.Enabled = False
''        optExpDelPhoto.Value = vbChecked
'
''        Call optExpDelPhoto_Click
'    Else
'        optImportPhoto.Enabled = True
'        frImpPhoto.Visible = True
''        frmDelExpPhotos.Visible = False
'
'        Call optImportPhoto_Click
'    End If
'
'    File1.Pattern = "*.jpg"
'
'    'chkDelete.Visible = True
'    'chkDeleteAll.Visible = True
'    'chkExportPhotos.Visible = True
'    'chkReplace.Value = False
'
'    'frmFile.Visible = False
'    'lblPath.Visible = False
'    'lblPath.Caption = "Export to Path"
'    'Drive1.Visible = False
'    'Dir1.Visible = False
'    'File1.Visible = False
'
'    'lblEENum(1).Visible = True
'    'elpEEID_Del.Visible = True
'
'End If
'
'End Sub

'Private Sub chkReplace_Click()
'    If chkReplace Then
'        chkDelete.Value = vbUnchecked
'    End If
'End Sub

'Private Sub Dir1_Change()
'    ChDir Dir1.Path
'    File1.Path = Dir1.Path
'    File1.Pattern = "*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pub;*.pdf;*.jpg" '|*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pub;*.pdf;*.jpg" '"*.JPG"
'End Sub

'Private Sub Drive1_Change()
'Dim xdir, xerror
'
'On Error GoTo CKERROR
'
'xerror = False
'Dir1.Path = Drive1.Drive
'
'Exit Sub
'CKERROR:
'    If Err = 68 Then
'         MsgBox "Invalid Drive Selected"
'         Drive1.Drive = App.Path
'         xerror = True
'         Resume Next
'    End If
'    MsgBox "ERROR " & Str(Err)
'    xerror = True
'    Resume Next
'End Sub

'Private Sub File1_Click()
'    Dim iit As Integer
'    Dim ii1 As Long
'    Dim sit As String
'
'    For iit = 0 To File1.ListCount - 1
'        If File1.selected(iit) Then
'            sit = File1.List(iit)
'            If chkImpWord Then
'                txtFileName.Text = UCase(File1.List(iit))
'            Else
'                ii1 = InStr(sit, ".")
'                If ii1 > 0 Then
'                    sit = Mid(sit, 1, ii1 - 1)
'                    txtFileName.Text = UCase(sit)
'                Else
'                    txtFileName.Text = UCase(File1.List(iit))
'                End If
'            End If
'        End If
'    Next
'
'    ' dkostka - 10/16/2001 - Shouldn't be able to select multiple files if you are choosing
'    '   pictures one by one.  Can't change this at runtime via control property so have to
'    '   do it in code.
'    If chkFile.Value = 0 Then
'        For ii1 = 0 To File1.ListCount - 1
'            If ii1 <> File1.ListIndex Then File1.selected(ii1) = False
'        Next ii1
'    End If
'End Sub

Private Sub Form_Activate()
Call INI_Controls(Me)
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Dim X%
Dim Y%

On Error GoTo Line_Err

glbOnTop = "FRMUASSIGNDOCTYPE"

Screen.MousePointer = HOURGLASS

Screen.MousePointer = DEFAULT

'Drive1.Drive = "c:"
'Dir1.Path = "c:\"

'Drive1.Drive = "G:"
'Dir1.Path = "G:\"
'FPath = Dir1.Path

'Load Employee Flags list
lstEmpFlagsList.Clear
For Y = 1 To 20
    If lStr("Employee Flag " & Y) <> "" Then
        lstEmpFlagsList.AddItem lStr("Employee Flag " & Y)
    End If
Next Y

'Label Master
optPerfReview.Caption = lStr(optPerfReview.Caption)
optAssociation.Caption = lStr(optAssociation.Caption)
optCounselling.Caption = lStr(optCounselling.Caption)
optComments.Caption = lStr(optComments.Caption)


''Ticket #22682 - Disable the Import Photo option if Employee Photo In Other Folder is checked (Company Preference - File Locations)
'If gsEMPLOYEEPHOTO Then
'    optImportPhoto.Enabled = False
''    optExpDelPhoto.Value = vbChecked
'
''    Call optExpDelPhoto_Click
'Else
'    optImportPhoto.Enabled = True
'    frImpPhoto.Visible = True
''    frmDelExpPhotos.Visible = False
'
'    Call optImportPhoto_Click
'End If

Exit Sub

Line_Err:
    If Err = "68" Then
        'MsgBox Err.Description
        Resume Next
    End If
    
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

Private Sub optAssociation_Click()
    'lblDate.Visible = True
    'lblDate.Caption = "Starting Date"
    'dlpDate.Visible = True
    lblCode.Visible = True
    lblCode.Top = 6330
    lblCode.Caption = lStr("Associations")
    clpCode(2).Visible = False
    clpCode(3).Visible = False
    clpCode(4).Visible = True
    clpCode(4).MaxLength = 4
    'clpCode(4).Top = clpCode(2).Top
    'clpCode(4).Left = clpCode(2).Left
    clpCode(5).Visible = False
    clpCode(6).Visible = False
    clpCode(7).Visible = False
    clpCode(8).Visible = False
    clpCode(9).Visible = False
    clpCode(10).Visible = False
    clpCode(11).Visible = False
    clpJob.Visible = False
    lblJob.Visible = False
End Sub

Private Sub optAttendance_Click()
    'lblDate.Visible = True
    'dlpDate.Visible = True
    'lblDate.Caption = "Date"
    lblCode.Visible = True
    lblCode.Top = 5970
    lblCode.Caption = "Reason"
    clpCode(2).Visible = False
    clpCode(3).Visible = True
    clpCode(3).MaxLength = 4
    'clpCode(3).Top = clpCode(2).Top
    'clpCode(3).Left = clpCode(2).Left
    clpCode(4).Visible = False
    clpCode(5).Visible = False
    clpCode(6).Visible = False
    clpCode(7).Visible = False
    clpCode(8).Visible = False
    clpCode(9).Visible = False
    clpCode(10).Visible = False
    clpCode(11).Visible = False
    clpJob.Visible = False
    lblJob.Visible = False
End Sub

Private Sub optComments_Click()
    'lblDate.Visible = True
    'lblDate.Caption = "Effective"
    'dlpDate.Visible = True
    lblCode.Visible = True
    lblCode.Top = 7770
    lblCode.Caption = "Type"
    clpCode(2).Visible = False
    clpCode(3).Visible = False
    clpCode(4).Visible = False
    clpCode(5).Visible = False
    clpCode(6).Visible = False
    clpCode(7).Visible = False
    clpCode(8).Visible = True
    clpCode(8).MaxLength = 4
    'clpCode(8).Top = clpCode(2).Top
    'clpCode(8).Left = clpCode(2).Left
    clpCode(9).Visible = False
    clpCode(10).Visible = False
    clpCode(11).Visible = False
    clpJob.Visible = False
    lblJob.Visible = False
End Sub

Private Sub optContEdu_Click()
    'lblDate.Visible = True
    'lblDate.Caption = "Start Date"
    'dlpDate.Visible = True
    lblCode.Visible = True
    lblCode.Top = 6690
    lblCode.Caption = "Course Code"
    clpCode(2).Visible = False
    clpCode(3).Visible = False
    clpCode(4).Visible = False
    clpCode(5).Visible = True
    clpCode(5).MaxLength = 8
    'clpCode(5).Top = clpCode(2).Top
    'clpCode(5).Left = clpCode(2).Left
    clpCode(6).Visible = False
    clpCode(7).Visible = False
    clpCode(8).Visible = False
    clpCode(9).Visible = False
    clpCode(10).Visible = False
    clpCode(11).Visible = False
    clpJob.Visible = False
    lblJob.Visible = False
End Sub

Private Sub optCounselling_Click()
    'lblDate.Visible = True
    'lblDate.Caption = "Counseling Date"
    'dlpDate.Visible = True
    lblCode.Visible = True
    lblCode.Top = 7410
    lblCode.Caption = "Type"
    clpCode(2).Visible = False
    clpCode(3).Visible = False
    clpCode(4).Visible = False
    clpCode(5).Visible = False
    clpCode(6).Visible = False
    clpCode(7).Visible = True
    clpCode(7).MaxLength = 4
    'clpCode(7).Top = clpCode(2).Top
    'clpCode(7).Left = clpCode(2).Left
    clpCode(8).Visible = False
    clpCode(9).Visible = False
    clpCode(10).Visible = False
    clpCode(11).Visible = False
    clpJob.Visible = False
    lblJob.Visible = False
End Sub

Private Sub optDollarEnt_Click()
    'lblDate.Visible = False
    'dlpDate.Visible = False
    lblCode.Visible = True
    lblCode.Top = 5610
    lblCode.Caption = "Entitlement"
    clpCode(2).Visible = True
    clpCode(2).MaxLength = 4
    clpCode(3).Visible = False
    clpCode(4).Visible = False
    clpCode(5).Visible = False
    clpCode(6).Visible = False
    clpCode(7).Visible = False
    clpCode(8).Visible = False
    clpCode(9).Visible = False
    clpCode(10).Visible = False
    clpCode(11).Visible = False
    clpJob.Visible = False
    lblJob.Visible = False
End Sub

Private Sub optEmpFlags_Click()
    'lblDate.Visible = False
    'dlpDate.Visible = False
    lblCode.Visible = False
    clpCode(2).Visible = False
    clpCode(3).Visible = False
    clpCode(4).Visible = False
    clpCode(5).Visible = False
    clpCode(6).Visible = False
    clpCode(7).Visible = False
    clpCode(8).Visible = False
    clpCode(9).Visible = False
    clpCode(10).Visible = False
    clpCode(11).Visible = False
    clpJob.Visible = False
    lblJob.Visible = False
End Sub

Private Sub optFormalEdu_Click()
    'lblDate.Visible = False
    'dlpDate.Visible = False
    lblCode.Visible = True
    lblCode.Top = 7050
    lblCode.Caption = "School"
    clpCode(2).Visible = False
    clpCode(3).Visible = False
    clpCode(4).Visible = False
    clpCode(5).Visible = False
    clpCode(6).Visible = True
    clpCode(6).MaxLength = 4
    'clpCode(6).Top = clpCode(2).Top
    'clpCode(6).Left = clpCode(2).Left
    clpCode(7).Visible = False
    clpCode(8).Visible = False
    clpCode(9).Visible = False
    clpCode(10).Visible = False
    clpCode(11).Visible = False
    clpJob.Visible = False
    lblJob.Visible = False
End Sub

Private Sub optHSF7Concerns_Click()
    lblCode.Visible = False
    clpCode(2).Visible = False
    clpCode(3).Visible = False
    clpCode(4).Visible = False
    clpCode(5).Visible = False
    clpCode(6).Visible = False
    clpCode(7).Visible = False
    clpCode(8).Visible = False
    clpCode(9).Visible = False
    clpCode(10).Visible = False
    clpCode(11).Visible = False
    clpJob.Visible = False
    lblJob.Visible = False
End Sub

Private Sub optHSF7WrittenOfr_Click()
    lblCode.Visible = False
    clpCode(2).Visible = False
    clpCode(3).Visible = False
    clpCode(4).Visible = False
    clpCode(5).Visible = False
    clpCode(6).Visible = False
    clpCode(7).Visible = False
    clpCode(8).Visible = False
    clpCode(9).Visible = False
    clpCode(10).Visible = False
    clpCode(11).Visible = False
    clpJob.Visible = False
    lblJob.Visible = False
End Sub

Private Sub optHSIncidents_Click()
    lblCode.Visible = False
    clpCode(2).Visible = False
    clpCode(3).Visible = False
    clpCode(4).Visible = False
    clpCode(5).Visible = False
    clpCode(6).Visible = False
    clpCode(7).Visible = False
    clpCode(8).Visible = False
    clpCode(9).Visible = False
    clpCode(10).Visible = False
    clpCode(11).Visible = False
    clpJob.Visible = False
    lblJob.Visible = False
End Sub

Private Sub optJobOffer_Click()
    'lblDate.Visible = False
    'dlpDate.Visible = False
    lblCode.Visible = True
    lblCode.Top = 4890
    lblCode.Caption = "Reason for Change"
    clpCode(2).Visible = False
    clpCode(3).Visible = False
    clpCode(4).Visible = False
    clpCode(5).Visible = False
    clpCode(6).Visible = False
    clpCode(7).Visible = False
    clpCode(8).Visible = False
    clpCode(9).Visible = False
    clpCode(10).Visible = True
    clpCode(10).MaxLength = 4
    clpCode(11).Visible = False
    clpJob.Visible = False
    lblJob.Visible = False
End Sub

Private Sub optOtherInfo_Click()
    'lblDate.Visible = False
    'dlpDate.Visible = False
    lblCode.Visible = False
    clpCode(2).Visible = False
    clpCode(3).Visible = False
    clpCode(4).Visible = False
    clpCode(5).Visible = False
    clpCode(6).Visible = False
    clpCode(7).Visible = False
    clpCode(8).Visible = False
    clpCode(9).Visible = False
    clpCode(10).Visible = False
    clpCode(11).Visible = False
    clpJob.Visible = False
    lblJob.Visible = False
End Sub

Private Sub optPerfReview_Click()
    'lblDate.Visible = False
    'dlpDate.Visible = False
    lblCode.Visible = True
    lblCode.Top = 5250
    lblCode.Caption = lStr("Performance Rating")
    clpCode(2).Visible = False
    clpCode(3).Visible = False
    clpCode(4).Visible = False
    clpCode(5).Visible = False
    clpCode(6).Visible = False
    clpCode(7).Visible = False
    clpCode(8).Visible = False
    clpCode(9).Visible = False
    clpCode(10).Visible = False
    clpCode(11).Visible = True
    clpCode(11).MaxLength = 4
    clpJob.Visible = False
    lblJob.Visible = False
End Sub

Private Sub optPosSkills_Click()
    'lblDate.Visible = True
    'lblDate.Caption = "Effective"
    'dlpDate.Visible = True
    lblCode.Visible = True
    lblCode.Top = 8490
    lblCode.Caption = "Skill"
    clpCode(2).Visible = False
    clpCode(3).Visible = False
    clpCode(4).Visible = False
    clpCode(5).Visible = False
    clpCode(6).Visible = False
    clpCode(7).Visible = False
    clpCode(8).Visible = False
    clpCode(9).Visible = True
    clpCode(9).MaxLength = 4
    clpCode(10).Visible = False
    clpCode(11).Visible = False
    clpJob.Visible = True
    lblJob.Visible = True
    'clpCode(8).Top = clpCode(2).Top
    'clpCode(8).Left = clpCode(2).Left
End Sub

Private Sub optResume_Click()
    'lblDate.Visible = False
    'dlpDate.Visible = False
    lblCode.Visible = False
    clpCode(2).Visible = False
    clpCode(3).Visible = False
    clpCode(4).Visible = False
    clpCode(5).Visible = False
    clpCode(6).Visible = False
    clpCode(7).Visible = False
    clpCode(8).Visible = False
    clpCode(9).Visible = False
    clpCode(10).Visible = False
    clpCode(11).Visible = False
    clpJob.Visible = False
    lblJob.Visible = False
End Sub

Private Sub optTermination_Click()
    'lblDate.Visible = False
    'dlpDate.Visible = False
    lblCode.Visible = False
    clpCode(2).Visible = False
    clpCode(3).Visible = False
    clpCode(4).Visible = False
    clpCode(5).Visible = False
    clpCode(6).Visible = False
    clpCode(7).Visible = False
    clpCode(8).Visible = False
    clpCode(9).Visible = False
    clpCode(10).Visible = False
    clpCode(11).Visible = False
    clpJob.Visible = False
    lblJob.Visible = False
End Sub

Private Sub txtDocDesc_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Sub txtFileName_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Sub txtFileName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

'Function modUpdateSelectionResume()
'Dim xxx, xx1, X%, XCNT
'Dim xEMPNBR, xShowEmpNbr
'Dim SQLQ
'Dim Msg$, Title$, DgDef As Variant
'Dim Response%, pct%, SSERIAL
'Dim rsEmp As New ADODB.Recordset
'Dim xPath, xFileName As String
'
'On Error GoTo modUpdateSelection_Err
'
'modUpdateSelectionResume = False
'
'UPDTCNT = 0
'MDIMain.panHelp(2).Caption = ""
'MDIMain.panHelp(1).Caption = " Please Wait"
'MDIMain.panHelp(0).FloodType = 1
'
'MDIMain.panHelp(0).FloodPercent = 0
'
'If False Then
'    Call AppendPhoto(getEmpnbr(elpEEID), ImportFile)
'Else
'    xPath = UCase(Dir1.Path) & UCase(IIf(Right(Dir1.Path, 1) = "\", "", "\"))
'    For X = 0 To File1.ListCount - 1
'        If X <> 0 Then
'            MDIMain.panHelp(0).FloodPercent = (X / (File1.ListCount - 1)) * 100
'        ElseIf (X = 0) And (File1.ListCount - 1) = 0 Then
'            MDIMain.panHelp(0).FloodPercent = 100
'        End If
'
'        If File1.selected(X) Then
'            xFileName = UCase(File1.List(X))
'            'xShowEmpNbr = elpEEID 'xFileName 'Left(xFileName, InStr(xFileName, ".JPG") - 1)
'            xShowEmpNbr = Left(xFileName, InStr(xFileName, ".") - 1)
'            xEMPNBR = getEmpnbr(xShowEmpNbr)
'            If Not IsNumeric(xEMPNBR) Then xEMPNBR = 0
'            If xEMPNBR <> 0 Then
'                rsEmp.Open "SELECT ED_EMPNBR FROM HREMP where ED_EMPNBR=" & xEMPNBR & " AND " & glbSeleDeptUn, gdbAdoIhr001, adOpenStatic
'                If Not rsEmp.EOF Then
'                    xFileName = xPath & xFileName
'                    Call AppendResume(xEMPNBR, xFileName, Right(xFileName, 3))
'                    File1.selected(X) = False
'                End If
'                rsEmp.Close
'            End If
'        End If
'        DoEvents
'    Next
'End If
'
'MDIMain.panHelp(0).Caption = ""
'modUpdateSelectionResume = True
'Screen.MousePointer = DEFAULT
'
'Exit Function
'
'modUpdateSelection_Err:
'
'Screen.MousePointer = DEFAULT
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update", "ImportPhoto", "Import")
'Screen.MousePointer = DEFAULT
'
'If gintRollBack% = False Then Resume Next Else Unload Me
'End Function

'Function modUpdateSelection()
'Dim xxx, xx1, X%, XCNT
'Dim xEMPNBR, xShowEmpNbr
'Dim SQLQ
'Dim Msg$, Title$, DgDef As Variant
'Dim Response%, pct%, SSERIAL
'Dim rsEmp As New ADODB.Recordset
'Dim xPath, xFileName As String
'On Error GoTo modUpdateSelection_Err
'modUpdateSelection = False
'
'UPDTCNT = 0
'
'MDIMain.panHelp(2).Caption = ""
'MDIMain.panHelp(1).Caption = " Please Wait"
'MDIMain.panHelp(0).FloodType = 1
'
'MDIMain.panHelp(0).FloodPercent = 0
'
'If chkFile = 0 Then
'    Call AppendPhoto(getEmpnbr(elpEEID), ImportFile)
'Else
'    xPath = UCase(Dir1.Path) & UCase(IIf(Right(Dir1.Path, 1) = "\", "", "\"))
'    For X = 0 To File1.ListCount - 1
'        If X <> 0 Then
'            MDIMain.panHelp(0).FloodPercent = (X / (File1.ListCount - 1)) * 100
'        ElseIf (X = 0) And (File1.ListCount - 1) = 0 Then
'            MDIMain.panHelp(0).FloodPercent = 100
'        End If
'
'        If File1.selected(X) Then
'            xFileName = UCase(File1.List(X))
'            xShowEmpNbr = Left(xFileName, InStr(xFileName, ".JPG") - 1)
'            xEMPNBR = getEmpnbr(xShowEmpNbr)
'            If Not IsNumeric(xEMPNBR) Then xEMPNBR = 0
'            If xEMPNBR <> 0 Then
'                rsEmp.Open "SELECT ED_EMPNBR FROM HREMP where ED_EMPNBR=" & xEMPNBR & " AND " & glbSeleDeptUn, gdbAdoIhr001, adOpenStatic
'                If Not rsEmp.EOF Then
'                    xFileName = xPath & xFileName
'                    Call AppendPhoto(xEMPNBR, xFileName)
'                    File1.selected(X) = False
'                End If
'                rsEmp.Close
'            End If
'        End If
'        DoEvents
'    Next
'End If
'
'MDIMain.panHelp(0).Caption = ""
'modUpdateSelection = True
'Screen.MousePointer = DEFAULT
'
'Exit Function
'
'modUpdateSelection_Err:
'
'Screen.MousePointer = DEFAULT
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update", "ImportPhoto", "Import")
'Screen.MousePointer = DEFAULT
'
'If gintRollBack% = False Then Resume Next Else Unload Me
'End Function

Function chkImpAttachment()
Dim Alphabet, xlen, I%, xwk, xok
Dim strDatelbl As String
Dim Y As Integer
Dim empFlagSel As Boolean

chkImpAttachment = False

On Error GoTo chkImpAttachment_Err

'If chkFile = 0 Then
            
    'Mandatory fields
    If Not clpCode(1).ListChecker Then Exit Function
    If Len(clpCode(1).Text) = 0 Then
        MsgBox "Document Type is required."
        clpCode(1).SetFocus
        Exit Function
    End If
    
    'Release 8.1
    'Send LostFocus on Document Type code so it is validated as per the Document Type Codes security
    flgWrongDocTypeCode = False
    Call clpCode_LostFocus(1)
    If flgWrongDocTypeCode = True Then
        clpCode(1).SetFocus
        Exit Function
    End If
    
    
    If optEmpFlags Then
        empFlagSel = False
        'Go through the flags list to look for selected Flag
        For Y = 0 To lstEmpFlagsList.ListCount - 1
            If lstEmpFlagsList.selected(Y) Then
                empFlagSel = True
                Exit For
            Else
                empFlagSel = False
            End If
        Next Y
        
        If empFlagSel = False Then
            MsgBox "Employee Flag must be selected"
            lstEmpFlagsList.SetFocus
            Exit Function
        End If
    End If
    
    If optJobOffer Then
        If Len(clpCode(10)) > 0 Then
            If clpCode(10).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode.Caption & " code"
                clpCode(10).SetFocus
                Exit Function
            End If
        End If
    End If
    
    If optPerfReview Then
        If Len(clpCode(11)) > 0 Then
            If clpCode(11).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode.Caption & " code"
                clpCode(11).SetFocus
                Exit Function
            End If
        End If
    End If
    
    If optDollarEnt Then
        'lblCode.Caption = "Entitlement"
        If Len(clpCode(2)) > 0 Then
            If clpCode(2).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode.Caption & " code"
                clpCode(2).SetFocus
                Exit Function
            End If
        Else
            MsgBox lblCode.Caption & " code is required"
            clpCode(2).SetFocus
            Exit Function
        End If
    End If
    
    'Attendance
    If optAttendance Then
'        strDatelbl = lStr("From Date")
'        If Len(dlpDate.Text) < 1 Then
'            MsgBox strDatelbl & " is Required Field"
'            dlpDate.SetFocus
'            Exit Function
'        Else
'            If Not IsDate(dlpDate.Text) Then
'                MsgBox strDatelbl & " is not a valid date."
'                dlpDate.SetFocus
'                Exit Function
'            End If
'        End If
        
        'lblCode.Caption = "Reason"
        If Len(clpCode(3)) > 0 Then
            If clpCode(3).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode.Caption & " code"
                clpCode(3).SetFocus
                Exit Function
            End If
        Else
            MsgBox lblCode.Caption & " code is required"
            clpCode(3).SetFocus
            Exit Function
        End If
    End If
        
    'Association
    If optAssociation Then
'        strDatelbl = lStr("Starting Date")
'        If Len(dlpDate.Text) < 1 Then
'            MsgBox strDatelbl & " is Required Field"
'            dlpDate.SetFocus
'            Exit Function
'        Else
'            If Not IsDate(dlpDate.Text) Then
'                MsgBox strDatelbl & " is not a valid date."
'                dlpDate.SetFocus
'                Exit Function
'            End If
'        End If
                
        'lblCode.Caption = "Associations"
        If Len(clpCode(4)) > 0 Then
            If clpCode(4).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode.Caption & " code"
                clpCode(4).SetFocus
                Exit Function
            End If
        Else
            MsgBox lblCode.Caption & " code is required"
            clpCode(4).SetFocus
            Exit Function
        End If
    End If
        
    'Continuing Education
    If optContEdu Then
'        strDatelbl = lStr("Start Date")
'        If Len(dlpDate.Text) < 1 Then
'            MsgBox strDatelbl & " is Required Field"
'            dlpDate.SetFocus
'            Exit Function
'        Else
'            If Not IsDate(dlpDate.Text) Then
'                MsgBox strDatelbl & " is not a valid date."
'                dlpDate.SetFocus
'                Exit Function
'            End If
'        End If
        
        'lblCode.Caption = "Course Code"
        If Len(clpCode(5)) > 0 Then
            If clpCode(5).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode.Caption & " code"
                clpCode(5).SetFocus
                Exit Function
            End If
        Else
            MsgBox lblCode.Caption & " code is required"
            clpCode(5).SetFocus
            Exit Function
        End If
    End If
    
    'Formal Education
    If optFormalEdu Then
        'lblCode.Caption = "School"
        If Len(clpCode(6)) > 0 Then
            If clpCode(6).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode.Caption & " code"
                clpCode(6).SetFocus
                Exit Function
            End If
        Else
            MsgBox lblCode.Caption & " code is required"
            clpCode(6).SetFocus
            Exit Function
        End If
    End If
    
    'Counseling
    If optCounselling Then
        If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
            strDatelbl = "Issuing"
        Else
            strDatelbl = lStr("Counseling")
        End If
'        If Len(dlpDate.Text) < 1 Then
'            MsgBox strDatelbl & " Date is Required Field"
'            dlpDate.SetFocus
'            Exit Function
'        Else
'            If Not IsDate(dlpDate.Text) Then
'                MsgBox strDatelbl & " Date is not a valid date."
'                dlpDate.SetFocus
'                Exit Function
'            End If
'        End If
        
        'lblCode.Caption = "Type"
        If Len(clpCode(7)) > 0 Then
            If clpCode(7).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode.Caption & " code"
                clpCode(7).SetFocus
                Exit Function
            End If
        Else
            MsgBox lblCode.Caption & " code is required"
            clpCode(7).SetFocus
            Exit Function
        End If
    End If
    
    'Comments
    If optComments Then
        'lblCode.Caption = "Type"
        If Len(clpCode(8)) > 0 Then
            If clpCode(8).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode.Caption & " code"
                clpCode(8).SetFocus
                Exit Function
            End If
        Else
            MsgBox lblCode.Caption & " code is required"
            clpCode(8).SetFocus
            Exit Function
        End If
    End If
    
    'Position Skills
    If optPosSkills Then
        'lblCode.Caption = "Type"
        If Len(clpCode(9)) > 0 Then
            If clpCode(9).Caption = "Unassigned" Then
                MsgBox "Invalid " & lblCode.Caption & " code"
                clpCode(9).SetFocus
                Exit Function
            End If
        Else
            MsgBox lblCode.Caption & " code is required"
            clpCode(9).SetFocus
            Exit Function
        End If
        
        If Len(Trim(clpJob)) > 0 Then
            If clpJob.Caption = "Unassigned" Then
                MsgBox "Invalid Position Code"
                clpJob.SetFocus
                Exit Function
            End If
        End If
    End If
    
    'Files
'    If Get_SelectedFilename = -1 Then
'        MsgBox "No File Name is selected to import. File Name is required."
'        File1.SetFocus
'        Exit Function
'    End If
'
'    If Len(txtFileName) = 0 Then
'        MsgBox "File Name is required."
'        File1.SetFocus
'        Exit Function
'    End If
'
'    txtFileName = LTrim(txtFileName)
'    xlen = Len(txtFileName)
'    ' dkostka - 10/16/2001 - Added space and -_()! to end of alphabet, filenames can have these chars
'    Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890-_()! "
'    xok = True
'    For I% = 1 To xlen
'        xwk = Mid(txtFileName, I%, 1)
'        If InStr(Alphabet, xwk) = 0 Then
'            xok = False
'            Exit For
'        End If
'    Next
'    If Not xok Then
'        MsgBox "Invalid File Name"
'        'txtFileName.SetFocus
'        File1.SetFocus
'        Exit Function
'    End If
    
    
'    ' dkostka - 10/16/2001 - A valid employee number is required.
'    If Len(elpEEID.Text) = 0 Then
'        MsgBox "Employee Number is required."
'        elpEEID.SetFocus
'        Exit Function
'    End If
'    If elpEEID.Caption = "Unassinged" Then
'        MsgBox "Employee Number is not valid."
'        elpEEID.SetFocus
'        Exit Function
'    End If
    
'    ImportFile = UCase(Dir1.Path) & UCase(IIf(Right(Dir1.Path, 1) = "\", "", "\")) & UCase(txtFileName & ".JPG")
'    'MsgBox ImportFile
'    If Dir(ImportFile) = "" Then
'        MsgBox "FILE not Found :" & Chr(10) & "[" & ImportFile & "]"
'        txtFileName.SetFocus
'        Exit Function
'    End If
    
'End If

chkImpAttachment = True

Exit Function

chkImpAttachment_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkImpAttachment", "Import Attachment", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

'Public Sub AppendResume(zEMPNBR, FileName As String, FileExtension As String)
'    Dim rsPhoto As New ADODB.Recordset
'
'    Dim byteChunk() As Byte
'    Dim X, xChr
'    Dim FileNumber As Integer
'    If Not IsNumeric(zEMPNBR) Then Exit Sub
'    rsPhoto.Open "select * from HRDOC_EMP WHERE RE_EMPNBR=" & zEMPNBR, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If Not rsPhoto.EOF Then
'        If chkReplace = 0 Then
'            Exit Sub
'        Else
'            rsPhoto.Delete
'        End If
'    End If
'    UPDTCNT = UPDTCNT + 1
'    FileNumber = FreeFile
'    Open FileName For Binary Access Read As FileNumber
'    ReDim byteChunk(FileLen(FileName))
'
'    rsPhoto.AddNew
'    rsPhoto("RE_EMPNBR") = zEMPNBR
'    rsPhoto("RE_COMPNO") = "001"
'    rsPhoto("RE_FILEEXT") = FileExtension
'    rsPhoto("RE_TYPE") = "RESUME"
'    rsPhoto("RE_LUSER") = glbUserID
'    rsPhoto("RE_LDATE") = Date
'    rsPhoto("RE_LTIME") = Time$
'    Get FileNumber, , byteChunk
'    rsPhoto!RE_DOC.AppendChunk byteChunk
'    Close FileNumber
'
'    If glbSQL Or glbOracle Then rsPhoto.Update
''    rsPHOTO.Requery
'    rsPhoto.Close
'
'End Sub

'Public Sub AppendPhoto(zEMPNBR, FileName As String)
'
'    Dim rsPhoto As New ADODB.Recordset
'
'    Dim byteChunk() As Byte
'    Dim X, xChr
'    Dim FileNumber As Integer
'    If Not IsNumeric(zEMPNBR) Then Exit Sub
'    rsPhoto.Open "select * from HR_PHOTO WHERE PT_EMPNBR=" & zEMPNBR, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If Not rsPhoto.EOF Then
'        If chkReplace = 0 Then
'            Exit Sub
'        Else
'            Do
'                rsPhoto.Delete
'                rsPhoto.MoveNext
'            Loop Until rsPhoto.EOF
'        End If
'    End If
'    UPDTCNT = UPDTCNT + 1
'    FileNumber = FreeFile
'    Open FileName For Binary Access Read As FileNumber
'    ReDim byteChunk(FileLen(FileName))
'
'    rsPhoto.AddNew
'    rsPhoto("PT_EMPNBR") = zEMPNBR
'    rsPhoto("PT_COMPNO") = "001"
'    rsPhoto("PT_LUSER") = glbUserID
'    rsPhoto("PT_LDATE") = Date
'    rsPhoto("PT_LTIME") = Time$
'    Get FileNumber, , byteChunk
'    rsPhoto!PT_PHOTO.AppendChunk byteChunk
'    Close FileNumber
'    If glbSQL Or glbOracle Then rsPhoto.Update
''    rsPHOTO.Requery
'    rsPhoto.Close
'
'End Sub

Public Property Get ChangeAction() As UpdateStateEnum
ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = MassChanges
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = GetMassUpdateSecurities("ImpAttachment_MassUpdate", glbUserID)
'UpdateRight = True
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = True
End Property

Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = False
End Property

Public Sub SET_UP_MODE()
Call set_Buttons
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

'Function Import_Attachment_Files()
'Dim xxx, xx1, X%, XCNT
'Dim xEmpnbr, xShowEmpNbr
'Dim SQLQ
'Dim Msg$, Title$, DgDef As Variant
'Dim Response%, pct%, SSERIAL
'Dim rsEmp As New ADODB.Recordset
'Dim xPath, xFileName As String
'Dim xFileExtension As String
'
'On Error GoTo Import_Attachment_Files_Err
'
'Import_Attachment_Files = False
'
'UPDTCNT = 0
'
'MDIMain.panHelp(2).Caption = ""
'MDIMain.panHelp(1).Caption = " Please Wait"
'MDIMain.panHelp(0).FloodType = 1
'
'MDIMain.panHelp(0).FloodPercent = 0
'
'If chkFile = 0 Then
'    'File names not equal to Employee #s
'Else
'    xPath = UCase(Dir1.Path) & UCase(IIf(Right(Dir1.Path, 1) = "\", "", "\"))
'    For X = 0 To File1.ListCount - 1
'        If X <> 0 Then
'            MDIMain.panHelp(0).FloodPercent = (X / (File1.ListCount - 1)) * 100
'        ElseIf (X = 0) And (File1.ListCount - 1) = 0 Then
'            MDIMain.panHelp(0).FloodPercent = 100
'        End If
'
'        If File1.selected(X) Then
'            xFileName = UCase(File1.List(X))
'            xShowEmpNbr = Left(xFileName, InStr(xFileName, ".") - 1)    'If want to add alpha values in the values then we should consider nnnnn_XXXXX.xxx format. In that case Instr will search for "_"
'            xEmpnbr = getEmpnbr(xShowEmpNbr)
'
'            If Not IsNumeric(xEmpnbr) Then xEmpnbr = 0
'
'            If xEmpnbr <> 0 Then
'                'Check if Employee exists
'                rsEmp.Open "SELECT ED_EMPNBR FROM HREMP where ED_EMPNBR=" & xEmpnbr & " AND " & glbSeleDeptUn, gdbAdoIhr001, adOpenStatic
'                If Not rsEmp.EOF Then
'                    'Employee Exists - Import attachment in the respective table
'
'                    'Full file path
'                    xFileName = xPath & xFileName
'
'                    'File extension
'                    xFileExtension = GetFileExtension(xFileName)
'
'                    'Import Attachment as per Attachment Type
'                    If optResume Then
'                        'Get other info. to update the Attachment table with
'                        glbDocNewRecord = False
'                        glbDocName = "Resume"
'
'                        'Check if attachment already exists and depending on 'Replace Existing Attachment' checkbox - delete or skip the import
'                        If AttachmentExists(xEmpnbr, glbDocName) Then
'                            If chkReplace = 0 Then
'                                'Do not replace - move to next document
'                                GoTo NextAttachment
'                            Else
'                                'Delete the Attachment
'                                Call DeleteAttachment(xEmpnbr, glbDocName)
'                            End If
'                        End If
'
'                        'Add new attachment
'                        Call AppendResume(xEmpnbr, xFileName, xFileExtension, clpCode(1).Text, txtDocDesc.Text)
'
'                    ElseIf optOtherInfo Then
'                        'Get other info. to update the Attachment table with
'                        glbDocNewRecord = False
'                        glbDocName = "OtherInfo"
'
'                        'Check if attachment already exists and depending on 'Replace Existing Attachment' checkbox - delete or skip the import
'                        If AttachmentExists(xEmpnbr, glbDocName) Then
'                            If chkReplace = 0 Then
'                                'Do not replace - move to next document
'                                GoTo NextAttachment
'                            Else
'                                'Delete the Attachment
'                                Call DeleteAttachment(xEmpnbr, glbDocName)
'                            End If
'                        End If
'
'                        'Add new attachment
'                        Call AppendOtherInfo(xEmpnbr, xFileName, xFileExtension, clpCode(1).Text, txtDocDesc.Text)
'
'                    ElseIf optEmpFlags Then
'                        'Get other info. to update the Attachment table with
'                        'glbEmpFlagNo, glbDocName, glbDocKey, glbEmpFlagDate
'                        glbDocKey = ""
'                        glbDocName = "EmployeeFlag"
'                        glbEmpFlagNo = Get_SelectedEmployeeFlag
'                        'glbEmpFlag = lstEmpFlagsList.ListIndex(glbEmpFlagNo)
'
'                        If glbEmpFlagNo <> -1 Then
'                            glbEmpFlagDate = GetEmpFlagData(xEmpnbr, "EF_FLAGDTE" & glbEmpFlagNo + 1, "01/01/1900")
'                            If glbEmpFlagDate = "01/01/1900" Then
'                                'This employee do not have the respective flag setup. Move to next employee import
'                                GoTo NextAttachment
'                            End If
'                            glbDocKey = GetEmpFlagData(xEmpnbr, "EF_ID", "")
'                        End If
'
'                        If glbDocKey <> "" Then
'                            'Check if attachment already exists and depending on 'Replace Existing Attachment' checkbox - delete or skip the import
'                            If AttachmentExists(xEmpnbr, glbDocName, glbDocKey) Then
'                                If chkReplace = 0 Then
'                                    'Do not replace - move to next document
'                                    GoTo NextAttachment
'                                Else
'                                    'Delete the Attachment
'                                    Call DeleteAttachment(xEmpnbr, glbDocName, glbDocKey)
'                                End If
'                            End If
'
'                            'Add new attachment
'                            Call AppendEmployeeFlag(xEmpnbr, xFileName, xFileExtension, clpCode(1).Text, txtDocDesc.Text)
'                        End If
'
'                    ElseIf optJobOffer Then
'                        'Get other info. to update the Attachment table with
'                        glbDocNewRecord = False
'                        glbDocName = "Offer"
'                        'Current glbJob & glbSDate
'                        glbJob = GetJHData(xEmpnbr, "JH_JOB", "")
'                        glbSDate = GetJHData(xEmpnbr, "JH_SDATE", "")
'                        If Len(glbJob) > 0 And IsDate(glbSDate) Then
'                            'Check if attachment already exists and depending on 'Replace Existing Attachment' checkbox - delete or skip the import
'                            If AttachmentExists(xEmpnbr, glbDocName, glbDocKey, glbJob, glbSDate) Then
'                                If chkReplace = 0 Then
'                                    'Do not replace - move to next document
'                                    GoTo NextAttachment
'                                Else
'                                    'Delete the Attachment
'                                    Call DeleteAttachment(xEmpnbr, glbDocName, glbDocKey, glbJob, glbSDate)
'                                End If
'                            End If
'
'                            'Add new attachment
'                            Call AppendOffer(xEmpnbr, xFileName, xFileExtension, clpCode(1).Text, txtDocDesc.Text)
'                        End If
'
'                    ElseIf optPerfReview Then
'                        'Get other info. to update the Attachment table with
'                        glbDocKey = ""
'                        glbDocNewRecord = False
'                        glbDocName = "Performance"
'
'                        'Current Performance Review glbJob, glbDocKey, glbSDate
'                        glbJob = GetPHData(xEmpnbr, "PH_JOB", "")
'                        glbSDate = GetPHData(xEmpnbr, "PH_PREVIEW", "")
'                        glbDocKey = GetPHData(xEmpnbr, "PH_ID", "")
'
'                        If Len(glbJob) > 0 And glbDocKey <> "" Then
'                            'Check if attachment already exists and depending on 'Replace Existing Attachment' checkbox - delete or skip the import
'                            If AttachmentExists(xEmpnbr, glbDocName, glbDocKey) Then
'                                If chkReplace = 0 Then
'                                    'Do not replace - move to next document
'                                    GoTo NextAttachment
'                                Else
'                                    'Delete the Attachment
'                                    Call DeleteAttachment(xEmpnbr, glbDocName, glbDocKey)
'                                End If
'                            End If
'
'                            'Add new attachment
'                            Call AppendPerformance(xEmpnbr, xFileName, xFileExtension, clpCode(1).Text, txtDocDesc.Text)
'                        End If
'
'                    ElseIf optDollarEnt Then
'                        'Get other info. to update the Attachment table with
'                        'glbDocName, glbDocKey
'                        glbDocKey = ""
'                        glbDocNewRecord = False
'                        glbDocName = "DollarEnt"
'
'                        If Len(clpCode(2).Text) > 0 Then
'                            glbDocKey = GetDollarEntData(xEmpnbr, clpCode(2).Text, "DE_TYPE", "DE_ENTITLE_ID", "")
'                        Else
'                            glbDocKey = GetDollarEntData(xEmpnbr, "", "", "DE_ENTITLE_ID", "") 'Top most Dollar Entitlement record
'                        End If
'
'                        If glbDocKey <> "" Then
'                            'Check if attachment already exists and depending on 'Replace Existing Attachment' checkbox - delete or skip the import
'                            If AttachmentExists(xEmpnbr, glbDocName, glbDocKey) Then
'                                If chkReplace = 0 Then
'                                    'Do not replace - move to next document
'                                    GoTo NextAttachment
'                                Else
'                                    'Delete the Attachment
'                                    Call DeleteAttachment(xEmpnbr, glbDocName, glbDocKey)
'                                End If
'                            End If
'
'                            'Add new attachment
'                            Call AppendDollarEnt(xEmpnbr, xFileName, xFileExtension, clpCode(1).Text, txtDocDesc.Text)
'                        End If
'
'                    ElseIf optAttendance Then
'                        'Get other info. to update the Attachment table with
'                        'glbDocName, glbDocKey, glbAttReason, glbAttDOA
'                        glbDocKey = ""
'                        glbDocNewRecord = False
'                        glbDocName = "Attendance"
'                        glbAttReason = clpCode(3).Text
'                        glbAttDOA = IIf(IsDate(dlpDate.Text), dlpDate.Text, "")
'
'                        If Len(clpCode(3).Text) > 0 And IsDate(dlpDate.Text) Then
'                            glbDocKey = GetAttendData(xEmpnbr, clpCode(3).Text, "AD_REASON", dlpDate.Text, "AD_DOA", "AD_ATT_ID", "")
'                        ElseIf Len(clpCode(3).Text) > 0 And Not IsDate(dlpDate.Text) Then
'                            glbDocKey = GetAttendData(xEmpnbr, clpCode(3).Text, "AD_REASON", "", "", "AD_ATT_ID", "")
'                        ElseIf Len(clpCode(3).Text) = 0 And IsDate(dlpDate.Text) Then
'                            glbDocKey = GetAttendData(xEmpnbr, "", "", dlpDate.Text, "AD_DOA", "AD_ATT_ID", "")
'                        Else
'                            glbDocKey = GetAttendData(xEmpnbr, "", "", "", "", "AD_ATT_ID", "") 'Top most Attendance record.
'                        End If
'
'                        If glbDocKey <> "" Then
'                            'Check if attachment already exists and depending on 'Replace Existing Attachment' checkbox - delete or skip the import
'                            If AttachmentExists(xEmpnbr, glbDocName, glbDocKey) Then
'                                If chkReplace = 0 Then
'                                    'Do not replace - move to next document
'                                    GoTo NextAttachment
'                                Else
'                                    'Delete the Attachment
'                                    Call DeleteAttachment(xEmpnbr, glbDocName, glbDocKey)
'                                End If
'                            End If
'
'                            'Add new attachment
'                            Call AppendAttendance(xEmpnbr, xFileName, xFileExtension, clpCode(1).Text, txtDocDesc.Text)
'                        End If
'
'                    ElseIf optAssociation Then
'                        'Get other info. to update the Attachment table with
'                        'glbDocName, glbDocKey, glbAssocCode, glbBeginDt
'                        glbDocKey = ""
'                        glbDocNewRecord = False
'                        glbDocName = "Associations"
'                        glbAssocCode = clpCode(4).Text
'                        glbBeginDt = IIf(IsDate(dlpDate.Text), dlpDate.Text, "")
'
'                        If Len(clpCode(4).Text) > 0 And IsDate(dlpDate.Text) Then
'                            glbDocKey = GetAssociationData(xEmpnbr, clpCode(4).Text, "TD_CODE", dlpDate.Text, "TD_BEGINDT", "TD_ID", "")
'                        ElseIf Len(clpCode(4).Text) > 0 And Not IsDate(dlpDate.Text) Then
'                            glbDocKey = GetAssociationData(xEmpnbr, clpCode(4).Text, "TD_CODE", "", "", "TD_ID", "")
'                        ElseIf Len(clpCode(4).Text) = 0 And IsDate(dlpDate.Text) Then
'                            glbDocKey = GetAssociationData(xEmpnbr, "", "", dlpDate.Text, "TD_BEGINDT", "TD_ID", "")
'                        Else
'                            glbDocKey = GetAssociationData(xEmpnbr, "", "", "", "", "TD_ID", "") 'Top most Associations record.
'                        End If
'
'                        If glbDocKey <> "" Then
'                            'Check if attachment already exists and depending on 'Replace Existing Attachment' checkbox - delete or skip the import
'                            If AttachmentExists(xEmpnbr, glbDocName, glbDocKey) Then
'                                If chkReplace = 0 Then
'                                    'Do not replace - move to next document
'                                    GoTo NextAttachment
'                                Else
'                                    'Delete the Attachment
'                                    Call DeleteAttachment(xEmpnbr, glbDocName, glbDocKey)
'                                End If
'                            End If
'
'                            'Add new attachment
'                            Call AppendAssociations(xEmpnbr, xFileName, xFileExtension, clpCode(1).Text, txtDocDesc.Text)
'                        End If
'
'                    ElseIf optContEdu Then
'                        'Get other info. to update the Attachment table with
'                        'glbDocName, glbDocKey
'                        glbDocKey = ""
'                        glbDocNewRecord = False
'                        glbDocName = "EdSem"
'                        glbBeginDt = IIf(IsDate(dlpDate.Text), dlpDate.Text, "")
'
'                        If Len(clpCode(5).Text) > 0 And IsDate(dlpDate.Text) Then
'                            glbDocKey = GetContEduData(xEmpnbr, clpCode(5).Text, "ES_CRSCODE", dlpDate.Text, "ES_START", "ES_ID", "") 'Most recent Continuing Education
'                        ElseIf Len(clpCode(5).Text) > 0 And Not IsDate(dlpDate.Text) Then
'                            glbDocKey = GetContEduData(xEmpnbr, clpCode(5).Text, "ES_CRSCODE", "", "", "ES_ID", "") 'Most recent Continuing Education
'                        ElseIf Len(clpCode(5).Text) = 0 And IsDate(dlpDate.Text) Then
'                            glbDocKey = GetContEduData(xEmpnbr, "", "", dlpDate.Text, "ES_START", "ES_ID", "") 'Most recent Continuing Education
'                        Else
'                            glbDocKey = GetContEduData(xEmpnbr, "", "", "", "", "ES_ID", "") 'Top most Continuing Education record.
'                        End If
'
'                        If glbDocKey <> "" Then
'                            'Check if attachment already exists and depending on 'Replace Existing Attachment' checkbox - delete or skip the import
'                            If AttachmentExists(xEmpnbr, glbDocName, glbDocKey) Then
'                                If chkReplace = 0 Then
'                                    'Do not replace - move to next document
'                                    GoTo NextAttachment
'                                Else
'                                    'Delete the Attachment
'                                    Call DeleteAttachment(xEmpnbr, glbDocName, glbDocKey)
'                                End If
'                            End If
'
'                            'Add new attachment
'                            Call AppendEdSem(xEmpnbr, xFileName, xFileExtension, clpCode(1).Text, txtDocDesc.Text)
'                        End If
'
'                    ElseIf optFormalEdu Then
'                        'Get other info. to update the Attachment table with
'                        'glbDocName, glbDocKey
'                        glbDocKey = ""
'                        glbDocNewRecord = False
'                        glbDocName = "FormalEdu"
'
'                        If Len(clpCode(6).Text) > 0 Then
'                            glbDocKey = GetFormEduData(xEmpnbr, clpCode(6).Text, "EU_SCHOOL", "EU_ID", "") 'Most recent Formal Education
'                        Else
'                            glbDocKey = GetFormEduData(xEmpnbr, "", "", "EU_ID", "") 'Top most Formal Education record.
'                        End If
'
'                        If glbDocKey <> "" Then
'                            'Check if attachment already exists and depending on 'Replace Existing Attachment' checkbox - delete or skip the import
'                            If AttachmentExists(xEmpnbr, glbDocName, glbDocKey) Then
'                                If chkReplace = 0 Then
'                                    'Do not replace - move to next document
'                                    GoTo NextAttachment
'                                Else
'                                    'Delete the Attachment
'                                    Call DeleteAttachment(xEmpnbr, glbDocName, glbDocKey)
'                                End If
'                            End If
'
'                            'Add new attachment
'                            Call AppendFormalEdu(xEmpnbr, xFileName, xFileExtension, clpCode(1).Text, txtDocDesc.Text)
'                        End If
'
'                    ElseIf optCounselling Then
'                        'Get other info. to update the Attachment table with
'                        'glbDocName, glbDocKey, glbSDate
'                        glbDocKey = ""
'                        glbDocNewRecord = False
'                        glbDocName = "Counsel"
'                        glbSDate = IIf(IsDate(dlpDate.Text), dlpDate.Text, "")
'
'                        If Len(clpCode(7).Text) > 0 And IsDate(dlpDate.Text) Then
'                            glbDocKey = GetCounsellingData(xEmpnbr, clpCode(7).Text, "CL_TYPE", dlpDate.Text, "CL_COUDATE", "CL_ID", "") 'Most recent Counselling
'                        ElseIf Len(clpCode(7).Text) > 0 And Not IsDate(dlpDate.Text) Then
'                            glbDocKey = GetCounsellingData(xEmpnbr, clpCode(7).Text, "CL_TYPE", "", "", "CL_ID", "")
'                        ElseIf Len(clpCode(7).Text) = 0 And IsDate(dlpDate.Text) Then
'                            glbDocKey = GetCounsellingData(xEmpnbr, "", "", dlpDate.Text, "CL_COUDATE", "CL_ID", "")
'                        Else
'                            glbDocKey = GetCounsellingData(xEmpnbr, "", "", "", "", "CL_ID", "") 'Top most Counselling record.
'                        End If
'
'                        If glbDocKey <> "" Then
'                            'Check if attachment already exists and depending on 'Replace Existing Attachment' checkbox - delete or skip the import
'                            If AttachmentExists(xEmpnbr, glbDocName, glbDocKey) Then
'                                If chkReplace = 0 Then
'                                    'Do not replace - move to next document
'                                    GoTo NextAttachment
'                                Else
'                                    'Delete the Attachment
'                                    Call DeleteAttachment(xEmpnbr, glbDocName, glbDocKey)
'                                End If
'                            End If
'
'                            'Add new attachment
'                            Call AppendCounsel(xEmpnbr, xFileName, xFileExtension, clpCode(1).Text, txtDocDesc.Text)
'                        End If
'
'                    ElseIf optComments Then
'                        'Get other info. to update the Attachment table with
'                        'glbDocName, glbDocKey,
'                        glbDocKey = ""
'                        glbDocNewRecord = False
'                        glbDocName = "Comments"
'
'                        If Len(clpCode(8).Text) > 0 And IsDate(dlpDate.Text) Then
'                            glbDocKey = GetCommentsData(xEmpnbr, clpCode(8).Text, "CO_TYPE", dlpDate.Text, "CO_EDATE", "CO_COMMENT_ID", "") 'Most recent Comments
'                        ElseIf Len(clpCode(8).Text) > 0 And Not IsDate(dlpDate.Text) Then
'                            glbDocKey = GetCommentsData(xEmpnbr, clpCode(8).Text, "CO_TYPE", "", "", "CO_COMMENT_ID", "")
'                        ElseIf Len(clpCode(8).Text) = 0 And IsDate(dlpDate.Text) Then
'                            glbDocKey = GetCommentsData(xEmpnbr, "", "", dlpDate.Text, "CO_EDATE", "CO_COMMENT_ID", "")
'                        Else
'                            glbDocKey = GetCommentsData(xEmpnbr, "", "", "", "", "CO_COMMENT_ID", "")  'Top most Comments record.
'                        End If
'
'                        If glbDocKey <> "" Then
'                            'Check if attachment already exists and depending on 'Replace Existing Attachment' checkbox - delete or skip the import
'                            If AttachmentExists(xEmpnbr, glbDocName, glbDocKey) Then
'                                If chkReplace = 0 Then
'                                    'Do not replace - move to next document
'                                    GoTo NextAttachment
'                                Else
'                                    'Delete the Attachment
'                                    Call DeleteAttachment(xEmpnbr, glbDocName, glbDocKey)
'                                End If
'                            End If
'
'                            'Add new attachment
'                            Call AppendComments(xEmpnbr, xFileName, xFileExtension, clpCode(1).Text, txtDocDesc.Text)
'                        End If
'                    End If
'NextAttachment:
'                    File1.selected(X) = False
'                End If
'                rsEmp.Close
'            End If
'        End If
'        DoEvents
'    Next
'End If
'
'MDIMain.panHelp(0).Caption = ""
'
'Import_Attachment_Files = True
'
'Screen.MousePointer = DEFAULT
'
'Exit Function
'
'Import_Attachment_Files_Err:
'
'Screen.MousePointer = DEFAULT
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update", "ImportAttachment", "ImportAttachment")
'Screen.MousePointer = DEFAULT
'
'If gintRollBack% = False Then Resume Next Else Unload Me
'
'End Function

Private Function Get_SelectedEmployeeFlag() As Integer
    Dim Y As Integer
    
    Get_SelectedEmployeeFlag = -1
    
    'Get the Employee Flag selected.
    If optEmpFlags Then
        'Go through the flags list to look for selected Flag
        For Y = 0 To lstEmpFlagsList.ListCount - 1
            If lstEmpFlagsList.selected(Y) Then
                Get_SelectedEmployeeFlag = Y
                Exit For
            Else
                Get_SelectedEmployeeFlag = -1
            End If
        Next Y
    End If
End Function

Public Function GetEmpFlagData(EmpNbr, Field As String, DEFAULT)
    Dim rsEmpFlagData As New ADODB.Recordset
    
    rsEmpFlagData.Open "SELECT " & Field & " FROM HREMP_FLAGS WHERE EF_EMPNBR=" & EmpNbr, gdbAdoIhr001, adOpenForwardOnly
    GetEmpFlagData = DEFAULT
    
    If Not rsEmpFlagData.EOF Then
        If Not IsNull(rsEmpFlagData(Field)) Then GetEmpFlagData = rsEmpFlagData(Field)
    End If
    rsEmpFlagData.Close
    Set rsEmpFlagData = Nothing
End Function

Public Function GetPHData(EmpNbr, Field As String, DEFAULT)
    Dim rsPHTEMP As New ADODB.Recordset
    rsPHTEMP.Open "SELECT " & Field & " FROM HR_PERFORM_HISTORY WHERE PH_CURRENT<>0 AND PH_EMPNBR=" & EmpNbr, gdbAdoIhr001, adOpenForwardOnly
    GetPHData = DEFAULT
    
    If Not rsPHTEMP.EOF Then
        If Not IsNull(rsPHTEMP(Field)) Then GetPHData = rsPHTEMP(Field)
    End If
    rsPHTEMP.Close
    Set rsPHTEMP = Nothing
End Function

Public Function GetDollarEntData(EmpNbr, selValue As String, selField As String, Field As String, DEFAULT, AllRecs As Boolean)
    Dim rsDollarEntData As New ADODB.Recordset
    
    If EmpNbr <> 0 Then
        If Len(selField) > 0 Then
            rsDollarEntData.Open "SELECT " & Field & " FROM HRDOLENT WHERE DE_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "'", gdbAdoIhr001, adOpenForwardOnly
        Else
            rsDollarEntData.Open "SELECT " & Field & " FROM HRDOLENT WHERE DE_EMPNBR=" & EmpNbr & " ORDER BY DE_ENTITLE_ID DESC", gdbAdoIhr001, adOpenForwardOnly
        End If
    Else
        If Len(selField) > 0 Then
            rsDollarEntData.Open "SELECT " & Field & " FROM HRDOLENT WHERE " & selField & "= '" & selValue & "'", gdbAdoIhr001, adOpenForwardOnly
        Else
            rsDollarEntData.Open "SELECT " & Field & " FROM HRDOLENT ORDER BY DE_ENTITLE_ID DESC", gdbAdoIhr001, adOpenForwardOnly
        End If
    End If
    GetDollarEntData = DEFAULT
    
    If Not AllRecs Then
        If Not rsDollarEntData.EOF Then
            If Not IsNull(rsDollarEntData(Field)) Then GetDollarEntData = rsDollarEntData(Field)
        End If
    Else
        Do While Not rsDollarEntData.EOF
            If IsNull(GetDollarEntData) Or Len(GetDollarEntData) = 0 Then
                If Not IsNull(rsDollarEntData(Field)) Then GetDollarEntData = rsDollarEntData(Field)
            Else
                If Not IsNull(rsDollarEntData(Field)) Then GetDollarEntData = GetDollarEntData & "," & rsDollarEntData(Field)
            End If
            
            rsDollarEntData.MoveNext
        Loop
    End If
    rsDollarEntData.Close
    Set rsDollarEntData = Nothing
End Function

Public Function GetDollarEntData_Term(EmpNbr, selValue As String, selField As String, Field As String, DEFAULT, AllRecs As Boolean)
    Dim rsDollarEntData As New ADODB.Recordset
    
    If EmpNbr <> 0 Then
        If Len(selField) > 0 Then
            rsDollarEntData.Open "SELECT " & Field & " FROM TERM_DOLENT WHERE DE_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "'", gdbAdoIhr001X, adOpenForwardOnly
        Else
            rsDollarEntData.Open "SELECT " & Field & " FROM TERM_DOLENT WHERE DE_EMPNBR=" & EmpNbr & " ORDER BY DE_ENTITLE_ID DESC", gdbAdoIhr001X, adOpenForwardOnly
        End If
    Else
        If Len(selField) > 0 Then
            rsDollarEntData.Open "SELECT " & Field & " FROM TERM_DOLENT WHERE " & selField & "= '" & selValue & "'", gdbAdoIhr001X, adOpenForwardOnly
        Else
            rsDollarEntData.Open "SELECT " & Field & " FROM TERM_DOLENT ORDER BY DE_ENTITLE_ID DESC", gdbAdoIhr001X, adOpenForwardOnly
        End If
    End If
    
    GetDollarEntData_Term = DEFAULT
    
    If Not AllRecs Then
        If Not rsDollarEntData.EOF Then
            If Not IsNull(rsDollarEntData(Field)) Then GetDollarEntData_Term = rsDollarEntData(Field)
        End If
    Else
        Do While Not rsDollarEntData.EOF
            If IsNull(GetDollarEntData_Term) Or Len(GetDollarEntData_Term) = 0 Then
                If Not IsNull(rsDollarEntData(Field)) Then GetDollarEntData_Term = rsDollarEntData(Field)
            Else
                If Not IsNull(rsDollarEntData(Field)) Then GetDollarEntData_Term = GetDollarEntData_Term & "," & rsDollarEntData(Field)
            End If
            
            rsDollarEntData.MoveNext
        Loop
    End If
    rsDollarEntData.Close
    Set rsDollarEntData = Nothing
End Function

Public Function GetAttendData(EmpNbr, selValue As String, selField As String, selDateValue As Date, selDateField As String, Field As String, DEFAULT)
    Dim rsAttendData As New ADODB.Recordset
    
    If Len(selField) > 0 And Len(selDateField) > 0 Then
        rsAttendData.Open "SELECT " & Field & " FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "' AND " & selDateField & "= " & Date_SQL(selDateValue), gdbAdoIhr001, adOpenForwardOnly
    ElseIf Len(selField) > 0 Then
        rsAttendData.Open "SELECT " & Field & " FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "'", gdbAdoIhr001, adOpenForwardOnly
    ElseIf Len(selDateField) > 0 Then
        rsAttendData.Open "SELECT " & Field & " FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & EmpNbr & " AND " & selDateField & "= '" & selDateValue & "'", gdbAdoIhr001, adOpenForwardOnly
    Else
        rsAttendData.Open "SELECT " & Field & " FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & EmpNbr & " ORDER BY AD_ATT_ID DESC", gdbAdoIhr001, adOpenForwardOnly
    End If
    GetAttendData = DEFAULT
    
    If Not rsAttendData.EOF Then
        If Not IsNull(rsAttendData(Field)) Then GetAttendData = rsAttendData(Field)
    End If
    rsAttendData.Close
    Set rsAttendData = Nothing
End Function

Public Function GetAssociationData(EmpNbr, selValue As String, selField As String, selDateValue As Date, selDateField As String, Field As String, DEFAULT)
    Dim rsAssociationData As New ADODB.Recordset
    
    If Len(selField) > 0 And Len(selDateField) > 0 Then
        rsAssociationData.Open "SELECT " & Field & " FROM HRTRADE WHERE TD_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "' AND " & selDateField & "= " & Date_SQL(selDateValue), gdbAdoIhr001, adOpenForwardOnly
    ElseIf Len(selField) > 0 Then
        rsAssociationData.Open "SELECT " & Field & " FROM HRTRADE WHERE TD_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "'", gdbAdoIhr001, adOpenForwardOnly
    ElseIf Len(selDateField) > 0 Then
        rsAssociationData.Open "SELECT " & Field & " FROM HRTRADE WHERE TD_EMPNBR=" & EmpNbr & " AND " & selDateField & "= '" & selDateValue & "'", gdbAdoIhr001, adOpenForwardOnly
    Else
        rsAssociationData.Open "SELECT " & Field & " FROM HRTRADE WHERE TD_EMPNBR=" & EmpNbr & " ORDER BY TD_ID DESC", gdbAdoIhr001, adOpenForwardOnly
    End If
    GetAssociationData = DEFAULT
    
    If Not rsAssociationData.EOF Then
        If Not IsNull(rsAssociationData(Field)) Then GetAssociationData = rsAssociationData(Field)
    End If
    rsAssociationData.Close
    Set rsAssociationData = Nothing
End Function

Public Function GetContEduData(EmpNbr, selValue As String, selField As String, selDateValue As Date, selDateField As String, Field As String, DEFAULT, AllRecs As Boolean)
    Dim rsContEduData As New ADODB.Recordset
    
    If EmpNbr <> 0 Then
        If Len(selField) > 0 And Len(selDateField) > 0 Then
            rsContEduData.Open "SELECT " & Field & " FROM HREDSEM WHERE ES_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "' AND " & selDateField & "= " & Date_SQL(selDateValue), gdbAdoIhr001, adOpenForwardOnly
        ElseIf Len(selField) > 0 Then
            rsContEduData.Open "SELECT " & Field & " FROM HREDSEM WHERE ES_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "'", gdbAdoIhr001, adOpenForwardOnly
        ElseIf Len(selDateField) > 0 Then
            rsContEduData.Open "SELECT " & Field & " FROM HREDSEM WHERE ES_EMPNBR=" & EmpNbr & " AND " & selDateField & "= '" & selDateValue & "'", gdbAdoIhr001, adOpenForwardOnly
        Else
            rsContEduData.Open "SELECT " & Field & " FROM HREDSEM WHERE ES_EMPNBR=" & EmpNbr & " ORDER BY ES_ID DESC", gdbAdoIhr001, adOpenForwardOnly
        End If
    Else
        If Len(selField) > 0 And Len(selDateField) > 0 Then
            rsContEduData.Open "SELECT " & Field & " FROM HREDSEM WHERE " & selField & "= '" & selValue & "' AND " & selDateField & "= " & Date_SQL(selDateValue), gdbAdoIhr001, adOpenForwardOnly
        ElseIf Len(selField) > 0 Then
            rsContEduData.Open "SELECT " & Field & " FROM HREDSEM WHERE " & selField & "= '" & selValue & "'", gdbAdoIhr001, adOpenForwardOnly
        ElseIf Len(selDateField) > 0 Then
            rsContEduData.Open "SELECT " & Field & " FROM HREDSEM WHERE " & selDateField & "= '" & selDateValue & "'", gdbAdoIhr001, adOpenForwardOnly
        Else
            rsContEduData.Open "SELECT " & Field & " FROM HREDSEM ORDER BY ES_ID DESC", gdbAdoIhr001, adOpenForwardOnly
        End If
    End If
    
    GetContEduData = DEFAULT
    
    If Not AllRecs Then
        If Not rsContEduData.EOF Then
            If Not IsNull(rsContEduData(Field)) Then GetContEduData = rsContEduData(Field)
        End If
    Else
        Do While Not rsContEduData.EOF
            If IsNull(GetContEduData) Or Len(GetContEduData) = 0 Then
                If Not IsNull(rsContEduData(Field)) Then GetContEduData = rsContEduData(Field)
            Else
                If Not IsNull(rsContEduData(Field)) Then GetContEduData = GetContEduData & "," & rsContEduData(Field)
            End If
            
            rsContEduData.MoveNext
        Loop
    End If
    rsContEduData.Close
    Set rsContEduData = Nothing
End Function

Public Function GetContEduData_Term(EmpNbr, selValue As String, selField As String, selDateValue As Date, selDateField As String, Field As String, DEFAULT, AllRecs As Boolean)
    Dim rsContEduData As New ADODB.Recordset
    
    If EmpNbr <> 0 Then
        If Len(selField) > 0 And Len(selDateField) > 0 Then
            rsContEduData.Open "SELECT " & Field & " FROM TERM_HREDSEM WHERE ES_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "' AND " & selDateField & "= " & Date_SQL(selDateValue), gdbAdoIhr001X, adOpenForwardOnly
        ElseIf Len(selField) > 0 Then
            rsContEduData.Open "SELECT " & Field & " FROM TERM_HREDSEM WHERE ES_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "'", gdbAdoIhr001X, adOpenForwardOnly
        ElseIf Len(selDateField) > 0 Then
            rsContEduData.Open "SELECT " & Field & " FROM TERM_HREDSEM WHERE ES_EMPNBR=" & EmpNbr & " AND " & selDateField & "= '" & selDateValue & "'", gdbAdoIhr001X, adOpenForwardOnly
        Else
            rsContEduData.Open "SELECT " & Field & " FROM TERM_HREDSEM WHERE ES_EMPNBR=" & EmpNbr & " ORDER BY ES_ID DESC", gdbAdoIhr001X, adOpenForwardOnly
        End If
    Else
        If Len(selField) > 0 And Len(selDateField) > 0 Then
            rsContEduData.Open "SELECT " & Field & " FROM TERM_HREDSEM WHERE " & selField & "= '" & selValue & "' AND " & selDateField & "= " & Date_SQL(selDateValue), gdbAdoIhr001X, adOpenForwardOnly
        ElseIf Len(selField) > 0 Then
            rsContEduData.Open "SELECT " & Field & " FROM TERM_HREDSEM WHERE " & selField & "= '" & selValue & "'", gdbAdoIhr001X, adOpenForwardOnly
        ElseIf Len(selDateField) > 0 Then
            rsContEduData.Open "SELECT " & Field & " FROM TERM_HREDSEM WHERE " & selDateField & "= '" & selDateValue & "'", gdbAdoIhr001X, adOpenForwardOnly
        Else
            rsContEduData.Open "SELECT " & Field & " FROM TERM_HREDSEM ORDER BY ES_ID DESC", gdbAdoIhr001X, adOpenForwardOnly
        End If
    End If
    
    GetContEduData_Term = DEFAULT
    
    If Not AllRecs Then
        If Not rsContEduData.EOF Then
            If Not IsNull(rsContEduData(Field)) Then GetContEduData_Term = rsContEduData(Field)
        End If
    Else
        Do While Not rsContEduData.EOF
            If IsNull(GetContEduData_Term) Or Len(GetContEduData_Term) = 0 Then
                If Not IsNull(rsContEduData(Field)) Then GetContEduData_Term = rsContEduData(Field)
            Else
                If Not IsNull(rsContEduData(Field)) Then GetContEduData_Term = GetContEduData_Term & "," & rsContEduData(Field)
            End If
            
            rsContEduData.MoveNext
        Loop
    End If
    rsContEduData.Close
    Set rsContEduData = Nothing
End Function

Public Function GetFormEduData(EmpNbr, selValue As String, selField As String, Field As String, DEFAULT, AllRecs As Boolean)
    Dim rsFormEduData As New ADODB.Recordset
    
    If EmpNbr <> 0 Then
        If Len(selField) > 0 Then
            rsFormEduData.Open "SELECT " & Field & " FROM HREDU WHERE EU_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "'", gdbAdoIhr001, adOpenForwardOnly
        Else
            rsFormEduData.Open "SELECT " & Field & " FROM HREDU WHERE EU_EMPNBR=" & EmpNbr & " ORDER BY EU_ID DESC", gdbAdoIhr001, adOpenForwardOnly
        End If
    Else
        If Len(selField) > 0 Then
            rsFormEduData.Open "SELECT " & Field & " FROM HREDU WHERE " & selField & "= '" & selValue & "'", gdbAdoIhr001, adOpenForwardOnly
        Else
            rsFormEduData.Open "SELECT " & Field & " FROM HREDU ORDER BY EU_ID DESC", gdbAdoIhr001, adOpenForwardOnly
        End If
    End If
    
    GetFormEduData = DEFAULT
    
    If Not AllRecs Then
        If Not rsFormEduData.EOF Then
            If Not IsNull(rsFormEduData(Field)) Then GetFormEduData = rsFormEduData(Field)
        End If
    Else
        Do While Not rsFormEduData.EOF
            If IsNull(GetFormEduData) Or Len(GetFormEduData) = 0 Then
                If Not IsNull(rsFormEduData(Field)) Then GetFormEduData = rsFormEduData(Field)
            Else
                If Not IsNull(rsFormEduData(Field)) Then GetFormEduData = GetFormEduData & "," & rsFormEduData(Field)
            End If
            
            rsFormEduData.MoveNext
        Loop
    End If
    rsFormEduData.Close
    Set rsFormEduData = Nothing
End Function

Public Function GetFormEduData_Term(EmpNbr, selValue As String, selField As String, Field As String, DEFAULT, AllRecs As Boolean)
    Dim rsFormEduData As New ADODB.Recordset
    
    If EmpNbr <> 0 Then
        If Len(selField) > 0 Then
            rsFormEduData.Open "SELECT " & Field & " FROM Term_EDU WHERE EU_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "'", gdbAdoIhr001X, adOpenForwardOnly
        Else
            rsFormEduData.Open "SELECT " & Field & " FROM Term_EDU WHERE EU_EMPNBR=" & EmpNbr & " ORDER BY EU_ID DESC", gdbAdoIhr001X, adOpenForwardOnly
        End If
    Else
        If Len(selField) > 0 Then
            rsFormEduData.Open "SELECT " & Field & " FROM Term_EDU WHERE " & selField & "= '" & selValue & "'", gdbAdoIhr001X, adOpenForwardOnly
        Else
            rsFormEduData.Open "SELECT " & Field & " FROM Term_EDU ORDER BY EU_ID DESC", gdbAdoIhr001X, adOpenForwardOnly
        End If
    End If
    
    GetFormEduData_Term = DEFAULT
    
    If Not AllRecs Then
        If Not rsFormEduData.EOF Then
            If Not IsNull(rsFormEduData(Field)) Then GetFormEduData_Term = rsFormEduData(Field)
        End If
    Else
        Do While Not rsFormEduData.EOF
            If IsNull(GetFormEduData_Term) Or Len(GetFormEduData_Term) = 0 Then
                If Not IsNull(rsFormEduData(Field)) Then GetFormEduData_Term = rsFormEduData(Field)
            Else
                If Not IsNull(rsFormEduData(Field)) Then GetFormEduData_Term = GetFormEduData_Term & "," & rsFormEduData(Field)
            End If
            
            rsFormEduData.MoveNext
        Loop
    End If
    rsFormEduData.Close
    Set rsFormEduData = Nothing
End Function

Public Function GetCounsellingData(EmpNbr, selValue As String, selField As String, selDateValue As Date, selDateField As String, Field As String, DEFAULT)
    Dim rsCounselData As New ADODB.Recordset
    
    If Len(selField) > 0 And Len(selDateField) > 0 Then
        rsCounselData.Open "SELECT " & Field & " FROM HR_COUNSEL WHERE CL_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "' AND " & selDateField & "= " & Date_SQL(selDateValue), gdbAdoIhr001, adOpenForwardOnly
    ElseIf Len(selField) > 0 Then
        rsCounselData.Open "SELECT " & Field & " FROM HR_COUNSEL WHERE CL_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "'", gdbAdoIhr001, adOpenForwardOnly
    ElseIf Len(selDateField) > 0 Then
        rsCounselData.Open "SELECT " & Field & " FROM HR_COUNSEL WHERE CL_EMPNBR=" & EmpNbr & " AND " & selDateField & "= '" & selDateValue & "'", gdbAdoIhr001, adOpenForwardOnly
    Else
        rsCounselData.Open "SELECT " & Field & " FROM HR_COUNSEL WHERE CL_EMPNBR=" & EmpNbr & " ORDER BY CL_ID DESC", gdbAdoIhr001, adOpenForwardOnly
    End If
    GetCounsellingData = DEFAULT
    
    If Not rsCounselData.EOF Then
        If Not IsNull(rsCounselData(Field)) Then GetCounsellingData = rsCounselData(Field)
    End If
    rsCounselData.Close
    Set rsCounselData = Nothing
End Function

Public Function GetCommentsData(EmpNbr, selValue As String, selField As String, selDateValue As Date, selDateField As String, Field As String, DEFAULT)
    Dim rsCommentsData As New ADODB.Recordset
    
    If Len(selField) > 0 And Len(selDateField) > 0 Then
        rsCommentsData.Open "SELECT " & Field & " FROM HR_COMMENTS WHERE CO_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "' AND " & selDateField & "= " & Date_SQL(selDateValue), gdbAdoIhr001, adOpenForwardOnly
    ElseIf Len(selField) > 0 Then
        rsCommentsData.Open "SELECT " & Field & " FROM HR_COMMENTS WHERE CO_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "'", gdbAdoIhr001, adOpenForwardOnly
    ElseIf Len(selDateField) > 0 Then
        rsCommentsData.Open "SELECT " & Field & " FROM HR_COMMENTS WHERE CO_EMPNBR=" & EmpNbr & " AND " & selDateField & "= '" & selDateValue & "'", gdbAdoIhr001, adOpenForwardOnly
    Else
        rsCommentsData.Open "SELECT " & Field & " FROM HR_COMMENTS WHERE CO_EMPNBR=" & EmpNbr & " ORDER BY CO_COMMENT_ID DESC", gdbAdoIhr001, adOpenForwardOnly
    End If
    GetCommentsData = DEFAULT
    
    If Not rsCommentsData.EOF Then
        If Not IsNull(rsCommentsData(Field)) Then GetCommentsData = rsCommentsData(Field)
    End If
    rsCommentsData.Close
    Set rsCommentsData = Nothing
End Function

Public Function GetEmpPerfData(EmpNbr, selValue As String, selField As String, Field As String, DEFAULT, AllRecs As Boolean)
    Dim rsEmpPerfData As New ADODB.Recordset
    
    If EmpNbr <> 0 Then
        If Len(selField) > 0 Then
            rsEmpPerfData.Open "SELECT " & Field & " FROM HR_PERFORM_HISTORY WHERE PH_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "'", gdbAdoIhr001, adOpenForwardOnly
        Else
            rsEmpPerfData.Open "SELECT " & Field & " FROM HR_PERFORM_HISTORY WHERE PH_EMPNBR=" & EmpNbr & " ORDER BY PH_ID DESC", gdbAdoIhr001, adOpenForwardOnly
        End If
    Else
        If Len(selField) > 0 Then
            rsEmpPerfData.Open "SELECT " & Field & " FROM HR_PERFORM_HISTORY WHERE " & selField & "= '" & selValue & "'", gdbAdoIhr001, adOpenForwardOnly
        Else
            rsEmpPerfData.Open "SELECT " & Field & " FROM HR_PERFORM_HISTORY ORDER BY PH_ID DESC", gdbAdoIhr001, adOpenForwardOnly
        End If
    End If
    
    GetEmpPerfData = DEFAULT
    
    If Not AllRecs Then
        If Not rsEmpPerfData.EOF Then
            If Not IsNull(rsEmpPerfData(Field)) Then GetEmpPerfData = rsEmpPerfData(Field)
        End If
    Else
        Do While Not rsEmpPerfData.EOF
            If IsNull(GetEmpPerfData) Or Len(GetEmpPerfData) = 0 Then
                If Not IsNull(rsEmpPerfData(Field)) And rsEmpPerfData(Field) <> "" Then GetEmpPerfData = rsEmpPerfData(Field)
            Else
                If Not IsNull(rsEmpPerfData(Field)) And rsEmpPerfData(Field) <> "" Then GetEmpPerfData = GetEmpPerfData & "," & rsEmpPerfData(Field)
            End If
            
            rsEmpPerfData.MoveNext
        Loop
    End If
    rsEmpPerfData.Close
    Set rsEmpPerfData = Nothing
End Function

Public Function GetEmpPerfData_Term(EmpNbr, selValue As String, selField As String, Field As String, DEFAULT, AllRecs As Boolean)
    Dim rsEmpPerfData As New ADODB.Recordset
    
    If EmpNbr <> 0 Then
        If Len(selField) > 0 Then
            rsEmpPerfData.Open "SELECT " & Field & " FROM Term_PERFORM_HISTORY WHERE PH_EMPNBR=" & EmpNbr & " AND " & selField & "= '" & selValue & "'", gdbAdoIhr001X, adOpenForwardOnly
        Else
            rsEmpPerfData.Open "SELECT " & Field & " FROM Term_PERFORM_HISTORY WHERE PH_EMPNBR=" & EmpNbr & " ORDER BY PH_ID DESC", gdbAdoIhr001X, adOpenForwardOnly
        End If
    Else
        If Len(selField) > 0 Then
            rsEmpPerfData.Open "SELECT " & Field & " FROM Term_PERFORM_HISTORY WHERE " & selField & "= '" & selValue & "'", gdbAdoIhr001X, adOpenForwardOnly
        Else
            rsEmpPerfData.Open "SELECT " & Field & " FROM Term_PERFORM_HISTORY ORDER BY PH_ID DESC", gdbAdoIhr001X, adOpenForwardOnly
        End If
    End If
    
    GetEmpPerfData_Term = DEFAULT
    
    If Not AllRecs Then
        If Not rsEmpPerfData.EOF Then
            If Not IsNull(rsEmpPerfData(Field)) Then GetEmpPerfData_Term = rsEmpPerfData(Field)
        End If
    Else
        Do While Not rsEmpPerfData.EOF
            If IsNull(GetEmpPerfData_Term) Or Len(GetEmpPerfData_Term) = 0 Then
                If Not IsNull(rsEmpPerfData(Field)) Then GetEmpPerfData_Term = rsEmpPerfData(Field)
            Else
                If Not IsNull(rsEmpPerfData(Field)) Then GetEmpPerfData_Term = GetEmpPerfData_Term & "," & rsEmpPerfData(Field)
            End If
            
            rsEmpPerfData.MoveNext
        Loop
    End If
    rsEmpPerfData.Close
    Set rsEmpPerfData = Nothing
End Function

Public Function AttachmentExists(xEmpnbr, xDocName, Optional xDocKey, Optional xJob, Optional xSDate) As Boolean
    'Check if the attachment exists for an employee in the respective table
    Dim rsDoc As New ADODB.Recordset
    
    AttachmentExists = False
    
    Select Case xDocName
        Case "Resume"
            rsDoc.Open "SELECT * FROM HRDOC_EMP WHERE RE_TYPE='" & UCase(xDocName) & "' AND RE_EMPNBR=" & xEmpnbr, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            If Not rsDoc.EOF Then AttachmentExists = True
            
        Case "OtherInfo"
            'If glbtermopen Then
            '    rsDoc.Open "SELECT * FROM Term_HRDOC_HREMP_OTHER WHERE ER_TYPE='" & UCase(xDocName) & "' AND ER_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            '    If Not rsDoc.EOF Then AttachmentExists = True
            'Else
                rsDoc.Open "SELECT * FROM HRDOC_HREMP_OTHER WHERE ER_TYPE='" & UCase(xDocName) & "' AND ER_EMPNBR=" & xEmpnbr, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
                If Not rsDoc.EOF Then AttachmentExists = True
            'End If
        
        Case "EmployeeFlag"
            rsDoc.Open "SELECT * FROM HRDOC_EMP_FLAGS WHERE EF_FLAG = " & glbEmpFlagNo & " AND EF_TYPE='" & UCase(xDocName) & "' AND EF_EMPNBR=" & xEmpnbr, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            If Not rsDoc.EOF Then AttachmentExists = True
        
        Case "Offer"
            rsDoc.Open "SELECT * FROM HRDOC_JOB_HISTORY WHERE DJ_TYPE='" & UCase(xDocName) & "' AND DJ_EMPNBR=" & xEmpnbr & " AND DJ_JOB= '" & xJob & "' AND DJ_SDATE =" & Date_SQL(xSDate), gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            If Not rsDoc.EOF Then AttachmentExists = True
        
        Case "Performance"
            rsDoc.Open "SELECT * FROM HRDOC_PERFORM_HISTORY WHERE DH_TYPE='" & UCase(xDocName) & "' AND DH_EMPNBR=" & xEmpnbr & " AND DH_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            If Not rsDoc.EOF Then AttachmentExists = True
        
        Case "DollarEnt"
            rsDoc.Open "SELECT * FROM HRDOC_HRDOLENT WHERE DE_TYPE='" & UCase(xDocName) & "' AND DE_EMPNBR=" & xEmpnbr & " AND DE_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            If Not rsDoc.EOF Then AttachmentExists = True
            
        Case "Attendance"
            'If glbtermopen Then
            '    rsDoc.Open "SELECT * FROM Term_HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(xDocName) & "' AND AD_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND AD_REASON='" & glbAttReason & "' AND AD_DOA=" & Date_SQL(glbAttDOA) & " AND AD_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            '    If Not rsDoc.EOF Then AttachmentExists = True
            'Else
                rsDoc.Open "SELECT * FROM HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(xDocName) & "' AND AD_EMPNBR=" & xEmpnbr & " AND AD_DOCKEY =" & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
                If Not rsDoc.EOF Then AttachmentExists = True
            'End If
        
        Case "Associations"
            'If glbtermopen Then
            '    rsDoc.Open "SELECT * FROM Term_HRDOC_TRADE WHERE TD_TYPE='" & UCase(xDocName) & "' AND TD_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND TD_CODE='" & glbAssocCode & "' AND TD_BEGINDT=" & Date_SQL(glbBeginDt)    '" AND TD_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            '    If Not rsDoc.EOF Then AttachmentExists = True
            'Else
                rsDoc.Open "SELECT * FROM HRDOC_TRADE WHERE TD_TYPE ='" & UCase(xDocName) & "' AND TD_EMPNBR=" & xEmpnbr & " AND TD_DOCKEY =" & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
                If Not rsDoc.EOF Then AttachmentExists = True
            'End If
        
        Case "EdSem"
            rsDoc.Open "SELECT * FROM HRDOC_EDSEM WHERE ES_TYPE='" & UCase(xDocName) & "' AND ES_EMPNBR=" & xEmpnbr & " AND ES_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            If Not rsDoc.EOF Then AttachmentExists = True
        
        Case "FormalEdu"
            'If glbtermopen Then
            '    rsDoc.Open "SELECT * FROM Term_HRDOC_HREDU WHERE EU_TYPE='" & UCase(xDocName) & "' AND EU_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND EU_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            '    If Not rsDoc.EOF Then AttachmentExists = True
            'Else
                rsDoc.Open "SELECT * FROM HRDOC_HREDU WHERE EU_TYPE='" & UCase(xDocName) & "' AND EU_EMPNBR=" & xEmpnbr & " AND EU_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
                If Not rsDoc.EOF Then AttachmentExists = True
            'End If
        
        Case "Counsel"
            'If glbtermopen Then
            '    rsDoc.Open "SELECT * FROM Term_HRDOC_COUNSEL WHERE DC_TYPE='" & UCase(xDocName) & "' AND DC_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND DC_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            '    If Not rsDoc.EOF Then AttachmentExists = True
            'Else
                rsDoc.Open "SELECT * FROM HRDOC_COUNSEL WHERE DC_TYPE='" & UCase(xDocName) & "' AND DC_EMPNBR=" & xEmpnbr & " AND DC_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
                If Not rsDoc.EOF Then AttachmentExists = True
            'End If
        
        Case "Comments"
            'If glbtermopen Then
            '    rsDoc.Open "SELECT * FROM Term_HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(xDocName) & "' AND DO_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND DO_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            '    If Not rsDoc.EOF Then AttachmentExists = True
            'Else
                rsDoc.Open "SELECT * FROM HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(xDocName) & "' AND DO_EMPNBR=" & xEmpnbr & " AND DO_DOCKEY= " & xDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
                If Not rsDoc.EOF Then AttachmentExists = True
            'End If
        
        Case "Jobdescription"
            rsDoc.Open "SELECT * FROM HRDOC_JOB WHERE DB_TYPE='" & UCase(xDocName) & "' AND DB_JOB= '" & xJob & "'", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            If Not rsDoc.EOF Then AttachmentExists = True
            
    End Select
End Function

Public Function DeleteAttachment(xEmpnbr, xDocName, Optional xDocKey, Optional xJob, Optional xSDate, Optional xPosSkill)
    'Delete the attachment for an employee in the respective table
    
    Select Case xDocName
        Case "Resume"
            gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_EMP WHERE RE_TYPE='" & UCase(xDocName) & "' AND RE_EMPNBR=" & xEmpnbr
                
        Case "OtherInfo"
            'If glbtermopen Then
            '    gdbAdoIhr001_DOC.Execute "DELETE FROM Term_HRDOC_HREMP_OTHER WHERE ER_TYPE='" & UCase(xDocName) & "' AND ER_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            'Else
                gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_HREMP_OTHER WHERE ER_TYPE='" & UCase(xDocName) & "' AND ER_EMPNBR=" & xEmpnbr
            'End If
        
        Case "EmployeeFlag"
            gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_EMP_FLAGS WHERE EF_FLAG = " & glbEmpFlagNo & " AND EF_TYPE='" & UCase(xDocName) & "' AND EF_EMPNBR=" & xEmpnbr
        
        Case "Offer"
            gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_JOB_HISTORY WHERE DJ_TYPE='" & UCase(xDocName) & "' AND DJ_EMPNBR=" & xEmpnbr & " AND DJ_JOB= '" & xJob & "' AND DJ_SDATE =" & Date_SQL(xSDate)
        
        Case "Performance"
            gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_PERFORM_HISTORY WHERE DH_TYPE='" & UCase(xDocName) & "' AND DH_EMPNBR=" & xEmpnbr & " AND DH_DOCKEY= " & xDocKey & " "
            
            'Ticket #25355 - Remove the link to the master table
            gdbAdoIhr001.Execute "UPDATE HR_PERFORM_HISTORY SET PH_DOCKEY = Null WHERE PH_EMPNBR=" & xEmpnbr & " AND PH_DOCKEY= " & xDocKey & " "
        
        Case "DollarEnt"
            gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_HRDOLENT WHERE DE_TYPE='" & UCase(xDocName) & "' AND DE_EMPNBR=" & xEmpnbr & " AND DE_DOCKEY= " & xDocKey & " "
            
            'Ticket #25355 - Remove the link to the master table
            gdbAdoIhr001.Execute "UPDATE HRDOLENT SET DE_DOCKEY = Null WHERE DE_EMPNBR=" & xEmpnbr & " AND DE_DOCKEY= " & xDocKey & " "
        
        Case "Attendance"
            'If glbtermopen Then
            '    gdbAdoIhr001_DOC.Execute "DELETE FROM Term_HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(xDocName) & "' AND AD_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND AD_REASON='" & glbAttReason & "' AND AD_DOA=" & Date_SQL(glbAttDOA) & " AND AD_DOCKEY= " & xDocKey & " "
            '
            '    'Ticket #25355 - Remove the link to the master table
            '    gdbAdoIhr001.Execute "UPDATE Term_ATTENDANCE SET AD_DOCKEY = Null WHERE AD_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND AD_REASON='" & glbAttReason & "' AND AD_DOA=" & Date_SQL(glbAttDOA) & " AND AD_DOCKEY= " & xDocKey & " "
            '
            'Else
                'gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(xDocName) & "' AND AD_EMPNBR=" & xEmpnbr & " AND AD_REASON='" & xCode & "' AND AD_DOA=" & Date_SQL(xDATE) & " AND AD_DOCKEY= " & xDocKey & " "
                gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(xDocName) & "' AND AD_EMPNBR=" & xEmpnbr & " AND AD_DOCKEY =" & xDocKey & " "
                
                'Ticket #25355 - Remove the link to the master table
                'gdbAdoIhr001.Execute "UPDATE HR_ATTENDANCE SET AD_DOCKEY = Null WHERE AD_EMPNBR=" & xEmpnbr & " AND AD_REASON='" & xCode & "' AND AD_DOA=" & Date_SQL(xDATE) & " AND AD_DOCKEY= " & xDocKey & " "
                gdbAdoIhr001.Execute "UPDATE HR_ATTENDANCE SET AD_DOCKEY = Null WHERE AD_EMPNBR=" & xEmpnbr & " AND AD_DOCKEY= " & xDocKey & " "
            'End If
        
        Case "Associations"
            'If glbtermopen Then
            '    gdbAdoIhr001_DOC.Execute "DELETE FROM Term_HRDOC_TRADE WHERE TD_TYPE='" & UCase(xDocName) & "' AND TD_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND TD_CODE='" & glbAssocCode & "' AND TD_BEGINDT=" & Date_SQL(glbBeginDt)    '" AND TD_DOCKEY= " & xDocKey & " "
            'Else
                'gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_TRADE WHERE TD_TYPE ='" & UCase(xDocName) & "' AND TD_EMPNBR=" & xEmpnbr & " AND TD_CODE='" & xCode & "' AND TD_BEGINDT=" & Date_SQL(xDATE)    '" AND TD_DOCKEY= " & xDocKey & " "
                gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_TRADE WHERE TD_TYPE ='" & UCase(xDocName) & "' AND TD_EMPNBR=" & xEmpnbr & " AND TD_DOCKEY =" & xDocKey & " "
            
                'Ticket #25355 - Remove the link to the master table
                'gdbAdoIhr001.Execute "UPDATE HRDOC_TRADE SET TD_DOCKEY = Null WHERE TD_EMPNBR=" & xEmpnbr & " AND TD_CODE='" & xCode & "' AND TD_BEGINDT=" & Date_SQL(xDATE) & " AND TD_DOCKEY= " & xDocKey & " "
                gdbAdoIhr001.Execute "UPDATE HRDOC_TRADE SET TD_DOCKEY = Null WHERE TD_EMPNBR=" & xEmpnbr & " AND TD_DOCKEY= " & xDocKey & " "
            'End If
        
        Case "EdSem"
            gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_EDSEM WHERE ES_TYPE='" & UCase(xDocName) & "' AND ES_EMPNBR=" & xEmpnbr & " AND ES_DOCKEY= " & xDocKey & " "
            
            'Ticket #25355 - Remove the link to the master table
            gdbAdoIhr001.Execute "UPDATE HREDSEM SET ES_DOCKEY = Null WHERE ES_EMPNBR=" & xEmpnbr & " AND ES_DOCKEY= " & xDocKey & " "
        
        Case "FormalEdu"
            'If glbtermopen Then
            '    gdbAdoIhr001_DOC.Execute "DELETE FROM Term_HRDOC_HREDU WHERE EU_TYPE='" & UCase(xDocName) & "' AND EU_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND EU_DOCKEY= " & xDocKey & " "
            '
            '    'Ticket #25355 - Remove the link to the master table
            '    gdbAdoIhr001.Execute "UPDATE Term_EDU SET EU_DOCKEY = Null WHERE EU_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND EU_DOCKEY= " & xDocKey & " "
            'Else
                gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_HREDU WHERE EU_TYPE='" & UCase(xDocName) & "' AND EU_EMPNBR=" & xEmpnbr & " AND EU_DOCKEY= " & xDocKey & " "
                
                'Ticket #25355 - Remove the link to the master table
                gdbAdoIhr001.Execute "UPDATE HREDU SET EU_DOCKEY = Null WHERE EU_EMPNBR=" & xEmpnbr & " AND EU_DOCKEY= " & xDocKey & " "
            'End If
        
        Case "Counsel"
            'If glbtermopen Then
            '    gdbAdoIhr001_DOC.Execute "DELETE FROM Term_HRDOC_COUNSEL WHERE DC_TYPE='" & UCase(xDocName) & "' AND DC_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND DC_DOCKEY= " & xDocKey & " "
            '
            '    'Ticket #25355 - Remove the link to the master table
            '    gdbAdoIhr001.Execute "UPDATE Term_HR_COUNSEL SET CL_DOCKEY = Null WHERE CL_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND CL_DOCKEY= " & xDocKey & " "
            'Else
                gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_COUNSEL WHERE DC_TYPE='" & UCase(xDocName) & "' AND DC_EMPNBR=" & xEmpnbr & " AND DC_DOCKEY= " & xDocKey & " "
                
                'Ticket #25355 - Remove the link to the master table
                gdbAdoIhr001.Execute "UPDATE HR_COUNSEL SET CL_DOCKEY = Null WHERE CL_EMPNBR=" & xEmpnbr & " AND CL_DOCKEY= " & xDocKey & " "
            'End If
        
        Case "Comments"
            'If glbtermopen Then
            '    gdbAdoIhr001_DOC.Execute "DELETE FROM Term_HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(xDocName) & "' AND DO_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND DO_DOCKEY= " & xDocKey & " "
            '
            '    'Ticket #25355 - Remove the link to the master table
            '    gdbAdoIhr001.Execute "UPDATE Term_COMMENTS SET CO_DOCKEY = Null WHERE CO_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND CO_DOCKEY= " & xDocKey & " "
            'Else
                gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(xDocName) & "' AND DO_EMPNBR=" & xEmpnbr & " AND DO_DOCKEY= " & xDocKey & " "
                
                'Ticket #25355 - Remove the link to the master table
                gdbAdoIhr001.Execute "UPDATE HR_COMMENTS SET CO_DOCKEY = Null WHERE CO_EMPNBR=" & xEmpnbr & " AND CO_DOCKEY= " & xDocKey & " "
            'End If

        
        Case "Jobdescription"
            gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_JOB WHERE DB_TYPE='" & UCase(xDocName) & "' AND DB_JOB= '" & xJob & "'"

'        Case "INCIDENT"
'            SQLQ = "DELETE FROM HRDOC_HEALTH_SAFETY_2 WHERE DE_TYPE='" & UCase(xDocName) & "' AND DE_EMPNBR=" & xEmpnbr
'            SQLQ = SQLQ & " AND DE_CASE= '" & glbJob & "'"
'            SQLQ = SQLQ & " AND DE_DOCNO= '" & glbDocTmp & "'"
'            gdbAdoIhr001_DOC.Execute SQLQ
'
'        Case "INJURYWF7"
'            'gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(xDocName) & "' AND DO_EMPNBR=" & xEmpnbr & " AND DO_DOCKEY= " & xDocKey & " "
'            'gdbAdoIhr001_DOC.Execute "Update HRDOC_HEALTH_SAFETY set DE_FILEEXT = null WHERE DE_TYPE='" & UCase(xDocName) & "' AND DE_EMPNBR=" & xEmpnbr & " AND DE_CASE= '" & glbJob & "' AND DE_DOCNO ='" & frmEHSAttach.txtDocNum & "'"
'            SQLQ = "DELETE FROM HRDOC_HEALTH_SAFETY_CONCERNSWF7 WHERE W7_TYPE='" & UCase(xDocName) & "' AND W7_EMPNBR=" & xEmpnbr
'            SQLQ = SQLQ & " AND W7_CASE = '" & glbJob & "'"
'            SQLQ = SQLQ & " AND W7_DOCKEY = '" & xDocKey & "'"
'            gdbAdoIhr001_DOC.Execute SQLQ
'
'            'Ticket #25355 - Remove the link to the master table
'            SQLQ = "UPDATE HR_OCC_HEALTH_SAFETY SET EC_DOCKEY = Null WHERE EC_EMPNBR=" & xEmpnbr
'            SQLQ = SQLQ & " AND EC_CASE = '" & glbJob & "'"
'            SQLQ = SQLQ & " AND EC_DOCKEY = '" & xDocKey & "'"
'            gdbAdoIhr001.Execute SQLQ
'
'        Case "INJURYWF7_WRITTENOFR"
'            SQLQ = "DELETE FROM HRDOC_OHS_WRITTEN_OFFER WHERE F7_TYPE='" & UCase(xDocName) & "' AND F7_EMPNBR=" & xEmpnbr
'            SQLQ = SQLQ & " AND F7_CASE = '" & glbJob & "'"
'            SQLQ = SQLQ & " AND F7_DOCKEY = '" & xDocKey & "'"
'            gdbAdoIhr001_DOC.Execute SQLQ
'
'            'Ticket #25355 - Remove the link to the master table
'            SQLQ = "UPDATE HR_OHS_FORM7_SECTIONS SET F7_DOCKEY = Null WHERE F7_EMPNBR=" & xEmpnbr
'            SQLQ = SQLQ & " AND F7_CASE = '" & glbJob & "'"
'            SQLQ = SQLQ & " AND F7_DOCKEY = '" & xDocKey & "'"
'            gdbAdoIhr001.Execute SQLQ
            
'        Case "EdSem_Retest"
'            gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_EDSEM_RETEST WHERE ES_TYPE='" & UCase(xDocName) & "' AND ES_EMPNBR=" & xEmpnbr & " AND ES_DOCKEY= " & xDocKey & " "
'
'            'Ticket #25355 - Remove the link to the master table
'            gdbAdoIhr001.Execute "UPDATE HREDSEM_RETEST SET ES_DOCKEY = Null WHERE ES_EMPNBR=" & xEmpnbr & " AND ES_DOCKEY= " & xDocKey & " "
        
'        Case "Termination"
'            gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_EMP WHERE RE_TYPE='" & UCase(xDocName) & "' AND RE_EMPNBR=" & xEmpnbr
                
'        Case "PositionSkill"
'            gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_JOBSKL WHERE DS_TYPE='" & UCase(xDocName) & "' AND DS_JOB= '" & xJob & "' AND DS_SKILL= '" & xPosSkill & "'"
    End Select

End Function

'Private Function Get_SelectedFilename() As Integer
'    Dim X As Integer
'
'    Get_SelectedFilename = -1
'
'    'Go through the list to look for File name selected.
'    For X = 0 To File1.ListCount - 1
'
'        If File1.selected(X) Then
'            Get_SelectedFilename = X
'            Exit For
'        Else
'            Get_SelectedFilename = -1
'        End If
'    Next X
'End Function

Public Function UpdateDocumentTypeInfo(xDocName, xDocType, xDocDesc)   'xEmpnbr, Optional xDocKey, Optional xJob, Optional xSDate)
    Dim SQLQ As String
    Dim xRowsAffected As Long
    Dim xDocKey
    
    'Update the document of an employee with Docunment Type and Description in the respective tables
    
    'Take the reference to Employee # and Term SEQ or any other fields apart from Document Name out - this update is for all Active and Term Employees.
    
    Select Case xDocName
        Case "Resume"
            gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_EMP SET RE_DOCTYPE = '" & xDocType & "', RE_USRDESC = '" & xDocDesc & "', RE_DOCTYPE_TABL = 'DOCT' WHERE RE_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (RE_DOCTYPE IS NULL OR RE_DOCTYPE = '')", ""), xRowsAffected ' AND RE_EMPNBR=" & xEmpnbr
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
            gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_EMP SET RE_DOCTYPE = '" & xDocType & "', RE_USRDESC = '" & xDocDesc & "', RE_DOCTYPE_TABL = 'DOCT' WHERE RE_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (RE_DOCTYPE IS NULL OR RE_DOCTYPE = '')", ""), xRowsAffected   ' AND RE_EMPNBR=" & xEmpnbr
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
                        
        Case "OtherInfo"
            gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_HREMP_OTHER SET ER_DOCTYPE = '" & xDocType & "', ER_USRDESC = '" & xDocDesc & "', ER_DOCTYPE_TABL = 'DOCT' WHERE ER_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (ER_DOCTYPE IS NULL OR ER_DOCTYPE = '')", ""), xRowsAffected   ' AND ER_EMPNBR=" & xEmpnbr
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
            gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_HREMP_OTHER SET ER_DOCTYPE = '" & xDocType & "', ER_USRDESC = '" & xDocDesc & "', ER_DOCTYPE_TABL = 'DOCT' WHERE ER_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (ER_DOCTYPE IS NULL OR ER_DOCTYPE = '')", ""), xRowsAffected      ' AND ER_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
        
        Case "EmployeeFlag"
            gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_EMP_FLAGS SET EF_DOCTYPE = '" & xDocType & "', EF_USRDESC = '" & xDocDesc & "', EF_DOCTYPE_TABL = 'DOCT' WHERE EF_FLAG = " & glbEmpFlagNo & " AND EF_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (EF_DOCTYPE IS NULL OR EF_DOCTYPE = '')", ""), xRowsAffected      ' AND EF_EMPNBR=" & xEmpnbr
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
            gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_EMP_FLAGS SET EF_DOCTYPE = '" & xDocType & "', EF_USRDESC = '" & xDocDesc & "', EF_DOCTYPE_TABL = 'DOCT' WHERE EF_FLAG = " & glbEmpFlagNo & " AND EF_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (EF_DOCTYPE IS NULL OR EF_DOCTYPE = '')", ""), xRowsAffected     ' AND EF_EMPNBR=" & xEmpnbr
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
            
        Case "Offer"
            If Len(clpCode(10).Text) > 0 Then
                gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_JOB_HISTORY SET DJ_DOCTYPE = '" & xDocType & "', DJ_USRDESC = '" & xDocDesc & "', DJ_DOCTYPE_TABL = 'DOCT' WHERE DJ_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DJ_DOCTYPE IS NULL OR DJ_DOCTYPE = '')", "") & " AND DJ_JOB IN (SELECT JH_JOB FROM " & SQLDatabaseName & ".DBO.HR_JOB_HISTORY WHERE JH_JREASON = '" & clpCode(10).Text & "') AND DJ_SDATE IN (SELECT JH_SDATE FROM " & SQLDatabaseName & ".DBO.HR_JOB_HISTORY WHERE JH_JREASON = '" & clpCode(10).Text & "')", xRowsAffected       ' AND DJ_EMPNBR=" & xEmpnbr & " AND DJ_JOB= '" & xJob & "' AND DJ_SDATE =" & Date_SQL(xSDate)
            Else
                gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_JOB_HISTORY SET DJ_DOCTYPE = '" & xDocType & "', DJ_USRDESC = '" & xDocDesc & "', DJ_DOCTYPE_TABL = 'DOCT' WHERE DJ_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DJ_DOCTYPE IS NULL OR DJ_DOCTYPE = '')", ""), xRowsAffected       ' AND DJ_EMPNBR=" & xEmpnbr & " AND DJ_JOB= '" & xJob & "' AND DJ_SDATE =" & Date_SQL(xSDate)
            End If
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
            
            If Len(clpCode(10).Text) > 0 Then
                gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_JOB_HISTORY SET DJ_DOCTYPE = '" & xDocType & "', DJ_USRDESC = '" & xDocDesc & "', DJ_DOCTYPE_TABL = 'DOCT' WHERE DJ_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DJ_DOCTYPE IS NULL OR DJ_DOCTYPE = '')", "") & " AND DJ_JOB IN (SELECT JH_JOB FROM " & SQLDatabaseName & ".DBO.Term_JOB_HISTORY WHERE JH_JREASON = '" & clpCode(10).Text & "') AND DJ_SDATE IN (SELECT JH_SDATE FROM " & SQLDatabaseName & ".DBO.Term_JOB_HISTORY WHERE JH_JREASON = '" & clpCode(10).Text & "')", xRowsAffected      ' AND DJ_EMPNBR=" & xEmpnbr & " AND DJ_JOB= '" & xJob & "' AND DJ_SDATE =" & Date_SQL(xSDate)
            Else
                gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_JOB_HISTORY SET DJ_DOCTYPE = '" & xDocType & "', DJ_USRDESC = '" & xDocDesc & "', DJ_DOCTYPE_TABL = 'DOCT' WHERE DJ_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DJ_DOCTYPE IS NULL OR DJ_DOCTYPE = '')", ""), xRowsAffected      ' AND DJ_EMPNBR=" & xEmpnbr & " AND DJ_JOB= '" & xJob & "' AND DJ_SDATE =" & Date_SQL(xSDate)
            End If
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
        
        Case "Performance"
            If Len(clpCode(11).Text) > 0 Then
                xDocKey = GetEmpPerfData(0, clpCode(11).Text, "PH_PCODE", "PH_DOCKEY", "", True)
                If Len(xDocKey) > 0 Then
                    gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_PERFORM_HISTORY SET DH_DOCTYPE = '" & xDocType & "', DH_USRDESC = '" & xDocDesc & "', DH_DOCTYPE_TABL = 'DOCT' WHERE DH_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DH_DOCTYPE IS NULL OR DH_DOCTYPE = '')", "") & " AND DH_DOCKEY IN (" & xDocKey & ")", xRowsAffected       ' AND DH_EMPNBR=" & xEmpnbr & " AND DH_DOCKEY= " & xDocKey & " "
                End If
            Else
                gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_PERFORM_HISTORY SET DH_DOCTYPE = '" & xDocType & "', DH_USRDESC = '" & xDocDesc & "', DH_DOCTYPE_TABL = 'DOCT' WHERE DH_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DH_DOCTYPE IS NULL OR DH_DOCTYPE = '')", ""), xRowsAffected       ' AND DH_EMPNBR=" & xEmpnbr & " AND DH_DOCKEY= " & xDocKey & " "
            End If
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
            
            If Len(clpCode(11).Text) > 0 Then
                xDocKey = GetEmpPerfData_Term(0, clpCode(11).Text, "PH_PCODE", "PH_DOCKEY", "", True)
                If Len(xDocKey) > 0 Then
                    gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_PERFORM_HISTORY SET DH_DOCTYPE = '" & xDocType & "', DH_USRDESC = '" & xDocDesc & "', DH_DOCTYPE_TABL = 'DOCT' WHERE DH_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DH_DOCTYPE IS NULL OR DH_DOCTYPE = '')", "") & " AND DH_DOCKEY IN (" & xDocKey & ")", xRowsAffected      ' AND DH_EMPNBR=" & xEmpnbr & " AND DH_DOCKEY= " & xDocKey & " "
                End If
            Else
                gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_PERFORM_HISTORY SET DH_DOCTYPE = '" & xDocType & "', DH_USRDESC = '" & xDocDesc & "', DH_DOCTYPE_TABL = 'DOCT' WHERE DH_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DH_DOCTYPE IS NULL OR DH_DOCTYPE = '')", ""), xRowsAffected      ' AND DH_EMPNBR=" & xEmpnbr & " AND DH_DOCKEY= " & xDocKey & " "
            End If
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
                    
        Case "DollarEnt"
            xDocKey = GetDollarEntData(0, clpCode(2).Text, "DE_TYPE", "DE_DOCKEY", "", True)
            If Len(xDocKey) > 0 Then
                gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_HRDOLENT SET DE_DOCTYPE = '" & xDocType & "', DE_USRDESC = '" & xDocDesc & "', DE_DOCTYPE_TABL = 'DOCT' WHERE DE_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DE_DOCTYPE IS NULL OR DE_DOCTYPE = '')", "") & " AND DE_DOCKEY IN (" & xDocKey & ")", xRowsAffected      ' AND DE_EMPNBR=" & xEmpnbr & " AND DE_DOCKEY= " & xDocKey & " "
                If xRowsAffected > 0 Then
                    glbUPDTCNT = glbUPDTCNT + xRowsAffected
                End If
            End If
            
            xDocKey = GetDollarEntData_Term(0, clpCode(2).Text, "DE_TYPE", "DE_DOCKEY", "", True)
            If Len(xDocKey) > 0 Then
                gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_DOLENT SET DE_DOCTYPE = '" & xDocType & "', DE_USRDESC = '" & xDocDesc & "', DE_DOCTYPE_TABL = 'DOCT' WHERE DE_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DE_DOCTYPE IS NULL OR DE_DOCTYPE = '')", "") & " AND DE_DOCKEY IN (" & xDocKey & ")", xRowsAffected       ' AND DE_EMPNBR=" & xEmpnbr & " AND DE_DOCKEY= " & xDocKey & " "
                If xRowsAffected > 0 Then
                    glbUPDTCNT = glbUPDTCNT + xRowsAffected
                End If
            End If
        Case "Attendance"
            gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_ATTENDANCE SET AD_DOCTYPE = '" & xDocType & "', AD_USRDESC = '" & xDocDesc & "', AD_DOCTYPE_TABL = 'DOCT' WHERE AD_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (AD_DOCTYPE IS NULL OR AD_DOCTYPE = '')", "") & " AND AD_REASON = '" & clpCode(3).Text & "'", xRowsAffected    ' AND AD_EMPNBR=" & xEmpnbr & " AND AD_DOCKEY =" & xDocKey & " "
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
            gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_ATTENDANCE SET AD_DOCTYPE = '" & xDocType & "', AD_USRDESC = '" & xDocDesc & "', AD_DOCTYPE_TABL = 'DOCT' WHERE AD_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (AD_DOCTYPE IS NULL OR AD_DOCTYPE = '')", "") & " AND AD_REASON = '" & clpCode(3).Text & "'", xRowsAffected       ' AND AD_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND AD_REASON='" & glbAttReason & "' AND AD_DOA=" & Date_SQL(glbAttDOA) & " AND AD_DOCKEY= " & xDocKey & " "
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
        
        Case "Associations"
            gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_TRADE SET TD_DOCTYPE = '" & xDocType & "', TD_USRDESC = '" & xDocDesc & "', TD_DOCTYPE_TABL = 'DOCT' WHERE TD_TYPE ='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (TD_DOCTYPE IS NULL OR TD_DOCTYPE = '')", "") & " AND TD_CODE = '" & clpCode(4).Text & "'", xRowsAffected    ' AND TD_EMPNBR=" & xEmpnbr & " AND TD_DOCKEY =" & xDocKey & " "
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
            
            gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_TRADE SET TD_DOCTYPE = '" & xDocType & "', TD_USRDESC = '" & xDocDesc & "', TD_DOCTYPE_TABL = 'DOCT' WHERE TD_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (TD_DOCTYPE IS NULL OR TD_DOCTYPE = '')", "") & " AND TD_CODE = '" & clpCode(4).Text & "'", xRowsAffected      ' AND TD_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND TD_CODE='" & glbAssocCode & "' AND TD_BEGINDT=" & Date_SQL(glbBeginDt)    '" AND TD_DOCKEY= " & xDocKey & " "
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
        
        Case "EdSem"
            xDocKey = GetContEduData(0, clpCode(5).Text, "ES_CRSCODE", "01/01/1900", "", "ES_DOCKEY", "", True)
            If Len(xDocKey) > 0 Then
                gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_EDSEM SET ES_DOCTYPE = '" & xDocType & "', ES_USRDESC = '" & xDocDesc & "', ES_DOCTYPE_TABL = 'DOCT' WHERE ES_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (ES_DOCTYPE IS NULL OR ES_DOCTYPE = '')", "") & " AND ES_DOCKEY IN (" & xDocKey & ")", xRowsAffected    ' AND ES_EMPNBR=" & xEmpnbr & " AND ES_DOCKEY= " & xDocKey & " "
                If xRowsAffected > 0 Then
                    glbUPDTCNT = glbUPDTCNT + xRowsAffected
                End If
            End If
            
            xDocKey = GetContEduData_Term(0, clpCode(5).Text, "ES_CRSCODE", "01/01/1900", "", "ES_DOCKEY", "", True)
            If Len(xDocKey) > 0 Then
                gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_EDSEM SET ES_DOCTYPE = '" & xDocType & "', ES_USRDESC = '" & xDocDesc & "', ES_DOCTYPE_TABL = 'DOCT' WHERE ES_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (ES_DOCTYPE IS NULL OR ES_DOCTYPE = '')", "") & " AND ES_DOCKEY IN (" & xDocKey & ")", xRowsAffected     ' AND ES_EMPNBR=" & xEmpnbr & " AND ES_DOCKEY= " & xDocKey & " "
                If xRowsAffected > 0 Then
                    glbUPDTCNT = glbUPDTCNT + xRowsAffected
                End If
            End If
            
        Case "FormalEdu"
            xDocKey = GetFormEduData(0, clpCode(6).Text, "EU_SCHOOL", "EU_DOCKEY", "", True)
            If Len(xDocKey) > 0 Then
                gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_HREDU SET EU_DOCTYPE = '" & xDocType & "', EU_USRDESC = '" & xDocDesc & "', EU_DOCTYPE_TABL = 'DOCT' WHERE EU_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (EU_DOCTYPE IS NULL OR EU_DOCTYPE = '')", "") & " AND EU_DOCKEY IN (" & xDocKey & ")", xRowsAffected    ' AND EU_EMPNBR=" & xEmpnbr & " AND EU_DOCKEY= " & xDocKey & " "
                If xRowsAffected > 0 Then
                    glbUPDTCNT = glbUPDTCNT + xRowsAffected
                End If
            End If
            
            xDocKey = GetFormEduData_Term(0, clpCode(6).Text, "EU_SCHOOL", "EU_DOCKEY", "", True)
            If Len(xDocKey) > 0 Then
                gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_HREDU SET EU_DOCTYPE = '" & xDocType & "', EU_USRDESC = '" & xDocDesc & "', EU_DOCTYPE_TABL = 'DOCT' WHERE EU_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (EU_DOCTYPE IS NULL OR EU_DOCTYPE = '')", "") & " AND EU_DOCKEY IN (" & xDocKey & ")", xRowsAffected   ' AND EU_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND EU_DOCKEY= " & xDocKey & " "
                If xRowsAffected > 0 Then
                    glbUPDTCNT = glbUPDTCNT + xRowsAffected
                End If
            End If
            
        Case "Counsel"
            gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_COUNSEL SET DC_DOCTYPE = '" & xDocType & "', DC_USRDESC = '" & xDocDesc & "', DC_DOCTYPE_TABL = 'DOCT' WHERE DC_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DC_DOCTYPE IS NULL OR DC_DOCTYPE = '')", "") & " AND DC_CLTYPE = '" & clpCode(7).Text & "'", xRowsAffected       ' AND DC_EMPNBR=" & xEmpnbr & " AND DC_DOCKEY= " & xDocKey & " "
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
            
            gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_COUNSEL SET DC_DOCTYPE = '" & xDocType & "', DC_USRDESC = '" & xDocDesc & "', DC_DOCTYPE_TABL = 'DOCT' WHERE DC_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DC_DOCTYPE IS NULL OR DC_DOCTYPE = '')", "") & " AND DC_CLTYPE = '" & clpCode(7).Text & "'", xRowsAffected      ' AND DC_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND DC_DOCKEY= " & xDocKey & " "
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
        
        Case "Comments"
            gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_COMMENTS SET DO_DOCTYPE = '" & xDocType & "', DO_USRDESC = '" & xDocDesc & "', DO_DOCTYPE_TABL = 'DOCT' WHERE DO_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DO_DOCTYPE IS NULL OR DO_DOCTYPE = '')", "") & " AND DO_COTYPE = '" & clpCode(8).Text & "'", xRowsAffected      ' AND DO_EMPNBR=" & xEmpnbr & " AND DO_DOCKEY= " & xDocKey & " "
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
            
            gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_COMMENTS SET DO_DOCTYPE = '" & xDocType & "', DO_USRDESC = '" & xDocDesc & "', DO_DOCTYPE_TABL = 'DOCT' WHERE DO_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DO_DOCTYPE IS NULL OR DO_DOCTYPE = '')", "") & " AND DO_COTYPE = '" & clpCode(8).Text & "'", xRowsAffected     ' AND DO_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND DO_DOCKEY= " & xDocKey & " "
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
        
'        Case "Jobdescription"
'            gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_JOB WHERE DB_TYPE='" & UCase(xDocName) & "' AND DB_JOB= '" & xJob & "'"

        'Incident Documents (includes Form 7 and Form 9)
        Case "INCIDENT"
            gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_HEALTH_SAFETY_2 SET DE_DOCTYPE = '" & xDocType & "', DE_USRDESC = '" & xDocDesc & "', DE_DOCTYPE_TABL = 'DOCT' WHERE DE_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DE_DOCTYPE IS NULL OR DE_DOCTYPE = '')", ""), xRowsAffected       ' AND DE_EMPNBR=" & xEmpnbr
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
            
            gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_HEALTH_SAFETY_2 SET DE_DOCTYPE = '" & xDocType & "', DE_USRDESC = '" & xDocDesc & "', DE_DOCTYPE_TABL = 'DOCT' WHERE DE_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DE_DOCTYPE IS NULL OR DE_DOCTYPE = '')", ""), xRowsAffected     ' AND DE_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If

        'H&S - Form 7 - Concerns Document
        Case "INJURYWF7"
            gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_HEALTH_SAFETY_CONCERNSWF7 SET W7_DOCTYPE = '" & xDocType & "', W7_USRDESC = '" & xDocDesc & "', W7_DOCTYPE_TABL = 'DOCT' WHERE W7_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (W7_DOCTYPE IS NULL OR W7_DOCTYPE = '')", ""), xRowsAffected       ' AND W7_EMPNBR=" & xEmpnbr
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
            
            gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_HEALTH_SAFETY_CONCERNSWF7 SET W7_DOCTYPE = '" & xDocType & "', W7_USRDESC = '" & xDocDesc & "', W7_DOCTYPE_TABL = 'DOCT' WHERE W7_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (W7_DOCTYPE IS NULL OR W7_DOCTYPE = '')", ""), xRowsAffected     ' AND W7_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
            
        'H&S - Form 7 - Written Offer
        Case "INJURYWF7_WRITTENOFR"
            gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_OHS_WRITTEN_OFFER SET F7_DOCTYPE = '" & xDocType & "', F7_USRDESC = '" & xDocDesc & "', F7_DOCTYPE_TABL = 'DOCT' WHERE F7_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (F7_DOCTYPE IS NULL OR F7_DOCTYPE = '')", ""), xRowsAffected       ' AND F7_EMPNBR=" & xEmpnbr
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
            
            gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_OHS_WRITTEN_OFFER SET F7_DOCTYPE = '" & xDocType & "', F7_USRDESC = '" & xDocDesc & "', F7_DOCTYPE_TABL = 'DOCT' WHERE F7_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (F7_DOCTYPE IS NULL OR F7_DOCTYPE = '')", ""), xRowsAffected     ' AND F7_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
            
            
'        Case "EdSem_Retest"
'            gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_EDSEM_RETEST SET RE_DOCTYPE = '" & xDocType & "', RE_USRDESC = '" & xDocDesc & "' WHERE ES_TYPE='" & UCase(xDocName) & "' AND ES_EMPNBR=" & xEmpnbr & " AND ES_DOCKEY= " & xDocKey & " "

'            'Ticket #25355 - Remove the link to the master table
'            gdbAdoIhr001.Execute "UPDATE HREDSEM_RETEST SET ES_DOCKEY = Null WHERE ES_EMPNBR=" & xEmpnbr & " AND ES_DOCKEY= " & xDocKey & " "
        
        Case "Termination"
            gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_EMP SET RE_DOCTYPE = '" & xDocType & "', RE_USRDESC = '" & xDocDesc & "', RE_DOCTYPE_TABL = 'DOCT' WHERE RE_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (RE_DOCTYPE IS NULL OR RE_DOCTYPE = '')", ""), xRowsAffected       ' AND RE_EMPNBR=" & xEmpnbr
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
            gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_EMP SET RE_DOCTYPE = '" & xDocType & "', RE_USRDESC = '" & xDocDesc & "', RE_DOCTYPE_TABL = 'DOCT' WHERE RE_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (RE_DOCTYPE IS NULL OR RE_DOCTYPE = '')", ""), xRowsAffected      ' AND RE_EMPNBR=" & xEmpnbr
            If xRowsAffected > 0 Then
                glbUPDTCNT = glbUPDTCNT + xRowsAffected
            End If
                
        Case "PositionSkill"
            If Len(Trim(clpJob.Text)) > 0 Then
                gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_JOBSKL SET DS_DOCTYPE = '" & xDocType & "', DS_USRDESC = '" & xDocDesc & "', DS_DOCTYPE_TABL = 'DOCT' WHERE DS_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DS_DOCTYPE IS NULL OR DS_DOCTYPE = '')", "") & " AND DS_JOB= '" & clpJob.Text & "' AND DS_SKILL = '" & clpCode(9).Text & "'", xRowsAffected    ' AND DS_JOB= '" & glbPos & "' AND DS_SKILL= '" & glbPosSkill & "'"
            Else
                gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_JOBSKL SET DS_DOCTYPE = '" & xDocType & "', DS_USRDESC = '" & xDocDesc & "', DS_DOCTYPE_TABL = 'DOCT' WHERE DS_TYPE='" & UCase(xDocName) & "'" & IIf(chkReplace.Value = 0, " AND (DS_DOCTYPE IS NULL OR DS_DOCTYPE = '')", "") & " AND DS_SKILL = '" & clpCode(9).Text & "'", xRowsAffected    ' AND DS_JOB= '" & glbPos & "' AND DS_SKILL= '" & glbPosSkill & "'"
            End If
            
    End Select

End Function

