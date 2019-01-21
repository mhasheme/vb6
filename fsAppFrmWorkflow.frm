VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSAppFrmWorkflow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Application Form Workflow"
   ClientHeight    =   9420
   ClientLeft      =   90
   ClientTop       =   1005
   ClientWidth     =   13920
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
   ScaleHeight     =   9420
   ScaleWidth      =   13920
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   6135
      LargeChange     =   300
      Left            =   11280
      Max             =   3000
      SmallChange     =   300
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame frWorkflow 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   10815
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "WF_LUSER"
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
         Left            =   9870
         MaxLength       =   10
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   5295
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "WF_LDATE"
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
         Height          =   285
         Index           =   0
         Left            =   8430
         MaxLength       =   12
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   5295
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "WF_LTIME"
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
         Left            =   9180
         MaxLength       =   8
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   5295
         Visible         =   0   'False
         Width           =   645
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   3270
         Top             =   5415
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   7710
         Top             =   5655
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
         GridSource      =   "vbxTrueGrid"
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "WF_POSTYPE"
         Height          =   285
         Index           =   0
         Left            =   1830
         TabIndex        =   0
         Tag             =   "00-Position Type Code"
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "POTY"
      End
      Begin MSComDlg.CommonDialog AttachmentDialog 
         Left            =   7080
         Top             =   5640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin INFOHR_Controls.CodeLookup clpJob1 
         DataField       =   "WF_AUTH_JOB1"
         Height          =   285
         Left            =   1830
         TabIndex        =   1
         Tag             =   "01-Position code"
         Top             =   720
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   5
      End
      Begin INFOHR_Controls.CodeLookup clpJob2 
         DataField       =   "WF_AUTH_JOB2"
         Height          =   285
         Left            =   1830
         TabIndex        =   2
         Tag             =   "01-Position code"
         Top             =   1080
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   5
      End
      Begin INFOHR_Controls.CodeLookup clpJob3 
         DataField       =   "WF_AUTH_JOB3"
         Height          =   285
         Left            =   1830
         TabIndex        =   3
         Tag             =   "01-Position code"
         Top             =   1440
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   5
      End
      Begin INFOHR_Controls.CodeLookup clpJob4 
         DataField       =   "WF_AUTH_JOB4"
         Height          =   285
         Left            =   1830
         TabIndex        =   4
         Tag             =   "01-Position code"
         Top             =   1800
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   5
      End
      Begin INFOHR_Controls.CodeLookup clpJob5 
         DataField       =   "WF_AUTH_JOB5"
         Height          =   285
         Left            =   1830
         TabIndex        =   5
         Tag             =   "01-Position code"
         Top             =   2160
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   5
      End
      Begin VB.Label lblAuthorization 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fifth Authorization"
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
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   1260
      End
      Begin VB.Label lblAuthorization 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fourth Authorization"
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
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label lblAuthorization 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Third Authorization"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1320
      End
      Begin VB.Label lblAuthorization 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Second Authorization"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label lblAuthorization 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "First Authorization"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label lblPosType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Type"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   285
         Width           =   1170
      End
      Begin VB.Label lblCNum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comp"
         DataField       =   "WF_COMPNO"
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
         Left            =   8580
         TabIndex        =   12
         Top             =   5745
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fsAppFrmWorkflow.frx":0000
      Height          =   2535
      Left            =   240
      OleObjectBlob   =   "fsAppFrmWorkflow.frx":0014
      TabIndex        =   6
      Top             =   240
      Width           =   11055
   End
End
Attribute VB_Name = "frmSAppFrmWorkflow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean
Dim fUPMode As Integer
Dim rsDATA As New ADODB.Recordset
Dim UpdateState As UpdateStateEnum
Dim fglbVSQLQ As String
Dim fglbESQLQ As String

Private Function chkAppFormWorkflow()
Dim Msg As String
Dim X%, xchk
Dim SQLQ As String
Dim rsWR As New ADODB.Recordset
Dim xID
Dim a As Integer
Dim xWorkSch

chkAppFormWorkflow = False

If clpCode(0).Caption = "Unassigned" Then
    MsgBox lStr("Position Type must be valid")
    clpCode(0).SetFocus
    Exit Function
End If

If Len(clpCode(0).Text) = 0 Then
    MsgBox lStr("Position Type is a required field")
    clpCode(0).SetFocus
    Exit Function
End If

If fglbNew Then
    If Duplicate_LetterByPosType(clpCode(0).Text) Then
        MsgBox lStr("Duplicate Position Type. Application Form Workflow already exists for Position Type: " & clpCode(0).Text & ".")
        clpCode(0).SetFocus
        Exit Function
    End If
End If

If clpJob1.Caption = "Unassigned" Then
    MsgBox lStr("First Authorization must be valid")
    clpJob1.SetFocus
    Exit Function
End If

If Len(clpJob1.Text) = 0 Then
    MsgBox lStr("First Authorization is a required field")
    clpJob1.SetFocus
    Exit Function
End If

If clpJob2.Caption = "Unassigned" Then
    MsgBox lStr("Second Authorization must be valid")
    clpJob2.SetFocus
    Exit Function
End If
If clpJob3.Caption = "Unassigned" Then
    MsgBox lStr("Third Authorization must be valid")
    clpJob3.SetFocus
    Exit Function
End If
If clpJob4.Caption = "Unassigned" Then
    MsgBox lStr("Fourth Authorization must be valid")
    clpJob4.SetFocus
    Exit Function
End If
If clpJob5.Caption = "Unassigned" Then
    MsgBox lStr("Fifth Authorization must be valid")
    clpJob5.SetFocus
    Exit Function
End If

chkAppFormWorkflow = True

End Function

Sub cmdCancel_Click()

On Error GoTo Can_Err

fglbNew = False

Me.vbxTrueGrid.Enabled = True
Me.vbxTrueGrid.Refresh

rsDATA.CancelUpdate

Call Display_Value

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HRA_APPLFORM_WRKFLOW", "Cancel")
Call RollBack '09June99 js

End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String

If data1.Recordset.BOF And data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
data1.Refresh

Call SET_UP_MODE
'Call ST_UPD_MODE(False)


Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRA_APPLFORM_WRKFLOW", "Delete")
Call RollBack '09June99 js

End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Call ST_UPD_MODE(True)

clpCode(0).SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HRA_APPLFORM_WRKFLOW", "Modify")
Call RollBack '09June99 js

End Sub

Sub cmdNew_Click()

On Error GoTo AddN_Err

Call Set_Control("B", Me)

rsDATA.AddNew

lblCNum.Caption = "001"

fglbNew = True

Call SET_UP_MODE

'Call ST_UPD_MODE(True)
clpCode(0).SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRA_APPLFORM_WRKFLOW", "Add")
Call RollBack '09June99 js

End Sub

Sub cmdOK_Click()
Dim X%
Dim bmk As Variant

On Error GoTo cmdOK_Err

If (data1.Recordset.EOF And data1.Recordset.BOF) Then
    bmk = 0
Else
    bmk = data1.Recordset.Bookmark
End If

If Not chkAppFormWorkflow() Then Exit Sub


Call UpdUStats(Me) ' update user's stats (who did it and when)

Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans

data1.Refresh
If Not bmk = 0 Then
    data1.Recordset.Bookmark = bmk
End If

fglbNew = False

Call Display_Value

Me.vbxTrueGrid.Enabled = True
Me.vbxTrueGrid.SetFocus
Screen.MousePointer = DEFAULT

Exit Sub

cmdOK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRA_APPLFORM_WRKFLOW", "Update")
Call RollBack '09June99 js

End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = "Application Form Workflow"
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

RHeading = "Application Form Workflow"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub clpCode_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpJob1_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpJob2_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpJob3_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpJob4_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpJob5_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRA_APPLFORM_WRKFLOW", "SELECT")

End Sub

Private Sub Form_Activate()

Call SET_UP_MODE

Me.cmdModify_Click

End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim SQLQ

'Me.Show

glbOnTop = "FRMSAPPFRMWORKFLOW"

Screen.MousePointer = HOURGLASS

data1.ConnectionString = glbAdoIHRDB
data1.RecordSource = "SELECT * FROM HRA_APPLFORM_WRKFLOW ORDER BY WF_POSTYPE"
data1.Refresh

Screen.MousePointer = DEFAULT

'Call setRptCaption(Me)
Call setCaption(lblPosType)

'Call setCaption(lblDiv)
'Call setCaption(lblDept)
'Call setCaption(lblLocation)
'Call setCaption(lblRegion)
'Call setCaption(lblAdmin)
'Call setCaption(lblSection)
'Call setCaption(lblUnion)
'Call setCaption(lblPT)

vbxTrueGrid.Columns(0).Caption = lStr("Position Type")

'Call Display_Value

Call ST_UPD_MODE(False)

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT                           '

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

Private Sub Form_Resize()
On Error GoTo Err_WorkScheduleRule_Scroll

If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    If Me.Height >= vbxTrueGrid.Height + frWorkflow.Height + 1000 Then
        scrControl.Value = 0
        frWorkflow.Top = vbxTrueGrid.Height + 520
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        scrControl.Left = Me.ScaleWidth - scrControl.Width
        scrControl.Height = Me.Height - vbxTrueGrid.Height - 1000
    End If
End If

Cont:
Exit Sub

Err_WorkScheduleRule_Scroll:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form_Resize", "Application Form Workflow", "Form Resize")
    Resume Cont
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
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

fUPMode = TF
'vbxTrueGrid.Enabled = FT
'cmdOK.Enabled = TF              '
'cmdCancel.Enabled = TF          '
'cmdClose.Enabled = FT           '
'cmdModify.Enabled = FT          '
'cmdNew.Enabled = FT             '
'cmdDelete.Enabled = FT          '
'cmdPrint.Enabled = FT           '

'clpDiv.Enabled = TF
'clpDept.Enabled = TF
'clpPT.Enabled = TF
clpCode(0).Enabled = TF
clpJob1.Enabled = TF
clpJob2.Enabled = TF
clpJob3.Enabled = TF
clpJob4.Enabled = TF
clpJob5.Enabled = TF

End Sub

Private Sub scrControl_Change()
    frWorkflow.Top = 3000 - scrControl.Value
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
    
    SQLQ = "SELECT * FROM HRA_APPLFORM_WRKFLOW "
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

    data1.RecordSource = SQLQ
    data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I%

On Error GoTo vbxTrueGrid_Err

Call Display_Value

If data1.Recordset.EOF Or data1.Recordset.BOF = 0 Then
    Exit Sub
End If

Exit Sub

vbxTrueGrid_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HRA_APPLFORM_WRKFLOW", "Select")
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

Private Sub Display_Value()
    Dim SQLQ
    
    If data1.Recordset.EOF Or data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
        Call SET_UP_MODE
        
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HRA_APPLFORM_WRKFLOW where WF_ID = '" & data1.Recordset!WF_ID & "'"
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
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_AppFormWorkFlow
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
ElseIf data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If

Call ST_UPD_MODE(TF)

Call set_Buttons(UpdateState)

If Not UpdateRight Then TF = False

End Sub

Private Sub getWSQLQ() 'xType)
Dim xDiv, xDept, xORG, xAsOf, xEMP, xEmpMode, xGRPCE
Dim xLoc, xSection
Dim xFromDate
Dim xToDate
Dim xID
Dim SQLQ As String

fglbESQLQ = "" 'glbSeleDeptUn
fglbVSQLQ = " (1=1) "

'If Len(clpDiv.Text) = 0 Then
'    fglbVSQLQ = fglbVSQLQ & "AND (WR_DIV IS NULL OR WR_DIV='') "
'Else
'    fglbVSQLQ = fglbVSQLQ & "AND WR_DIV = '" & clpDiv.Text & "' "
'End If

'If Len(clpDept.Text) = 0 Then
'    fglbVSQLQ = fglbVSQLQ & " AND (WR_DEPT IS NULL OR WR_DEPT='') "
'Else
'    fglbVSQLQ = fglbVSQLQ & " AND WR_DEPT = '" & clpDept.Text & "' "
'End If

'If Len(clpCode(1).Text) = 0 Then
'    fglbVSQLQ = fglbVSQLQ & " AND (WR_ORG IS NULL OR WR_ORG='') "
'Else
'    fglbVSQLQ = fglbVSQLQ & " AND WR_ORG = '" & clpCode(1).Text & "' "
'End If

'If Len(clpCode(3).Text) = 0 Then
'    fglbVSQLQ = fglbVSQLQ & " AND (WR_EMP IS NULL OR WR_EMP='') "
'Else
'    fglbVSQLQ = fglbVSQLQ & " AND WR_EMP = '" & clpCode(3).Text & "' "
'End If

'If Len(clpPT.Text) = 0 Then
'    fglbVSQLQ = fglbVSQLQ & " AND (WR_PT IS NULL OR WR_PT='') "
'Else
'    fglbVSQLQ = fglbVSQLQ & " AND WR_PT = '" & clpPT.Text & "' "
'End If

'If Len(clpCode(2).Text) = 0 Then
'    fglbVSQLQ = fglbVSQLQ & " AND (WR_ADMINBY IS NULL OR WR_ADMINBY='') "
'Else
'    fglbVSQLQ = fglbVSQLQ & " AND WR_ADMINBY = '" & clpCode(2).Text & "' "
'End If

'If Len(clpCode(0).Text) = 0 Then
'    fglbVSQLQ = fglbVSQLQ & " AND (WR_SECTION IS NULL OR WR_SECTION='') "
'Else
'    fglbVSQLQ = fglbVSQLQ & " AND WR_SECTION = '" & clpCode(0).Text & "' "
'End If

'If Len(clpCode(4).Text) = 0 Then
'    fglbVSQLQ = fglbVSQLQ & " AND (WR_LOC IS NULL OR WR_LOC='') "
'Else
'    fglbVSQLQ = fglbVSQLQ & " AND WR_LOC = '" & clpCode(4).Text & "' "
'End If

If fglbNew Then
    xID = 0
Else
    If Not rsDATA.EOF Then
        xID = rsDATA("WF_ID")
    Else
        xID = 0
    End If
End If
If xID > 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND NOT WF_ID = " & xID & " "
End If
'getWSQLQ = fglbVSQLQ

End Sub

Private Function Duplicate_LetterByPosType(xPosType)
    Dim rsLetPosType As New ADODB.Recordset
    Dim SQLQ As String
    
    Duplicate_LetterByPosType = False
    
    SQLQ = "SELECT * FROM HRA_APPLFORM_WRKFLOW WHERE WF_POSTYPE = '" & xPosType & "'"
    rsLetPosType.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsLetPosType.EOF Then
        Duplicate_LetterByPosType = True
    Else
        Duplicate_LetterByPosType = False
    End If
    rsLetPosType.Close
    Set rsLetPosType = Nothing
End Function
