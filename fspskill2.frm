VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmPosSkills 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Skills for Position"
   ClientHeight    =   5925
   ClientLeft      =   150
   ClientTop       =   915
   ClientWidth     =   8970
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
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5925
   ScaleWidth      =   8970
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   270
      Left            =   5460
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5700
      Top             =   3090
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   105
      Left            =   0
      TabIndex        =   10
      Top             =   5820
      Width           =   8970
      _Version        =   65536
      _ExtentX        =   15822
      _ExtentY        =   185
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.74
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
   End
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      DataField       =   "JS_COMMENT"
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
      Left            =   2160
      MaxLength       =   25
      TabIndex        =   2
      Tag             =   "00-Comments on this Skill"
      Top             =   4230
      Width           =   4125
   End
   Begin VB.TextBox txtExperience 
      Appearance      =   0  'Flat
      DataField       =   "JS_EXPFACT"
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
      Left            =   6120
      MaxLength       =   4
      TabIndex        =   5
      Tag             =   "10-Experience"
      Top             =   3900
      Visible         =   0   'False
      Width           =   870
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "JS_SKILL"
      Height          =   285
      Index           =   1
      Left            =   1845
      TabIndex        =   0
      Tag             =   "EDSK-Skill"
      Top             =   3570
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSK"
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JS_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   5340
      MaxLength       =   25
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2460
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JS_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   5820
      MaxLength       =   25
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2460
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JS_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   6300
      MaxLength       =   25
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2460
      Visible         =   0   'False
      Width           =   420
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8970
      _Version        =   65536
      _ExtentX        =   15822
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
         Left            =   2400
         TabIndex        =   19
         Top             =   150
         Width           =   630
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   165
         Width           =   690
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
         TabIndex        =   11
         Top             =   135
         Width           =   630
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Align           =   1  'Align Top
      Bindings        =   "fspskill2.frx":0000
      Height          =   2625
      Left            =   0
      OleObjectBlob   =   "fspskill2.frx":0014
      TabIndex        =   15
      Tag             =   "Skills Lookup"
      Top             =   495
      Width           =   8970
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6300
      Top             =   2580
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
   Begin Threed.SSCheck chkRequired 
      DataField       =   "JS_ISSKLREQ"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Tag             =   "40-Skill Required"
      Top             =   4560
      Width           =   2215
      _Version        =   65536
      _ExtentX        =   3907
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Skill Required                      "
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin INFOHR_Controls.CodeLookup clpExpFactor 
      DataField       =   "JS_EXPF"
      Height          =   285
      Left            =   1845
      TabIndex        =   1
      Tag             =   "EDSK-Experience Factor"
      Top             =   3900
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SKEF"
   End
   Begin VB.Image imgNoSec 
      Height          =   240
      Left            =   5040
      Picture         =   "fspskill2.frx":3ABC
      Top             =   4920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblImport 
      Alignment       =   1  'Right Justify
      Caption         =   "Position Skill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3240
      TabIndex        =   20
      Top             =   4920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image imgSec 
      Height          =   240
      Left            =   5040
      Picture         =   "fspskill2.frx":3C06
      Top             =   4920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      Caption         =   "Comments"
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
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   4230
      Width           =   870
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      Caption         =   "Experience Factor"
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
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   3900
      Width           =   1560
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Skill"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   3570
      Width           =   375
   End
   Begin VB.Label lblPOSID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblPOSID"
      DataField       =   "JS_CODE"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   5070
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CompNo"
      DataField       =   "JS_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5700
      TabIndex        =   13
      Top             =   2865
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "frmPosSkills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim fglbRecords%, fglbEditMode%
'Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim fglbNew As Boolean
'Dim fglbUnload As Boolean
'Dim fglbDeac As Boolean
'Dim fglbBeforeChange As Boolean
'Dim ANumber As Integer
'
'Dim fglbLoad As Boolean
'Dim fglbIncorrect
'Dim OCode1, OExp, OComm

Dim rsDATA As New ADODB.Recordset

Private Function chkPosSkills()
Dim SQLQ As String, Msg As String, dd#, PID&, Expr#, Skill$

chkPosSkills = False

On Error GoTo chkPosSkills_Err

If Len(clpCode(1)) < 1 Then
    MsgBox "Skill Code is a required field."
    clpCode(1).SetFocus
    Exit Function
End If
If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Skill Code must be valid."
    clpCode(1).SetFocus
    Exit Function
End If

'Release 8.1
'If txtExperience = "" Then
'    txtExperience = 0
'Else
'    If Not IsNumeric(txtExperience) Then
'        MsgBox "Experience must be numeric."
'        txtExperience.SetFocus
'        Exit Function
'    End If
'End If
If clpExpFactor.Caption = "Unassigned" Then
    MsgBox "Experience Factor is invalid"
    clpExpFactor.SetFocus
    Exit Function
End If

If modISDupSkill() Then
    MsgBox "Skill must be unique"
    clpCode(1).SetFocus
    Exit Function
End If

chkPosSkills = True

Exit Function

chkPosSkills_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HRJOBSKL", "edit/Add")
Call RollBack

End Function

Public Sub cmdCancel_Click()
On Error GoTo Can_Err

fglbNew = False

Call Display_Value

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRJOBSKL", "Cancel")
Call RollBack

End Sub

Public Sub cmdClose_Click()
glbUserUploadMode = SwitchForm: Unload Me
End Sub

Public Sub cmdDelete_Click()
Call clkDelete
End Sub

Public Sub cmdNew_Click()
Dim SQLQ As String

On Error GoTo AddN_Err

fglbNew = True
Call Set_Control("B", Me)
Call SET_UP_MODE

If gsAttachment_DB Then
    lblImport.Visible = True
    imgSec.Visible = False
    imgNoSec.Visible = True
    cmdImport.Visible = True
    glbPosSkill = ""
End If

lblCNum.Caption = "001"
lblPOSID.Caption = glbPos$
clpCode(1).SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "clkNew", "HRJOBSKL", "Add")
Call RollBack
End Sub

Public Sub cmdOK_Click()
Call clkOK
End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String

RHeading = Me.Caption
RHeading = Mid(RHeading, 1, InStr(RHeading, "-"))
RHeading = RHeading & " " & lblPosDesc.Caption
'RHeading = Me.Caption & lblPosDesc.Caption

Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1
End Sub

Public Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = Me.Caption
RHeading = Mid(RHeading, 1, InStr(RHeading, "-"))
RHeading = RHeading & " " & lblPosDesc.Caption
'RHeading = Me.Caption & lblPosDesc.Caption

Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub

Private Sub clpExpFactor_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdImport_Click()
    glbDocNewRecord = fglbNew
    glbDocName = "PositionSkill"
    glbPosSkill = clpCode(1).Text
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmPosSkills")
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMPOSSKILLS"
End Sub

Private Sub Form_Deactivate()
glbUserUploadMode = SwitchForm: Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub

Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)

End Sub

Private Sub Form_Load()
Dim SQLQ
On Error GoTo FLErr

glbOnTop = "FRMPOSSKILLS"

Me.Height = 6000
Me.Width = 7000

Data1.ConnectionString = glbAdoIHRDB

If glbWFC Then 'Ticket #25911 Franks 10/21/2014
    If glbPos = "" Then frmJOBSWFC.Show 1
Else
    If glbPos = "" Then frmJOBS.Show 1
End If

If glbPos = "" Then glbUserUploadMode = UploadFormWithoutCheck: Unload Me: Exit Sub

If EERetrieve() = False Then
    MsgBox "Sorry, Position can not be found"
    If glbWFC Then 'Ticket #25911 Franks 10/21/2014
        frmJOBSWFC.Show 1
    Else
        frmJOBS.Show 1
    End If
Else
    Me.Show
    lblPOSID = glbPos
End If

Screen.MousePointer = DEFAULT

Call Display_Value

Call INI_Controls(Me)

If glbLinamar Then clpCode(1).TextBoxWidth = 2000

Screen.MousePointer = DEFAULT

Exit Sub

FLErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form load Error", "SKills", "Select")
Call RollBack

End Sub

Private Sub Form_LostFocus()
'MDIMain.MainToolBar.ButtonS(2).Enabled = True
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Public Function EERetrieve()
Dim SQLQ$

EERetrieve = False
Screen.MousePointer = HOURGLASS

On Error GoTo EERetrieveErr

SQLQ$ = "SELECT * FROM HRJOBSKL "
SQLQ$ = SQLQ$ & "WHERE JS_CODE = '" & glbPos & "'"
SQLQ$ = SQLQ$ & "ORDER BY JS_CODE"

Data1.RecordSource = SQLQ$
Data1.Refresh


EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERetrieveErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Pos Skills", "HRJOBSK", "SELECT")
Call RollBack

End Function

Private Function modISDupSkill()
Dim SQLQ$
Dim snapSkill As New ADODB.Recordset

modISDupSkill = True

On Error GoTo modISDupSkill_Err

Screen.MousePointer = HOURGLASS
SQLQ$ = "SELECT * FROM HRJOBSKL WHERE JS_CODE = '" & glbPos$ & "' "
SQLQ$ = SQLQ$ & "AND JS_SKILL = '" & clpCode(1).Text & "' "
'Release 8.1
'SQLQ$ = SQLQ$ & "AND JS_EXPFACT = " & txtExperience & " "
'If Len(clpExpFactor.Text) > 0 Then
    SQLQ$ = SQLQ$ & "AND JS_EXPF = '" & clpExpFactor.Text & "' "
'End If


If Not fglbNew Then SQLQ$ = SQLQ$ & "AND JS_ID <> " & Data1.Recordset("JS_ID")

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

Private Sub imgSec_Click()
    Dim SQLQ
    glbPosSkill = clpCode(1).Text
    SQLQ = getSQL("frmPosSkills")
    Call FillMemoFile(SQLQ, "PositionSkill")
End Sub

Private Sub lblPOSID_Change()
lblPOSID.Caption = glbPos$
lblPosition.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$
Me.Caption = "Position Skills - " & lblPosition
End Sub

Private Sub txtComments_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtExperience_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtExperience_LostFocus()

If Len(txtExperience) > 0 Then
    If IsNumeric(txtExperience) Then
        txtExperience = (Int(txtExperience * 100) / 100)
    End If
Else
    txtExperience = 0
End If

End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
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
    
    SQLQ$ = "SELECT * FROM HRJOBSKL "
    SQLQ$ = SQLQ$ & "WHERE JS_CODE = '" & glbPos & "'"
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
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

Public Function clkOK()
Dim SQLQ
Dim xID

On Error GoTo OK_Err

clkOK = False

If Not chkPosSkills() Then Exit Function

If fglbNew Then rsDATA.AddNew: fglbNew = False

Call UpdUStats(Me)
Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
xID = rsDATA("JS_ID")
gdbAdoIhr001.CommitTrans

Data1.Refresh

Data1.Recordset.Find "JS_ID=" & xID

Call Display_Value

clkOK = True

fglbNew = False

If gsAttachment_DB Then
    If glbDocNewRecord Then 'New Record only
        If Len(glbDocImpFile) > 0 Then
            'glbDocKey = xID
            'glbPos = rsDATA("JB_CODE") 'Data1.Recordset("JB_CODE")
            glbPosSkill = rsDATA("JS_SKILL")
            Call AttachmentAdd(glbLEE_ID, glbDocImpFile, glbDocType, glbDocDesc)
            Call DispimgIcon(Me, "frmPosSkills")
        End If
    End If
    glbDocImpFile = ""
End If

Exit Function

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRJOBSKL", "Update")
Call RollBack   '1
End Function

Public Sub clkCancel()
On Error GoTo Can_Err

fglbNew = False

Call Display_Value

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRJOBSKL", "Cancel")
Call RollBack

End Sub

Public Sub clkNew()
Dim SQLQ As String

On Error GoTo AddN_Err

fglbNew = True

Call Set_Control("B", Me)
Call SET_UP_MODE

If gsAttachment_DB Then
    lblImport.Visible = True
    imgSec.Visible = False
    imgNoSec.Visible = True
    cmdImport.Visible = True
    glbPosSkill = ""
End If

lblCNum.Caption = "001"
lblPOSID.Caption = glbPos$
clpCode(1).SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "clkNew", "HRJOBSKL", "Add")
Call RollBack

End Sub

Public Function clkDelete()
Dim a As Integer, Msg As String, INo&

If rsDATA.BOF And rsDATA.EOF Then
    MsgBox "Nothing to Delete"
    Exit Function
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Record?  "
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Function

gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans

If gsAttachment_DB Then
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute "Delete from HRDOC_JOBSKL where DS_TYPE='" & UCase(glbDocName) & "' and DS_JOB='" & glbPos & "' and DS_SKILL='" & clpCode(1).Text & "'"
    gdbAdoIhr001_DOC.CommitTrans
End If

Data1.Refresh

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If
 
fglbNew = False

Call SET_UP_MODE

Exit Function

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "clkDelete", "HRJOBSKL", "Delete")
Call RollBack
End Function

Public Sub clkReport(Destination As DestinationConstants)
Dim RHeading As String

RHeading = Me.Caption
RHeading = Mid(RHeading, 1, InStr(RHeading, "-"))
RHeading = RHeading & " " & lblPosDesc.Caption

Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = Destination
Me.vbxCrystal.Action = 1
End Sub

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


clpCode(1).Enabled = TF
'Release 8.1
'txtExperience.Enabled = TF
clpExpFactor.Enabled = TF
txtComments.Enabled = TF

'Ticket #25273 - '8.0 - Ticket #25273 - ATS changes
chkRequired.Enabled = TF

glbDocName = "PositionSkill"
If gsAttachment_DB Then
    'glbPos = Data1.Recordset("JB_CODE")
    Call DispimgIcon(Me, "frmPosSkills")
    
    If gSec_Upd_Job_Skills Then
        If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
            cmdImport.Visible = False 'George on Jan 26,2006 #10266
        Else
            cmdImport.Visible = True 'George on Jan 26,2006 #10266
        End If
    End If
End If

End Sub

Public Sub Display_Value()
Dim SQLQ

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then rsDATA.Close
    rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "SELECT * FROM HRJOBSKL "
    SQLQ = SQLQ & " WHERE JS_ID = " & Data1.Recordset!JS_ID
    If rsDATA.State <> 0 Then rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
    
    glbPosSkill = rsDATA("JS_SKILL")
End If

Call SET_UP_MODE

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
UpdateRight = gSec_Upd_Job_Skills
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

