VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmPlanData 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Plan Data"
   ClientHeight    =   7095
   ClientLeft      =   90
   ClientTop       =   1005
   ClientWidth     =   9690
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000017&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7095
   ScaleWidth      =   9690
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.DateLookup dlpDueDate 
      DataField       =   "PP_DUEDATE"
      Height          =   285
      Left            =   4950
      TabIndex        =   4
      Tag             =   "40-Enter Due Date"
      Top             =   3000
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpSurvDate 
      DataField       =   "PP_SURVEYD"
      Height          =   285
      Left            =   1350
      TabIndex        =   3
      Tag             =   "41-Enter Survey Date"
      Top             =   3000
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin VB.TextBox txtDescr 
      Appearance      =   0  'Flat
      DataField       =   "PP_DESC"
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
      Left            =   5280
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "01-Enter Plan Description"
      Top             =   2580
      Width           =   3825
   End
   Begin VB.TextBox txtQues8 
      Appearance      =   0  'Flat
      DataField       =   "PP_Q8"
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
      Left            =   1500
      MaxLength       =   30
      TabIndex        =   12
      Tag             =   "00-Enter Question #8"
      Top             =   6040
      Width           =   5955
   End
   Begin VB.TextBox txtQues7 
      Appearance      =   0  'Flat
      DataField       =   "PP_Q7"
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
      Left            =   1500
      MaxLength       =   30
      TabIndex        =   11
      Tag             =   "00-Enter Question #7"
      Top             =   5680
      Width           =   5955
   End
   Begin VB.TextBox txtQues6 
      Appearance      =   0  'Flat
      DataField       =   "PP_Q6"
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
      Left            =   1500
      MaxLength       =   30
      TabIndex        =   10
      Tag             =   "00-Enter Question #6"
      Top             =   5320
      Width           =   5955
   End
   Begin VB.TextBox txtQues5 
      Appearance      =   0  'Flat
      DataField       =   "PP_Q5"
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
      Left            =   1500
      MaxLength       =   30
      TabIndex        =   9
      Tag             =   "00-Enter Question #5"
      Top             =   4960
      Width           =   5955
   End
   Begin VB.TextBox txtQues4 
      Appearance      =   0  'Flat
      DataField       =   "PP_Q4"
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
      Left            =   1500
      MaxLength       =   30
      TabIndex        =   8
      Tag             =   "00-Enter Question #4"
      Top             =   4600
      Width           =   5955
   End
   Begin VB.TextBox txtQues3 
      Appearance      =   0  'Flat
      DataField       =   "PP_Q3"
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
      Left            =   1500
      MaxLength       =   30
      TabIndex        =   7
      Tag             =   "00-Enter Question #3"
      Top             =   4240
      Width           =   5955
   End
   Begin VB.TextBox txtQues2 
      Appearance      =   0  'Flat
      DataField       =   "PP_Q2"
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
      Left            =   1500
      MaxLength       =   30
      TabIndex        =   6
      Tag             =   "00-Enter Question #2"
      Top             =   3880
      Width           =   5955
   End
   Begin VB.TextBox txtQues1 
      Appearance      =   0  'Flat
      DataField       =   "PP_Q1"
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
      Left            =   1500
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "00-Enter Question #1"
      Top             =   3520
      Width           =   5955
   End
   Begin VB.TextBox txtPlanNbr 
      Appearance      =   0  'Flat
      DataField       =   "PP_PLAN"
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
      Left            =   1650
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "01-Enter Plan Number"
      Top             =   2550
      Width           =   1095
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PP_LUSER"
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
      Index           =   2
      Left            =   8040
      TabIndex        =   16
      Top             =   4480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PP_LDATE"
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
      Left            =   7920
      MaxLength       =   12
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4465
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PP_LTIME"
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
      Left            =   8640
      MaxLength       =   8
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4480
      Visible         =   0   'False
      Width           =   645
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fsplan.frx":0000
      Height          =   2145
      Left            =   0
      OleObjectBlob   =   "fsplan.frx":0014
      TabIndex        =   0
      Tag             =   "Plan Data"
      Top             =   240
      Width           =   9105
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6840
      Top             =   6600
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
      Left            =   7680
      Top             =   6600
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
   Begin VB.Label lblDueDate 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
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
      Height          =   255
      Left            =   4080
      TabIndex        =   28
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblDescr 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   27
      Top             =   2610
      Width           =   1095
   End
   Begin VB.Label lblQues8 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Question #8"
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
      Height          =   255
      Left            =   90
      TabIndex        =   26
      Top             =   6075
      Width           =   1215
   End
   Begin VB.Label lblQues7 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Question #7"
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
      Height          =   255
      Left            =   90
      TabIndex        =   25
      Top             =   5715
      Width           =   1095
   End
   Begin VB.Label lblQues6 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Question #6"
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
      Height          =   255
      Left            =   90
      TabIndex        =   24
      Top             =   5355
      Width           =   1215
   End
   Begin VB.Label lblQues5 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Question #5"
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
      Height          =   255
      Left            =   90
      TabIndex        =   23
      Top             =   4995
      Width           =   1215
   End
   Begin VB.Label lblQues4 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Question #4"
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
      Height          =   255
      Left            =   90
      TabIndex        =   22
      Top             =   4635
      Width           =   1215
   End
   Begin VB.Label lblQues3 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Question #3"
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
      Height          =   255
      Left            =   90
      TabIndex        =   21
      Top             =   4275
      Width           =   1215
   End
   Begin VB.Label lblQues2 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Question #2"
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
      Height          =   255
      Left            =   90
      TabIndex        =   20
      Top             =   3915
      Width           =   1215
   End
   Begin VB.Label lblQues1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Question #1"
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
      Height          =   255
      Left            =   90
      TabIndex        =   19
      Top             =   3555
      Width           =   1215
   End
   Begin VB.Label lblSurveyDate 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Survey Date"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   18
      Top             =   3045
      Width           =   1215
   End
   Begin VB.Label lblPlanNbr 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Plan Number"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   17
      Top             =   2565
      Width           =   1215
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      DataField       =   "PP_CO"
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
      Left            =   4065
      TabIndex        =   15
      Top             =   4470
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "frmPlanData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbSDate As Variant
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim DefType(0 To 3)
Dim SystType(0 To 3)
Dim fglbNewRec%
Dim fglbNew As Boolean
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control



Private Sub chkEmpl_Survey()
Dim SQLQ As String, Message As String
Dim dynEmplSurvey As New ADODB.Recordset

Message = "Plan Number " & txtPlanNbr & "is assigned to employees!"
Message = Message & "First Delete the Survey Data!"

SQLQ = "SELECT * FROM HREMPEQU WHERE HREMPEQU.EQ_PLAN = '" & txtPlanNbr & "'"
dynEmplSurvey.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If dynEmplSurvey.RecordCount > 0 Then
  MsgBox Message
  Exit Sub
End If

End Sub

Private Function chkPlanData()
Dim PlanNO As String, SQLQ As String, Msg$
Dim snapPlanNo As New ADODB.Recordset

chkPlanData = False

If Len(txtPlanNbr) <= 0 Then
    MsgBox "Plan Number field is required!"
    txtPlanNbr.SetFocus
    Exit Function
End If

If fglbNewRec Then
    PlanNO$ = CStr(txtPlanNbr)
    SQLQ = "SELECT PP_PLAN FROM HRPARCOP "
    SQLQ = SQLQ & "WHERE PP_PLAN = '" & PlanNO & "'"
    
    snapPlanNo.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If snapPlanNo.BOF And snapPlanNo.EOF Then
        snapPlanNo.Close
    Else
        Msg$ = "This Plan Number already exists"
        MsgBox Msg$
        snapPlanNo.Close
        txtPlanNbr.SetFocus
        Exit Function
    End If
End If

If Len(txtDescr) <= 0 Then
    MsgBox "Description Plan field is required!"
    txtDescr.SetFocus
    Exit Function
End If

If Len(dlpSurvDate.Text) <= 0 Then
    MsgBox "Survey Date field is required!"
    dlpSurvDate.SetFocus
    Exit Function
Else
    If Not IsDate(dlpSurvDate.Text) Then
        MsgBox "Enter a valid Survey Date!"
        dlpSurvDate.SetFocus
        Exit Function
    End If
End If

If Len(dlpDueDate.Text) > 0 Then
  If Not IsDate(dlpDueDate.Text) Then
      MsgBox "Enter a valid Due Date!"
      dlpDueDate.SetFocus
      Exit Function
  End If
End If

chkPlanData = True

End Function



Sub cmdCancel_Click()

On Error GoTo Can_Err
fglbNew = False

rsDATA.CancelUpdate

Call Display_Value
'Call ST_UPD_MODE(False) 'May99 js

fglbNewRec% = False

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HRMATRIX", "Cancel")
Call RollBack

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String
Dim SQLQ As String, Message As String
Dim dynEmplSurvey As New ADODB.Recordset

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
Else
  Message = "Employee Surveys exist for this Plan! " & Chr(10)
  Message = Message & "Delete Employee Surveys for this Plan "
  Message = Message & "before deleting Plan!"
  
  SQLQ = "SELECT * FROM HREMPEQU WHERE HREMPEQU.EQ_PLAN = '" & txtPlanNbr & "'"
  
  dynEmplSurvey.Open SQLQ, gdbAdoIhr001, adOpenKeyset
  
  If Not dynEmplSurvey.EOF And Not dynEmplSurvey.BOF Then
     dynEmplSurvey.MoveLast
  End If
  
  If dynEmplSurvey.RecordCount > 0 Then
    MsgBox Message
    dynEmplSurvey.MoveFirst
    Exit Sub
  End If
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If


Call ST_UPD_MODE(False)


Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRMATRIX", "Delete")
Call RollBack  '08June99 js

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Call ST_UPD_MODE(True) 'May99 js
txtPlanNbr.SetFocus      'Jaddy 10/25/99
Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HRMATRIX", "Modify")
Call RollBack  '08June99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()

On Error GoTo AddN_Err



'Data1.Recordset.AddNew
''' Sam add July 2002 * Remove Binding Control
Call Set_Control("B", Me)
rsDATA.AddNew


lblCNum.Caption = "001"
fglbNew = True
Call SET_UP_MODE
'Call ST_UPD_MODE(True)
txtPlanNbr.SetFocus      'Jaddy 10/25/99
fglbNewRec% = True

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRMATRIX", "Add")
Call RollBack  '08June99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim x%
Dim xChange
On Error GoTo cmdOK_Err

If Not chkPlanData() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)


gdbAdoIhr001.BeginTrans
Call Set_Control("U", Me, rsDATA)
rsDATA.Update
gdbAdoIhr001.CommitTrans

Data1.Refresh

fglbNewRec% = False
fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(False) 'May99 js

Me.vbxTrueGrid.SetFocus

Screen.MousePointer = DEFAULT

Exit Sub

cmdOK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRMATRIX", "Update")
Call RollBack  '08June99 js

End Sub
'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String, xReport
RHeading = "Plan Data"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1
End Sub
Sub cmdView_Click()
Dim RHeading As String, xReport

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = "Plan Data"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub
'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "PAYROLL", "SELECT")

End Sub


Private Sub Form_Activate()
Call SET_UP_MODE
Me.cmdModify_Click
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim I%
glbOnTop = "FRMPLANDATA"
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
Me.Show

'Data1.DatabaseName = glbIHRDB
Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "select * from HRPARCOP"
Data1.Refresh
Screen.MousePointer = DEFAULT

Call ST_UPD_MODE(False)
If Not gSec_Matrix Then                                    'May99 js
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False                             '
End If                                                  '
Call INI_Controls(Me)
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
'If gSec_Upd_EmploymentEQT Then
    Keepfocus = Not isUpdated(Me)
'End If
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
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

fUPMode = TF    ' update mode

'cmdOK.Enabled = TF          'May99 js
'cmdCancel.Enabled = TF      '
'cmdClose.Enabled = FT       '
'cmdModify.Enabled = FT      '
'cmdNew.Enabled = FT         '
'cmdDelete.Enabled = FT      '
'cmdPrint.Enabled = FT       '
txtDescr.Enabled = TF       '
dlpDueDate.Enabled = TF     '
txtPlanNbr.Enabled = TF     '
txtQues1.Enabled = TF       '
txtQues2.Enabled = TF       '
txtQues3.Enabled = TF       '
txtQues4.Enabled = TF       '
txtQues5.Enabled = TF       '
txtQues6.Enabled = TF       '
txtQues7.Enabled = TF       '
txtQues8.Enabled = TF       '
dlpSurvDate.Enabled = TF    '
'vbxTrueGrid.Enabled = FT
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If

End Sub

Private Sub txtDescr_GotFocus()
  Call SetPanHelp(ActiveControl)
End Sub
'Private Sub txtDueDate_Change()
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtDueDate_DblClick()
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub txtDueDate_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub txtDueDate_KeyPress(KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub
Private Sub txtPlanNbr_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtQues1_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub
Private Sub txtQues2_GotFocus()
  Call SetPanHelp(ActiveControl)
End Sub
Private Sub txtQues3_GotFocus()
  Call SetPanHelp(ActiveControl)
End Sub
Private Sub txtQues4_GotFocus()
  Call SetPanHelp(ActiveControl)
End Sub
Private Sub txtQues5_GotFocus()
  Call SetPanHelp(ActiveControl)
End Sub
Private Sub txtQues6_GotFocus()
  Call SetPanHelp(ActiveControl)
End Sub
Private Sub txtQues7_GotFocus()
  Call SetPanHelp(ActiveControl)
End Sub
Private Sub txtQues8_GotFocus()
  Call SetPanHelp(ActiveControl)
End Sub
'Private Sub txtSurvDate_Change()
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtSurvDate_DblClick()
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub txtSurvDate_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
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

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "select * from HRPARCOP"
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

'Private Sub txtSurvDate_KeyPress(KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value

If Data1.Recordset.RecordCount <> 0 Then
    If Not IsNull(Data1.Recordset("PP_DUEDATE")) Then
        dlpDueDate.Text = Data1.Recordset("PP_DUEDATE")
    Else
        dlpDueDate.Text = ""
    End If
End If

End Sub

''' Sam add July 2002 * Remove Binding Control
Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbtermopen Then
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        Call SET_UP_MODE
        Exit Sub
    End If

    
    SQLQ = "SELECT * FROM HRPARCOP "
    
    SQLQ = SQLQ & " WHERE  PP_ID= " & Data1.Recordset!PP_ID
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
UpdateRight = gSec_Upd_EmploymentEQT 'True
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
ElseIf Data1.Recordset.EOF Then
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



