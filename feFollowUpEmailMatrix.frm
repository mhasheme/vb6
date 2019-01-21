VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmFollowUpEMailMatrix 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Follow Up Code Email Matrix"
   ClientHeight    =   7905
   ClientLeft      =   90
   ClientTop       =   1005
   ClientWidth     =   13530
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
   ScaleHeight     =   7905
   ScaleWidth      =   13530
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRepeatFreq 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "FM_RPEAT_FREQ"
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
      Left            =   4560
      MaxLength       =   4
      TabIndex        =   30
      Top             =   4080
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txtPT 
      Appearance      =   0  'Flat
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
      Left            =   8280
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txtOtherEmails 
      Appearance      =   0  'Flat
      DataField       =   "FM_OTHR_EMAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2130
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Tag             =   "00-Other Email Address(es) to receive Email"
      Top             =   6915
      Width           =   7215
   End
   Begin VB.ComboBox cmbRepeatFreq 
      Appearance      =   0  'Flat
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
      Left            =   2130
      TabIndex        =   2
      Tag             =   "11-Choose Repeat Frequency"
      Top             =   4085
      Width           =   2355
   End
   Begin VB.TextBox txtEmployee 
      Appearance      =   0  'Flat
      DataField       =   "FM_EMP"
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
      Left            =   4560
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4545
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkRA1 
      Alignment       =   1  'Right Justify
      Caption         =   "Rept. Authority 1"
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
      Left            =   330
      TabIndex        =   4
      Tag             =   "Reporting Authority 1 to receive Email"
      Top             =   4915
      Width           =   1990
   End
   Begin VB.TextBox txtRA1 
      Appearance      =   0  'Flat
      DataField       =   "FM_RA1"
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
      Left            =   4560
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkRA2 
      Alignment       =   1  'Right Justify
      Caption         =   "Rept. Authority 2"
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
      Left            =   330
      TabIndex        =   5
      Tag             =   "Reporting Authority 2 to receive Email"
      Top             =   5315
      Width           =   1990
   End
   Begin VB.TextBox txtRA2 
      Appearance      =   0  'Flat
      DataField       =   "FM_RA2"
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
      Left            =   4560
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5295
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkRA3 
      Alignment       =   1  'Right Justify
      Caption         =   "Rept. Authority 3"
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
      Left            =   330
      TabIndex        =   6
      Tag             =   "Reporting Authority 3 to receive Email"
      Top             =   5715
      Width           =   1990
   End
   Begin VB.TextBox txtRA3 
      Appearance      =   0  'Flat
      DataField       =   "FM_RA3"
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
      Left            =   4560
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5655
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkRA4 
      Alignment       =   1  'Right Justify
      Caption         =   "Rept. Authority 4"
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
      Left            =   330
      TabIndex        =   7
      Tag             =   "Reporting Authority 4 to receive Email"
      Top             =   6115
      Width           =   1990
   End
   Begin VB.TextBox txtRA4 
      Appearance      =   0  'Flat
      DataField       =   "FM_RA4"
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
      Left            =   4560
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6030
      Visible         =   0   'False
      Width           =   495
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "FM_FREAS"
      Height          =   285
      Index           =   0
      Left            =   1815
      TabIndex        =   0
      Tag             =   "01-Followup Reason Code"
      Top             =   3285
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "FURE"
      MaxLength       =   7
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   11280
      Top             =   7080
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "FM_LUSER"
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
      Left            =   8040
      MaxLength       =   10
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "FM_LDATE"
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
      Left            =   6600
      MaxLength       =   12
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "FM_LTIME"
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
      Left            =   7320
      MaxLength       =   8
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   645
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   10680
      Top             =   7080
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
   End
   Begin VB.CheckBox chkEmployee 
      Alignment       =   1  'Right Justify
      Caption         =   "Employee"
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
      Left            =   330
      TabIndex        =   3
      Tag             =   "Employee to receive Email"
      Top             =   4515
      Width           =   1990
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feFollowUpEmailMatrix.frx":0000
      Height          =   2895
      Left            =   0
      OleObjectBlob   =   "feFollowUpEmailMatrix.frx":0014
      TabIndex        =   11
      Top             =   120
      Width           =   13215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   7920
      TabIndex        =   10
      Tag             =   "00-Section - Code"
      Top             =   4080
      Visible         =   0   'False
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   8295
      TabIndex        =   27
      Tag             =   "EDPT-Category"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDPT"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin MSMask.MaskEdBox medDayInAdv 
      DataField       =   "FM_DAYSIN_ADV"
      Height          =   285
      Left            =   2130
      TabIndex        =   1
      Tag             =   "10-Days in Advance to receive Follow Up Emails"
      Top             =   3685
      Width           =   865
      _ExtentX        =   1535
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
   Begin INFOHR_Controls.EmployeeLookup elpEEIDOthr 
      DataField       =   "FM_OTHR"
      Height          =   285
      Left            =   1815
      TabIndex        =   8
      Tag             =   "10-Other Employee(s) to receive Emails"
      Top             =   6515
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Email Address(es)"
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
      Left            =   360
      TabIndex        =   29
      Top             =   6915
      Width           =   1680
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Employees"
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
      Left            =   360
      TabIndex        =   28
      Top             =   6560
      Width           =   1200
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6840
      TabIndex        =   25
      Top             =   5565
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
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
      Left            =   6480
      TabIndex        =   24
      Top             =   4080
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Repeat Frequency"
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
      Left            =   330
      TabIndex        =   23
      Top             =   4145
      Width           =   1320
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   16
      Top             =   3330
      Width           =   1275
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Days In Advance"
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
      Index           =   0
      Left            =   330
      TabIndex        =   15
      Top             =   3730
      Width           =   1230
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      DataField       =   "FM_COMPNO"
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
      Left            =   6720
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "frmFollowUpEMailMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fGLBNew As Boolean
Dim fglbSDate As Variant
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim DefType(0 To 3)
Dim SystType(0 To 3)
Dim RSDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim UpdateState As UpdateStateEnum

Private Function chkAttMatrix()
Dim Msg As String
Dim X%, xchk
Dim rsFollowUpMat As New ADODB.Recordset
Dim rsHRTabl As New ADODB.Recordset
Dim SQLQ As String
Dim xCatCount As Integer
Dim xCatSource As String
Dim xCatSearch() As String
Dim I As Integer

chkAttMatrix = False

If Len(clpCode(0).Text) < 1 Then
    MsgBox "Reason Code must be entered"
    clpCode(0).SetFocus
    Exit Function
End If

If Len(clpCode(0).Text) > 0 And clpCode(0).Caption = "Unassigned" Then
    MsgBox "Invalid Reason Code"
    clpCode(0).SetFocus
    Exit Function
End If
       
'Check if the duplicate Reason Code already exists
SQLQ = "SELECT * FROM HR_FOLLOWUP_MATRIX"
SQLQ = SQLQ$ & " WHERE FM_FREAS = '" & clpCode(0).Text & "'"
If Not fGLBNew Then
    SQLQ = SQLQ & " AND FM_ID <> " & Data1.Recordset!FM_ID
End If
SQLQ = SQLQ & " ORDER BY FM_FREAS "
rsFollowUpMat.Open SQLQ$, gdbAdoIhr001, adOpenStatic
If rsFollowUpMat.EOF Then
    rsFollowUpMat.Close
    Set rsFollowUpMat = Nothing
Else
    MsgBox "The " & lStr("Follow-ups Code Email Matrix") & " for this Reason already exists."
    rsFollowUpMat.Close
    Set rsFollowUpMat = Nothing
    clpCode(0).SetFocus
    Exit Function
End If

If cmbRepeatFreq.ListIndex = -1 Then
    cmbRepeatFreq.ListIndex = 0
    'MsgBox "Please select the Repeat Frequency"
    'cmbRepeatFreq.SetFocus
    'Exit Function
End If

If Not elpEEIDOthr.ListChecker Then
    Exit Function
End If

'Check if the Follow Up code has Send Email checked
SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'FURE' AND TB_KEY = '" & clpCode(0).Text & "'"
SQLQ = SQLQ & " AND TB_USR3 <> 0"
rsHRTabl.Open SQLQ$, gdbAdoIhr001, adOpenStatic
If rsHRTabl.EOF Then
    MsgBox "The Reason code is not set to 'Send Email'. The " & lStr("Follow-ups Code Email Matrix") & " cannot be defined for this Reason code.", vbExclamation, "Reason code not set to Send out Emails"
    rsHRTabl.Close
    Set rsHRTabl = Nothing
    clpCode(0).SetFocus
    Exit Function
Else
    rsHRTabl.Close
    Set rsHRTabl = Nothing
End If

If Not IsNumeric(medDayInAdv) Then medDayInAdv = 0
If IsNull(txtEmployee) Then txtEmployee = 0
If IsNull(txtRA1) Then txtRA1 = 0
If IsNull(txtRA2) Then txtRA2 = 0
If IsNull(txtRA3) Then txtRA3 = 0
If IsNull(txtRA4) Then txtRA4 = 0

chkAttMatrix = True

End Function

Sub cmdCancel_Click()

On Error GoTo Can_Err

fGLBNew = False

If fglbEmptyNew Then
    Me.vbxTrueGrid.Enabled = True
    Me.vbxTrueGrid.Refresh
End If

RSDATA.CancelUpdate

Call Display_Value

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HR_FOLLOWUP_MATRIX", "Cancel")
Call RollBack '09June99 js

End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

gdbAdoIhr001.BeginTrans
RSDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

Call SET_UP_MODE
'Call ST_UPD_MODE(False)

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_FOLLOWUP_MATRIX", "Delete")
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_FOLLOWUP_MATRIX", "Modify")
Call RollBack '09June99 js

End Sub

Sub cmdNew_Click()

On Error GoTo AddN_Err

Call Set_Control("B", Me)

RSDATA.AddNew

lblCNum.Caption = "001"

cmbRepeatFreq.ListIndex = -1

fGLBNew = True

Call SET_UP_MODE

'Call ST_UPD_MODE(True)
clpCode(0).SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_FOLLOWUP_MATRIX", "Add")
Call RollBack '09June99 js

End Sub

Sub cmdOK_Click()
Dim X%
Dim bmk As Variant

On Error GoTo cmdOK_Err

If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    bmk = 0
Else
    bmk = Data1.Recordset.Bookmark
End If

If Not chkAttMatrix() Then Exit Sub

'If glbCompSerial = "S/N - 2411W" Then   'Ticket #24586 - WDGPHU
'    txtPT.Text = "'" & Replace(clpPT.Text, ",", "','") & "'"
'    clpPT.DataField = ""
'End If

Call UpdUStats(Me) ' update user's stats (who did it and when)
Call Set_Control("U", Me, RSDATA)

gdbAdoIhr001.BeginTrans
RSDATA.Update
gdbAdoIhr001.CommitTrans

Data1.Refresh
If Not bmk = 0 Then
    Data1.Recordset.Bookmark = bmk
End If

fGLBNew = False

Call Display_Value

Me.vbxTrueGrid.Enabled = True
Me.vbxTrueGrid.SetFocus
Screen.MousePointer = DEFAULT

Exit Sub

cmdOK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_FOLLOWUP_MATRIX", "Update")
Call RollBack '09June99 js

End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = "Follow Up Code Email Matrix"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Sub cmdView_Click()
Dim RHeading As String

RHeading = "Follow Up Code Email Matrix"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmbRepeatFreq_AddItems()
    cmbRepeatFreq.Clear
    cmbRepeatFreq.AddItem ""
    cmbRepeatFreq.AddItem "Daily"
    cmbRepeatFreq.AddItem "Weekly"
    cmbRepeatFreq.AddItem "Monthly"
End Sub

Private Sub chkEmployee_Click()
    If chkEmployee.Value = 1 Then
        txtEmployee.Text = "1"
    Else
        txtEmployee.Text = "0"
    End If
End Sub

Private Sub chkEmployee_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkRA1_Click()
    If chkRA1.Value = 1 Then
        txtRA1.Text = "1"
    Else
        txtRA1.Text = "0"
    End If
End Sub

Private Sub chkRA1_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkRA2_Click()
    If chkRA2.Value = 1 Then
        txtRA2.Text = "1"
    Else
        txtRA2.Text = "0"
    End If
End Sub

Private Sub chkRA2_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkRA3_Click()
    If chkRA3.Value = 1 Then
        txtRA3.Text = "1"
    Else
        txtRA3.Text = "0"
    End If
End Sub

Private Sub chkRA3_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkRA4_Click()
    If chkRA4.Value = 1 Then
        txtRA4.Text = "1"
    Else
        txtRA4.Text = "0"
    End If
End Sub

Private Sub chkRA4_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpCode_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpPT_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbRepeatFreq_Click()
    Select Case cmbRepeatFreq.ListIndex
        Case 0: txtRepeatFreq.Text = ""
        Case 1: txtRepeatFreq.Text = "D"
        Case 2: txtRepeatFreq.Text = "W"
        Case 3: txtRepeatFreq.Text = "M"
    End Select
End Sub

Private Sub cmbRepeatFreq_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "FOLLOW UP EMAIL MATRIX", "SELECT")

End Sub

Private Sub elpEEIDOthr_LostFocus()
    elpEEIDOthr = Left(elpEEIDOthr, 500)
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
Me.cmdModify_Click
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim I%, SQLQ

'Me.Show

glbOnTop = "FRMFOLLOWUPEMAILMATRIX"

Screen.MousePointer = HOURGLASS

Call cmbRepeatFreq_AddItems

Data1.ConnectionString = glbAdoIHRDB
SQLQ = "SELECT HR_FOLLOWUP_MATRIX.*, HRTABL.TB_DESC AS FREAS_DESC FROM (HR_FOLLOWUP_MATRIX"
SQLQ = SQLQ & " LEFT JOIN HRTABL ON (HR_FOLLOWUP_MATRIX.FM_FREAS_TABL = HRTABL.TB_NAME) AND (HR_FOLLOWUP_MATRIX.FM_FREAS = HRTABL.TB_KEY)) "
SQLQ = SQLQ & " WHERE HRTABL.TB_USR3 <> 0 ORDER BY FM_FREAS, FREAS_DESC"
Data1.RecordSource = SQLQ
Data1.Refresh

Screen.MousePointer = DEFAULT

'Call Display_Value

Call ST_UPD_MODE(False)

vbxTrueGrid.Columns(5).Caption = lStr("Rept. Authority 1")
vbxTrueGrid.Columns(6).Caption = lStr("Rept. Authority 2")
vbxTrueGrid.Columns(7).Caption = lStr("Rept. Authority 3")
vbxTrueGrid.Columns(8).Caption = lStr("Rept. Authority 4")

chkRA1.Caption = lStr("Rept. Authority 1")
chkRA2.Caption = lStr("Rept. Authority 2")
chkRA3.Caption = lStr("Rept. Authority 3")
chkRA4.Caption = lStr("Rept. Authority 4")

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

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
Dim I As Integer
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
clpCode(0).Enabled = TF
medDayInAdv.Enabled = TF
cmbRepeatFreq.Enabled = TF            '
chkEmployee.Enabled = TF      '
chkRA1.Enabled = TF      '
chkRA2.Enabled = TF      '
chkRA3.Enabled = TF      '
chkRA4.Enabled = TF
elpEEIDOthr.Enabled = TF
txtOtherEmails.Enabled = TF

End Sub

Private Sub medDayInAdv_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtOtherEmails_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtRepeatFreq_Change()
    If txtRepeatFreq = "D" Then
        cmbRepeatFreq.ListIndex = 1
    End If
    
    If txtRepeatFreq = "W" Then
        cmbRepeatFreq.ListIndex = 2
    End If
    
    If txtRepeatFreq = "M" Then
        cmbRepeatFreq.ListIndex = 3
    End If
    
    If txtRepeatFreq = "" Then
        cmbRepeatFreq.ListIndex = 0
    End If
End Sub

Private Sub txtEmployee_Change()
    If txtEmployee = "-1" Or txtEmployee = "1" Then
        chkEmployee = 1
    Else
        chkEmployee = 0
    End If
End Sub

Private Sub txtRA1_Change()
    If txtRA1 = "-1" Or txtRA1 = "1" Then
        chkRA1 = 1
    Else
        chkRA1 = 0
    End If
End Sub

Private Sub txtRA2_Change()
    If txtRA2 = "-1" Or txtRA2 = "1" Then
        chkRA2 = 1
    Else
        chkRA2 = 0
    End If
End Sub

Private Sub txtRA3_Change()
    If txtRA3 = "-1" Or txtRA3 = "1" Then
        chkRA3 = 1
    Else
        chkRA3 = 0
    End If
End Sub

Private Sub txtRA4_Change()
    If txtRA4 = "-1" Or txtRA4 = "1" Then
        chkRA4 = 1
    Else
        chkRA4 = 0
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
    
    SQLQ = "SELECT HR_FOLLOWUP_MATRIX.*, HRTABL.TB_DESC AS FREAS_DESC FROM (HR_FOLLOWUP_MATRIX"
    SQLQ = SQLQ & " LEFT JOIN HRTABL ON (HR_FOLLOWUP_MATRIX.FM_FREAS_TABL = HRTABL.TB_NAME) AND (HR_FOLLOWUP_MATRIX.FM_FREAS = HRTABL.TB_KEY)) "
    SQLQ = SQLQ & " WHERE HRTABL.TB_USR3 <> 0 "
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I%

On Error GoTo vbxTrueGrid_Err

Call Display_Value

If Data1.Recordset.EOF Or Data1.Recordset.BOF = 0 Then
    Exit Sub
End If


Exit Sub

vbxTrueGrid_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_FOLLOWUP_MATRIX", "Select")
Call RollBack '09June99 js

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
    
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
        RSDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Call SET_UP_MODE
        Exit Sub
    End If
    
    
    SQLQ = "SELECT HR_FOLLOWUP_MATRIX.*, HRTABL.TB_DESC AS FREAS_DESC FROM (HR_FOLLOWUP_MATRIX"
    SQLQ = SQLQ & " LEFT JOIN HRTABL ON (HR_FOLLOWUP_MATRIX.FM_FREAS_TABL = HRTABL.TB_NAME) AND (HR_FOLLOWUP_MATRIX.FM_FREAS = HRTABL.TB_KEY)) "
    SQLQ = SQLQ & " WHERE HRTABL.TB_USR3 <> 0 AND FM_ID = " & Data1.Recordset!FM_ID
    SQLQ = SQLQ & " ORDER BY FM_FREAS, FREAS_DESC"
    
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If RSDATA.EOF Or RSDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, RSDATA)
    Call SET_UP_MODE
    
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fGLBNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fGLBNew = True
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_FollowUpEmail_Matrix
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

If fGLBNew Then
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

