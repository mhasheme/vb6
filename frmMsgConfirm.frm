VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmMsgConfirm 
   Caption         =   "Confirm Transfer"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   5805
      _Version        =   65536
      _ExtentX        =   10239
      _ExtentY        =   979
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
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
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
         Left            =   240
         TabIndex        =   2
         Tag             =   "Save changes made"
         Top             =   30
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8490
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   3
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "ED_LUSER"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   5580
      MaxLength       =   25
      TabIndex        =   5
      Top             =   5250
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "ED_LTIME"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   5220
      MaxLength       =   25
      TabIndex        =   4
      Top             =   5250
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "ED_LDATE"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   4950
      MaxLength       =   25
      TabIndex        =   3
      Top             =   5250
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame frmBasic 
      BorderStyle     =   0  'None
      Height          =   4305
      Left            =   -90
      TabIndex        =   0
      Top             =   0
      Width           =   8235
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ED_ADMINBY"
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   6
         Top             =   720
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ED_PT"
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   8
         Top             =   360
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDPT"
      End
      Begin INFOHR_Controls.DateLookup dlpTermDate 
         Height          =   285
         Left            =   2760
         TabIndex        =   11
         Tag             =   "41-Date Terminated"
         Top             =   1440
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFollowupDate 
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   1080
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Follow-up Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Tag             =   "41-Date Terminated"
         Top             =   1110
         Width           =   1470
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Termination Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   12
         Tag             =   "41-Date Terminated"
         Top             =   1470
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "FT/PT/CASUAL"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Administered By"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmMsgConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'    If Len(dlpTermDate.Text) < 1 Then
'        MsgBox ("Termination Date is a required field")
'        dlpTermDate.SetFocus
'        Exit Sub
'    End If
'
'    If Not IsDate(dlpTermDate.Text) Then
'        MsgBox ("Termination Date is not a valid date.")
'        dlpTermDate.SetFocus
'        Exit Sub
'    End If
    If Len(dlpFollowupDate.Text) < 1 Then
        MsgBox ("Follow up Date is a required field")
        dlpFollowupDate.SetFocus
        Exit Sub
    End If

    If Not IsDate(dlpFollowupDate.Text) Then
        MsgBox ("Follow up Date is not a valid date.")
        dlpFollowupDate.SetFocus
        Exit Sub
    End If
    If clpCode(0).Caption = "Unassigned" Or Len(clpCode(0).Text) = 0 Then
        MsgBox ("FT/PT/CASUAL is not a valid field")
        clpCode(0).SetFocus
        Exit Sub
    End If
    If clpCode(1).Caption = "Unassigned" Or Len(clpCode(1).Text) = 0 Then
        MsgBox (lblTitle(1).Caption & " is not a valid field")
        clpCode(1).SetFocus
        Exit Sub
    End If

    'glbChgTermDate = dlpTermDate
    glbChgPT = clpCode(0)
    glbChgUseProfile = clpCode(1)
    Call updFollowClientTransfer
    Unload Me
End Sub

Private Sub Form_Load()
MDIMain.panHelp(0).Caption = "info:HR Message"
Call INI_Controls(Me)
lblTitle(1).Caption = lStr(lblTitle(1).Caption)
End Sub

Private Sub updFollowClientTransfer()   'CCAC London
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
On Error GoTo CrFollow_Err
SQLQ = "SELECT * FROM HR_FOLLOW_UP "
SQLQ = SQLQ & " WHERE EF_COMPLETED=0 AND EF_EMPNBR=" & glbLEE_ID
SQLQ = SQLQ & " AND EF_FREAS='CTRS' "
SQLQ = SQLQ & " AND EF_FDATE=" & Date_SQL(dlpFollowupDate)

rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
rsTB.AddNew
Msg = "A Follow Up Record was created!"
rsTB("EF_COMPNO") = "001"
rsTB("EF_EMPNBR") = glbLEE_ID
rsTB("EF_FDATE") = CVDate(dlpFollowupDate)
rsTB("EF_FREAS_TABL") = "FURE"
'Ticket #24257 - Do not update Admin By for them only
If glbCompSerial <> "S/N - 2262W" Then
    rsTB("EF_ADMINBY_TABL") = "EDAB"
    rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
End If

rsTB("EF_FREAS") = "CTRS"
rsTB("EF_COMMENTS") = "To change the Seniority Date that will be counted when the pay is completed."
rsTB("EF_LDATE") = Date
rsTB("EF_LTIME") = Time$
rsTB("EF_LUSER") = glbUserID
rsTB.Update


Dim rsTT As New ADODB.Recordset
rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='CTRS'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
If rsTT.EOF Then
    rsTT.AddNew
    rsTT("TB_COMPNO") = "001"
    rsTT("TB_NAME") = "FURE"
    rsTT("TB_KEY") = "CTRS"
    rsTT("TB_DESC") = "Client Transfer"
    rsTT("TB_LUSER") = glbUserID
    rsTT("TB_LDATE") = Date
    rsTT("TB_LTIME") = Time$
    rsTT.Update
End If
rsTT.Close

'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
'follow up record
Call Grant_FollowUpCode_Security(glbUserID, "CTRS", "Client Transfer")

TheEnd:
rsTB.Close
MsgBox Msg
 
Exit Sub

CrFollow_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Sub


