VERSION 5.00
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmUFollow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Mass Update Follow-Ups"
   ClientHeight    =   7260
   ClientLeft      =   525
   ClientTop       =   1515
   ClientWidth     =   8880
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
   ScaleHeight     =   7260
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.DateLookup dlpFUDate 
      Height          =   285
      Index           =   1
      Left            =   3930
      TabIndex        =   10
      Tag             =   "41-Follow-up Date"
      Top             =   4320
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpFUDate 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   9
      Tag             =   "41-Follow-up Date"
      Top             =   4320
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   8
      Tag             =   "01-Follow-up Reason"
      Top             =   3960
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "FURE"
   End
   Begin VB.Frame Frame1 
      Caption         =   "For Mass Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6600
      TabIndex        =   25
      Top             =   3900
      Width           =   1935
      Begin VB.OptionButton SeleCompleted 
         Caption         =   "Incomplete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton SeleCompleted 
         Caption         =   "Completed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton SeleCompleted 
         Caption         =   "All Records"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Tag             =   "00-Comments - free form Memo field"
      Top             =   5280
      Width           =   8310
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   630
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   300
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   945
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   3
      Top             =   1275
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   1605
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Tag             =   "10-Enter Employee Number"
      Top             =   2580
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   6355
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   1920
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1920
      TabIndex        =   6
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   2250
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
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
      TabIndex        =   30
      Top             =   1980
      Width           =   615
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Left            =   240
      TabIndex        =   29
      Top             =   2310
      Width           =   540
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      TabIndex        =   28
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
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
      Left            =   240
      TabIndex        =   27
      Top             =   2580
      Width           =   1290
   End
   Begin VB.Label lblToDate 
      Caption         =   "to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3570
      TabIndex        =   26
      Top             =   4350
      Width           =   375
   End
   Begin VB.Label lblComments 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   300
      TabIndex        =   24
      Top             =   4980
      Width           =   735
   End
   Begin VB.Label lblReviewDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   4380
      Width           =   420
   End
   Begin VB.Label lblReason 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   4050
      Width           =   660
   End
   Begin VB.Label lblFollowUps 
      Caption         =   "Follow-ups"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3810
      Width           =   1095
   End
   Begin VB.Label lblSelCri 
      Caption         =   "Selection Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label textMulti 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "The Union Code and FT/PT/SE/TR/OT will be validated from the Employee Basic Data"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.Label lblEStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status"
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
      TabIndex        =   18
      Top             =   1350
      Width           =   1350
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union Code"
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
      TabIndex        =   17
      Top             =   1020
      Width           =   840
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      Top             =   690
      Width           =   825
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
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
      TabIndex        =   15
      Top             =   360
      Width           =   555
   End
End
Attribute VB_Name = "frmUFollow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fglbAdd%    ' it is a global add request
Dim fglbDelete%, fglbNoDept&
Dim fglbModify%, fglb_FindDept
Dim fglbWSQLQ, fglbESQLQ
Dim fglbSDate As Variant


Private Function chkFUComment()

Dim SQLQ As String, Msg$, dd&, Response%, x%
Dim DgDef As Variant, Title$, DCurPDate As Variant
Dim rs As New ADODB.Recordset
Dim xTemplate As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbUserID)


chkFUComment = False

On Error GoTo chkFUComment_Err

For x% = 0 To 4
If Len(clpCode(x%).Text) > 0 And clpCode(x%).Caption = "Unassigned" Then
    MsgBox "If code entered it must be known"
    clpCode(x%).SetFocus
    Exit Function
End If
Next x%

If Len(clpCode(3).Text) < 1 Then
    MsgBox lStr("Follow-Up Reason is a required field")
    clpCode(3).SetFocus
    Exit Function
Else
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        SQLQ = "SELECT MAINTAINABLE from HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(glbUserID, "'", "''") & "'"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        SQLQ = "SELECT MAINTAINABLE from HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    End If
    SQLQ = SQLQ & " AND CODENAME='" & clpCode(3).Text & "'"
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF = False And rs.BOF = False Then
        If rs("MAINTAINABLE") = 0 Then
            MsgBox "You do not have Authority for '" & clpCode(3).Text & "' Reason code.", vbOKOnly + vbInformation, "Authorization failed"
            rs.Close
            Set rs = Nothing
            clpCode(3).SetFocus
            Exit Function
        End If
    Else
        MsgBox "You do not have Authority for '" & clpCode(3).Text & "' Reason code.", vbOKOnly + vbInformation, "Authorization failed"
        rs.Close
        Set rs = Nothing
        clpCode(3).SetFocus
        Exit Function
    End If
    rs.Close
    Set rs = Nothing
End If


If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
     clpDept.SetFocus
    Exit Function
End If


If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    MsgBox lStr("If Division Entered - it must be known")
     clpDiv.SetFocus
    Exit Function
End If
If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    MsgBox lStr("Category code must be valid")
     clpPT.SetFocus
    Exit Function
End If
If Len(dlpFUDate(0).Text) >= 1 Then
    If Not IsDate(dlpFUDate(0).Text) Then
        MsgBox lStr("Follow-Up Date is not a valid date.")
        dlpFUDate(0).SetFocus
        Exit Function
    End If
Else
    MsgBox lStr("Follow-Up Date is required.")
    dlpFUDate(0).SetFocus
    Exit Function
End If

If Len(dlpFUDate(1).Text) >= 1 Then
    If Not IsDate(dlpFUDate(1).Text) Then
        MsgBox lStr("Follow-Up Date is not a valid date.")
        dlpFUDate(1).SetFocus
        Exit Function
    Else
        If DateDiff("d", dlpFUDate(0).Text, dlpFUDate(1).Text) < 0 Then
            MsgBox lStr("Follow-Up Date range is not a valid date range.")
            dlpFUDate(1).SetFocus
        End If
    End If
End If
' Frank 4/25/2000 Follow Jaddy's SQL Version
If Not elpEEID.ListChecker Then
    Exit Function
End If

chkMUOK:
chkFUComment = True

Exit Function

chkFUComment_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkFUComment", "HR Attendance", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Public Sub cmdClose_Click()
Unload Me

End Sub

Public Sub cmdDelete_Click()
Dim a As Integer
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, rc%, DtTm As Variant, x%
Dim DgDef, Title$, Msg$, Response%
Dim rsTF As New ADODB.Recordset 'Frank 4/25/2000
Dim CntRec

If Not gSec_Upd_Follow_Ups Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Title$ = lStr("Mass Follow-Up Record Delete")
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.

fglbDelete% = True
fglbAdd% = False
fglbModify% = False

If Not chkFUComment() Then Exit Sub

Call getWSQLQ

SQLQ = "SELECT EF_EMPNBR FROM HR_FOLLOW_UP WHERE " & fglbWSQLQ
SQLQ = SQLQ & " AND EF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"

'Friesens - Ticket #16591
If glbCompSerial = "S/N - 2279W" Then
    SQLQ = SQLQ & " AND (EF_FREAS <> 'EDUC' or EF_COMPLETED = 1) "
End If

rsTF.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If rsTF.EOF Or rsTF.BOF Then
    CntRec = 0
Else
    CntRec = rsTF.RecordCount
End If
rsTF.Close

If CntRec = 0 Then
    Msg = "Records for this criteria do not exist."
    MsgBox Msg$, , Title
    Exit Sub
Else
    Msg$ = CntRec & " record" & IIf(CntRec = 1, "", "s") & " to be Deleted." & vbCrLf & vbCrLf & "Are you sure you want to delete Records for this criteria?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then    ' Evaluate response
        Exit Sub
    End If
End If

x% = modDelRecs()

Screen.MousePointer = DEFAULT
MsgBox "Records Deleted Successfully."

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "ATTEND", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Public Sub cmdNew_Click()
Dim SQLQ As String, Msg$, x%
Dim Title$, DgDef As Variant, Response%
On Error GoTo AddN_Err
Dim recCount As Integer

If Not gSec_Upd_Follow_Ups Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

dlpFUDate(1).Text = ""
fglbAdd% = True

If Not chkFUComment() Then Exit Sub

Title$ = lStr("Mass Follow-Up Add")
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to add Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Add
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Follow Up Record " Else Msg$ = Msg$ & " Follow Up Records "
    Msg$ = Msg$ & "will be Added. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, Title)     ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Employee record found to add the Follow Up record."
    Exit Sub
End If

If Not modInsRecs() Then Exit Sub

Screen.MousePointer = DEFAULT
MsgBox "Records Added Successfully."

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "ATTEND", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Form_Activate()
Call SET_UP_MODE

glbOnTop = "FRMUFOLLOW"

End Sub

Private Sub Form_Load()

glbOnTop = "FRMUFOLLOW"

Me.Caption = lStr("Mass Update Follow-ups")
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Screen.MousePointer = HOURGLASS

If glbMulti Then textMulti.Visible = True

textMulti.Caption = "The " & lStr("Union") & " and " & lStr("Category") & " will be validated from the Employee Basic Data"

Call setCaption(lblFollowUps)
Call setRptCaption(Me)
Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."

Set frmUFollow = Nothing   'carmen apr 2000
End Sub

Private Function GetRecs()
Dim BD As Integer
Dim SQLQ As String, countr As Integer
Dim Dat1 As Variant, Dat2 As Variant
Dim iOneWhere As Integer, NxtSQL As String

GetRecs = False
On Error GoTo GetRecs_Err

Screen.MousePointer = HOURGLASS

iOneWhere = True
' first do for departments
SQLQ = " Where " & glbSeleDeptUn
If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND ED_DEPTNO = '" & clpDept.Text & "'"
If Len(clpCode(1).Text) > 0 Then SQLQ = SQLQ & " AND ED_DIV = '" & clpCode(1).Text & "' "
If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND ED_ORG = '" & clpCode(2).Text & "' "
If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND ED_EMP= '" & clpCode(3).Text & "' "
If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND ED_PT = '" & clpPT.Text & "' "
If Len(elpEEID.Text) > 0 Then SQLQ = SQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND ED_LOC = '" & clpCode(0).Text & "' "
If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(4).Text & "' "


Screen.MousePointer = DEFAULT


GetRecs = True

Exit Function

GetRecs_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "attend", "DATA1-5table", "Select")
GetRecs = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub memComments_GotFocus()
Call SetPanHelp(ActiveControl)
MDIMain.panHelp(2).Caption = " "    'laura jan 08, 1998
End Sub

Private Function modDelRecs()
Dim BD As Integer
Dim SQLQ As String, SQL1 As String, countr As Integer
Dim Dat1 As Variant, Dat2 As Variant
Dim iOneWhere As Integer, NxtSQL As String, strReas$

modDelRecs = False
On Error GoTo modDelRecs_Err

Screen.MousePointer = HOURGLASS
SQLQ = "DELETE FROM HR_FOLLOW_UP WHERE " & fglbWSQLQ
SQLQ = SQLQ & " AND EF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"

'Friesens - Ticket #16591
If glbCompSerial = "S/N - 2279W" Then
    SQLQ = SQLQ & " AND (EF_FREAS <> 'EDUC' or EF_COMPLETED = 1) "
End If

gdbAdoIhr001.Execute SQLQ

modDelRecs = True

Exit Function

modDelRecs_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modDelRecs", "DeleteAttend", "Delete")
modDelRecs = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function modInsRecs()

Dim BD As Integer
Dim SQLQ As String, countr As Integer
Dim Dat1 As Variant, Dat2 As Variant
Dim iOneWhere As Integer, NxtSQL As String
Dim strdt$, strTm$, strUser&, strReas$

modInsRecs = False
On Error GoTo modInsRecs_Err

strTm$ = Time$

Call getWSQLQ
Screen.MousePointer = HOURGLASS
SQLQ = "INSERT INTO HR_FOLLOW_UP  "
SQLQ = SQLQ & "(EF_EMPNBR, EF_COMPNO, EF_COMPLETED, EF_FDATE, "
SQLQ = SQLQ & "EF_FREAS, EF_COMMENTS, EF_LDATE, EF_LTIME, EF_LUSER )  "

SQLQ = SQLQ & "SELECT HREMP.ED_EMPNBR AS EF_EMPNBR, '001' AS EF_COMPNO, "
SQLQ = SQLQ & 0 & " AS EF_COMPLETED, "
SQLQ = SQLQ & Date_SQL(dlpFUDate(0).Text) & " AS EF_FDATE,"
SQLQ = SQLQ & "'" & clpCode(3).Text & "' AS EF_FREAS, "
SQLQ = SQLQ & "'" & Replace(memComments, "'", "'+chr(39)+'") & "'" & " AS EF_COMMENTS, "
SQLQ = SQLQ & Date_SQL(Date) & " AS EF_LDATE, "
SQLQ = SQLQ & "'" & strTm$ & "' AS EF_LTIME, "
SQLQ = SQLQ & "'" & glbUserID & "' AS EF_LUSER FROM HREMP "
SQLQ = SQLQ & " WHERE " & fglbESQLQ

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

modInsRecs = True

Exit Function

modInsRecs_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modInsRecs", "Attendance", "Insert")
modInsRecs = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function getWSQLQ()

fglbESQLQ = glbSeleDeptUn
If Len(clpDept.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DEPTNO = '" & clpDept.Text & "'"
If Len(clpDiv.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & clpDiv.Text & "' "
If Len(clpCode(1).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & clpCode(1).Text & "' "
If Len(clpCode(2).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMP= '" & clpCode(2).Text & "' "
If Len(clpCode(0).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & clpCode(0).Text & "' "
If Len(clpCode(4).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION= '" & clpCode(4).Text & "' "

If Len(clpPT.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_PT = '" & clpPT.Text & "' "
If Len(elpEEID.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
If glbNoNONE And clpCode(3).Text = "SREV" Then fglbESQLQ = fglbESQLQ & " AND ED_ORG <> 'NONE' "

fglbWSQLQ = " HR_FOLLOW_UP.EF_FREAS = '" & clpCode(3).Text & "' "
If Len(dlpFUDate(0).Text) >= 1 Then
    If Len(dlpFUDate(1).Text) >= 1 Then
        fglbWSQLQ = fglbWSQLQ & " AND (HR_FOLLOW_UP.EF_FDATE) >= " & Date_SQL(dlpFUDate(0).Text)
    Else
        fglbWSQLQ = fglbWSQLQ & " AND (HR_FOLLOW_UP.EF_FDATE)= " & Date_SQL(dlpFUDate(0).Text)
    End If
End If
If Len(dlpFUDate(1).Text) >= 1 Then fglbWSQLQ = fglbWSQLQ & " AND (HR_FOLLOW_UP.EF_FDATE) <=" & Date_SQL(dlpFUDate(1).Text)
If SeleCompleted(1) Then fglbWSQLQ = fglbWSQLQ & " AND HR_FOLLOW_UP.EF_COMPLETED <> 0 "
If SeleCompleted(2) Then fglbWSQLQ = fglbWSQLQ & " AND HR_FOLLOW_UP.EF_COMPLETED = 0 "
End Function

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum

TF = True
UpdateState = OPENING
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False

'alpAPPNBR.Enabled = TF
End Sub

Public Property Get RelateMode() As RelateModeEnum
RelateMode = MassChanges
End Property

Public Property Get UpdateRight() As Boolean
'UpdateRight = gSec_Upd_Follow_Ups
UpdateRight = GetMassUpdateSecurities("Follow_Ups_MassUpdate", glbUserID)
End Property

Public Property Get Addable() As Boolean
Addable = True
End Property

Public Property Get Updateble() As Boolean
Updateble = False
End Property

Public Property Get Deleteble() As Boolean
Deleteble = True
End Property

Public Property Get Printable() As Boolean
Printable = False
End Property

Private Function getRecordCount_Add()
    Dim SQLQ As String
    Dim rsEMP As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Add = 0
    recCount = 0

    Call getWSQLQ
    
    SQLQ = "SELECT COUNT(HREMP.ED_EMPNBR) AS TOT_REC FROM HREMP "
    SQLQ = SQLQ & " WHERE " & fglbESQLQ
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEMP.EOF Then
        recCount = rsEMP("TOT_REC")
    Else
        recCount = 0
    End If
    rsEMP.Close
    Set rsEMP = Nothing
    
    getRecordCount_Add = recCount

End Function

