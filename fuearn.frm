VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmUOtherEarn 
   Appearance      =   0  'Flat
   Caption         =   "Other Earnings Mass Update"
   ClientHeight    =   9060
   ClientLeft      =   1305
   ClientTop       =   2625
   ClientWidth     =   10470
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9060
   ScaleWidth      =   10470
   Tag             =   "Other Earnings Mass Update"
   WindowState     =   2  'Maximized
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
      DataField       =   "COMMENTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Tag             =   "00-Enter Comments"
      Top             =   7680
      Width           =   8805
   End
   Begin INFOHR_Controls.DateLookup dlpTo 
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Tag             =   "40-Update to date"
      Top             =   6000
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpFrom 
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Tag             =   "40-Update from date"
      Top             =   5640
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   11
      Tag             =   "01-Earnings Type"
      Top             =   5280
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EARN"
   End
   Begin Threed.SSCheck chkCOEFlag 
      DataField       =   "select hrearn.cost_of_employment from hrearn where hrearn.fdate = hrparco.pc_fdate and hrearn.earn_type_table = hrtable.tb_name"
      Height          =   255
      Left            =   195
      TabIndex        =   15
      Tag             =   "Check for Cost of Employment"
      Top             =   6705
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Cost of Employment      "
      ForeColor       =   0
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   2010
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1842
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   2010
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2175
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   2010
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1509
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   2010
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1176
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   2010
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   843
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   2010
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   510
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   2010
      TabIndex        =   7
      Tag             =   "00-Enter Administered By Code"
      Top             =   2841
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   2010
      TabIndex        =   8
      Tag             =   "00-Enter Section Code"
      Top             =   3174
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   2010
      TabIndex        =   6
      Tag             =   "00-Enter Region Code"
      Top             =   2508
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   2010
      TabIndex        =   9
      Tag             =   "10-Enter Employee Number"
      Top             =   3510
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin MSMask.MaskEdBox medEntitleAmnt 
      DataSource      =   "Data2"
      Height          =   285
      Left            =   1875
      TabIndex        =   14
      Tag             =   "20-Amount of entitlement during the period"
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
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
      Format          =   "$##,##0.00;($##,##0.00)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Left            =   2010
      TabIndex        =   10
      Tag             =   "00-Enter Position Code"
      Top             =   3850
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   5
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpPosition 
      Height          =   285
      Left            =   1560
      TabIndex        =   16
      Tag             =   "01-Position code"
      Top             =   7035
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   6
      LookupType      =   5
   End
   Begin VB.Label lblPosTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Code"
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
      TabIndex        =   36
      Top             =   7080
      Width           =   1185
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Position Code"
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
      Left            =   330
      TabIndex        =   35
      Top             =   3895
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   34
      Top             =   7440
      Width           =   990
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Entitlement Amount"
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
      Left            =   240
      TabIndex        =   33
      Top             =   6405
      Width           =   1530
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
      Left            =   330
      TabIndex        =   32
      Top             =   3219
      Width           =   540
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
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
      Left            =   330
      TabIndex        =   31
      Top             =   2553
      Width           =   510
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
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
      Left            =   330
      TabIndex        =   30
      Top             =   2886
      Width           =   1125
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   330
      TabIndex        =   29
      Top             =   1221
      Width           =   615
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Left            =   330
      TabIndex        =   28
      Top             =   1887
      Width           =   450
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
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
      Left            =   330
      TabIndex        =   27
      Top             =   1554
      Width           =   420
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
      Left            =   330
      TabIndex        =   26
      Top             =   888
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
      Left            =   330
      TabIndex        =   25
      Top             =   555
      Width           =   555
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
      Left            =   330
      TabIndex        =   24
      Top             =   3555
      Width           =   1290
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
      Left            =   330
      TabIndex        =   23
      Top             =   2220
      Width           =   630
   End
   Begin VB.Label lblEarningsType 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   5295
      Width           =   1575
   End
   Begin VB.Label lblCostEmp 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Earnings Update"
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
      Left            =   60
      TabIndex        =   21
      Top             =   4860
      Width           =   3525
   End
   Begin VB.Label lblFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   5685
      Width           =   1050
   End
   Begin VB.Label lblTo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   6045
      Width           =   870
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
      Left            =   150
      TabIndex        =   18
      Top             =   210
      Width           =   1575
   End
End
Attribute VB_Name = "frmUOtherEarn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fglbAdd%    ' it is a global add request
Dim fglbDelete% ', fglbNoDept&
Dim fglbModify% ', fglb_FindDept
Dim xEmpNo
Dim XUpdCount, Actn

Private Function chkOtherEarnings()
Dim dd&
Dim Msg$, DgDef As Variant, Response%
Dim x%

chkOtherEarnings = False

On Error GoTo chkOtherEarnings_Err

If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    MsgBox lStr("If Division Entered - it must be known")
     clpDiv.SetFocus
    Exit Function
End If

If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
     clpDept.SetFocus
    Exit Function
End If

For x% = 0 To 6
    If Len(clpCode(x%).Text) > 0 And clpCode(x%).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(x%).SetFocus
        Exit Function
    End If
Next x%

If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    MsgBox lStr("Category code must be valid")
    clpPT.SetFocus
    Exit Function
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

If Not clpJob.ListChecker Then Exit Function    'Release 8.0

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Earnings Type is a required field"
    clpCode(1).SetFocus
    Exit Function
End If

If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Earnings Type must be valid"
    clpCode(1).SetFocus
    Exit Function
End If

If dlpFrom.Text = "" Then
    MsgBox "From Date needed"
    dlpFrom.SetFocus
    Exit Function
Else
    If Not IsDate(dlpFrom.Text) Then
        MsgBox "From Date must be valid"
        dlpFrom.SetFocus
        Exit Function
    End If
End If

If dlpTo.Text = "" Then
    MsgBox "To Date needed"
    dlpTo.SetFocus
    Exit Function
Else
    If Not IsDate(dlpTo.Text) Then
        MsgBox "To Date must be valid"
        dlpTo.SetFocus
        Exit Function
    End If
End If

dd& = DateDiff("d", CVDate(dlpFrom.Text), CVDate(dlpTo.Text))
If dd& < 0 Then
    MsgBox "To Date cannot precede From Date"
    dlpFrom.SetFocus
    Exit Function
End If

If Len(Trim(medEntitleAmnt)) <= 0 Then medEntitleAmnt = 0

chkOtherEarnings = True
Exit Function

chkOtherEarnings_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkOtherEarnings", "HREARN", "Update")
Resume Next

End Function

Private Sub chkCOEFlag_GotFocus()
  Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdNew_Click()
Dim Msg$, DgDef As Variant, Response%
Dim dd&
Dim Title$
Dim recCount As Integer

If Not gSec_Upd_Earnings Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Actn = "A"

If Not chkOtherEarnings() Then Exit Sub

Title$ = "Mass Records Other Earnings"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to Add Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Add
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Other Earning Record " Else Msg$ = Msg$ & " Other Earning Records "
    Msg$ = Msg$ & "will be Added. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, Title)     ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Employee record found to add the Other Earning record."
    Exit Sub
End If

Screen.MousePointer = HOURGLASS

If Not modInsRecs() Then
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

Screen.MousePointer = DEFAULT
If XUpdCount > 0 Then
    MsgBox Str(XUpdCount) & " Records Added Successfully."
Else
    MsgBox "No Records Added."
End If

End Sub

Public Sub cmdDelete_Click()
Dim a As Integer
'Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, rc%, DtTm As Variant, x%
Dim DgDef, Title$, Msg$, Response%
Dim recCount As Integer

If Not gSec_Upd_Earnings Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

fglbDelete% = True

If Not chkOtherEarnings() Then Exit Sub

Title$ = "Mass Other Earnings Delete"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are You Sure You Want To Delete ALL records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Delete
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Other Earning Record " Else Msg$ = Msg$ & " Other Earning Records "
    Msg$ = Msg$ & "will be Deleted. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, Title)     ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Other Earning record found to delete."
    Exit Sub
End If

If Not modDelRecs Then Exit Sub

Screen.MousePointer = DEFAULT

If XUpdCount > 0 Then
    MsgBox Str(XUpdCount) & " Records Deleted Successfully."
Else
    MsgBox "No Records Deleted."
End If

Screen.MousePointer = DEFAULT

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "Other Earnings", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Public Sub cmdModify_Click()
'Dim Msg$, DgDef As Variant, Response%
'
'If Not chekCOEFlag() Then
'  Exit Sub
'End If
'If Not gSec_Upd_Earnings Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If
'
'Call modCOEFlag
''dlpFrom = ""
''dlpto = ""
'
'End Sub

Public Sub cmdModify_Click()
Dim Msg$, DgDef As Variant, Response%
Dim dd&
Dim Title$
Dim recCount As Integer

If Not gSec_Upd_Earnings Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If
Actn = "M"

If Not chkOtherEarnings() Then Exit Sub

Title$ = "Mass Records Other Earnings"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to Update Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Modify
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Other Earning Record " Else Msg$ = Msg$ & " Other Earning Records "
    Msg$ = Msg$ & "will be Updated. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, Title)     ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Other Earning record found to update."
    Exit Sub
End If

Screen.MousePointer = HOURGLASS

If Not modUptRecs() Then Exit Sub

Screen.MousePointer = DEFAULT
If XUpdCount > 0 Then
    MsgBox Str(XUpdCount) & " Records Updated Successfully."
Else
    MsgBox "No Records Updated."
End If

End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMUOTHEREARN"
End Sub

Private Sub Form_Load()
glbOnTop = "FRMUOTHEREARN"

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Screen.MousePointer = HOURGLASS
Call setRptCaption(Me)
If glbCompSerial = "S/N - 2227W" Then clpCode(4).MaxLength = 6

Call INI_Controls(Me)
Screen.MousePointer = DEFAULT
End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Function modUptRecs()

Dim Msg$, DgDef As Variant, Response%, noRecs&
Dim rsTA As New ADODB.Recordset
Dim SQLQ, x%, strFld

modUptRecs = False

On Error GoTo cmdUpdErr

SQLQ = "SELECT * FROM HREARN WHERE EARN_TYPE = '" & clpCode(1).Text & "' AND "
SQLQ = SQLQ & "FDATE = " & Date_SQL(dlpFrom.Text) & " And "
SQLQ = SQLQ & "TDATE = " & Date_SQL(dlpTo.Text)
SQLQ = SQLQ & " AND EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ("M") & ")"
rsTA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
XUpdCount = rsTA.RecordCount

'rsTA.Close
'Hemu
Do While Not rsTA.EOF
    rsTA("FDATE") = dlpFrom.Text
    rsTA("TDATE") = dlpTo.Text
    rsTA("ACT_DOLLAR") = medEntitleAmnt
    rsTA("COST_OF_EMPLOYMENT") = IIf(chkCOEFlag, 1, 0)
    rsTA("COMMENTS") = memComments.Text
    rsTA("LDATE") = Date
    rsTA("LTIME") = Time$
    rsTA("LUSER") = glbUserID
    
    'Ticket #24410 - City of Sarnia - Position Code added
    rsTA("OE_JOB") = clpPosition.Text
    
    rsTA.Update
    
    xEmpNo = rsTA("EMPNBR")
    Call AUDITOEAR("M")
    
    rsTA.MoveNext
Loop
rsTA.Close
'Hemu

'SQLQ = " UPDATE HREARN SET"
'SQLQ = SQLQ & " EARN_TYPE='" & clpCode(1).Text & "'"
'SQLQ = SQLQ & ",FDATE = " & Date_SQL(dlpFrom.Text)
'SQLQ = SQLQ & ",TDATE = " & Date_SQL(dlpTo.Text)
'SQLQ = SQLQ & ",ACT_DOLLAR=" & medEntitleAmnt
'SQLQ = SQLQ & ",COST_OF_EMPLOYMENT=" & IIf(chkCOEFlag, 1, 0)
'SQLQ = SQLQ & ",COMMENTS='" & memComments.Text & "'"
'SQLQ = SQLQ & ",LDATE=" & Date_SQL(Date)
'SQLQ = SQLQ & ",LTIME='" & Time$ & "'"
'SQLQ = SQLQ & ",LUSER=" & glbUserID
'SQLQ = SQLQ & " WHERE EARN_TYPE = '" & clpCode(1).Text & "' "
'SQLQ = SQLQ & " AND EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ("M") & ")"
'gdbAdoIhr001.Execute SQLQ

modUptRecs = True

Exit Function
cmdUpdErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Mass change", "HREARN", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If

End Function

Private Function modInsRecs()

Dim Msg$, DgDef As Variant, Response%, noRecs&
Dim rsTA As New ADODB.Recordset
Dim rsEarn As New ADODB.Recordset
Dim SQLQ, x%, strFld

modInsRecs = False
On Error GoTo cmdInsErr

rsTA.Open "SELECT ED_EMPNBR,ED_COMPNO FROM HREMP " & WSQLQ("A"), gdbAdoIhr001, adOpenStatic
XUpdCount = rsTA.RecordCount

'Hemu
'rsTA.Close
rsEarn.Open "HREARN", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
Do While Not rsTA.EOF
    rsEarn.AddNew
    rsEarn("COMPNO") = rsTA("ED_COMPNO")
    rsEarn("EMPNBR") = rsTA("ED_EMPNBR")
    rsEarn("EARN_TYPE") = clpCode(1).Text
    rsEarn("FDATE") = dlpFrom.Text
    rsEarn("TDATE") = dlpTo.Text
    rsEarn("ACT_DOLLAR") = medEntitleAmnt
    rsEarn("COST_OF_EMPLOYMENT") = IIf(chkCOEFlag, 1, 0)
    rsEarn("COMMENTS") = memComments.Text
    rsEarn("LDATE") = Date
    rsEarn("LTIME") = Time$
    rsEarn("LUSER") = glbUserID
    
    'Ticket #24410 - City of Sarnia - Position Code added
    rsEarn("OE_JOB") = clpPosition.Text
    
    rsEarn.Update
        
    xEmpNo = rsTA("ED_EMPNBR")
    
    Call AUDITOEAR("A")
    
    rsTA.MoveNext
Loop
rsEarn.Close
rsTA.Close

'SQLQ = "INSERT INTO HREARN ("
'SQLQ = SQLQ & " COMPNO"
'SQLQ = SQLQ & ",EMPNBR"
'SQLQ = SQLQ & ",EARN_TYPE"
'SQLQ = SQLQ & ",FDATE"
'SQLQ = SQLQ & ",TDATE"
'SQLQ = SQLQ & ",ACT_DOLLAR"
'SQLQ = SQLQ & ",COST_OF_EMPLOYMENT"
'SQLQ = SQLQ & ",COMMENTS"
'SQLQ = SQLQ & ",LDATE"
'SQLQ = SQLQ & ",LTIME"
'SQLQ = SQLQ & ",LUSER "
'SQLQ = SQLQ & ")"
'SQLQ = SQLQ & " SELECT "
'SQLQ = SQLQ & " ED_COMPNO"
'SQLQ = SQLQ & ",ED_EMPNBR"
'SQLQ = SQLQ & ",'" & clpCode(1).Text & "'"
'SQLQ = SQLQ & "," & Date_SQL(dlpFrom.Text)
'SQLQ = SQLQ & "," & Date_SQL(dlpTo.Text)
'SQLQ = SQLQ & "," & medEntitleAmnt
'SQLQ = SQLQ & "," & IIf(chkCOEFlag, 1, 0)
'SQLQ = SQLQ & ",'" & memComments.Text & "'"
'SQLQ = SQLQ & "," & Date_SQL(Date)
'SQLQ = SQLQ & ",'" & Time$ & "'"
'SQLQ = SQLQ & "," & glbUserID
'SQLQ = SQLQ & " FROM HREMP "
'SQLQ = SQLQ & WSQLQ("A")
'gdbAdoIhr001.Execute SQLQ
'Hemu

modInsRecs = True

Exit Function
cmdInsErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
If glbErrNum& = -2147467259 Then
    MsgBox "The changes were not successful because it would create duplicate values."
    Exit Function
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Mass Add", "HREARN", "Insert")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        RollBack
        Resume Next
    Else
        Unload Me
    End If
End If
End Function

Private Function WSQLQ(FSTR) As String
Dim countr As Integer

WSQLQ = WSQLQ & " WHERE " & glbSeleDeptUn

If Len(clpDept.Text) > 0 Then WSQLQ = WSQLQ & " AND ED_DEPTNO = '" & clpDept.Text & "'"
If Len(clpDiv.Text) > 0 Then WSQLQ = WSQLQ & " AND ED_DIV = '" & clpDiv.Text & "' "
If Len(clpCode(0).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_LOC = '" & clpCode(0).Text & "' "
If Len(clpCode(2).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_ORG = '" & clpCode(2).Text & "' "
If Len(clpCode(3).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_EMP = '" & clpCode(3).Text & "' "
If Len(clpCode(4).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_REGION = '" & clpCode(4).Text & "' "
If Len(clpCode(5).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_ADMINBY = '" & clpCode(5).Text & "' "
If Len(clpCode(6).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_SECTION = '" & clpCode(6).Text & "' "
If Len(clpPT) > 0 Then WSQLQ = WSQLQ & " AND ED_PT = '" & clpPT.Text & "' "
If Len(elpEEID.Text) > 0 Then WSQLQ = WSQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "

'If FSTR <> "A" Then
'    WSQLQ = WSQLQ & " AND DE_TYPE = '" & clpCode(1) & "' "
'End If

'Ticket #22682 - Release 8.0 - Add Position Code Selection Criteria
If clpJob.Text <> "" Then WSQLQ = WSQLQ & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_JOB IN ('" & Replace(clpJob.Text, ",", "','") & "') )"

End Function

Private Function AUDITOEAR(ACTX)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsTC As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String

On Error GoTo AUDIT_ERR

AUDITOEAR = False

rsTB.Open "SELECT ED_PT,ED_DIV,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR = " & xEmpNo, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    'xPT = rsTB("ED_PT")
    'xDiv = rsTB("ED_DIV")
    If IsNull(rsTB("ED_PT")) Then
        xPT = ""
    Else
        xPT = rsTB("ED_PT")
    End If
    If IsNull(rsTB("ED_DIV")) Then
        xDiv = ""
    Else
        xDiv = rsTB("ED_DIV")
    End If
Else
    xPT = ""
    xDiv = ""
End If
'strFields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, "
strFields = strFields & "AU_TREAS_TABL, AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_SERVICE, AU_EARN, "
strFields = strFields & "AU_COEFLAG, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID, AU_ADOLLAR, AU_JOB "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False

If ACTX = "D" Or ACTX = "A" Then GoTo MODUPD
If ACTX = "M" Then
    rsTC.Open "SELECT * FROM HREARN WHERE EMPNBR = " & xEmpNo & "", gdbAdoIhr001, adOpenStatic
    If Not rsTC.EOF Then
        If rsTC("COST_OF_EMPLOYMENT") <> IIf(chkCOEFlag, 1, 0) Then
            rsTC.Close
            GoTo MODUPD
        End If
    End If
    rsTC.Close
End If
GoTo MODNOUPD

MODUPD:
rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM"
rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP"
rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL"
rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv

If ACTX = "D" Then
    rsTA("AU_EARN") = clpCode(1).Text
    If clpCode(1).Text = "SERV" Then rsTA("AU_SERVICE") = "Y"
Else
    rsTA("AU_EARN") = clpCode(1).Text
    If ACTX = "A" Then If clpCode(1).Text = "SERV" Then rsTA("AU_SERVICE") = "Y"
    
    If medEntitleAmnt <> "" Then                       '16Aug99 js
        rsTA("AU_ADOLLAR") = medEntitleAmnt  '
    Else                                         '
        rsTA("AU_ADOLLAR") = 0                   '
    End If
    
    If chkCOEFlag = True Then
        rsTA("AU_COEFLAG") = "Y"
    Else
        rsTA("AU_COEFLAG") = "N"
    End If
End If

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = xEmpNo
rsTA("AU_LDATE") = Date
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
'If glbSoroc Or glbSyndesis Then
    If Not IsNull(rsTB("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsTB("ED_PAYROLL_ID")
'End If

'Ticket #24410 - City of Sarnia - Position Code added in Earnings table so updating the Audit table as well
rsTA("AU_JOB") = clpPosition.Text

rsTA.Update

MODNOUPD:
AUDITOEAR = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '26July99 js

End Function

Private Sub modCOEFlag()

Dim Msg$, DgDef As Variant, Response%, noRecs&
Dim dyn_HREARN As New ADODB.Recordset
Dim SQLQ, x%, strFld

On Error GoTo cmdUpdErr


Screen.MousePointer = HOURGLASS

SQLQ = "SELECT * FROM HREARN where EARN_TYPE = '" & clpCode(1).Text & "'"
SQLQ = SQLQ & " and FDATE = " & Date_SQL(dlpFrom.Text)
SQLQ = SQLQ & " and TDATE = " & Date_SQL(dlpTo.Text)

SQLQ = SQLQ & " AND  EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn & ")"

dyn_HREARN.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If Not dyn_HREARN.EOF And Not dyn_HREARN.BOF Then
   
    noRecs& = dyn_HREARN.RecordCount
    Screen.MousePointer = DEFAULT
    Msg$ = "Are you sure you wish to update the "
    Msg$ = Msg$ & Chr(10) & "Cost of Employment to"
    If chkCOEFlag Then
        Msg$ = Msg$ & " Yes"
    Else
        Msg$ = Msg$ & " No"
    End If
    

    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
    Response% = MsgBox(Msg$, DgDef, "Warning!")
    If Response% = IDNO Then
        Exit Sub
    End If
    Screen.MousePointer = HOURGLASS
    BeginTrans
    While Not dyn_HREARN.EOF
        'dyn_HREARN.Edit
        xEmpNo = dyn_HREARN("EMPNBR")
        Call AUDITOEAR("M")
        dyn_HREARN("COST_OF_EMPLOYMENT") = IIf(chkCOEFlag, 1, 0)
        dyn_HREARN("LDATE") = Now
        dyn_HREARN("LTIME") = Time$
        dyn_HREARN("LUSER") = glbLEE_ID
        dyn_HREARN.Update
        dyn_HREARN.MoveNext
    Wend
    CommitTrans
    MsgBox "Update completed"
Else
    MsgBox "Employees for this selection do not exist!"
End If
Screen.MousePointer = DEFAULT
Exit Sub
cmdUpdErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Error", "HREARN", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If

End Sub

Private Function modDelRecs()
Dim BD As Integer
Dim SQLQ As String, countr As Integer
Dim Dat1 As Variant, Dat2 As Variant
Dim iOneWhere As Integer, NxtSQL As String, strReas$
Dim dyn_HREARN As New ADODB.Recordset

On Error GoTo cmdDel_Err

modDelRecs = False

Screen.MousePointer = HOURGLASS

SQLQ = "SELECT HREARN.* FROM HREARN "
SQLQ = SQLQ & "WHERE EARN_TYPE = '" & clpCode(1).Text & "' AND "
SQLQ = SQLQ & "FDATE = " & Date_SQL(dlpFrom.Text) & " And "
SQLQ = SQLQ & "TDATE = " & Date_SQL(dlpTo.Text)
SQLQ = SQLQ & " AND EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ("M") & ")"

dyn_HREARN.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic ', adOpenStatic
XUpdCount = dyn_HREARN.RecordCount

'If dyn_HREARN.BOF And dyn_HREARN.EOF Then
'  modDelRecs = False
'Else
'  modDelRecs = True
'End If

Screen.MousePointer = HOURGLASS
While Not dyn_HREARN.EOF
    'dyn_HREARN.Edit
    xEmpNo = dyn_HREARN("EMPNBR")
    Call AUDITOEAR("D")
    dyn_HREARN.Delete
    dyn_HREARN.MoveNext
Wend
dyn_HREARN.Close

modDelRecs = True

Screen.MousePointer = DEFAULT

'SQLQ = "Delete FROM HREARN "
'SQLQ = SQLQ & "WHERE EARN_TYPE = '" & clpCode(1).Text & "' AND "
'SQLQ = SQLQ & "FDATE = " & Date_SQL(dlpFrom.Text) & " And "
'SQLQ = SQLQ & "TDATE = " & Date_SQL(dlpTo.Text)
'gdbAdoIhr001.Execute SQLQ
''modDelRecs = True

Exit Function
cmdDel_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HREARN", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frmUOtherEarn = Nothing 'carmen apr 2000
End Sub

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
'UpdateRight = gSec_Upd_Earnings
UpdateRight = GetMassUpdateSecurities("Other_Earnings_MassUpdate", glbUserID)
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
Printable = False
End Property

Private Function getRecordCount_Add()
    Dim SQLQ As String
    Dim rsEMP As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Add = 0
    recCount = 0

    rsEMP.Open "SELECT COUNT(ED_EMPNBR) AS TOT_REC FROM HREMP " & WSQLQ("A"), gdbAdoIhr001, adOpenStatic
    If Not rsEMP.EOF Then
        recCount = rsEMP("TOT_REC")
    Else
        recCount = 0
    End If
    rsEMP.Close
    Set rsEMP = Nothing
    
    getRecordCount_Add = recCount

End Function

Private Function getRecordCount_Modify()
    Dim SQLQ As String
    Dim rsEMP As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Modify = 0
    recCount = 0

    SQLQ = "SELECT COUNT(EMPNBR) AS TOT_REC FROM HREARN WHERE EARN_TYPE = '" & clpCode(1).Text & "' AND "
    SQLQ = SQLQ & "FDATE = " & Date_SQL(dlpFrom.Text) & " AND "
    SQLQ = SQLQ & "TDATE = " & Date_SQL(dlpTo.Text)
    SQLQ = SQLQ & " AND EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ("M") & ")"
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEMP.EOF Then
        recCount = rsEMP("TOT_REC")
    Else
        recCount = 0
    End If
    rsEMP.Close
    Set rsEMP = Nothing
    
    getRecordCount_Modify = recCount

End Function

Private Function getRecordCount_Delete()
    Dim SQLQ As String
    Dim rsEMP As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Delete = 0
    recCount = 0

    SQLQ = "SELECT COUNT(EMPNBR) AS TOT_REC FROM HREARN "
    SQLQ = SQLQ & "WHERE EARN_TYPE = '" & clpCode(1).Text & "' AND "
    SQLQ = SQLQ & "FDATE = " & Date_SQL(dlpFrom.Text) & " AND "
    SQLQ = SQLQ & "TDATE = " & Date_SQL(dlpTo.Text)
    SQLQ = SQLQ & " AND EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ("M") & ")"
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEMP.EOF Then
        recCount = rsEMP("TOT_REC")
    Else
        recCount = 0
    End If
    rsEMP.Close
    Set rsEMP = Nothing
    
    getRecordCount_Delete = recCount

End Function

