VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmDolEntit 
   Appearance      =   0  'Flat
   Caption         =   "Dollar Entitlements Mass Update"
   ClientHeight    =   7710
   ClientLeft      =   1560
   ClientTop       =   2325
   ClientWidth     =   10410
   DrawMode        =   1  'Blackness
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7710
   ScaleWidth      =   10410
   Tag             =   "Dollar Entitlements Mass Update"
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.DateLookup dlpDate 
      Height          =   285
      Index           =   1
      Left            =   4320
      TabIndex        =   12
      Tag             =   "40-Date upto and including this date forward"
      Top             =   4560
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   11
      Tag             =   "40-Date from and including this date forward"
      Top             =   4560
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   10
      Tag             =   "11-Enter type of entitlements"
      Top             =   4200
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOL"
   End
   Begin MSMask.MaskEdBox medEntitleAmnt 
      DataSource      =   "Data2"
      Height          =   285
      Left            =   2220
      TabIndex        =   13
      Tag             =   "20-Amount of entitlement during the period"
      Top             =   4920
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
   Begin Threed.SSCheck chkCOEFlag 
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Tag             =   "Check for Cost of Employment"
      Top             =   5310
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Cost of Employment            "
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
      Left            =   1920
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1680
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2010
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1350
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1020
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   690
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
      Left            =   1920
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   360
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
      Left            =   1920
      TabIndex        =   7
      Tag             =   "00-Enter Administered By Code"
      Top             =   2700
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
      Left            =   1920
      TabIndex        =   8
      Tag             =   "00-Enter Section Code"
      Top             =   3030
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1920
      TabIndex        =   6
      Tag             =   "00-Enter Region Code"
      Top             =   2370
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Tag             =   "10-Enter Employee Number"
      Top             =   3360
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   6835
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5400
      Top             =   6960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      TabIndex        =   30
      Top             =   2100
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
      TabIndex        =   29
      Top             =   3450
      Width           =   1290
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
      TabIndex        =   28
      Top             =   4950
      Width           =   1650
   End
   Begin VB.Label lblSelectCrit 
      BackStyle       =   0  'Transparent
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
      Left            =   30
      TabIndex        =   27
      Top             =   210
      Width           =   1695
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
      TabIndex        =   26
      Top             =   450
      Width           =   555
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
      TabIndex        =   25
      Top             =   780
      Width           =   825
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
      Left            =   240
      TabIndex        =   24
      Top             =   1440
      Width           =   420
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
      Left            =   240
      TabIndex        =   23
      Top             =   1770
      Width           =   450
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
      Left            =   240
      TabIndex        =   22
      Top             =   1110
      Width           =   615
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
      Left            =   240
      TabIndex        =   21
      Top             =   2760
      Width           =   1125
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
      Left            =   240
      TabIndex        =   20
      Top             =   2430
      Width           =   510
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
      TabIndex        =   19
      Top             =   3090
      Width           =   540
   End
   Begin VB.Label lblTo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
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
      Left            =   3570
      TabIndex        =   18
      Top             =   4590
      Width           =   840
   End
   Begin VB.Label lblFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
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
      Top             =   4590
      Width           =   990
   End
   Begin VB.Label lblType 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   4230
      Width           =   435
   End
   Begin VB.Label lblSelCri 
      AutoSize        =   -1  'True
      Caption         =   "Dollar Entitlements Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   15
      Top             =   3840
      Width           =   2280
   End
End
Attribute VB_Name = "frmDolEntit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim EECri As String, OneSet%, X%
Dim XUpdCount, Actn

Private Sub chkCOEFlag_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Function chkFUDEntit()

Dim SQLQ As String, Msg$, dd&, Response%, X%
Dim DgDef As Variant, Title$, DCurPDate As Variant

chkFUDEntit = False

On Error GoTo chkFUDEntit_Err

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Entitlement Type is a required field"
    clpCode(1).SetFocus
    Exit Function
End If

If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Entitlement Type must be valid"
    clpCode(1).SetFocus
    Exit Function
End If
If Actn = "A" Then
    If Len(dlpDate(0).Text) < 1 Then
        MsgBox "Date From must be entered"
        dlpDate(0).SetFocus
        Exit Function
    End If
    If Len(dlpDate(1).Text) < 1 Then
        MsgBox "Date To must be entered"
        dlpDate(1).SetFocus
        Exit Function
    End If
End If
If Len(dlpDate(0).Text) > 0 Then
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "Date From must be valid"
        dlpDate(0).SetFocus
        Exit Function
    End If
End If
If Len(dlpDate(1).Text) > 0 Then
    If Not IsDate(dlpDate(1).Text) Then
        MsgBox "Date To must be valid"
        dlpDate(1).SetFocus
        Exit Function
    End If
End If
If IsDate(dlpDate(1).Text) And IsDate(dlpDate(0).Text) Then
    dd& = DateDiff("d", CVDate(dlpDate(0).Text), CVDate(dlpDate(1).Text))
    If dd& < 1 Then
        MsgBox "From date must be earlier than To Date"
        dlpDate(0).SetFocus
        Exit Function
    End If
End If
If Len(clpDIV.Text) > 0 And clpDIV.Caption = "Unassigned" Then
    MsgBox lStr("If Division Entered - it must be known")
     clpDIV.SetFocus
    Exit Function
End If

If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
     clpDept.SetFocus
    Exit Function
End If

For X% = 0 To 6
    If Len(clpCode(X%).Text) > 0 And clpCode(X%).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(X%).SetFocus
        Exit Function
    End If
Next X%

If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    MsgBox lStr("Category code must be valid")
     clpPT.SetFocus
    Exit Function
End If
If Len(Trim(medEntitleAmnt)) <= 0 Then medEntitleAmnt = 0
If Not elpEEID.ListChecker Then
    Exit Function
End If

chkFUDEntit = True

Exit Function

chkFUDEntit_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkFUDEntit", "HRDOLENT", "Delete")
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
Dim SQLQ As String, rc%, DtTm As Variant, X%
Dim DgDef, Title$, Msg$, Response%
Dim recCount As Integer

If Not gSec_Upd_Other_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Actn = "D"
If Not chkFUDEntit() Then Exit Sub       'laura jan 09, 1998

Title$ = "Mass Dollar Entitlements Delete"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are You Sure You Want To Delete ALL records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Delete
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Dollar Entitlement Record " Else Msg$ = Msg$ & " Dollar Entitlement Records "
    Msg$ = Msg$ & "will be Deleted. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Dollar Entitlement record found to delete."
    Exit Sub
End If

Screen.MousePointer = HOURGLASS

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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "Dollar Entitlements", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub



Public Sub cmdNew_Click()
Dim Msg$, DgDef As Variant, Response%
Dim dd&
Dim Title$
Dim recCount As Integer

If Not gSec_Upd_Other_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Actn = "A"
If Not chkFUDEntit() Then Exit Sub

Title$ = "Mass Records Dollar Entitlements "
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to Add Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Add
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Dollar Entitlement Record " Else Msg$ = Msg$ & " Dollar Entitlement Records "
    Msg$ = Msg$ & "will be Added. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Employee record found to add the Dollar Entitlement record."
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

Public Sub cmdModify_Click()
Dim Msg$, DgDef As Variant, Response%
Dim dd&
Dim Title$
Dim recCount As Integer

If Not gSec_Upd_Other_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Actn = "M"
If Not chkFUDEntit() Then Exit Sub

Title$ = "Mass Records Dollar Entitlements "
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to Update Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Modify
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Dollar Entitlement Record " Else Msg$ = Msg$ & " Dollar Entitlement Records "
    Msg$ = Msg$ & "will be Updated. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Dollar Entitlement record found to update."
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
glbOnTop = "FRMDOLENTIT"
End Sub

Private Sub Form_Load()
glbOnTop = "FRMDOLENTIT"
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
Dim SQLQ, X%, strFld
modUptRecs = False
On Error GoTo cmdUpdErr

SQLQ = "SELECT DE_EMPNBR FROM HRDOLENT WHERE DE_TYPE = '" & clpCode(1).Text & "' "
SQLQ = SQLQ & " AND DE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ("M") & ")"
rsTA.Open SQLQ, gdbAdoIhr001, adOpenStatic
XUpdCount = rsTA.RecordCount
rsTA.Close

SQLQ = " UPDATE HRDOLENT SET"
SQLQ = SQLQ & " DE_TYPE='" & clpCode(1).Text & "'"
SQLQ = SQLQ & ",DE_FDATE = " & Date_SQL(dlpDate(0).Text)
SQLQ = SQLQ & ",DE_TDATE = " & Date_SQL(dlpDate(1).Text)
SQLQ = SQLQ & ",DE_ENTITLE=" & medEntitleAmnt
SQLQ = SQLQ & ",DE_COST_OF_EMPLOYMENT=" & IIf(chkCOEFlag, 1, 0)
SQLQ = SQLQ & ",DE_LDATE=" & Date_SQL(Date)
SQLQ = SQLQ & ",DE_LTIME='" & Time$ & "'"
SQLQ = SQLQ & ",DE_LUSER=" & glbLEE_ID
SQLQ = SQLQ & " WHERE DE_TYPE = '" & clpCode(1).Text & "' "
SQLQ = SQLQ & " AND DE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ("M") & ")"
gdbAdoIhr001.Execute SQLQ

modUptRecs = True

Exit Function
cmdUpdErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "COE Error", "HRDOLENT", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If

End Function

Private Function modDelRecs()
Dim BD As Integer
Dim SQLQ As String, countr As Integer
Dim Dat1 As Variant, Dat2 As Variant
Dim iOneWhere As Integer, NxtSQL As String, strReas$
Dim rsTA As New ADODB.Recordset

modDelRecs = False
On Error GoTo cmdDel_Err

SQLQ = "SELECT DE_EMPNBR FROM HRDOLENT WHERE DE_TYPE = '" & clpCode(1).Text & "' "
If IsDate(dlpDate(0).Text) Then SQLQ = SQLQ & " AND DE_FDATE = " & Date_SQL(dlpDate(0).Text)
If IsDate(dlpDate(1).Text) Then SQLQ = SQLQ & " AND DE_TDATE = " & Date_SQL(dlpDate(1).Text)
SQLQ = SQLQ & " AND DE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ("M") & ")"
rsTA.Open SQLQ, gdbAdoIhr001, adOpenStatic
XUpdCount = rsTA.RecordCount
rsTA.Close

SQLQ = " DELETE FROM HRDOLENT "
SQLQ = SQLQ & " WHERE DE_TYPE = '" & clpCode(1).Text & "' "
If IsDate(dlpDate(0).Text) Then SQLQ = SQLQ & " AND DE_FDATE = " & Date_SQL(dlpDate(0).Text)
If IsDate(dlpDate(1).Text) Then SQLQ = SQLQ & " AND DE_TDATE = " & Date_SQL(dlpDate(1).Text)
SQLQ = SQLQ & " AND DE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ("M") & ")"
gdbAdoIhr001.Execute SQLQ
modDelRecs = True
Exit Function

cmdDel_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRDOLENT", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If
End Function

Private Sub Form_Unload(Cancel As Integer)
Set frmDolEntit = Nothing
End Sub


Private Function modInsRecs()

Dim Msg$, DgDef As Variant, Response%, noRecs&
Dim rsTA As New ADODB.Recordset
Dim SQLQ, X%, strFld

modInsRecs = False
On Error GoTo cmdInsErr



rsTA.Open "SELECT ED_EMPNBR FROM HREMP " & WSQLQ("A"), gdbAdoIhr001, adOpenStatic
XUpdCount = rsTA.RecordCount
rsTA.Close

SQLQ = "INSERT INTO HRDOLENT ("
SQLQ = SQLQ & " DE_COMPNO"
SQLQ = SQLQ & ",DE_EMPNBR"
SQLQ = SQLQ & ",DE_TYPE"
SQLQ = SQLQ & ",DE_FDATE"
SQLQ = SQLQ & ",DE_TDATE"
SQLQ = SQLQ & ",DE_ENTITLE"
SQLQ = SQLQ & ",DE_COST_OF_EMPLOYMENT"
SQLQ = SQLQ & ",DE_LDATE"
SQLQ = SQLQ & ",DE_LTIME"
SQLQ = SQLQ & ",DE_LUSER "
SQLQ = SQLQ & ")"
SQLQ = SQLQ & " SELECT "
SQLQ = SQLQ & " ED_COMPNO"
SQLQ = SQLQ & ",ED_EMPNBR"
SQLQ = SQLQ & ",'" & clpCode(1).Text & "'"
SQLQ = SQLQ & "," & Date_SQL(dlpDate(0).Text)
SQLQ = SQLQ & "," & Date_SQL(dlpDate(1).Text)
SQLQ = SQLQ & "," & medEntitleAmnt
SQLQ = SQLQ & "," & IIf(chkCOEFlag, 1, 0)
SQLQ = SQLQ & "," & Date_SQL(Date)
SQLQ = SQLQ & ",'" & Time$ & "'"
SQLQ = SQLQ & "," & glbLEE_ID
SQLQ = SQLQ & " FROM HREMP "
SQLQ = SQLQ & WSQLQ("A")
gdbAdoIhr001.Execute SQLQ

modInsRecs = True

Exit Function
cmdInsErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
If glbErrNum& = -2147467259 Then
    MsgBox "The changes were not successful because it would create duplicate values"
    Exit Function
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "COE Error", "HRDOLENT", "Insert")
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
If Len(clpDIV.Text) > 0 Then WSQLQ = WSQLQ & " AND ED_DIV = '" & clpDIV.Text & "' "
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
UpdateRight = gSec_Upd_Other_Entitlements
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

    SQLQ = "SELECT COUNT(DE_EMPNBR) AS TOT_REC FROM HRDOLENT WHERE DE_TYPE = '" & clpCode(1).Text & "' "
    SQLQ = SQLQ & " AND DE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ("M") & ")"
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

    SQLQ = "SELECT COUNT(DE_EMPNBR) AS TOT_REC FROM HRDOLENT WHERE DE_TYPE = '" & clpCode(1).Text & "' "
    If IsDate(dlpDate(0).Text) Then SQLQ = SQLQ & " AND DE_FDATE = " & Date_SQL(dlpDate(0).Text)
    If IsDate(dlpDate(1).Text) Then SQLQ = SQLQ & " AND DE_TDATE = " & Date_SQL(dlpDate(1).Text)
    SQLQ = SQLQ & " AND DE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ("M") & ")"
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

