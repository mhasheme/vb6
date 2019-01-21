VERSION 5.00
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmUDOORS 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Mass Update Door Access"
   ClientHeight    =   8055
   ClientLeft      =   -210
   ClientTop       =   1350
   ClientWidth     =   10395
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
   ScaleHeight     =   8055
   ScaleWidth      =   10395
   WindowState     =   2  'Maximized
   Begin VB.Frame frmDetail 
      BorderStyle     =   0  'None
      Height          =   3315
      Left            =   180
      TabIndex        =   20
      Top             =   3510
      Width           =   9765
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door1"
         DataField       =   "door1"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   40
         Top             =   90
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door2"
         DataField       =   "door2"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   360
         TabIndex        =   39
         Top             =   390
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door3"
         DataField       =   "door3"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   360
         TabIndex        =   38
         Top             =   690
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door4"
         DataField       =   "door4"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   360
         TabIndex        =   37
         Top             =   990
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door5"
         DataField       =   "door5"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   360
         TabIndex        =   36
         Top             =   1290
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door6"
         DataField       =   "door6"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   360
         TabIndex        =   35
         Top             =   1590
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door7"
         DataField       =   "door7"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   360
         TabIndex        =   34
         Top             =   1890
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door8"
         DataField       =   "door8"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   7
         Left            =   360
         TabIndex        =   33
         Top             =   2190
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door9"
         DataField       =   "door9"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   8
         Left            =   360
         TabIndex        =   32
         Top             =   2490
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door10"
         DataField       =   "door10"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   9
         Left            =   360
         TabIndex        =   31
         Top             =   2790
         Width           =   2985
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door11"
         DataField       =   "door11"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   10
         Left            =   5580
         TabIndex        =   30
         Top             =   90
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door12"
         DataField       =   "door12"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   11
         Left            =   5580
         TabIndex        =   29
         Top             =   390
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door13"
         DataField       =   "door13"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   12
         Left            =   5580
         TabIndex        =   28
         Top             =   690
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door14"
         DataField       =   "door14"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   13
         Left            =   5580
         TabIndex        =   27
         Top             =   990
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door15"
         DataField       =   "door15"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   14
         Left            =   5580
         TabIndex        =   26
         Top             =   1290
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door16"
         DataField       =   "door16"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   15
         Left            =   5580
         TabIndex        =   25
         Top             =   1590
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door17"
         DataField       =   "door17"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   16
         Left            =   5580
         TabIndex        =   24
         Top             =   1890
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door18"
         DataField       =   "door18"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   17
         Left            =   5580
         TabIndex        =   23
         Top             =   2190
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door19"
         DataField       =   "door19"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   18
         Left            =   5580
         TabIndex        =   22
         Top             =   2490
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door20"
         DataField       =   "door20"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   300
         Index           =   19
         Left            =   5580
         TabIndex        =   21
         Top             =   2790
         Width           =   2955
      End
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1500
      TabIndex        =   3
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   1350
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1500
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   690
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1500
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   360
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   20
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1500
      TabIndex        =   2
      Tag             =   "00-Enter Union Code"
      Top             =   1020
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1500
      TabIndex        =   4
      Tag             =   "EDPT-Category"
      Top             =   1680
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1500
      TabIndex        =   5
      Tag             =   "00-Enter Region Code"
      Top             =   2010
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1500
      TabIndex        =   8
      Tag             =   "10-Enter Employee Number"
      Top             =   2340
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   6715
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   5880
      TabIndex        =   6
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   360
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   5880
      TabIndex        =   7
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   690
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin VB.Label lblSection 
      Alignment       =   1  'Right Justify
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
      Left            =   5160
      TabIndex        =   19
      Top             =   720
      Width           =   540
   End
   Begin VB.Label lblLocation 
      Alignment       =   1  'Right Justify
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
      Left            =   5100
      TabIndex        =   18
      Top             =   360
      Width           =   615
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
      Left            =   60
      TabIndex        =   17
      Top             =   1710
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
      Left            =   60
      TabIndex        =   16
      Top             =   2400
      Width           =   1290
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Left            =   60
      TabIndex        =   15
      Top             =   2040
      Width           =   510
   End
   Begin VB.Label lblSelCri 
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
      Left            =   0
      TabIndex        =   14
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblAttendance 
      BackStyle       =   0  'Transparent
      Caption         =   "Door Access"
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
      Left            =   0
      TabIndex        =   13
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblTitle 
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
      Index           =   3
      Left            =   60
      TabIndex        =   12
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
      Left            =   60
      TabIndex        =   11
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
      Left            =   60
      TabIndex        =   10
      Top             =   720
      Width           =   825
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   420
      Width           =   690
   End
End
Attribute VB_Name = "frmUDOORS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fglbSDate As Variant
Dim fglbESQLQ, fglbWSQLQ
Dim xASL As String


Private Function chkMUAttend()

Dim SQLQ As String, Msg$, dd&, Response%, x%
Dim DgDef As Variant, Title$, DCurPDate As Variant

chkMUAttend = False

On Error GoTo chkMUAttend_Err

For x% = 0 To 4
If Len(clpCode(x%).Text) > 0 And clpCode(x%).Caption = "Unassigned" Then
    MsgBox "If code entered it must be known"
    clpCode(x%).SetFocus
    Exit Function
End If
Next x%



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

If Not elpEEID.ListChecker Then
    Exit Function
End If

chkMUOK:
chkMUAttend = True

Exit Function

chkMUAttend_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkMUAttend", "HR Attendance", "edit/Add")
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

If Not gSec_Upd_Attendance Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

If Not chkMUAttend() Then Exit Sub

Title$ = "Mass Attendance Records Delete"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are You Sure You Want To Delete ALL records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

If Not modDelRecs() Then
  Screen.MousePointer = DEFAULT
  Exit Sub
End If

Screen.MousePointer = DEFAULT
MsgBox "Records Deleted Successfully"

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



Public Sub cmdModify_Click()
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String
Dim Title$, Msg$, DgDef As Variant, Response%

On Error GoTo Mod_Err
If Not gSec_Upd_Attendance Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

If Not chkMUAttend() Then Exit Sub

Title$ = "Mass Update Attendance"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to update all Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

If Not modUpdRecs() Then Exit Sub

Screen.MousePointer = DEFAULT
MsgBox "Records Updated Successfully"

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
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


If Not chkMUAttend() Then Exit Sub

Title$ = "Mass Records Door Access"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to add Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

If Not modInsRecs() Then Exit Sub
Screen.MousePointer = DEFAULT
MsgBox "Records Added Successfully"


Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "ATTEND", "Add")
Resume Next
End Sub








Private Sub clpDIV_Change()
Call SETUPLABEL
End Sub


Private Sub Form_Activate()
Call SET_UP_MODE

glbOnTop = "FRMUDOORS"

End Sub

Private Sub Form_Load()

glbOnTop = "FRMUDOORS"

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Screen.MousePointer = HOURGLASS
Call setRptCaption(Me)
If glbLinamar Then
    lblRegion.Visible = True
    clpCode(4).Visible = True
    clpCode(4).MaxLength = 8
End If



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
Set frmUDOORS = Nothing 'carmen apr 2000
End Sub










Private Function modDelRecs()
Dim SQLQ As String, countr As Integer
Dim rsDoor As New ADODB.Recordset

modDelRecs = False
On Error GoTo modDelRecs_Err

Screen.MousePointer = HOURGLASS

Call getWSQLQ
SQLQ = "SELECT USERID FROM LN_DOORS "
SQLQ = SQLQ & " WHERE EMP<>0 "
SQLQ = SQLQ & " AND USERID IN (SELECT RIGHT(ED_EMPNBR,3)+'-'+LEFT(ED_EMPNBR, LEN(ED_EMPNBR)-3) FROM HREMP WHERE " & fglbESQLQ & ")"
rsDoor.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
If rsDoor.EOF Or rsDoor.BOF Then
    modDelRecs = False
    MsgBox "Records Selection Not Found!"
Else
    SQLQ = "DELETE FROM LN_DOORS "
    SQLQ = SQLQ & " WHERE EMP<>0 "
    SQLQ = SQLQ & " AND USERID IN (SELECT RIGHT(ED_EMPNBR,3)+'-'+LEFT(ED_EMPNBR, LEN(ED_EMPNBR)-3) FROM HREMP WHERE " & fglbESQLQ & ")"
    gdbAdoIhr001.Execute SQLQ
    modDelRecs = True
End If
rsDoor.Close

Exit Function

modDelRecs_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modDelRecs", "DeleteAttend", "Delete")
modDelRecs = False
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    
    Resume Next
Else
    Unload Me
End If

End Function

Private Function modInsRecs()
Dim SQLQ As String
Dim rsEMP As New ADODB.Recordset, rsDup As New ADODB.Recordset, rsDoor As New ADODB.Recordset
Dim x, xDup
Dim Result
Dim Msg$
modInsRecs = False
On Error GoTo modInsRecs_Err


Screen.MousePointer = HOURGLASS

Call getWSQLQ

SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ
rsEMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If rsEMP.EOF And rsEMP.BOF Then
    modInsRecs = False
    MsgBox "Records for this selection do not exist!"
    Screen.MousePointer = DEFAULT
    Exit Function
End If
'rsEMP.Close

   
SQLQ = "SELECT USERID FROM LN_DOORS "
SQLQ = SQLQ & " WHERE EMP<>0 "
SQLQ = SQLQ & " AND USERID IN (SELECT RIGHT(ED_EMPNBR,3)+'-'+LEFT(ED_EMPNBR, LEN(ED_EMPNBR)-3) FROM HREMP WHERE " & fglbESQLQ & ")"
rsDup.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsDup.EOF Then
    Msg$ = rsDup.RecordCount & " duplicates found in Door Access. " & Chr(10) & Chr(10)
    Msg$ = Msg$ & "Click Yes to post non-duplicate records and update existing records." & Chr(10)
    Msg$ = Msg$ & "Click No to post non-duplicate records." & Chr(10)
    Result = MsgBox(Msg$, vbYesNo, "Duplicates Found")
    If Result = vbYes Then
        xDup = False
    Else
        xDup = True
    End If
End If
rsDup.Close

Do Until rsEMP.EOF
    rsDoor.Open "SELECT * FROM LN_DOORS WHERE EMP<>0 AND USERID='" & ShowEmpnbr(rsEMP("ED_EMPNBR")) & "'", gdbAdoIhr001, adOpenKeyset, adLockPessimistic
    If rsDoor.EOF Then
        rsDoor.AddNew
    Else
        If xDup Then GoTo NotUpdate
    End If
    rsDoor("COMPNO") = "001"
    rsDoor("USERID") = ShowEmpnbr(rsEMP("ED_EMPNBR"))
    rsDoor("EMP") = 1
    rsDoor("DIV") = clpDiv.Text
    For x = 1 To 20
        rsDoor("DOOR" & x) = chkDoors(x - 1)
    Next
    rsDoor.Update
NotUpdate:
    rsDoor.Close
    rsEMP.MoveNext
Loop
rsEMP.Close



modInsRecs = True

Exit Function

modInsRecs_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modInsRecs", "Attendance", "Insert")
modInsRecs = False
Resume Next




End Function

Private Function modUpdRecs()
Dim x
Dim SQLQ As String
Dim rsEMP As New ADODB.Recordset
Dim rsDoor As New ADODB.Recordset

modUpdRecs = False
On Error GoTo modUpdRecs2_Err

Screen.MousePointer = HOURGLASS

Call getWSQLQ

SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ
rsEMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If rsEMP.EOF And rsEMP.BOF Then
    modInsRecs = False
    MsgBox "Records for this selection do not exist!"
    Screen.MousePointer = DEFAULT
    Exit Function
End If


Do Until rsEMP.EOF
    rsDoor.Open "SELECT * FROM LN_DOORS WHERE EMP<>0 AND USERID='" & ShowEmpnbr(rsEMP("ED_EMPNBR")) & "'", gdbAdoIhr001, adOpenKeyset, adLockPessimistic
    If Not rsDoor.EOF Then
        rsDoor("COMPNO") = "001"
        rsDoor("USERID") = ShowEmpnbr(rsEMP("ED_EMPNBR"))
        rsDoor("EMP") = 1
        rsDoor("DIV") = clpDiv.Text
        For x = 1 To 20
            rsDoor("DOOR" & x) = chkDoors(x - 1)
        Next
        rsDoor.Update
    End If
    rsDoor.Close
    rsEMP.MoveNext
Loop
rsEMP.Close


modUpdRecs = True

Exit Function

modUpdRecs2_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modUpdRecs", "Attendance Reason", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function






Private Sub getWSQLQ()
fglbESQLQ = glbSeleDeptUn
If Len(clpDept.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DEPTNO = '" & clpDept.Text & "'"
If Len(clpDiv.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & clpDiv.Text & "' "
If Len(clpCode(2).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & clpCode(2).Text & "' "
If Len(clpCode(3).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMP = '" & clpCode(3).Text & "' "
If Len(clpCode(0).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & clpCode(0).Text & "' "
If Len(clpCode(1).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & clpCode(1).Text & "' "

If Len(clpPT.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_PT = '" & clpPT.Text & "' "
If glbLinamar Then
    If Len(clpCode(4).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND (ED_REGION = '" & clpDiv.Text & clpCode(4).Text & "' or  ED_REGION= 'ALL" & clpCode(4).Text & "')"
End If
If Len(elpEEID.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
End Sub



Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
TF = True
UpdateState = OPENING
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False


End Sub
Public Property Get RelateMode() As RelateModeEnum
RelateMode = MassChanges
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_DoorAccess
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


Private Sub SETUPLABEL()
Dim rsTD As New ADODB.Recordset
Dim SQLQ, x
rsTD.Open "SELECT * FROM LN_DOORS_NAME WHERE DIV='" & clpDiv & "'", gdbAdoIhr001, adOpenStatic
If rsTD.EOF Then
    For x = 1 To 20
        chkDoors(x - 1).Caption = "Door " & x
        chkDoors(x - 1).Value = 0
        chkDoors(x - 1).Enabled = False
    Next
Else
    For x = 1 To 20
        If IsNull(rsTD("DOORNAME" & x)) Then
            chkDoors(x - 1).Caption = "Door" & x
            chkDoors(x - 1).Value = 0
            chkDoors(x - 1).Enabled = False
        Else
            If Len(rsTD("DOORNAME" & x)) = 0 Then
                chkDoors(x - 1).Caption = "Door" & x
                chkDoors(x - 1).Value = 0
                chkDoors(x - 1).Enabled = False
            Else
                chkDoors(x - 1).Caption = rsTD("DOORNAME" & x)
                chkDoors(x - 1).Enabled = True
                If Not IsNull(rsTD("DOORCTRL" & x)) Then
                    If rsTD("DOORCTRL" & x) = "9560" Then
                        If x Mod 4 = 0 Then
                            chkDoors(x - 1).Enabled = False
                        End If
                    End If
                End If
            End If
        End If
    Next
End If
rsTD.Close
End Sub
