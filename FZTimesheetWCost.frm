VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRTimesheetWCost 
   Caption         =   "Timesheet With Equipment Cost"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1995
      TabIndex        =   7
      Top             =   2684
      Width           =   1335
   End
   Begin VB.TextBox txtWeek 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1980
      TabIndex        =   8
      Top             =   3016
      Width           =   1335
   End
   Begin INFOHR_Controls.CodeLookup clpAtt 
      Height          =   285
      Left            =   1680
      TabIndex        =   15
      Top             =   4676
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowDescription =   0   'False
      TABLName        =   "ADRE"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1995
      MaxLength       =   4
      TabIndex        =   16
      Tag             =   "00-Employee Position Shift"
      Top             =   5020
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1688
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDEM"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2020
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDPT"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1356
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDOR"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1024
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   692
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   360
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   1
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1680
      TabIndex        =   13
      Tag             =   "00-Enter Administered By Code"
      Top             =   4012
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   14
      Tag             =   "00-Enter Section Code"
      Top             =   4344
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   12
      Tag             =   "00-Enter Region Code"
      Top             =   3680
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2352
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Tag             =   "00-Enter Position Code"
      Top             =   3348
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   6
      LookupType      =   5
   End
   Begin INFOHR_Controls.CodeLookup clpPosGroup 
      Height          =   285
      Left            =   6540
      TabIndex        =   18
      Tag             =   "00-Position Group  Code"
      Top             =   2684
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGC"
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   480
      Top             =   6000
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
   Begin Threed.SSCheck chkShowEmp 
      Height          =   225
      Left            =   1995
      TabIndex        =   17
      Tag             =   "If X-Show All Employees"
      Top             =   5610
      Width           =   2835
      _Version        =   65536
      _ExtentX        =   5001
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   " Show All Employees"
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
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   7830
      TabIndex        =   10
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3016
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   6210
      TabIndex        =   9
      Tag             =   "40-Date from and including this date forward"
      Top             =   3016
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
      Enabled         =   0   'False
   End
   Begin VB.Label lblFromTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From / To Date"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3840
      TabIndex        =   36
      Top             =   3061
      Width           =   1095
   End
   Begin VB.Label lblYear 
      Caption         =   "Year"
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   2729
      Width           =   1395
   End
   Begin VB.Label lblWeek 
      Caption         =   "Pay Period #"
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   3061
      Width           =   1395
   End
   Begin VB.Image imgIcon 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1680
      Picture         =   "FZTimesheetWCost.frx":0000
      Top             =   3030
      Width           =   240
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Group"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5400
      TabIndex        =   33
      Top             =   2729
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   32
      Top             =   3393
      Width           =   975
   End
   Begin VB.Label lblBCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Attendance Codes"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   31
      Top             =   4721
      Width           =   1320
   End
   Begin VB.Label lblShift 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   5065
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   2065
      Width           =   630
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   405
      Width           =   555
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   737
      Width           =   825
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   1401
      Width           =   420
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   1733
      Width           =   450
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
      TabIndex        =   24
      Top             =   2397
      Width           =   1290
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
      TabIndex        =   23
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   1069
      Width           =   615
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   3725
      Width           =   510
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   4057
      Width           =   1125
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   4389
      Width           =   540
   End
End
Attribute VB_Name = "frmRTimesheetWCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CodeCodes(10, 2)

Dim ODIV, ODivD, xGlbDiv, xGlbDivDesc, xKeyPress, xTxtDiv, xLblDivDesc
Dim LastlastID, LastlastNme, LastFirstNme, xTxtEEID, xLblEEName
Dim PosFlag As Boolean, strShift, strPosCode, strPosGrp
Dim strReqCourses As String
Dim strStartDate, strEndDate

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdView_Click()
Dim x%
Dim strWHand As String

On Error GoTo CRW_Err

If CriCheck() Then
'    cmdView.Enabled = False
    MDIMain.MainToolBar.ButtonS("preview").Enabled = False
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    Screen.MousePointer = HOURGLASS
    x% = Cri_SetAll()
    Screen.MousePointer = DEFAULT
    MDIMain.Timer1.Enabled = True
'    cmdView.Enabled = True
MDIMain.MainToolBar.ButtonS("preview").Enabled = True
End If

Exit Sub

CRW_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox "CRW ERROR : " & Chr(10) & "[" & str(Err) & "] : " & Me.vbxCrystal.LastErrorString
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ENTITLEMENTS", "VIEW")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub dlpDateRange_LostFocus(Index As Integer)
If IsDate(dlpDateRange(0)) Then
    dlpDateRange(1) = DateAdd("d", 13, dlpDateRange(0))
Else
    dlpDateRange(1) = ""
End If
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
MDIMain.MainToolBar.ButtonS("print").Enabled = False
MDIMain.mnu_F_Print.Enabled = False
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = Me.name

Screen.MousePointer = HOURGLASS

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

Call setRptCaption(Me)
Call setCaption(lblBCode)

If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

If glbWFC Then
    lblFromTo.FontBold = True
    lblSection.FontBold = True
    If Len(glbPlantCode) > 0 Then
        clpCode(5).Text = glbPlantCode
        If Not (glbPlantCode = "MISS" Or glbPlantCode = "TROY") Then
            clpCode(5).Enabled = False
        End If
    End If
End If

Call INI_Controls(Me)

txtYear = Year(Date)

Screen.MousePointer = DEFAULT

End Sub

Private Function Cri_SetAll()
Dim x%, strRName$

Cri_SetAll = False
On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
PosFlag = False: strShift = "": strPosCode = "": strPosGrp = ""
Call XLSwriter_All

' window title if appropriate
'Me.vbxCrystal.WindowTitle = "Training Matrix Report"

Cri_SetAll = True
Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "Comp Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function CriCheck()
Dim x%

CriCheck = False

If Not clpDiv.ListChecker Then
'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    'MsgBox lStr("If Division Entered - it must be known")
    'clpDiv.SetFocus
    Exit Function
End If

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known"
    'clpDept.SetFocus
    Exit Function
End If

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

For x% = 0 To 5
    If Not clpCode(x).ListChecker Then Exit Function
Next x%

If Len(txtWeek.Text) = 0 Then
    MsgBox "Pay Period # is requried field"
    txtWeek.SetFocus
    Exit Function
End If

'For x% = 0 To 1
' If Len(dlpDateRange(x%).Text) = 0 Then
'    MsgBox "Date Range is requried field"
'    dlpDateRange(x%).Text = ""
'    dlpDateRange(x%).SetFocus
'    Exit Function
' End If
' If Len(dlpDateRange(x%).Text) > 0 Then
'    If Not IsDate(dlpDateRange(x%).Text) Then
'        MsgBox "Not a valid date"
'        dlpDateRange(x%).Text = ""
'        dlpDateRange(x%).SetFocus
'        Exit Function
'    End If
' End If
'Next x%
'If Weekday(dlpDateRange(0)) <> 1 Then
'    MsgBox "From Date must be Sunday"
'    dlpDateRange(0).Text = ""
'    dlpDateRange(0).SetFocus
'    Exit Function
'End If

If Len(clpJOB.Text) > 0 And clpJOB.Caption = "Unassigned" Then
    MsgBox "Job code must be valid"
    clpJOB.SetFocus
    Exit Function
End If

'Hemu - 05/13/2003 Begin - From Date and To Date
If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
    End If
End If
'Hemu - 05/13/2003 End

If glbWFC Then
    If Len(clpCode(5)) = 0 Then
        MsgBox lStr("Section is required.")
        clpCode(5).SetFocus
        Exit Function
    End If
    If Not IsDate(dlpDateRange(0)) Then
        MsgBox "From Date is required!"                       '
        Me.dlpDateRange(0).SetFocus
        Exit Function
    End If
    If Not IsDate(dlpDateRange(1)) Then
        MsgBox "To Date is required!"                       '
        Me.dlpDateRange(1).SetFocus
        Exit Function
    End If
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

CriCheck = True
End Function

Function LanchXlsW98(xFileName)
On Error GoTo Error_Deal
    LanchXlsW98 = False
    Shell "Start " & GetShortName(xFileName)
    LanchXlsW98 = True
Exit Function
Error_Deal:

End Function

Public Function GetShortName(ByVal sLongFileName As String) As String
    Dim lRetVal As Long, sShortPathName As String, iLen As Integer
    'Set up buffer area for API function call return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)

    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    'Strip away unwanted characters.
    GetShortName = Left(sShortPathName, lRetVal)
End Function

Public Property Get ChangeAction() As UpdateStateEnum
ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = Reports
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = False
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = False
End Property

Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Call set_Buttons
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

' 11/23/2003 Begin
Private Sub XLSwriter_All()
Dim CoJobCodeS As New Collection
Dim CoJobCode As New Collection
Dim rsHours As New ADODB.Recordset
Dim RsEdEmp As New ADODB.Recordset
Dim rsProjectCode As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim strCourseCodes As String
Dim SQLQ, I, J, k, L, M, x, Y, Q, xMax As Integer, xRecNum As Long
Dim Iloop, Jloop As Integer
Dim exApp As Object
Dim exBook As Object
Dim exSheet As Object
Dim exApp1 As Excel.Application
'Dim exBook As excel.Workbooks
'Dim exSheet As excel.Worksheets

Dim xlsFileTmp, xlsFileMat, xEmpnbr, xNA, xJobDesc, xJobCode
Dim StartLine As Long, strTemp As String
Dim StartRow As Integer

Dim FStopRow As Integer
Dim SStopRow As Integer
Dim TStopRow As Integer

Dim FStopFlag As Boolean
Dim SStopFlag As Boolean
Dim TStopFlag As Boolean

Dim TotalRow As Integer
Dim OVStartRow As Integer

Dim NewDateFormat
Dim ARCode(260, 3)
'Dim ARDesc(260, 3)
'Dim NAflag(260, 3)
Dim flgReqC As Boolean, strDisp As String  'strReqCourses As String,

Dim z As Integer
Dim iRepeat, iOTOV As Integer
Dim iNoAccount, iOTNoAccount, iOVNoAccount As Integer
Dim iARCODE As Integer
Dim iMod As Integer
Dim iLine As Integer
Dim iLLine As Integer
Dim FFFRowChr As String  'First Period First Row first character ( A for AB)
Dim FFSRowChr As String  'First Period First Row second character( B for AB)
Dim FLFRowChr As String  'First Period Last Row first character ( A for AB)
Dim FLSRowChr As String  'First Period Last Row second character( B for AB)
Dim SFFRowChr As String  'Second Period First Row first character ( A for AB)
Dim SFSRowChr As String  'Second Period First Row second character( B for AB)
Dim SLFRowChr As String  'Second Period Last Row first character ( A for AB)
Dim SLSRowChr As String  'Second Period Last Row second character( B for AB)
Dim TFFRowChr As String  'Third Period First Row first character ( A for AB)
Dim TFSRowChr As String  'Third Period First Row second character( B for AB)
Dim TLFRowChr As String  'Third Period Last Row first character ( A for AB)
Dim TLSRowChr As String  'Third Period Last Row second character( B for AB)
Dim strCourse As String
Dim strAndWhere As String
Dim QStr, QStr1
Dim xRes
Dim strEMPNBRList As String
Dim strEMPNBR As String
Dim iWorksheet As Integer
Dim iDate, DayNum As Integer
Dim strRANGE As String
Dim sFlag, sOTFlag, sOVFlag As Boolean
Dim xRegularRate
Dim dtDate As Date
Dim strTableName
Dim strStatus
Dim machinehours
Dim gdbESS As New ADODB.Connection

Dim xExcelRptPath  As String

If glbSQL Or glbOracle Then
    Set gdbESS = gdbAdoIhr001
Else
    gdbESS.Open Replace(glbAdoIHRDB, "IHR001", "IHRESS")
End If
    
On Error GoTo Err_XLS

    'Ticket #22034 - Get Excel reports path
    If gsTRAININGMATRIX Then
        xExcelRptPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xExcelRptPath) = 0 Then
        xExcelRptPath = glbIHRREPORTS
    End If

    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "TimesheetWCostTmp.xls"
    
    'Ticket #22034 - May need to save the report in different path
    'xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "TimesheetWCostMat.xls"
    xlsFileMat = xExcelRptPath & IIf(Right(xExcelRptPath, 1) = "\", "", "\") & "TimesheetWCostMat.xls"

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."

    FileCopy xlsFileTmp, xlsFileMat

    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)

'With exBook.PageSetup
'    .LeftMargin = Application.InchesToPoints(0.5)
'    .RightMargin = Application.InchesToPoints(0.75)
'    .TopMargin = Application.InchesToPoints(1.5)
'    .BottomMargin = Application.InchesToPoints(1)
'    .HeaderMargin = Application.InchesToPoints(0.5)
'    .FooterMargin = Application.InchesToPoints(0.5)
'End With
    strStartDate = Date_SQL(dlpDateRange(0).Text)
    strEndDate = Date_SQL(dlpDateRange(1).Text)
    'Get Attendace code....
    SQLQ = "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME FROM HREMP where "
    SQLQ = SQLQ & getWSQLQ(False)
    SQLQ = SQLQ & " ORDER BY ED_SURNAME,ED_FNAME"

    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic

    If rsTemp.EOF Then
         MsgBox ("There is no employee based on the search cretira.")
    End If
     
    iWorksheet = 0
    exBook.Worksheets.Copy After:=exBook.Worksheets(exBook.Worksheets.count)
    Do While Not rsTemp.EOF
        MDIMain.panHelp(0).FloodPercent = 10
        strStatus = getStatus(rsTemp.Fields("ED_EMPNBR"))
        Select Case strStatus
        Case "APPROVED", ""
            strTableName = "HR_ATTENDANCE"
        Case Else
            strTableName = "HR_TIMESHEET"
        End Select
        

        QStr = "SELECT DISTINCT AD_PROJECT_CODE,HR_PROJECT_CODE.DESCRIPTION,HR_PROJECT_CODE.GL_NUMBER FROM "
        QStr = QStr & strTableName & " LEFT JOIN HR_PROJECT_CODE "
        QStr = QStr & " ON " & strTableName & ".AD_PROJECT_CODE=HR_PROJECT_CODE.PROJECT_CODE "
        QStr = QStr & " WHERE AD_EMPNBR =" & rsTemp.Fields("ED_EMPNBR")
        QStr = QStr & " AND AD_DOA>=" & strStartDate
        QStr = QStr & " AND AD_DOA<=" & strEndDate
        If rsProjectCode.State <> 0 Then rsProjectCode.Close
        
        If glbSQL Or glbOracle Then
            rsProjectCode.Open QStr, gdbAdoIhr001, adOpenForwardOnly
        Else
            If strTableName = "HR_TIMESHEET" Then
                rsProjectCode.Open QStr, gdbESS, adOpenForwardOnly
            Else
                rsProjectCode.Open QStr, gdbAdoIhr001, adOpenForwardOnly
            End If
        End If
        
        If Not rsProjectCode.EOF Then
            rsProjectCode.Close
            If iWorksheet <> 0 Then exBook.Worksheets(exBook.Worksheets.count).Copy After:=exBook.Worksheets(exBook.Worksheets.count)
          
            iWorksheet = iWorksheet + 1
            Set exSheet = exBook.Worksheets(iWorksheet)
            exSheet.Cells(2, 1) = "Pay Period: " & Format(dlpDateRange(1).Text, "mmm dd, yyyy")
            If Not (IsDate(dlpDateRange(0).Text) Or IsDate(dlpDateRange(1).Text)) Then
                exSheet.Cells(2, 5) = "No date entered"
            Else
                strTemp = ""
                If IsDate(dlpDateRange(0).Text) Then
                    strTemp = "From Date: " & Format(dlpDateRange(0).Text, "mmm dd, yyyy") & "  "
                End If
                If IsDate(dlpDateRange(1).Text) Then
                    strTemp = strTemp & "To Date: " & Format(dlpDateRange(1).Text, "mmm dd, yyyy")
                End If
                exSheet.Cells(2, 3) = strTemp
            End If
            
            exSheet.Cells(3, 1) = ""
            exSheet.Cells(3, 2) = ""
    
            exSheet.Cells(4, 1) = "NAME:"
            exSheet.Cells(4, 2) = rsTemp.Fields("ED_SURNAME") & "," & rsTemp.Fields("ED_FNAME")
            exSheet.Cells(4, 4) = "Status:" & strStatus
            exSheet.Cells(5, 1) = "EMPL. #:"
            exSheet.Cells(5, 2) = rsTemp.Fields("ED_EMPNBR")
            StartLine = 7
    
            exSheet.Cells(StartLine, 1) = "Date"
            exSheet.Cells(StartLine, 2) = "Job Acct. #"
    
            FStopFlag = True
            SStopFlag = True
            TStopFlag = True
    
            'step2-6
            'Retrieve title begin ???????????????
            'every paid hours not including vac,sick and ots
            SQLQ = QStr & " AND AD_REASON NOT IN ('BT','FL','VAC','PUB','SICK','STD','STDP','OT','OT15','OT20','STDU','USIC','UVAC','WSIU','LWOP','LTD','FAML','EMRG','MAT','ADJ','BFHR','BFSH','LATE','LTDC','MEAL','OV1','OV2','PCR','SUSP','WCR')"
            rsProjectCode.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            StartRow = 3
            Do While Not rsProjectCode.EOF 'getting REG titles
                'create LIne 5, line 6 and project code or Reason code
                exSheet.Rows(StartRow).ShrinkToFit = True
                exSheet.Cells(StartLine - 1, StartRow) = IIf(rsProjectCode.Fields("GL_NUMBER") & "" = "", "NO ACCOUNT CODE", rsProjectCode.Fields("GL_NUMBER"))
                exSheet.Cells(StartLine, StartRow) = rsProjectCode.Fields("AD_PROJECT_CODE") & ""
                ARCode(StartRow, 0) = rsProjectCode.Fields("AD_PROJECT_CODE") & ""
                StartRow = StartRow + 1
                rsProjectCode.MoveNext
            Loop
            rsProjectCode.Close
            
            'for vacation and public holiday
            SQLQ = QStr & " AND AD_REASON in ('BT','FL','PUB','VAC')"
            rsProjectCode.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            Do While Not rsProjectCode.EOF 'getting Vacation titles
                exSheet.Rows(StartRow).ShrinkToFit = True
                exSheet.Cells(StartLine - 1, StartRow) = IIf(rsProjectCode.Fields("GL_NUMBER") & "" = "", "NO ACCOUNT CODE", rsProjectCode.Fields("GL_NUMBER"))
                exSheet.Cells(StartLine, StartRow) = rsProjectCode.Fields("AD_PROJECT_CODE") & ""
                ARCode(StartRow, 0) = rsProjectCode.Fields("AD_PROJECT_CODE") & ""
                StartRow = StartRow + 1
                rsProjectCode.MoveNext
            Loop
            rsProjectCode.Close
            'for sick time
            
            SQLQ = QStr & " AND AD_REASON in ('STD','STDP','SICK')"
            rsProjectCode.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            Do While Not rsProjectCode.EOF  'getting Sick titles
                exSheet.Rows(StartRow).ShrinkToFit = True
                exSheet.Cells(StartLine - 1, StartRow) = IIf(rsProjectCode.Fields("GL_NUMBER") & "" = "", "NO ACCOUNT CODE", rsProjectCode.Fields("GL_NUMBER"))
                exSheet.Cells(StartLine, StartRow) = rsProjectCode.Fields("AD_PROJECT_CODE") & ""
                ARCode(StartRow, 0) = rsProjectCode.Fields("AD_PROJECT_CODE") & ""
                StartRow = StartRow + 1
                rsProjectCode.MoveNext
            Loop
            rsProjectCode.Close
            
            'Daily Total
            If FStopFlag Then
                FStopFlag = False
                FStopRow = StartRow
                exSheet.Cells(StartLine, StartRow) = "Daily Total"
    
                FFFRowChr = ""
                FFSRowChr = "C"
                FLFRowChr = ""
                FLSRowChr = getRowChar(FStopRow - 1)
    
            End If
            StartRow = StartRow + 1
    
            sFlag = False
            SQLQ = QStr & " AND left(AD_REASON,2)='OT'"
            rsProjectCode.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    
            If Not rsProjectCode.EOF Then
                Do While Not rsProjectCode.EOF
                    exSheet.Rows(StartRow).ShrinkToFit = True
                    exSheet.Cells(StartLine - 1, StartRow) = IIf(rsProjectCode.Fields("GL_NUMBER") & "" = "", "NO ACCOUNT CODE", rsProjectCode.Fields("GL_NUMBER"))
                    exSheet.Cells(StartLine, StartRow) = rsProjectCode.Fields("AD_PROJECT_CODE") & ""
                    ARCode(StartRow, 1) = rsProjectCode.Fields("AD_PROJECT_CODE") & ""
                    StartRow = StartRow + 1
                    rsProjectCode.MoveNext
                Loop
                sFlag = True
                sOTFlag = True
            End If
            rsProjectCode.Close
            
            'For Banked Overtime
            If sOTFlag = True Then
'                strRANGE = getRowChar(FStopRow + 1) & "5:" & getRowChar(StartRow - 1) & 5
'                exSheet.Range(strRANGE).Merge (True)
'                exSheet.Cells(StartLine - 2, FStopRow + 1).HorizontalAlignment = xlCenter
'                exSheet.Cells(StartLine - 2, FStopRow + 1).Font.Size = 7
'                exSheet.Cells(StartLine - 2, FStopRow + 1).Font.Bold = True
'                exSheet.Cells(StartLine - 2, FStopRow + 1) = "OT" 'For OT
            End If
            
             
            If sFlag Then
                If SStopFlag Then
                    SStopFlag = False
                    SStopRow = StartRow
                    strRANGE = getRowChar(StartRow) & "5:" & getRowChar(StartRow) & 7
                    exSheet.Range(strRANGE).Merge (True)
                    exSheet.Range(strRANGE).VerticalAlignment = xlCenter
                    exSheet.Cells(StartLine, StartRow) = "Daily Total"
        
                    SFFRowChr = ""
                    SFSRowChr = getRowChar(FStopRow + 1)
                    
                    SLFRowChr = ""
                    SLSRowChr = getRowChar(SStopRow - 1)
        
                End If
                
'                strRANGE = SFSRowChr & "4:" & getRowChar(SStopRow) & 4
'                exSheet.Range(strRANGE).Merge (True)
'                exSheet.Cells(StartLine - 3, FStopRow + 1).HorizontalAlignment = xlCenter
'                exSheet.Cells(StartLine - 3, FStopRow + 1).Font.Size = 7
'                exSheet.Cells(StartLine - 3, FStopRow + 1).Font.Bold = True
'                exSheet.Cells(StartLine - 3, FStopRow + 1) = "Banked Overtime" 'For OVERTIME
'
'                strRANGE = SFSRowChr & "4:" & getRowChar(SStopRow) & 5
'                exSheet.Range(strRANGE).Borders.LineStyle = xlContinuous
'
                
                strRANGE = SFSRowChr & "5:" & getRowChar(SStopRow) & 5
                exSheet.Range(strRANGE).Merge (True)
                exSheet.Cells(StartLine - 2, FStopRow + 1).HorizontalAlignment = xlCenter
                exSheet.Cells(StartLine - 2, FStopRow + 1).Font.Size = 7
                exSheet.Cells(StartLine - 2, FStopRow + 1).Font.Bold = True
                exSheet.Cells(StartLine - 2, FStopRow + 1) = "Banked Overtime" 'For OVERTIME
        
                strRANGE = SFSRowChr & "5:" & getRowChar(SStopRow) & 5
                exSheet.Range(strRANGE).Borders.LineStyle = xlContinuous
                
                StartRow = StartRow + 1
            End If
            
            StartLine = StartLine + 1
            TotalRow = StartRow - 1
            'Retrieve title end
    
    
           'Retrieve data begin ???????????????
            DayNum = DateDiff("d", dlpDateRange(0).Text, dlpDateRange(1).Text)
    
            For iDate = 0 To DayNum
                dtDate = DateAdd("d", iDate, dlpDateRange(0).Text)
                exSheet.Cells(StartLine, 1) = WeekdayName(Weekday(dtDate), True) & "."
                exSheet.Cells(StartLine, 1).HorizontalAlignment = xlLeft
                exSheet.Cells(StartLine + 1, 1) = MonthName(month(dtDate), True) & ". " & Day(dtDate)
                exSheet.Cells(StartLine + 1, 1).HorizontalAlignment = xlLeft
                exSheet.Cells(StartLine, 2) = "Hours"
                exSheet.Cells(StartLine + 1, 2) = "Rate"
                exSheet.Cells(StartLine + 2, 2) = "Wages $"
                exSheet.Cells(StartLine + 3, 2) = "Mach Acct. #"
                exSheet.Cells(StartLine + 4, 2) = "Hours"
                exSheet.Cells(StartLine + 5, 2) = "Rate"
                exSheet.Cells(StartLine + 6, 2) = "Machine $"
    
                StartRow = 3
                SQLQ = "SELECT AD_HRS,AD_REASON,AD_LTIME,AD_PROJECT_CODE,AD_JOB,AD_SALARY,AD_MACHINE_NUM,AD_MACHINE_HRS,AD_MACHINE_RATE,AD_MACHINE_HRS,HR_SALARY_HISTORY.SH_SALARY,HR_SALARY_HISTORY.SH_WHRS,HR_SALARY_HISTORY.SH_SALCD "
                If glbSQL Or glbOracle Then
                    SQLQ = SQLQ & "FROM " & strTableName & " INNER JOIN HR_SALARY_HISTORY "
                    SQLQ = SQLQ & " ON " & strTableName & ".AD_EMPNBR=HR_SALARY_HISTORY.SH_EMPNBR AND HR_SALARY_HISTORY.SH_CURRENT<>0 "
                    SQLQ = SQLQ & " WHERE " & strTableName & ".AD_EMPNBR = " & rsTemp.Fields("ED_EMPNBR")
                Else
                    SQLQ = SQLQ & "FROM " & strTableName & ", HR_SALARY_HISTORY "
                    SQLQ = SQLQ & " WHERE " & strTableName & ".AD_EMPNBR=HR_SALARY_HISTORY.SH_EMPNBR AND HR_SALARY_HISTORY.SH_CURRENT<>0 "
                    SQLQ = SQLQ & " AND " & strTableName & ".AD_EMPNBR = " & rsTemp.Fields("ED_EMPNBR")
                End If
                
                SQLQ = SQLQ & " AND AD_REASON NOT IN ('STDU','USIC','UVAC','WSIU','LWOP','LTD','FAML','EMRG','MAT','ADJ','BFHR','BFSH','LATE','LTDC','MEAL','OV1','OV2','PCR','SUSP','WCR')" '' #9389 added "MAT"; ON dEC 15,2005 ADD ,|ADJ,|BFHR,|BFSH,|LATE,|LTDC,|MEAL,|OV1,|OV2,|PCR,|SUSP,|WCR,| #9521
                If IsDate(dlpDateRange(0)) Then SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(DateAdd("d", iDate, dlpDateRange(0).Text))
        
                rsHours.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                If Not rsHours.EOF Then
                    Do While Not rsHours.EOF
                        iOTOV = 0
                        If Left(rsHours.Fields("AD_REASON"), 2) = "OT" Then iOTOV = 1
                '        If Left(rsHours.Fields("AD_REASON"), 2) = "OV" Then iOTOV = 2
                        
                        If Not IsNull(rsHours.Fields("AD_PROJECT_CODE")) Then
                            iRepeat = 0
                            For iARCODE = 3 To 260
                                If rsHours.Fields("AD_PROJECT_CODE") = ARCode(iARCODE, iOTOV) Then
                                    iRepeat = iRepeat + 1
                                    StartRow = iARCODE
                                    If rsHours.Fields("AD_LTIME") = iRepeat Then Exit For
                                End If
                            Next
                        Else
                            Select Case iOTOV
                            Case 0
                                StartRow = iNoAccount
                            Case 1
                                StartRow = iOTNoAccount
                            Case 2
                                StartRow = iOVNoAccount
                            End Select
                        End If
                        'Put cell for hours
                        If Len(Trim(exSheet.Cells(StartLine, StartRow))) > 0 Then
                            exSheet.Cells(StartLine, StartRow) = exSheet.Cells(StartLine, StartRow) + rsHours.Fields("AD_HRS")
                        Else
                            exSheet.Cells(StartLine, StartRow) = rsHours.Fields("AD_HRS")
                        End If
                        'Put cell for rate
                        xRegularRate = CDbl(rsHours("AD_SALARY")) 'for #8578
                        exSheet.Cells(StartLine + 1, StartRow) = xRegularRate
                        'put cell for wages
                        If Len(Trim(exSheet.Cells(StartLine + 2, StartRow))) > 0 Then
                            exSheet.Cells(StartLine + 2, StartRow) = exSheet.Cells(StartLine + 2, StartRow) + Round(rsHours.Fields("AD_HRS") * xRegularRate + 0.0001, 2)
                        Else
                            exSheet.Cells(StartLine + 2, StartRow) = Round(rsHours.Fields("AD_HRS") * xRegularRate + 0.0001, 2)
                        End If
'                        If Len(Trim(exSheet.Cells(StartLine + 2, StartRow))) > 0 Then
'                            exSheet.Cells(StartLine + 2, StartRow) = exSheet.Cells(StartLine + 2, StartRow) + rsHours.Fields("AD_HRS") * xRegularRate
'                        Else
'                            exSheet.Cells(StartLine + 2, StartRow) = rsHours.Fields("AD_HRS") * xRegularRate
'                        End If
                        'Put cell for machine number
                        If Len(Trim(exSheet.Cells(StartLine + 3, StartRow))) > 0 And rsHours.Fields("AD_MACHINE_NUM") & "" <> "" Then
                            exSheet.Cells(StartLine + 3, StartRow) = exSheet.Cells(StartLine + 3, StartRow) & "|" & rsHours.Fields("AD_MACHINE_NUM")
                        Else
                            If Not IsNull(rsHours.Fields("AD_MACHINE_NUM")) Then
                                exSheet.Cells(StartLine + 3, StartRow) = rsHours.Fields("AD_MACHINE_NUM") & ""
                            End If
                        End If
                        'Put cell for machine hours
                        If Len(Trim(exSheet.Cells(StartLine + 4, StartRow))) > 0 And Not IsNull(rsHours.Fields("AD_MACHINE_HRS")) Then
                            exSheet.Cells(StartLine + 4, StartRow) = exSheet.Cells(StartLine + 4, StartRow) & "|" & rsHours.Fields("AD_MACHINE_HRS")
                        Else
                            If Not IsNull(rsHours.Fields("AD_MACHINE_HRS")) Then
                                exSheet.Cells(StartLine + 4, StartRow) = rsHours.Fields("AD_MACHINE_HRS") & ""
                            End If
                        End If
                        machinehours = machinehours + rsHours.Fields("AD_MACHINE_HRS")

                        'Put cell for machine wages
                        If Len(Trim(exSheet.Cells(StartLine + 5, StartRow))) > 0 And Not IsNull(rsHours.Fields("AD_MACHINE_RATE")) Then
                            exSheet.Cells(StartLine + 5, StartRow) = exSheet.Cells(StartLine + 5, StartRow) & "|" & rsHours.Fields("AD_MACHINE_RATE")
                        Else
                            If Not IsNull(rsHours.Fields("AD_MACHINE_RATE")) Then
                                exSheet.Cells(StartLine + 5, StartRow) = rsHours.Fields("AD_MACHINE_RATE")
                            End If
                        End If
                        'put cell for machine wages
                        If Not IsNull(rsHours.Fields("AD_MACHINE_RATE")) And Not IsNull(rsHours.Fields("AD_MACHINE_HRS")) Then
                            exSheet.Cells(StartLine + 6, StartRow) = exSheet.Cells(StartLine + 6, StartRow) + Round(rsHours.Fields("AD_MACHINE_RATE") * rsHours.Fields("AD_MACHINE_HRS") + 0.0001, 2)
                            'exSheet.Cells(StartLine + 6, StartRow) = exSheet.Cells(StartLine + 6, StartRow) + rsHours.Fields("AD_MACHINE_RATE") * rsHours.Fields("AD_MACHINE_HRS")
                        End If
                        
                        exSheet.Cells(StartLine, StartRow).HorizontalAlignment = xlCenter
                        exSheet.Cells(StartLine + 1, StartRow).HorizontalAlignment = xlCenter
                        exSheet.Cells(StartLine + 2, StartRow).HorizontalAlignment = xlCenter
                        exSheet.Cells(StartLine + 3, StartRow).HorizontalAlignment = xlCenter
                        exSheet.Cells(StartLine + 4, StartRow).HorizontalAlignment = xlCenter
                        exSheet.Cells(StartLine + 5, StartRow).HorizontalAlignment = xlCenter
                        exSheet.Cells(StartLine + 6, StartRow).HorizontalAlignment = xlCenter
                        rsHours.MoveNext
                    Loop
    
                End If
                exSheet.Cells(StartLine, FStopRow) = "=sum(C" & CStr(StartLine) & ":" & FLFRowChr & FLSRowChr & CStr(StartLine) & ")"
                exSheet.Cells(StartLine + 2, FStopRow) = "=sum(C" & CStr(StartLine + 2) & ":" & FLFRowChr & FLSRowChr & CStr(StartLine + 2) & ")"
                exSheet.Cells(StartLine + 4, FStopRow) = "=sum(C" & CStr(StartLine + 4) & ":" & FLFRowChr & FLSRowChr & CStr(StartLine + 4) & ")"
                exSheet.Cells(StartLine + 6, FStopRow) = "=sum(C" & CStr(StartLine + 6) & ":" & FLFRowChr & FLSRowChr & CStr(StartLine + 6) & ")"
'                exSheet.Cells(StartLine + 6, FStopRow).FixedDecimal = True
'                exSheet.Cells(StartLine + 6, FStopRow).FixedDecimalPlaces = 2
                If sFlag Then
                    exSheet.Cells(StartLine, SStopRow) = "=sum(" & SFFRowChr & SFSRowChr & CStr(StartLine) & ":" & SLFRowChr & SLSRowChr & CStr(StartLine) & ")"
                    exSheet.Cells(StartLine + 2, SStopRow) = "=sum(" & SFFRowChr & SFSRowChr & CStr(StartLine + 2) & ":" & SLFRowChr & SLSRowChr & CStr(StartLine + 2) & ")"
                    exSheet.Cells(StartLine + 4, SStopRow) = "=sum(" & SFFRowChr & SFSRowChr & CStr(StartLine + 4) & ":" & SLFRowChr & SLSRowChr & CStr(StartLine + 4) & ")"
                    exSheet.Cells(StartLine + 6, SStopRow) = "=sum(" & SFFRowChr & SFSRowChr & CStr(StartLine + 6) & ":" & SLFRowChr & SLSRowChr & CStr(StartLine + 6) & ")"
                End If
                
                StartLine = StartLine + 7
                rsHours.Close
            Next
            exSheet.Cells(StartLine, 2) = "Total Mach."
            exSheet.Cells(StartLine, 2).Font.Size = 6
            exSheet.Cells(StartLine, 2).Font.Bold = True
            exSheet.Cells(StartLine + 1, 2) = "Total Hours"
            exSheet.Cells(StartLine + 1, 2).Font.Size = 6
            exSheet.Cells(StartLine + 1, 2).Font.Bold = True
            exSheet.Cells(StartLine + 2, 2) = "Total Wages"
            exSheet.Cells(StartLine + 2, 2).Font.Size = 6
            exSheet.Cells(StartLine + 2, 2).Font.Bold = True
            For iRepeat = 3 To TotalRow
                FFFRowChr = ""
                If Int(iRepeat / 26) >= 1 And ((iRepeat / 26) <> Int(iRepeat / 26)) Then
                    FFFRowChr = Chr(64 + Int(iRepeat / 26))
                End If
                iMod = iRepeat Mod 26
                If iMod = 0 Then
                    FFSRowChr = "Z"
                Else
                    FFSRowChr = Chr(64 + iMod)
                End If
    
    
                exSheet.Cells(StartLine, iRepeat) = "sum("
                For iLine = 14 To StartLine - 1 Step 7
                    exSheet.Cells(StartLine, iRepeat) = exSheet.Cells(StartLine, iRepeat) & FFFRowChr & FFSRowChr & CStr(iLine) & ","
                Next
                exSheet.Cells(StartLine, iRepeat) = "=" & Left(exSheet.Cells(StartLine, iRepeat), Len(exSheet.Cells(StartLine, iRepeat)) - 1) & ")"
                exSheet.Cells(StartLine + 1, iRepeat) = "sum("
                For iLine = 8 To StartLine - 1 Step 7
                    exSheet.Cells(StartLine + 1, iRepeat) = exSheet.Cells(StartLine + 1, iRepeat) & FFFRowChr & FFSRowChr & CStr(iLine) & ","
                Next
                exSheet.Cells(StartLine + 1, iRepeat) = "=" & Left(exSheet.Cells(StartLine + 1, iRepeat), Len(exSheet.Cells(StartLine + 1, iRepeat)) - 1) & ")"
                exSheet.Cells(StartLine + 2, iRepeat) = "sum("
                For iLine = 10 To StartLine - 1 Step 7
                    exSheet.Cells(StartLine + 2, iRepeat) = exSheet.Cells(StartLine + 2, iRepeat) & FFFRowChr & FFSRowChr & CStr(iLine) & ","
                Next
                exSheet.Cells(StartLine + 2, iRepeat) = "=" & Left(exSheet.Cells(StartLine + 2, iRepeat), Len(exSheet.Cells(StartLine + 2, iRepeat)) - 1) & ")"
                
                
            Next
            exSheet.Cells(StartLine + 3, 1) = exSheet.Cells(7, 1)
            exSheet.Cells(StartLine + 3, 1).HorizontalAlignment = xlCenter
            exSheet.Cells(StartLine + 3, 1).Font.Size = 6
            exSheet.Cells(StartLine + 3, 1).Font.Bold = True
            exSheet.Cells(StartLine + 3, 2) = exSheet.Cells(7, 2)
            exSheet.Cells(StartLine + 3, 2).HorizontalAlignment = xlCenter
            exSheet.Cells(StartLine + 3, 2).Font.Size = 6
            exSheet.Cells(StartLine + 3, 2).Font.Bold = True
            For iRepeat = 3 To TotalRow
                exSheet.Cells(StartLine + 3, iRepeat) = exSheet.Cells(7, iRepeat)
                exSheet.Cells(StartLine + 3, iRepeat).HorizontalAlignment = xlCenter
                exSheet.Cells(StartLine + 3, iRepeat).Font.Size = 6
                exSheet.Cells(StartLine + 3, iRepeat).Font.Bold = True
            Next

            strRANGE = "A1:" & getRowChar(TotalRow) & 1
            exSheet.Range(strRANGE).Merge (True)
            exSheet.Cells(1, 1).HorizontalAlignment = xlCenter
            exSheet.Cells(1, 1).Font.Size = 11
            exSheet.Cells(1, 1).Font.Bold = True
            exSheet.Cells(1, 1) = "Timesheet With Equipment Cost"
            
            strRANGE = getRowChar(TotalRow + 1) & "1:AG109"
            exSheet.Range(strRANGE).Delete (True)

                        
            exSheet.name = rsTemp.Fields("ED_SURNAME") & "," & rsTemp.Fields("ED_FNAME")
        End If
Loopend:
        For Iloop = 0 To 260
            For Jloop = 0 To 3
                ARCode(Iloop, Jloop) = ""
            Next
        Next

        rsTemp.MoveNext
        DoEvents
    Loop
    rsTemp.Close

    exBook.Worksheets(exBook.Worksheets.count).Visible = False  'delete the template worksheet
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Call Pause(1)
    'launch Excel file
    'Shell "Start " & GetShortName(xlsFileMat)
    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If

    Exit Sub
Err_XLS:
'    If Err.Number = 91 Then
'        MsgBox Err.Number
'        Resume Next
'    End If
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT

    If Err = 1004 Then
        Resume Next
    End If

    If Err = 75 Then
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Sub
    End If
    If Not exBook Is Nothing Then
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
    End If
    If Err = 70 Then
        Set exApp = Nothing
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Sub
    End If

    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "XLSwriter", "", "Select")
Resume Next
End Sub

Private Function getWSQLQ(WithAtt As Boolean)
Dim QStr
QStr = glbSeleDeptUn
If clpDiv <> "" Then QStr = QStr & " AND ED_DIV in ('" & Replace(clpDiv, ",", "','") & "')"
If clpCode(0) <> "" Then QStr = QStr & " AND ED_LOC='" & clpCode(0) & "'"

If clpCode(1) <> "" Then
    'QStr = QStr & " AND ED_ORG='" & clpCode(1) & "'"
    QStr = QStr & " AND ED_ORG in  ('" & Replace(clpCode(1).Text, ",", "','") & "')"
End If

If clpCode(2) <> "" Then QStr = QStr & " AND ED_EMP in ('" & Replace(clpCode(2), ",", "','") & "')"
If clpCode(3) <> "" Then QStr = QStr & " AND ED_REGION='" & clpCode(3) & "'"
If clpCode(4) <> "" Then QStr = QStr & " AND ED_ADMINBY='" & clpCode(4) & "'"
If clpCode(5) <> "" Then QStr = QStr & " AND ED_SECTION='" & clpCode(5) & "'"
If clpPT <> "" Then QStr = QStr & " AND ED_PT in ('" & Replace(clpPT, ",", "','") & "')"
If elpEEID.Text <> "" Then QStr = QStr & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
If WithAtt Then
    If IsDate(dlpDateRange(0)) Then QStr = QStr & " AND AD_DOA>=" & strStartDate
    If IsDate(dlpDateRange(1)) Then QStr = QStr & " AND AD_DOA<=" & strEndDate
'    If clpAtt <> "" Then QStr = QStr & " AND ES_CTYPE IN ('" & Replace(clpAtt, ",", "','") & "') "
'    If clpPayP <> "" Then QStr = QStr & " AND ES_CRSCODE IN ('" & Replace(clpPayP, ",", "','") & "') "
'    If txtShift <> "" Or clpJob <> "" Or clpPosGroup <> "" Then
'        QStr = QStr & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0"
'        If txtShift <> "" Then QStr = QStr & " AND JH_SHIFT='" & txtShift & "'"
'        If clpJob <> "" Then QStr = QStr & " AND JH_JOB='" & clpJob & "'"
'        If clpPosGroup <> "" Then QStr = QStr & " AND JH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD='" & clpPosGroup & "')"
'        QStr = QStr & ")"
'    End If
End If

getWSQLQ = QStr
End Function

Private Function getRowChar(RowNum As Integer)
Dim iMod As Integer
Dim FChar, SChar As String
    iMod = RowNum Mod 26
    FChar = ""
    SChar = ""
    If iMod <> 0 Then
        If Int(RowNum / 26) = 0 Then 'A-Y
            SChar = Chr(64 + iMod)
        Else 'AA-AY,BA-BY,CA-CY,...
            FChar = Chr(64 + Int(RowNum / 26))
            SChar = Chr(64 + iMod)
        End If
    Else
        If Int(RowNum / 26) > 1 Then '(A)Z,(B)Z,(C)Z,...
            FChar = Chr(64 + Int(RowNum / 26) - 1)
        End If
        SChar = "Z" 'Chr(64 + 26)
    End If
getRowChar = FChar & SChar
End Function

Private Sub lblFrom_Click()

End Sub

Private Sub txtWeek_Change()
Dim DateRange
DateRange = Split(getDateRange("", txtWeek, txtYear), "|")
    dlpDateRange(0) = DateRange(0)
    dlpDateRange(1) = DateRange(1)
    lblFromTo = DateRange(0) & " To " & DateRange(1)
End Sub

Private Sub txtWeek_DblClick()
Call imgIcon_Click
End Sub

Private Sub imgIcon_Click()
frmPayPeriodList.SelectedYear = Val(txtYear)
'frmPayPeriodList.PayPeriodCode = clpPayP.Text
frmPayPeriodList.Show 1
txtWeek = glbWeek
dlpDateRange(0) = glbFrom
dlpDateRange(1) = glbTo
End Sub

Private Sub txtWeek_LostFocus()
If txtWeek = "" Then
    dlpDateRange(0) = ""
    dlpDateRange(1) = ""
Else
    'FIND THE DATA RANGE FROM THE DATABASE FOR THAT WEEK #
End If
End Sub

Private Sub txtYear_Change()
Dim DateRange
DateRange = Split(getDateRange("", txtWeek, txtYear), "|")
    dlpDateRange(0) = DateRange(0)
    dlpDateRange(1) = DateRange(1)
End Sub

Function getDateRange(theClientNumber, thePayNbr, theYear)
Dim rsPayPeriod As New ADODB.Recordset
Dim SQLQ, intNum
On Error Resume Next
getDateRange = "|"
If Not IsNumeric(thePayNbr) Then Exit Function
If Not IsNumeric(theYear) Then Exit Function
SQLQ = "SELECT PP_NBR,PP_YEAR,PP_Start,PP_End FROM HR_payperiod "
SQLQ = SQLQ & " WHERE PP_NBR = " & thePayNbr
SQLQ = SQLQ & " and PP_YEAR = '" & theYear & "'"
rsPayPeriod.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

If Not rsPayPeriod.EOF Then
    getDateRange = rsPayPeriod("PP_Start") & "|" & rsPayPeriod("PP_End")
   
End If
rsPayPeriod.Close
Exit Function

End Function

Public Function getStatus(xEmpnbr)
Dim SQLQ, statusFlag
Dim rsDS As New ADODB.Recordset
Dim gdbESS As New ADODB.Connection

If glbSQL Or glbOracle Then
    Set gdbESS = gdbAdoIhr001
Else
    gdbESS.Open Replace(glbAdoIHRDB, "IHR001", "IHRESS")
End If

    On Error Resume Next
    SQLQ = "SELECT DISTINCT AD_APPROVED,AD_UPLOAD FROM HR_TIMESHEET "
    SQLQ = SQLQ & " WHERE AD_EMPNBR =" & xEmpnbr
    SQLQ = SQLQ & " AND AD_DOA >=" & strStartDate
    SQLQ = SQLQ & " AND AD_DOA <=" & strEndDate

    If glbSQL Or glbOracle Then
        rsDS.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Else
        rsDS.Open SQLQ, gdbESS, adOpenForwardOnly
    End If
    getStatus = ""
    statusFlag = True
    Do While Not rsDS.EOF
        If statusFlag Then
            If IsNull(rsDS("AD_APPROVED")) Then
                If rsDS("AD_UPLOAD") & "" = "Y" Then
                    getStatus = "SUBMITTED"
                Else
                    getStatus = "SAVED"
                End If
            Else
                getStatus = rsDS("AD_APPROVED")
                If getStatus = "RESUBMIT" Then getStatus = "RESUBMITTED"
            End If
            statusFlag = False
        Else
            'getStatus="Inconsistent"
            getStatus = "SAVED"
        End If
        rsDS.MoveNext

    Loop
    'if not rsDS.EOF then
    '   if isnull(rsDS("AD_APPROVED")) then
    '       if rsDS("AD_UPLOAD") & "" ="Y" then
    '           getStatus="SUBMITTED"
    '       else
    '           getStatus="SAVED"
    '       end if
    '   else
    '       getStatus=rsDS("AD_APPROVED")
    '       if getStatus="RESUBMIT" THEN getStatus="RESUBMITTED"
    '   end if
    'end if
    rsDS.Close
    Set rsDS = Nothing
    If Err.Number <> 0 Then
    End If
End Function

