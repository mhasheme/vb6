VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRHCASCsRpt 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Hamilton C.C.A.S. - Custom Reports"
   ClientHeight    =   8235
   ClientLeft      =   180
   ClientTop       =   825
   ClientWidth     =   10140
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8235
   ScaleWidth      =   10140
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkVacation 
      Caption         =   "Show Vacation"
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   6720
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.CheckBox chkOvertime 
      Caption         =   "Show Overtime"
      Height          =   285
      Left            =   2520
      TabIndex        =   18
      Top             =   6720
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.ComboBox comCountryOfEmp 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2115
      TabIndex        =   11
      Tag             =   "00-Country of Employment"
      Top             =   3960
      Width           =   1440
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2115
      MaxLength       =   4
      TabIndex        =   10
      Tag             =   "00-Employee Position Shift"
      Top             =   3630
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1650
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
      Left            =   1800
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   1980
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
      Left            =   1800
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1320
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
      Left            =   1800
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   990
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDLC"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   660
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
      Left            =   1800
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   330
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
      Left            =   1800
      TabIndex        =   8
      Tag             =   "00-Enter Administered By Code"
      Top             =   2970
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
      Left            =   1800
      TabIndex        =   9
      Tag             =   "00-Enter Section Code"
      Top             =   3300
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDSE"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   7
      Tag             =   "00-Enter Region Code"
      Top             =   2640
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2310
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   9000
      Top             =   7440
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
   Begin INFOHR_Controls.DateLookup dlpAsOf 
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Tag             =   "40-Month Ending"
      Top             =   5580
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3630
      TabIndex        =   15
      Tag             =   "40-Date upto and including this date forward"
      Top             =   5940
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   14
      Tag             =   "40-Date from and including this date forward"
      Top             =   5940
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.EmployeeLookup elpRept 
      Height          =   285
      Left            =   1800
      TabIndex        =   16
      Tag             =   "00-Reporting Authority 1"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   503
      ShowDescription =   0   'False
      RefreshDescriptionWhen=   2
   End
   Begin VB.ComboBox cmbReports 
      Height          =   315
      Left            =   2115
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4920
      Width           =   4395
   End
   Begin VB.Label lblRptAuth 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reporting Authority"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   6360
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblAsOf 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Month Ending"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   33
      Top             =   5625
      Width           =   990
   End
   Begin VB.Label lblCountry 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Country of Employment"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   32
      Top             =   4020
      Width           =   1620
   End
   Begin VB.Label lblShift 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   31
      Top             =   3675
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
      TabIndex        =   30
      Top             =   2010
      Width           =   630
   End
   Begin VB.Label lblReports 
      AutoSize        =   -1  'True
      Caption         =   "Reports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   4980
      Width           =   915
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   2340
      Width           =   1290
   End
   Begin VB.Label lblSection 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   3345
      Width           =   540
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
      TabIndex        =   26
      Top             =   2670
      Width           =   510
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   3000
      Width           =   1125
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
      TabIndex        =   24
      Top             =   1020
      Width           =   615
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
      TabIndex        =   23
      Top             =   1680
      Width           =   450
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
      TabIndex        =   22
      Top             =   1350
      Width           =   420
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
      TabIndex        =   21
      Top             =   690
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
      Left            =   120
      TabIndex        =   20
      Top             =   360
      Width           =   555
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
      Left            =   0
      TabIndex        =   19
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblDates 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "For the Period"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   5985
      Visible         =   0   'False
      Width           =   990
   End
End
Attribute VB_Name = "frmRHCASCsRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbEmpTable As String
Dim rsRPT As New ADODB.Recordset
Dim fglbFileName
Dim fglbDateTable
Dim fglbDateField
Dim IsFenwick As Boolean
Dim HisSQL  As String
Dim xWFCBilling As Boolean, xMonNum, xLIndex

Private Sub cmbReports_Click()
Dim Sch

If cmbReports = "Attendance Report" Then
    lblAsOf.Visible = True
    dlpAsOf.Visible = True
    lblDates.Visible = True
    dlpDateRange(0).Visible = True
    dlpDateRange(1).Visible = True
    chkVacation.Visible = False
    chkOvertime.Visible = False
    lblRptAuth.Visible = False
    elpRept.Visible = False
ElseIf cmbReports = "Overtime Report" Then
    lblAsOf.Visible = True
    dlpAsOf.Visible = True
    lblDates.Visible = False
    dlpDateRange(0).Visible = False
    dlpDateRange(1).Visible = False
    chkVacation.Visible = False
    chkOvertime.Visible = False
    lblRptAuth.Visible = False
    elpRept.Visible = False
ElseIf cmbReports = "Vacation Report" Then
    lblAsOf.Visible = True
    dlpAsOf.Visible = True
    lblDates.Visible = False
    dlpDateRange(0).Visible = False
    dlpDateRange(1).Visible = False
    chkVacation.Visible = False
    chkOvertime.Visible = False
    lblRptAuth.Visible = False
    elpRept.Visible = False
ElseIf cmbReports = "Request Report" Then
    lblAsOf.Visible = False
    dlpAsOf.Visible = False
    lblDates.Visible = True
    dlpDateRange(0).Visible = True
    dlpDateRange(1).Visible = True
    chkVacation.Visible = True
    chkOvertime.Visible = True
    lblRptAuth.Visible = True
    elpRept.Visible = True
End If
fglbEmpTable = "HREMP"
Me.Caption = "Hamilton C.C.A.S. - " & cmbReports.Text

'If glbSQL Then
'    Sch = Replace(cmbReports, "'", "'+CHAR(39)+'")
'Else
'    Sch = Replace(cmbReports, "'", "'+CHR(39)+'")
'End If
'If rsRPT.State <> 0 Then rsRPT.Close
'rsRPT.Open "SELECT * FROM HR_CUSTOMRPT WHERE RT_RPTNAME='" & Sch & "'", gdbAdoIhr001, adOpenForwardOnly
'
'If Not rsRPT.EOF Then
'    If Not IsNull(rsRPT("RT_DATETABLE")) And Not IsNull(rsRPT("RT_DATEFIELD")) Then
'        If Trim(rsRPT("RT_DATETABLE")) <> "" And Trim(rsRPT("RT_DATEFIELD")) <> "" Then
'            If glbCompSerial = "S/N - 2369W" And InStr(1, fglbFileName, "sn2369AttBonus.rpt") > 0 Then
'                lblFromTo.FontBold = True
'            End If
'            fglbDateTable = Trim(rsRPT("RT_DATETABLE"))
'            fglbDateField = Trim(rsRPT("RT_DATEFIELD"))
'        End If
'    End If
'    If InStr(rsRPT("RT_FILENAME"), ":") = 0 Then
'        fglbFileName = glbIHRREPORTS & rsRPT("RT_FILENAME")
'    Else
'        fglbFileName = rsRPT("RT_FILENAME")
'    End If
'
'    If rsRPT("RT_TERMINATION") Then
'        fglbEmpTable = "TERM_HREMP"
'        elpEEID.LookupType = 1 'TERM
'    Else
'        fglbEmpTable = "HREMP"
'        elpEEID.LookupType = 0 'ACTIVE
'    End If
'    Call INI_Controls(Me)
'    cmdView.Enabled = True
'    cmdPrint.Enabled = True
'End If

'If glbWFC Then
'    xWFCBilling = False
'    If InStr(1, fglbFileName, "Benefit_OptLife_Billing.rpt") > 0 Or InStr(1, fglbFileName, "Benefit_OptLife_Billing_sum.rpt") > 0 Or InStr(1, fglbFileName, "Benefit_cost_details.rpt") > 0 Or InStr(1, fglbFileName, "Benefit_Costs_summary.rpt") > 0 Then
'        xWFCBilling = True
'    End If
'    frmBeneBilling.Visible = xWFCBilling
'End If

''Collectcorp Inc. - Display appropriate selection criteria based on the Report selection.
''Ticket #14437
'If glbCompSerial = "S/N - 2390W" Then
'    lblAsOf.Visible = False
'    dlpAsOf.Visible = False
'    lblDateS.Visible = False
'    dlpDateRange(2).Visible = False
'    dlpDateRange(3).Visible = False
'    lblCode(0).Visible = False
'    lblCode(1).Visible = False
'    clpUser(0).Visible = False
'    clpUser(1).Visible = False
'    lblMonth.Visible = False
'    comDateMonth.Visible = False
'
'    If InStr(1, fglbFileName, "SN2390_Anniversary.rpt") > 0 Then
'        'Show As of Date and DOH
'        lblAsOf.Visible = True
'        dlpAsOf.Visible = True
'        lblAsOf.FontBold = True
'        dlpAsOf.Text = Date
'        lblMonth.Visible = True
'        lblMonth.Caption = lStr("Original Hire Date")
'        comDateMonth.Visible = True
'    ElseIf InStr(1, fglbFileName, "SN2390_Birthday.rpt") > 0 Then
'        'Show Date of Birth
'        lblMonth.Visible = True
'        comDateMonth.Visible = True
'    ElseIf InStr(1, fglbFileName, "SN2390_LicenseAddr.rpt") > 0 Then
'        'Show DOH
'        lblDateS.Visible = True
'        lblDateS.Caption = lStr("Original Hire Date")
'        dlpDateRange(2).Visible = True
'        dlpDateRange(3).Visible = True
'    ElseIf InStr(1, fglbFileName, "SN2390_LicenseTerm.rpt") > 0 Then
'        'Show Termination Date, License Prov/State, License Status
'        lblDateS.Visible = True
'        lblDateS.Caption = "Termination Date"
'        dlpDateRange(2).Visible = True
'        dlpDateRange(3).Visible = True
'        lblCode(0).Visible = True
'        lblCode(0).Caption = lStr("Code 1")
'        lblCode(1).Visible = True
'        lblCode(1).Caption = lStr("Code 2")
'        clpUser(0).Visible = True
'        clpUser(1).Visible = True
'    ElseIf InStr(1, fglbFileName, "SN2390_LicenseAdditions.rpt") > 0 Then
'        'Show Date Submitted, License Prov/State, License Status
'        lblDateS.Visible = True
'        lblDateS.Caption = lStr("Date 1")
'        dlpDateRange(2).Visible = True
'        dlpDateRange(3).Visible = True
'        lblCode(0).Visible = True
'        lblCode(0).Caption = lStr("Code 1")
'        lblCode(1).Visible = True
'        lblCode(1).Caption = lStr("Code 2")
'        clpUser(0).Visible = True
'        clpUser(1).Visible = True
'    End If
'Else
'    lblAsOf.Visible = False
'    dlpAsOf.Visible = False
'    lblDateS.Visible = False
'    dlpDateRange(2).Visible = False
'    dlpDateRange(3).Visible = False
'    lblCode(0).Visible = False
'    lblCode(1).Visible = False
'    clpUser(0).Visible = False
'    clpUser(1).Visible = False
'    lblMonth.Visible = False
'    comDateMonth.Visible = False
'End If

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm(frmRHCASCsRpt.Caption & " Criteria", Me) Then Exit Sub
    Call set_PrintState(False)
    x% = Cri_SetAll()
    'Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    'Me.vbxCrystal.Action = 1
    'vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

PrntErr:
MsgBox "Error Printing - check your Windows Printer setup"


Resume Next

End Sub

Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
    Call set_PrintState(False)
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    x% = Cri_SetAll()
    'Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    'Me.vbxCrystal.Action = 1
    'vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", fglbEmpTable, "SELECT")
Resume Next

End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%)) > 0 Then
    Select Case intIdx%
    Case 0: strCd$ = fglbEmpTable & ".ED_LOC"
    Case 1: strCd$ = fglbEmpTable & ".ED_ORG"
    Case 2: strCd$ = fglbEmpTable & ".ED_EMP"
    Case 3: strCd$ = fglbEmpTable & ".ED_REGION"
    Case 4: strCd$ = fglbEmpTable & ".ED_ADMINBY"
    Case 5: strCd$ = fglbEmpTable & ".ED_SECTION"  'Lucy June 29, 2000
    Case 6: strCd$ = "HRJOB.JB_GRPCD" 'Fenwick only
    End Select
    CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    If glbLinamar And (strCd$ = fglbEmpTable & ".ED_REGION" Or strCd$ = fglbEmpTable & ".ED_SECTION") Then
        CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%) & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%) & "') )"
    End If
End If

If Len(CodeCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CodeCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CodeCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_Div()

Dim DivCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level


If Len(clpDiv.Text) > 0 Then
    DivCri = "({" & fglbEmpTable & ".ED_DIV} in  ['" & Replace(clpDiv.Text, ",", "','") & "'])"
End If

If Len(DivCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = DivCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & DivCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_EE()
Dim EECri As String

If Len(elpEEID.Text) > 0 Then
    EECri = "{" & fglbEmpTable & ".ED_EMPNBR} IN [" & getEmpnbr(elpEEID.Text) & "] "
End If

If Len(EECri) >= 1 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_Shift()
Dim EECri As String, OneSet%, x%

If Len(txtShift.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Sub Cri_PT()
Dim EECri As String, OneSet%, x%

If Len(clpPT.Text) < 1 Then Exit Sub
EECri = "{" & fglbEmpTable & ".ED_PT}= '" & clpPT.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_CountryOfEmployment()
Dim CountryCri As String

If Len(comCountryOfEmp.Text) > 0 Then
    CountryCri = "({HREMP.ED_WORKCOUNTRY} = '" & comCountryOfEmp.Text & "')"
End If

If Len(CountryCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CountryCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CountryCri
    End If
    glbiOneWhere = True
End If
End Sub

Private Sub Cri_RepAuth()
Dim EECri As String

If Len(elpRept.Text) > 0 Then
    EECri = "{JH_REPTAU}= " & Trim(elpRept.Text) & " "
   
End If

If Len(EECri) >= 1 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
End If
End Sub

Private Function Cri_SetAll()
Dim x%, strRName$

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

Call glbCri_DeptUN(clpDept.Text)
glbstrSelCri = Replace(glbstrSelCri, "HREMP.", fglbEmpTable & ".")


Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_PT
Call Cri_Shift
Call Cri_EE
Call Cri_CountryOfEmployment
For x% = 0 To 5
    Call Cri_Code(x%)
Next x%

'If IsFenwick Then 'Ticket #12505
'    Call Cri_Code(6)
'End If

'If glbCompSerial = "S/N - 2390W" Then       'Collectcorp Inc. Ticket #14437
'    If dlpDateRange(2).Visible Or dlpDateRange(3).Visible Then
'        Call Cri_Dates
'    End If
'    If clpUser(0).Visible Or clpUser(1).Visible Then
'        Call Cri_License_Codes(0)
'        Call Cri_License_Codes(1)
'    End If
'    If comDateMonth.Visible Then
'        If InStr(1, fglbFileName, "SN2390_Anniversary.rpt") > 0 Then
'            Call Cri_DOHMonth
'        ElseIf InStr(1, fglbFileName, "SN2390_Birthday.rpt") > 0 Then
'            Call Cri_BirthMonth
'        End If
'    End If
'End If

'If glbWFC Then 'Ticket #13957
'    If InStr(1, fglbFileName, "Benefit_OptLife_Billing.rpt") > 0 Then
'        Call WFCOptLifeBilling
'        'Exit Function
'    End If
'    If InStr(1, fglbFileName, "Annual_Manulife_Benefits.rpt") > 0 Then
'        Call WFC_Annual_Benefits
'    End If
'    If xWFCBilling Then 'Ticket #14061
'        Me.vbxCrystal.Formulas(10) = "BillYear=" & txtFiscal.Text & ""
'        Me.vbxCrystal.Formulas(11) = "BillMonth=" & ComMTH.ListIndex + 1 & ""
'    End If
'End If

'If (glbCompSerial <> "S/N - 2257W" And InStr(1, fglbFileName, "Request Report.rpt") = 0) Then
'    If frmDate.Visible Then Cri_FTDates
'End If

'Testing Excel reports
If cmbReports = "Attendance Report" Then
    Call Attendance_Report_XLS_HCAS
ElseIf cmbReports = "Overtime Report" Then
    Call Overtime_Report_XLS_HCAS
ElseIf cmbReports = "Vacation Report" Then
    Call Vacation_Report_XLS_HCAS
ElseIf cmbReports = "Request Report" Then
    Call Request_Report_XLS_HCAS
End If

'Exit Function

'Me.vbxCrystal.ReportFileName = fglbFileName

'If Len(glbstrSelCri) >= 0 Then
'    If glbWFC Then
'        If InStr(1, fglbFileName, "Benefit_OptLife_Billing.rpt") > 0 Then 'Ticket #13957
'            glbstrSelCri = "{WFC_MANULIFE_BENE_WRK.WRKEMP}='" & glbUserID & "'"
'        End If
'        'Ticket #14031
'        If InStr(1, fglbFileName, "Benefit_cost_details.rpt") > 0 Or InStr(1, fglbFileName, "Benefit_OptLife_Billing.rpt") > 0 Then
'            If glbNoNONE And glbNoEXEC Then
'                glbstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG }<> 'NONE' AND {HREMP.ED_ORG }<> 'EXEC'))"    'Hemu -EXE
'            ElseIf glbNoNONE Then
'                glbstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'NONE')"
'            ElseIf glbNoEXEC Then
'                glbstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'EXEC')"
'            End If
'        End If
'    End If
'
'    Me.vbxCrystal.SelectionFormula = glbstrSelCri
'
'    If glbCompSerial = "S/N - 2257W" And InStr(1, fglbFileName, "Request Report.rpt") > 0 Then
'        Call HCAS_Request_Report
'        Me.vbxCrystal.SelectionFormula = "{HR_REQUEST_RPT.REQ_WRKEMP}='" & glbUserID & "'"
'    End If
'End If

'If glbSQL Or glbOracle Then
'    Me.vbxCrystal.Connect = RptODBC_SQL
'    If glbCompSerial = "S/N - 2288W" And InStr(1, fglbFileName, "SN2288_1.rpt") > 0 Then
'        Me.vbxCrystal.SubreportToChange = "counselcount"
'        Me.vbxCrystal.Connect = RptODBC_SQL
'        Me.vbxCrystal.SubreportToChange = ""
'    End If
'        'If glbWFC And InStr(1, fglbFileName, "mzINCIDENTWSUB.rpt") > 0 Then
'        '    Me.vbxCrystal.SubreportToChange = "Root Causes"
'        '    Me.vbxCrystal.Connect = RptODBC_SQL
'        '    Me.vbxCrystal.SubreportToChange = "Corrective Action"
'        '    Me.vbxCrystal.Connect = RptODBC_SQL
'        '    Me.vbxCrystal.SubreportToChange = "Term Root Causes"
'        '    Me.vbxCrystal.Connect = RptODBC_SQL
'        '    Me.vbxCrystal.SubreportToChange = "Term Corrective Action"
'        '    Me.vbxCrystal.Connect = RptODBC_SQL
'        '    Me.vbxCrystal.SubreportToChange = ""
'        'End If
'
'    If glbCompSerial = "S/N - 2385W" And InStr(1, fglbFileName, "SN2385_1.rpt") > 0 Then   'Conservation Halton
'        Me.vbxCrystal.Formulas(1) = "lblPT='" & lStr("Category") & "'"
'        Me.vbxCrystal.Formulas(2) = "lblUnion='" & lStr("Union") & "'"
'    End If
'
'    If glbCompSerial = "S/N - 2390W" Then       'Collectcorp Inc. Ticket #14437
'        If InStr(1, fglbFileName, "SN2390_Anniversary.rpt") > 0 Then
'            Me.vbxCrystal.Formulas(1) = "lblDept='" & lStr("Department") & "'"
'            Me.vbxCrystal.Formulas(2) = "lblHireDate='" & lStr("Original Hire Date") & "'"
'            Me.vbxCrystal.Formulas(3) = "AsOfDate=Date('" & Format(dlpAsOf.Text, "mm/dd/yyyy") & "')"
'            Me.vbxCrystal.Formulas(4) = "DOHMonth='" & lStr("Department") & "'"
'        ElseIf InStr(1, fglbFileName, "SN2390_Birthday.rpt") > 0 Then
'            Me.vbxCrystal.Formulas(1) = "lblDept='" & lStr("Department") & "'"
'        ElseIf InStr(1, fglbFileName, "SN2390_CommPayroll.rpt") > 0 Then
'            Me.vbxCrystal.Formulas(1) = "lblDept='" & lStr("Department") & "'"
'            Me.vbxCrystal.Formulas(2) = "lblHireDate='" & lStr("Original Hire Date") & "'"
'        ElseIf InStr(1, fglbFileName, "SN2390_LicenseAddr.rpt") > 0 Then
'            Me.vbxCrystal.Formulas(1) = "lblDept='" & lStr("Department") & "'"
'            Me.vbxCrystal.Formulas(2) = "lblHireDate='" & lStr("Original Hire Date") & "'"
'            Me.vbxCrystal.Formulas(3) = "lblLocation='" & lStr("Location") & "'"
'        ElseIf InStr(1, fglbFileName, "SN2390_LicenseTerm.rpt") > 0 Then
'            Me.vbxCrystal.Formulas(1) = "lblDept='" & lStr("Department") & "'"
'            Me.vbxCrystal.Formulas(2) = "lblLicNumber='" & lStr("UText 1") & "'"
'        ElseIf InStr(1, fglbFileName, "SN2390_LicenseAdditions.rpt") > 0 Then
'            Me.vbxCrystal.Formulas(1) = "lblDept='" & lStr("Department") & "'"
'            Me.vbxCrystal.Formulas(2) = "lblHireDate='" & lStr("Original Hire Date") & "'"
'            Me.vbxCrystal.Formulas(3) = "lblLicStatus='" & lStr("Code 2") & "'"
'            Me.vbxCrystal.Formulas(4) = "lblLicProvState='" & lStr("Code 1") & "'"
'        End If
'    End If
'Else
'    Me.vbxCrystal.Connect = "PWD=petman;"
'
'
'    If glbCompSerial = "S/N - 2211W" And InStr(1, fglbFileName, "Vac and sick ent.rpt") > 0 Then
'        Me.vbxCrystal.SubreportToChange = "HRFDTaken"
'        Me.vbxCrystal.Connect = "PWD=petman;"
'        Me.vbxCrystal.SubreportToChange = ""
'    End If
'    If glbCompSerial = "S/N - 2211W" And InStr(1, fglbFileName, "Vac and sick ent.rpt") > 0 Then
'        Me.vbxCrystal.SubreportToChange = "OTHours"
'        Me.vbxCrystal.Connect = "PWD=petman;"
'        Me.vbxCrystal.SubreportToChange = ""
'    End If
'    If glbCompSerial = "S/N - 2211W" And InStr(1, fglbFileName, "Vac and sick ent.rpt") > 0 Then
'        Me.vbxCrystal.SubreportToChange = "VacationTimeUsed"
'        Me.vbxCrystal.Connect = "PWD=petman;"
'        Me.vbxCrystal.SubreportToChange = ""
'    End If
'    If glbCompSerial = "S/N - 2211W" And InStr(1, fglbFileName, "Vac and sick ent.rpt") > 0 Then
'        Me.vbxCrystal.SubreportToChange = "VacBooked"
'        Me.vbxCrystal.Connect = "PWD=petman;"
'        Me.vbxCrystal.SubreportToChange = ""
'    End If
'    If glbCompSerial = "S/N - 2211W" And InStr(1, fglbFileName, "Vac and sick ent.rpt") > 0 Then
'        Me.vbxCrystal.SubreportToChange = "SickTimeUsed"
'        Me.vbxCrystal.Connect = "PWD=petman;"
'        Me.vbxCrystal.SubreportToChange = ""
'    End If
'
'    If glbCompSerial = "S/N - 2330W" Then   'Town of Marathon
'        Me.vbxCrystal.Formulas(1) = "lblTypeVehicle='" & lStr("Type of Vehicle") & "'"
'    End If
'End If
'
'Me.vbxCrystal.WindowTitle = cmbReports


Cri_SetAll = True



Screen.MousePointer = DEFAULT
Exit Function


modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "Comp Report", "Select")
Cri_SetAll = False
Resume Next

End Function

Private Function CriCheck()
Dim x%

CriCheck = False

'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
'    MsgBox lStr("If Division Entered - it must be known")
'     clpDiv.SetFocus
'    Exit Function
'End If

'If glbCompSerial = "S/N - 2347W" And InStr(1, fglbFileName, "SN234718.rpt") > 0 Then
'    If Len(Trim(clpDiv.Text)) = 0 Then
'        MsgBox lStr("Division cannot be left blank")
'        clpDiv.SetFocus
'        Exit Function
'    End If
'End If

'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
'    MsgBox "If Department Entered - it must be known"
'     clpDept.SetFocus
'    Exit Function
'End If

For x% = 0 To 5
    If Not clpCode(x).ListChecker Then Exit Function
Next x%

For x% = 0 To 1
 If Len(dlpDateRange(x%).Text) > 0 Then
    If Not IsDate(dlpDateRange(x%).Text) Then
        MsgBox "Not a valid date"
        dlpDateRange(x%).Text = ""
        dlpDateRange(x%).SetFocus
        Exit Function
    End If
 End If
Next x%


If Len(clpPT.Text) > 0 Then
    If Len(clpPT) > 0 And clpPT.Caption = "Unassigned" Then
        MsgBox lStr("Category code must be valid")
        clpPT.SetFocus
        Exit Function
    End If
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

'If xWFCBilling Then
'    If Len(txtFiscal) < 1 Then
'        MsgBox "Year is a required field"
'        txtFiscal.SetFocus
'        Exit Function
'    Else
'        If Val(txtFiscal) < 2000 Then
'            MsgBox "Year must be greater than 2000"
'            txtFiscal.SetFocus
'            Exit Function
'        End If
'    End If
'End If

''Collectcorp Inc. - Ticket #14437
'If glbCompSerial = "S/N - 2390W" Then
'    If InStr(1, fglbFileName, "SN2390_Anniversary.rpt") > 0 Then
'        'As of Date
'        If Len(dlpAsOf.Text) > 0 Then
'            If Not IsDate(dlpAsOf.Text) Then
'                MsgBox "Not a valid As Of Date"
'                dlpAsOf.SetFocus
'                Exit Function
'            End If
'        Else
'            MsgBox "As of Date is a required field"
'            dlpAsOf.SetFocus
'            Exit Function
'        End If
'    ElseIf InStr(1, fglbFileName, "SN2390_LicenseAddr.rpt") > 0 Then
'        'Date of Hire
'        For X% = 2 To 3
'            If Len(dlpDateRange(X%).Text) > 0 Then
'                If Not IsDate(dlpDateRange(X%).Text) Then
'                    MsgBox "Not a valid " & lStr("Original Hire Date")
'                    dlpDateRange(X%).SetFocus
'                    Exit Function
'                End If
'            End If
'        Next X%
'        If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
'            If dlpDateRange(3).Text < dlpDateRange(2).Text Then
'                MsgBox "Invalid " & lStr("Original Hire Date") & " range. To " & lStr("Original Hire Date") & " cannot be less than From Date"
'                dlpDateRange(3).SetFocus
'                Exit Function
'            End If
'        End If
'
'    ElseIf InStr(1, fglbFileName, "SN2390_LicenseTerm.rpt") > 0 Then
'        'Termination Date, License Prov/State, License Status
'        For X% = 2 To 3
'            If Len(dlpDateRange(X%).Text) > 0 Then
'                If Not IsDate(dlpDateRange(X%).Text) Then
'                    MsgBox "Not a valid Termination Date"
'                    dlpDateRange(X%).SetFocus
'                    Exit Function
'                End If
'            End If
'        Next X%
'
'        If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
'            If dlpDateRange(3).Text < dlpDateRange(2).Text Then
'                MsgBox "Invalid Termination Date range. To Termination Date cannot be less than From Date"
'                dlpDateRange(3).SetFocus
'                Exit Function
'            End If
'        End If
'
'        For X% = 0 To 1
'            If Not clpUser(X).ListChecker Then Exit Function
'        Next X%
'
'    ElseIf InStr(1, fglbFileName, "SN2390_LicenseAdditions.rpt") > 0 Then
'        'Date Submitted, License Prov/State, License Status
'        For X% = 2 To 3
'            If Len(dlpDateRange(X%).Text) > 0 Then
'                If Not IsDate(dlpDateRange(X%).Text) Then
'                    MsgBox "Not a valid " & lStr("Date 1")
'                    dlpDateRange(X%).SetFocus
'                    Exit Function
'                End If
'            End If
'        Next X%
'        If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
'            If dlpDateRange(3).Text < dlpDateRange(2).Text Then
'                MsgBox "Invalid " & lStr("Date 1") & " range. To" & lStr("Date 1") & " cannot be less than From Date"
'                dlpDateRange(3).SetFocus
'                Exit Function
'            End If
'        End If
'
'        For X% = 0 To 1
'            If Not clpUser(X).ListChecker Then Exit Function
'        Next X%
'    End If
'
'End If

CriCheck = True
End Function

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = "FRMRHCASCSRPT"

Dim rsALLRPT As New ADODB.Recordset

Dim SQLQ
Screen.MousePointer = HOURGLASS
'If glbOracle Then
'    SQLQ = "SELECT * FROM HR_CUSTOMRPT, HR_SECRPT"
'    SQLQ = SQLQ & " WHERE HR_CUSTOMRPT.RT_RPTNAME(+)= HR_SECRPT.FUNCTION "
'    SQLQ = SQLQ & " AND USERID='" & glbUserID & "'"
'    SQLQ = SQLQ & " AND ACCESSABLE<>0 "
'    SQLQ = SQLQ & " ORDER BY UPPER(FUNCTION)"
'Else
'    SQLQ = "SELECT * FROM HR_CUSTOMRPT LEFT JOIN HR_SECRPT"
'    SQLQ = SQLQ & " ON HR_CUSTOMRPT.RT_RPTNAME= HR_SECRPT.[FUNCTION] "
'    SQLQ = SQLQ & " WHERE USERID='" & glbUserID & "'"
'    SQLQ = SQLQ & " AND ACCESSABLE<>0 "
'    SQLQ = SQLQ & " ORDER BY [FUNCTION]"
'End If
'rsALLRPT.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

cmbReports.Clear
cmbReports.AddItem "Attendance Report"
cmbReports.AddItem "Overtime Report"
cmbReports.AddItem "Vacation Report"
cmbReports.AddItem "Request Report"

'Do Until rsALLRPT.EOF
'    cmbReports.AddItem rsALLRPT("RT_RPTNAME")
'    rsALLRPT.MoveNext
'Loop
'rsALLRPT.Close

If cmbReports.ListCount <> 0 Then
    cmbReports.ListIndex = 0
End If

'If glbCompSerial = "S/N - 2369W" And InStr(1, fglbFileName, "sn2369AttBonus.rpt") > 0 Then
'    lblFromTo.FontBold = True
'End If
            
Call setRptCaption(Me)
'If glbLinamar Then clpCode(3).MaxLength = 8

Call addCountryItems 'Frank 09/07/2007 Ticket #13621

Call INI_Controls(Me)

'If glbWFC Then 'Ticket #14061
'    frmBeneBilling.Top = 5400
'    ComMTH.AddItem "Jan"
'    ComMTH.AddItem "Feb"
'    ComMTH.AddItem "Mar"
'    ComMTH.AddItem "Apr"
'    ComMTH.AddItem "May"
'    ComMTH.AddItem "Jun"
'    ComMTH.AddItem "Jul"
'    ComMTH.AddItem "Aug"
'    ComMTH.AddItem "Sep"
'    ComMTH.AddItem "Oct"
'    ComMTH.AddItem "Nov"
'    ComMTH.AddItem "Dec"
'    txtFiscal.Text = Year(Date)
'    xMonNum = Month(Date)
'    ComMTH.ListIndex = xMonNum - 1
'End If
'If glbCompSerial = "S/N - 2390W" Then       'Collectcorp Inc. Ticket #14437
'    comDateMonth.AddItem "January"
'    comDateMonth.AddItem "February"
'    comDateMonth.AddItem "March"
'    comDateMonth.AddItem "April"
'    comDateMonth.AddItem "May"
'    comDateMonth.AddItem "June"
'    comDateMonth.AddItem "July"
'    comDateMonth.AddItem "August"
'    comDateMonth.AddItem "September"
'    comDateMonth.AddItem "October"
'    comDateMonth.AddItem "November"
'    comDateMonth.AddItem "December"
'    comDateMonth.ListIndex = Month(Date) - 1
'End If

Screen.MousePointer = DEFAULT

'Hemu - Add Serial # Control
'lblAsOf.Visible = True
'dlpAsOf.Visible = True
'lblAsOf.FontBold = True
'dlpAsOf.Text = Date


End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub


Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%
If glbOttawaCCAC And InStr(fglbFileName, "Uptodate_Entitlement") <> 0 Then
    For x% = 0 To 1
        If Len(dlpDateRange(x).Text) > 0 Then
            TempCri = "({" & fglbDateTable & "." & fglbDateField & "} "
            If x% = 0 Then
                TempCri = "FromDate="
            Else
                TempCri = "ToDate="
            End If
            dtYYY% = Year(dlpDateRange(x).Text)
            dtMM% = month(dlpDateRange(x).Text)
            dtDD% = Day(dlpDateRange(x).Text)
            Me.vbxCrystal.Formulas(100 + x%) = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        End If
    Next x%
   Exit Sub
Else
    
    If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
        TempCri = "({" & fglbDateTable & "." & fglbDateField & "} "
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
        TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
        If glbCompSerial = "S/N - 2369W" And (InStr(1, fglbFileName, "sn2369AttBonus.rpt") > 0 Or InStr(1, fglbFileName, "sn2369CBonus.rpt") > 0) Then
            Me.vbxCrystal.Formulas(0) = "dteStart= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
        End If
        dtYYY% = Year(dlpDateRange(1).Text)
        dtMM% = month(dlpDateRange(1).Text)
        dtDD% = Day(dlpDateRange(1).Text)
        TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        If glbCompSerial = "S/N - 2369W" And (InStr(1, fglbFileName, "sn2369AttBonus.rpt") > 0 Or InStr(1, fglbFileName, "sn2369CBonus.rpt") > 0) Then
            Me.vbxCrystal.Formulas(1) = "dteEnd= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
        End If
        GoTo Cri_FTDatst
    End If
    
    For x% = 0 To 1
        If Len(dlpDateRange(x).Text) > 0 Then
            TempCri = "({" & fglbDateTable & "." & fglbDateField & "} "
            If x% = 0 Then
                TempCri = TempCri & " >= "
            Else
                TempCri = TempCri & " <= "
            End If
            dtYYY% = Year(dlpDateRange(x).Text)
            dtMM% = month(dlpDateRange(x).Text)
            dtDD% = Day(dlpDateRange(x).Text)
            TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
            GoTo Cri_FTDatst
        End If
    Next x%
    
Cri_FTDatst:
    If Len(TempCri) >= 1 Then
        If Not glbiOneWhere Then
            glbstrSelCri = TempCri
        Else
            glbstrSelCri = glbstrSelCri & " AND " & TempCri
        End If
        glbiOneWhere = True
    End If
End If
End Sub

Private Sub txtShift_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

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

If cmbReports.Text = "" Then
Printable = False
Else
Printable = True
End If

End Property

Public Sub SET_UP_MODE()
Call set_Buttons
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub addCountryItems()
Dim ctylist, x

ctylist = CountryList
x = 1
Do While x > 0
    x = InStr(ctylist, "&")
    If x > 0 Then
        comCountryOfEmp.AddItem Left(ctylist, x - 1)
        ctylist = Mid(ctylist, x + 1)
    Else
        comCountryOfEmp.AddItem ctylist
    End If
Loop

End Sub

Private Function CountryList() As String
Dim xCountryList As String, ctyFile
xCountryList = ""
ctyFile = glbIHRREPORTS & "CountryList.MTF"

On Error GoTo ErrorHandler

If File(ctyFile) Then
    Open ctyFile For Input As #1
    Input #1, xCountryList
    Close #1
End If

ResumeHere:
'If InStr(xCountryList, BasicCountry) = 0 Then
'    xCountryList = BasicCountry
'End If
If InStr(xCountryList, comCountryOfEmp) = 0 And comCountryOfEmp <> "" Then
    xCountryList = xCountryList & "&" & comCountryOfEmp
    comCountryOfEmp.AddItem comCountryOfEmp
End If
Open ctyFile For Output As #1
Print #1, xCountryList
Close #1
CountryList = xCountryList
Exit Function

ErrorHandler:
If Err.Number = 62 Then
    ' Corrupted CountryList.MTF, kill it and regenerate
    Close #1
    MsgBox "Found corrupt CountryList.MTF.  info:HR will re-create this file.", vbInformation + vbOKOnly, "Corrupted Country List"
    Kill ctyFile
    Resume ResumeHere
Else
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number & " in CountryList"
    Resume Next
End If
End Function

'Private Sub WFC_Annual_Benefits()
'Dim rsBeneOpt As New ADODB.Recordset
'Dim rsWRK As New ADODB.Recordset
'Dim SQLQ As String, xEMPNBR, xCode
'Dim I, totNum
'Dim xPolicyNo, xOptAccount
'Dim rsDep As New ADODB.Recordset
'Dim xBillingDate
'Dim BenefitCodeList As String
'
'    Screen.MousePointer = HOURGLASS
'    MDIMain.panHelp(0).FloodType = 1
'    MDIMain.panHelp(1).Caption = " Please Wait"
'    MDIMain.panHelp(2).Caption = ""
'    MDIMain.panHelp(0).FloodPercent = 0
'    BenefitCodeList = " ('LIF', 'LIF1', 'LIF2', 'LIF3', 'LFR', 'LFR1', 'LFR2', 'LFR3','OPLF','OPLS','LTD','AD&D','EHC','DENT')"
'    xBillingDate = CVDate(ComMTH.Text & " 1, " & txtFiscal)
'    gdbAdoIhr001W.BeginTrans
'    gdbAdoIhr001W.Execute "DELETE FROM WFC_REPORT_WRK WHERE WRKEMP='" & glbUserID & "'"
'    gdbAdoIhr001W.CommitTrans
'
'    HisSQL = " BF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
'    SQLQ = "SELECT * FROM HRBENFT WHERE " & HisSQL & " "
'    SQLQ = SQLQ & " AND (BF_EMPNBR in (select ED_EMPNBR from HREMP WHERE ED_COUNTRY = 'CANADA' AND NOT (ED_USER_TEXT1 IS NULL OR ED_USER_TEXT1 = '' ) AND NOT (ED_USER_TEXT2 IS NULL OR ED_USER_TEXT2 = '') AND NOT (ED_USER_NUM1 IS NULL) ))"
'    'SQLQ = SQLQ & "AND (BF_BCODE = 'OPLF' OR BF_BCODE = 'OPLS' OR BF_BCODE = 'OPLC') "
'    SQLQ = SQLQ & "AND (BF_BCODE IN " & BenefitCodeList & " )"
'    'SQLQ = SQLQ & "AND NOT (BF_COVER = 'W' )"
'    'SQLQ = SQLQ & "AND (BF_AMT > 0 )"
'    'SQLQ = SQLQ & "AND NOT (BF_POLICY IS NULL OR BF_POLICY = '' ) "
'    SQLQ = SQLQ & "ORDER BY BF_EMPNBR, BF_BCODE, BF_EDATE "
'    rsBeneOpt.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    If Not rsBeneOpt.EOF Then
'       totNum = rsBeneOpt.RecordCount: I = 0
'    End If
'    Do While Not rsBeneOpt.EOF
'        If (I / totNum) <= 1 Then
'            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
'            I = I + 1
'        End If
'        DoEvents
'        'If Not IsNull(rsBeneOpt("BF_COVER")) Then
'        '    If UCase(rsBeneOpt("BF_COVER")) = "W" Then
'        '        GoTo NexeRec
'        '    End If
'        'End If
'        'If Not IsNull(rsBeneOpt("BF_CEASEDATE")) Then
'        '    If IsDate(rsBeneOpt("BF_CEASEDATE")) Then
'        '        If CVDate(xBillingDate) >= CVDate(rsBeneOpt("BF_CEASEDATE")) Then
'        '            GoTo NexeRec
'        '        End If
'        '    End If
'        'End If
'
'        xEMPNBR = rsBeneOpt("BF_EMPNBR")
'        xCode = rsBeneOpt("BF_BCODE")
'        xPolicyNo = rsBeneOpt("BF_POLICY")
'        'xOptAccount = ""
'        'If Len(xPolicyNo) <> 9 Then
'        '    GoTo NexeRec 'Invalid Policy number format, it's "#####-###"
'        'Else
'        '    xOptAccount = Mid(xPolicyNo, 7, 3)
'        'End If
'        SQLQ = "SELECT * FROM WFC_REPORT_WRK WHERE WRKEMP='" & glbUserID & "'"
'        SQLQ = SQLQ & "AND R_EMPNBR = " & xEMPNBR & " "
'        If rsWRK.State <> 0 Then rsWRK.Close
'        rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'        If rsWRK.EOF Then
'            rsWRK.AddNew
'            rsWRK("R_COMPNO") = "001"
'            rsWRK("R_EMPNBR") = xEMPNBR
'            rsWRK("WRKEMP") = glbUserID
'            rsWRK("R_NUM2") = 0
'        End If
'        If xCode = "LIF" Or xCode = "LIF1" Or xCode = "LIF2" Or xCode = "LIF3" Or xCode = "LFR" Or xCode = "LFR" Or xCode = "LFR1" Or xCode = "LFR2" Or xCode = "LIF3" Then
'            If Not IsNull(rsBeneOpt("BF_AMT")) Then
'                rsWRK("R_NUM2") = rsWRK("R_NUM2") + rsBeneOpt("BF_AMT")
'            End If
'        End If
'        If xCode = "OPLF" Then
'            If IsNull(rsBeneOpt("BF_AMT")) Then
'                rsWRK("R_NUM3") = 0
'            Else
'                rsWRK("R_NUM3") = rsBeneOpt("BF_AMT")
'            End If
'        End If
'        If xCode = "AD&D" Then
'            If IsNull(rsBeneOpt("BF_AMT")) Then
'                rsWRK("R_NUM4") = 0
'            Else
'                rsWRK("R_NUM4") = rsBeneOpt("BF_AMT")
'            End If
'        End If
'        If xCode = "STD" Then
'            If IsNull(rsBeneOpt("BF_AMT")) Then
'                rsWRK("R_NUM5") = 0
'            Else
'                rsWRK("R_NUM5") = rsBeneOpt("BF_AMT")
'            End If
'        End If
'        If xCode = "LTD" Then
'            If IsNull(rsBeneOpt("BF_AMT")) Then
'                rsWRK("R_NUM6") = 0
'            Else
'                rsWRK("R_NUM6") = rsBeneOpt("BF_AMT")
'            End If
'        End If
'        If xCode = "OPLS" Then
'            If IsNull(rsBeneOpt("BF_AMT")) Then
'                rsWRK("R_NUM7") = 0
'            Else
'                rsWRK("R_NUM7") = rsBeneOpt("BF_AMT")
'            End If
'        End If
'        If xCode = "EHC" Then
'            If Not IsNull(rsBeneOpt("BF_COVER")) Then
'                rsWRK("R_TEXT1") = rsBeneOpt("BF_COVER")
'            End If
'        End If
'        If xCode = "DENT" Then
'            If Not IsNull(rsBeneOpt("BF_COVER")) Then
'                rsWRK("R_TEXT2") = rsBeneOpt("BF_COVER")
'            End If
'        End If
'        If Not IsNull(rsBeneOpt("BF_GROUP")) Then
'            rsWRK("R_TEXT3") = rsBeneOpt("BF_GROUP")
'        End If
'
'        rsWRK.Update
'NexeRec:
'        rsBeneOpt.MoveNext
'    Loop
'    rsBeneOpt.Close
'    MDIMain.panHelp(0).FloodType = 0
'    MDIMain.panHelp(1).Caption = " "
'    Screen.MousePointer = DEFAULT
'End Sub
'
'Private Sub WFCOptLifeBilling()
'Dim rsBeneOpt As New ADODB.Recordset
'Dim rsWRK As New ADODB.Recordset
'Dim SQLQ As String, xEMPNBR, xCode
'Dim I, totNum
'Dim xPolicyNo, xOptAccount
'Dim rsDep As New ADODB.Recordset
'Dim xBillingDate
'    Screen.MousePointer = HOURGLASS
'    MDIMain.panHelp(0).FloodType = 1
'    MDIMain.panHelp(1).Caption = " Please Wait"
'    MDIMain.panHelp(2).Caption = ""
'    MDIMain.panHelp(0).FloodPercent = 0
'
'    xBillingDate = CVDate(ComMTH.Text & " 1, " & txtFiscal)
'    gdbAdoIhr001W.BeginTrans
'    gdbAdoIhr001W.Execute "DELETE FROM WFC_MANULIFE_BENE_WRK WHERE WRKEMP='" & glbUserID & "'"
'    gdbAdoIhr001W.CommitTrans
'
'    HisSQL = " BF_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
'    SQLQ = "SELECT * FROM HRBENFT WHERE " & HisSQL & " "
'    SQLQ = SQLQ & " AND (BF_EMPNBR in (select ED_EMPNBR from HREMP WHERE ED_COUNTRY = 'CANADA' AND NOT (ED_USER_TEXT1 IS NULL OR ED_USER_TEXT1 = '' ) AND NOT (ED_USER_TEXT2 IS NULL OR ED_USER_TEXT2 = '') AND NOT (ED_USER_NUM1 IS NULL) ))"
'    SQLQ = SQLQ & "AND (BF_BCODE = 'OPLF' OR BF_BCODE = 'OPLS' OR BF_BCODE = 'OPLC') "
'    SQLQ = SQLQ & "AND NOT (BF_POLICY IS NULL OR BF_POLICY = '' ) "
'    SQLQ = SQLQ & "ORDER BY BF_EMPNBR, BF_BCODE, BF_EDATE "
'    rsBeneOpt.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    If Not rsBeneOpt.EOF Then
'       totNum = rsBeneOpt.RecordCount: I = 0
'    End If
'    Do While Not rsBeneOpt.EOF
'        If (I / totNum) <= 1 Then
'            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
'            I = I + 1
'        End If
'        DoEvents
'        If Not IsNull(rsBeneOpt("BF_COVER")) Then
'            If UCase(rsBeneOpt("BF_COVER")) = "W" Then
'                GoTo NexeRec
'            End If
'        End If
'        If Not IsNull(rsBeneOpt("BF_CEASEDATE")) Then
'            If IsDate(rsBeneOpt("BF_CEASEDATE")) Then
'                If CVDate(xBillingDate) >= CVDate(rsBeneOpt("BF_CEASEDATE")) Then
'                    GoTo NexeRec
'                End If
'            End If
'        End If
'
'        xEMPNBR = rsBeneOpt("BF_EMPNBR")
'        xCode = rsBeneOpt("BF_BCODE")
'        xPolicyNo = rsBeneOpt("BF_POLICY")
'        xOptAccount = ""
'        If Len(xPolicyNo) <> 9 Then
'            GoTo NexeRec 'Invalid Policy number format, it's "#####-###"
'        Else
'            xOptAccount = Mid(xPolicyNo, 7, 3)
'        End If
'        SQLQ = "SELECT * FROM WFC_MANULIFE_BENE_WRK WHERE WRKEMP='" & glbUserID & "'"
'        SQLQ = SQLQ & "AND WB_EMPNBR = " & xEMPNBR & " "
'        'SQLQ = SQLQ & "AND WB_TEXT1 = '" & xOptAccount & "' "
'        If rsWRK.State <> 0 Then rsWRK.Close
'        rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'        If rsWRK.EOF Then
'            rsWRK.AddNew
'            rsWRK("WB_COMPNO") = "001"
'            rsWRK("WB_EMPNBR") = xEMPNBR
'            rsWRK("WB_LUSER") = glbUserID
'            rsWRK("WB_LDATE") = Date
'            rsWRK("WB_LTIME") = Time$
'            rsWRK("WRKEMP") = glbUserID
'            rsWRK("WB_TEXT1") = xOptAccount
'        End If
'        If xCode = "OPLF" Then
'            rsWRK("WB_BCODE1") = xCode
'            If IsNull(rsBeneOpt("BF_AMT")) Then
'                rsWRK("WB_AMT1") = 0
'            Else
'                rsWRK("WB_AMT1") = rsBeneOpt("BF_AMT")
'            End If
'            If IsNull(rsBeneOpt("BF_ECOST")) Then
'                rsWRK("WB_PREMIUM1") = 0
'            Else
'                rsWRK("WB_PREMIUM1") = rsBeneOpt("BF_ECOST")
'            End If
'            rsWRK("WB_TAX1") = rsWRK("WB_PREMIUM1") * 0.08
'        End If
'        If xCode = "OPLS" Then
'            rsWRK("WB_BCODE2") = xCode
'            If IsNull(rsBeneOpt("BF_AMT")) Then
'                rsWRK("WB_AMT2") = 0
'            Else
'                rsWRK("WB_AMT2") = rsBeneOpt("BF_AMT")
'            End If
'            If IsNull(rsBeneOpt("BF_ECOST")) Then
'                rsWRK("WB_PREMIUM2") = 0
'            Else
'                rsWRK("WB_PREMIUM2") = rsBeneOpt("BF_ECOST")
'            End If
'            rsWRK("WB_TAX2") = rsWRK("WB_PREMIUM2") * 0.08
'            'Get Spouse Sex and Smoker - Begin
'            SQLQ = "SELECT * FROM HRDEPEND where DP_EMPNBR = " & xEMPNBR & " "
'            SQLQ = SQLQ & "AND (DP_RELATE = 'Wife' OR DP_RELATE = 'Husband' OR DP_RELATE = 'Spouse') "
'            If rsDep.State <> 0 Then rsDep.Close
'            rsDep.Open SQLQ, gdbAdoIhr001, adOpenStatic
'            If Not rsDep.EOF Then
'                If Not IsNull(rsDep("DP_SEX")) Then
'                    rsWRK("WB_TEXT2") = rsDep("DP_SEX")
'                End If
'                If Not IsNull(rsDep("DP_SMOKER")) Then
'                    If rsDep("DP_SMOKER") Then
'                        rsWRK("WB_TEXT3") = "Y"
'                    Else
'                        rsWRK("WB_TEXT3") = "N"
'                    End If
'                End If
'                If IsDate(rsDep("DP_DOB")) Then 'Ticket #14364
'                    rsWRK("WB_DATE") = rsDep("DP_DOB")
'                End If
'            End If
'
'            rsDep.Close
'            'Get Spouse Sex and Smoker - End
'        End If
'        If xCode = "OPLC" Then
'            rsWRK("WB_BCODE3") = xCode
'            If IsNull(rsBeneOpt("BF_AMT")) Then
'                rsWRK("WB_AMT3") = 0
'            Else
'                rsWRK("WB_AMT3") = rsBeneOpt("BF_AMT")
'            End If
'            If IsNull(rsBeneOpt("BF_ECOST")) Then
'                rsWRK("WB_PREMIUM3") = 0
'            Else
'                rsWRK("WB_PREMIUM3") = rsBeneOpt("BF_ECOST")
'            End If
'            rsWRK("WB_TAX3") = rsWRK("WB_PREMIUM3") * 0.08
'        End If
'        rsWRK.Update
'NexeRec:
'        rsBeneOpt.MoveNext
'    Loop
'    rsBeneOpt.Close
'    MDIMain.panHelp(0).FloodType = 0
'    MDIMain.panHelp(1).Caption = " "
'    Screen.MousePointer = DEFAULT
'End Sub

'Private Sub HCAS_Request_Report()
'    Dim rsHREmp As New ADODB.Recordset
'    Dim rsAttend As New ADODB.Recordset
'    Dim rsRequest As New ADODB.Recordset
'    Dim SQLQ As String
'    Dim WRKSQL As String
'    Dim I, totNum, xVacAcc
'    Dim xEmpNo
'
'    Screen.MousePointer = HOURGLASS
'    MDIMain.panHelp(0).FloodType = 1
'    MDIMain.panHelp(1).Caption = " Please Wait"
'    MDIMain.panHelp(2).Caption = ""
'    MDIMain.panHelp(0).FloodPercent = 0
'
'    gdbAdoIhr001W.BeginTrans
'    gdbAdoIhr001W.Execute "DELETE FROM HR_REQUEST_RPT WHERE REQ_WRKEMP='" & glbUserID & "'"
'    gdbAdoIhr001W.CommitTrans
'
'    rsRequest.Open "SELECT * FROM HR_REQUEST_RPT WHERE REQ_WRKEMP='" & glbUserID & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'
'    WRKSQL = " AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
'    SQLQ = "SELECT  AD_EMPNBR, MONTH(AD_DOA) AS MNTH,"
'    SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='OT' THEN AD_HRS ELSE 0 END) AS OTEARN,"
'    SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='CT' THEN AD_HRS ELSE 0 END) AS CTTAKEN,"
'    SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,3)='VAC' THEN AD_HRS ELSE 0 END) AS VACTAKEN"
'    SQLQ = SQLQ & " FROM HR_ATTENDANCE"
'    SQLQ = SQLQ & " WHERE AD_DOA >= " & Date_SQL(dlpDateRange(0).Text) & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text)
'    SQLQ = SQLQ & " AND " & WRKSQL
'    SQLQ = SQLQ & " GROUP BY AD_EMPNBR,MONTH(AD_DOA)"
'    SQLQ = SQLQ & " ORDER BY AD_EMPNBR, MONTH(AD_DOA)"
'    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    If Not rsAttend.EOF Then
'        totNum = rsAttend.RecordCount: I = 0
'        rsAttend.MoveFirst
'    End If
'
'    xEmpNo = 0
'    Do While Not rsAttend.EOF
'        If (I / totNum) <= 1 Then
'            MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
'            I = I + 1
'        End If
'        DoEvents
'
'
'        'Adding records to request report table
'        If xEmpNo <> rsAttend("AD_EMPNBR") Then
'
'            'Retrieve Vacation Accrued for the year
'            Set rsHREmp = Nothing
'            SQLQ = "SELECT ED_EMPNBR, ED_VAC, ED_EFDATE, ED_ETDATE FROM HREMP WHERE ED_EMPNBR = " & rsAttend("AD_EMPNBR")
'            'SQLQ = SQLQ & " AND ED_EFDATE <= " & " AND ED_ETDATE >="
'            rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'            If Not rsHREmp.EOF Then
'                xVacAcc = rsHREmp("ED_VAC")
'            Else
'                xVacAcc = 0
'            End If
'            rsHREmp.Close
'
'            rsRequest.AddNew
'            xEmpNo = rsAttend("AD_EMPNBR")
'        End If
'
'        rsRequest("EMPNBR") = rsAttend("AD_EMPNBR")
'        If rsAttend("MNTH") = 1 Then
'            rsRequest("OT_ACC_JAN") = rsAttend("OTEARN")
'            rsRequest("OT_TKN_JAN") = rsAttend("CTTAKEN")
'            rsRequest("VAC_TKN_JAN") = rsAttend("VACTAKEN")
'        End If
'
'        If rsAttend("MNTH") = 2 Then
'            rsRequest("OT_ACC_FEB") = rsAttend("OTEARN")
'            rsRequest("OT_TKN_FEB") = rsAttend("CTTAKEN")
'            rsRequest("VAC_TKN_FEB") = rsAttend("VACTAKEN")
'        End If
'
'        If rsAttend("MNTH") = 3 Then
'            rsRequest("OT_ACC_MAR") = rsAttend("OTEARN")
'            rsRequest("OT_TKN_MAR") = rsAttend("CTTAKEN")
'            rsRequest("VAC_TKN_MAR") = rsAttend("VACTAKEN")
'        End If
'
'        If rsAttend("MNTH") = 4 Then
'            rsRequest("OT_ACC_APR") = rsAttend("OTEARN")
'            rsRequest("OT_TKN_APR") = rsAttend("CTTAKEN")
'            rsRequest("VAC_TKN_APR") = rsAttend("VACTAKEN")
'        End If
'
'        If rsAttend("MNTH") = 5 Then
'            rsRequest("OT_ACC_MAY") = rsAttend("OTEARN")
'            rsRequest("OT_TKN_MAY") = rsAttend("CTTAKEN")
'            rsRequest("VAC_TKN_MAY") = rsAttend("VACTAKEN")
'        End If
'
'        If rsAttend("MNTH") = 6 Then
'            rsRequest("OT_ACC_JUN") = rsAttend("OTEARN")
'            rsRequest("OT_TKN_JUN") = rsAttend("CTTAKEN")
'            rsRequest("VAC_TKN_JUN") = rsAttend("VACTAKEN")
'        End If
'
'        If rsAttend("MNTH") = 7 Then
'            rsRequest("OT_ACC_JUL") = rsAttend("OTEARN")
'            rsRequest("OT_TKN_JUL") = rsAttend("CTTAKEN")
'            rsRequest("VAC_TKN_JUL") = rsAttend("VACTAKEN")
'        End If
'
'        If rsAttend("MNTH") = 8 Then
'            rsRequest("OT_ACC_AUG") = rsAttend("OTEARN")
'            rsRequest("OT_TKN_AUG") = rsAttend("CTTAKEN")
'            rsRequest("VAC_TKN_AUG") = rsAttend("VACTAKEN")
'        End If
'
'        If rsAttend("MNTH") = 9 Then
'            rsRequest("OT_ACC_SEP") = rsAttend("OTEARN")
'            rsRequest("OT_TKN_SEP") = rsAttend("CTTAKEN")
'            rsRequest("VAC_TKN_SEP") = rsAttend("VACTAKEN")
'        End If
'
'        If rsAttend("MNTH") = 10 Then
'            rsRequest("OT_ACC_OCT") = rsAttend("OTEARN")
'            rsRequest("OT_TKN_OCT") = rsAttend("CTTAKEN")
'            rsRequest("VAC_TKN_OCT") = rsAttend("VACTAKEN")
'        End If
'
'        If rsAttend("MNTH") = 11 Then
'            rsRequest("OT_ACC_NOV") = rsAttend("OTEARN")
'            rsRequest("OT_TKN_NOV") = rsAttend("CTTAKEN")
'            rsRequest("VAC_TKN_NOV") = rsAttend("VACTAKEN")
'        End If
'
'        If rsAttend("MNTH") = 12 Then
'            rsRequest("OT_ACC_DEC") = rsAttend("OTEARN")
'            rsRequest("OT_TKN_DEC") = rsAttend("CTTAKEN")
'            rsRequest("VAC_TKN_DEC") = rsAttend("VACTAKEN")
'        End If
'
'
'        If xVacAcc <> 0 And xVacAcc <> "" Then
'            rsRequest("VAC_ACC_JAN") = xVacAcc / 12
'            rsRequest("VAC_ACC_FEB") = xVacAcc / 12
'            rsRequest("VAC_ACC_MAR") = xVacAcc / 12
'            rsRequest("VAC_ACC_APR") = xVacAcc / 12
'            rsRequest("VAC_ACC_MAY") = xVacAcc / 12
'            rsRequest("VAC_ACC_JUN") = xVacAcc / 12
'            rsRequest("VAC_ACC_JUL") = xVacAcc / 12
'            rsRequest("VAC_ACC_AUG") = xVacAcc / 12
'            rsRequest("VAC_ACC_SEP") = xVacAcc / 12
'            rsRequest("VAC_ACC_OCT") = xVacAcc / 12
'            rsRequest("VAC_ACC_NOV") = xVacAcc / 12
'            rsRequest("VAC_ACC_DEC") = xVacAcc / 12
'        End If
'
'        rsRequest("REQ_WRKEMP") = glbUserID
'        rsAttend.MoveNext
'
'        If rsAttend.EOF Then
'            rsRequest.Update
'        Else
'            If xEmpNo <> rsAttend("AD_EMPNBR") Then
'                rsRequest.Update
'            End If
'        End If
'    Loop
'
'    rsAttend.Close
'    MDIMain.panHelp(0).FloodType = 0
'    MDIMain.panHelp(1).Caption = " "
'    Screen.MousePointer = DEFAULT
'
'End Sub

'Private Sub Cri_Dates()
'Dim TempCri As String
'Dim dtYYY%, dtMM%, dtDD%
'Dim X%
'    If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
'        'Collectcorp Inc. - Ticket #14437
'        If InStr(1, fglbFileName, "SN2390_LicenseAddr.rpt") > 0 Then
'            TempCri = "({HREMP.ED_DOH} "
'        ElseIf InStr(1, fglbFileName, "SN2390_LicenseTerm.rpt") > 0 Then
'            TempCri = "({TERM_HRTRMEMP.TERM_DOT} "
'        ElseIf InStr(1, fglbFileName, "SN2390_LicenseAdditions.rpt") > 0 Then
'            TempCri = "({HR_USERDEFINE_TABLE.UD_DATE1} "
'        End If
'
'        dtYYY% = Year(dlpDateRange(2).Text)
'        dtMM% = Month(dlpDateRange(2).Text)
'        dtDD% = Day(dlpDateRange(2).Text)
'        TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
'        dtYYY% = Year(dlpDateRange(3).Text)
'        dtMM% = Month(dlpDateRange(3).Text)
'        dtDD% = Day(dlpDateRange(3).Text)
'        TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
'        GoTo Cri_FTDatst
'    End If
'
'    For X% = 2 To 3
'        If Len(dlpDateRange(X).Text) > 0 Then
'            'Collectcorp Inc. - Ticket #14437
'            If InStr(1, fglbFileName, "SN2390_LicenseAddr.rpt") > 0 Then
'                TempCri = "({HREMP.ED_DOH} "
'            ElseIf InStr(1, fglbFileName, "SN2390_LicenseTerm.rpt") > 0 Then
'                TempCri = "({TERM_HRTRMEMP.TERM_DOT} "
'            ElseIf InStr(1, fglbFileName, "SN2390_LicenseAdditions.rpt") > 0 Then
'                TempCri = "({HR_USERDEFINE_TABLE.UD_DATE1} "
'            End If
'
'            If X% = 2 Then
'                TempCri = TempCri & " >= "
'            Else
'                TempCri = TempCri & " <= "
'            End If
'            dtYYY% = Year(dlpDateRange(X).Text)
'            dtMM% = Month(dlpDateRange(X).Text)
'            dtDD% = Day(dlpDateRange(X).Text)
'            TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
'            GoTo Cri_FTDatst
'        End If
'    Next X%
'
'Cri_FTDatst:
'    If Len(TempCri) >= 1 Then
'        If Not glbiOneWhere Then
'            glbstrSelCri = TempCri
'        Else
'            glbstrSelCri = glbstrSelCri & " AND " & TempCri
'        End If
'        glbiOneWhere = True
'    End If
'
'End Sub

'Private Sub Cri_License_Codes(intIdx%)
'Dim CodeCri As String
'Dim countr   As Integer
'Dim strCd$
'
'If Len(clpUser(intIdx%)) > 0 Then
'    If InStr(1, fglbFileName, "SN2390_LicenseTerm.rpt") > 0 Then
'        Select Case intIdx%
'            Case 0: strCd$ = "TERM_USERDEFINE_TABLE.UD_CODE1"
'            Case 1: strCd$ = "TERM_USERDEFINE_TABLE.UD_CODE2"
'        End Select
'    Else
'        Select Case intIdx%
'            Case 0: strCd$ = "HR_USERDEFINE_TABLE.UD_CODE1"
'            Case 1: strCd$ = "HR_USERDEFINE_TABLE.UD_CODE2"
'        End Select
'    End If
'    CodeCri = "(Ucase({" & strCd$ & "}) in  ['" & Replace(clpUser(intIdx%).Text, ",", "','") & "'])"
'End If
'
'If Len(CodeCri) >= 1 Then
'    If Not glbiOneWhere Then
'        glbstrSelCri = CodeCri
'    Else
'        glbstrSelCri = glbstrSelCri & " AND " & CodeCri
'    End If
'    glbiOneWhere = True
'End If
'
'End Sub

'Private Sub Cri_BirthMonth()
'Dim EECri As String, OneSet%, X%
'Dim strMonth As String
'
'If Len(comDateMonth.Text) < 1 Then Exit Sub
'
'Select Case comDateMonth.Text
'    Case "January": strMonth = "1"
'    Case "February": strMonth = "2"
'    Case "March": strMonth = "3"
'    Case "April": strMonth = "4"
'    Case "May": strMonth = "5"
'    Case "June": strMonth = "6"
'    Case "July": strMonth = "7"
'    Case "August": strMonth = "8"
'    Case "September": strMonth = "9"
'    Case "October": strMonth = "10"
'    Case "November": strMonth = "11"
'    Case "December": strMonth = "12"
'End Select
'
''EECri = "{@BirthMonth}= " & Val(intMonth)
'EECri = "Month({HREMP.ED_DOB})= " & Val(strMonth)
'
'If glbiOneWhere Then
'    glbstrSelCri = glbstrSelCri & " AND " & EECri
'Else
'    glbstrSelCri = EECri
'End If
'glbiOneWhere = True
'
'End Sub

'Private Sub Cri_DOHMonth()
'Dim EECri As String, OneSet%, X%
'Dim strMonth As String
'
'If Len(comDateMonth.Text) < 1 Then Exit Sub
'
'Select Case comDateMonth.Text
'    Case "January": strMonth = "1"
'    Case "February": strMonth = "2"
'    Case "March": strMonth = "3"
'    Case "April": strMonth = "4"
'    Case "May": strMonth = "5"
'    Case "June": strMonth = "6"
'    Case "July": strMonth = "7"
'    Case "August": strMonth = "8"
'    Case "September": strMonth = "9"
'    Case "October": strMonth = "10"
'    Case "November": strMonth = "11"
'    Case "December": strMonth = "12"
'End Select
'
''EECri = "{@DOHMonth}= " & Val(strMonth)
'EECri = "Month({HREMP.ED_DOH})= " & Val(strMonth)
'
'If glbiOneWhere Then
'    glbstrSelCri = glbstrSelCri & " AND " & EECri
'Else
'    glbstrSelCri = EECri
'End If
'glbiOneWhere = True
'
'End Sub

Private Sub Vacation_Report_XLS_HCAS()
    On Error GoTo Vacation_Report_XLS_HCAS_Err

    Dim rsHREmp As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim i, totNum
    Dim xHourlyRate
    Dim xTotVacOut

    'Check the Month End Date is within the Current Vacation Year
    If Len(Trim(dlpAsOf.Text)) = 0 Then
        MsgBox "Month Ending Date cannot be blank"
        dlpAsOf.SetFocus
        Exit Sub
    ElseIf Not IsDate(dlpAsOf.Text) Then
        MsgBox "Invalid Month Ending Date"
        dlpAsOf.SetFocus
        Exit Sub
    End If
    
    sSQLQ = Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")")
    SQLQ = "SELECT HREMP.ED_EMPNBR, HREMP.ED_FNAME AS EFNAME, HREMP.ED_SURNAME AS ESURNAME, HREMP.ED_EFDATE, HREMP.ED_ETDATE, "
    SQLQ = SQLQ & " (CASE WHEN HREMP.ED_PVAC IS NULL THEN 0 ELSE HREMP.ED_PVAC END) + "
    'SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AC_HRS) FROM HR_ACCRUAL WHERE AC_EMPNBR = HREMP.ED_EMPNBR AND AC_TYPE = 'VAC' AND AC_ACTION IN ('U','C') AND AC_COMMENTS LIKE 'Current Vac.%' AND AC_EDATE >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AC_EDATE <=" & Date_SQL(dlpAsOf.Text) & ") IS NULL THEN 0 ELSE (SELECT SUM(AC_HRS) FROM HR_ACCRUAL WHERE AC_EMPNBR = HREMP.ED_EMPNBR AND AC_TYPE = 'VAC' AND AC_ACTION IN ('U','C') AND AC_COMMENTS LIKE 'Current Vac.%' AND AC_EDATE >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AC_EDATE <=" & Date_SQL(dlpAsOf.Text) & ") END) - "
    SQLQ = SQLQ & " HREMP.ED_VAC - "
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,3) = 'VAC' AND AD_DOA >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,3) = 'VAC' AND AD_DOA >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") END) AS ED_VACOUTS, "
    SQLQ = SQLQ & " JH_JOB, JB_DESCR, JH_REPTAU, SH_SALARY, HREMP.ED_PVAC,JH_DHRS,JH_WHRS,SH_SALCD,SUPER.ED_SURNAME AS SSURNAME,SUPER.ED_FNAME AS SFNAME,"
    'SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AC_HRS) FROM HR_ACCRUAL WHERE AC_EMPNBR = HREMP.ED_EMPNBR AND AC_TYPE = 'VAC' AND AC_ACTION IN ('U','C') AND AC_COMMENTS LIKE 'Current Vac.%' AND AC_EDATE >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AC_EDATE <=" & Date_SQL(dlpAsOf.Text) & ") IS NULL THEN 0 ELSE (SELECT SUM(AC_HRS) FROM HR_ACCRUAL WHERE AC_EMPNBR = HREMP.ED_EMPNBR AND AC_TYPE = 'VAC' AND AC_ACTION IN ('U','C') AND AC_COMMENTS LIKE 'Current Vac.%' AND AC_EDATE >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AC_EDATE <=" & Date_SQL(dlpAsOf.Text) & ") END) AS VAC,"
    SQLQ = SQLQ & " HREMP.ED_VAC,"
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,3) = 'VAC' AND AD_DOA >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,3) = 'VAC' AND AD_DOA >=" & Date_SQL("01/01/" & Year(dlpAsOf.Text)) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") END) AS VACT"
    SQLQ = SQLQ & " FROM ((((HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR AND JH_CURRENT <>0)"
    SQLQ = SQLQ & " INNER JOIN HR_SALARY_HISTORY ON HR_SALARY_HISTORY.SH_EMPNBR = ED_EMPNBR AND SH_CURRENT <>0)"
    SQLQ = SQLQ & " INNER JOIN HRJOB ON HRJOB.JB_CODE = HR_JOB_HISTORY.JH_JOB)"
    SQLQ = SQLQ & " INNER JOIN HREMP SUPER ON SUPER.ED_EMPNBR = HR_JOB_HISTORY.JH_REPTAU)"
    
    SQLQ = SQLQ & " WHERE " & sSQLQ
    
    SQLQ = SQLQ & " GROUP BY HREMP.ED_EMPNBR,HREMP.ED_FNAME,HREMP.ED_SURNAME,HREMP.ED_EFDATE,HREMP.ED_ETDATE,JH_JOB,JB_DESCR,JH_REPTAU,SH_SALARY,HREMP.ED_PVAC,JH_DHRS,JH_WHRS,SH_SALCD,SUPER.ED_FNAME,SUPER.ED_SURNAME,HREMP.ED_VAC"
    SQLQ = SQLQ & " ORDER BY SSURNAME, SFNAME, ED_VACOUTS ASC"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: i = 0
        rsHREmp.MoveFirst

        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "VacationRptTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "VacationRpt" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
    
        FileCopy xlsFileTmp, xlsFileMat
    
        Screen.MousePointer = HOURGLASS
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
    
        exSheet.Cells(1, 2) = Format(Now, "mmm dd, yyyy")
        exSheet.Cells(2, 2) = Time$
        exSheet.Cells(5, 1) = "Report for the Month Ending: " & Format(dlpAsOf.Text, "dd-mmm-yy")
        
        xTotVacOut = 0
        xRow = 11
        'Columns: 1 - Name, 3 - Job Title, 4 - Supervisor, 5 - Previous Vac, 6 - Current Vac, 7 - Taken Vac, 8 - Outstanding Vac, 9 - Salary, 10 - Cost of Oustanding Vacation
        Do While Not rsHREmp.EOF
            If (i / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (i / totNum) * 100
                i = i + 1
            End If
            DoEvents
            
            xHourlyRate = 0
            If rsHREmp("SH_SALCD") = "H" Then
                xHourlyRate = rsHREmp("SH_SALARY")
            ElseIf rsHREmp("SH_SALCD") = "A" And rsHREmp("JH_WHRS") <> 0 Then
                xHourlyRate = Round((rsHREmp("SH_SALARY") / 52) / rsHREmp("JH_WHRS"), 2)
            End If
            
            exSheet.Cells(xRow, 1) = rsHREmp("ESURNAME") & ", " & rsHREmp("EFNAME")
            exSheet.Cells(xRow, 3) = rsHREmp("JB_DESCR")
            'If rsHREmp("JH_REPTAU") <> "" And Not IsNull(rsHREmp("JH_REPTAU")) Then
                'exSheet.Cells(xRow, 4) = GetEmpData(rsHREmp("JH_REPTAU"), "ED_SURNAME") & ", " & GetEmpData(rsHREmp("JH_REPTAU"), "ED_FNAME")
                exSheet.Cells(xRow, 4) = rsHREmp("SSURNAME") & ", " & rsHREmp("SFNAME")
            'End If
            If rsHREmp("JH_DHRS") <> 0 And Not IsNull(rsHREmp("JH_DHRS")) Then
                exSheet.Cells(xRow, 5) = Round(rsHREmp("ED_PVAC") / rsHREmp("JH_DHRS"), 2)
                'exSheet.Cells(xRow, 6) = Round(rsHREmp("VAC") / rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 6) = Round(rsHREmp("ED_VAC") / rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 7) = Round(rsHREmp("VACT") / rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 8) = Round(rsHREmp("ED_VACOUTS") / rsHREmp("JH_DHRS"), 2)
                xTotVacOut = xTotVacOut + Round(rsHREmp("ED_VACOUTS") / rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 9) = Format(Round((rsHREmp("ED_VACOUTS") / rsHREmp("JH_DHRS")), 2) * (xHourlyRate * rsHREmp("JH_DHRS")), "#,##0.00")
                'exSheet.Cells(xRow, 10) = Format(xHourlyRate * rsHREmp("JH_DHRS"), "#,##0.00")
                exSheet.Cells(xRow, 10) = Format(rsHREmp("SH_SALARY"), "#,##0.00")
            End If
            
            rsHREmp.MoveNext
            xRow = xRow + 1
        Loop
        
        exSheet.Cells(xRow + 2, 1) = "Total Number of Employees Reported: " & totNum
        exSheet.Cells(xRow + 3, 1) = "Total Number of Days of Vacation Outstanding as at Current Month End: " & xTotVacOut
        exSheet.Rows(xRow + 2).Font.Bold = True
        exSheet.Rows(xRow + 3).Font.Bold = True
        
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing

    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    rsHREmp.Close
    
Exit Sub

Vacation_Report_XLS_HCAS_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
Resume Next

End Sub

Function LanchXlsW98(xFileName)
On Error GoTo Error_Deal
    LanchXlsW98 = False
    Shell "Start " & GetShortName(xFileName)
    LanchXlsW98 = True
Exit Function
Error_Deal:

End Function


Private Sub Overtime_Report_XLS_HCAS()
    On Error GoTo Overtime_Report_XLS_HCAS_Err

    Dim rsHREmp As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim i, totNum
    Dim xHourlyRate
    Dim xPrvMonthDate, xMonthBeg As Date

    'Check the Month End Date is within the Current Vacation Year
    If Len(Trim(dlpAsOf.Text)) = 0 Then
        MsgBox "Month Ending Date cannot be blank"
        dlpAsOf.SetFocus
        Exit Sub
    ElseIf Not IsDate(dlpAsOf.Text) Then
        MsgBox "Invalid Month Ending Date"
        dlpAsOf.SetFocus
        Exit Sub
    End If
    
    xPrvMonthDate = DateAdd("d", -1, CVDate(month(dlpAsOf.Text) & " 1," & Year(dlpAsOf.Text)))
    xMonthBeg = DateAdd("d", 1, xPrvMonthDate)
    sSQLQ = Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")")
    SQLQ = "SELECT HREMP.ED_EMPNBR AS EEMPNBR, HREMP.ED_FNAME AS EFNAME, HREMP.ED_SURNAME AS ESURNAME, "
    SQLQ = SQLQ & " ((CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA <=" & Date_SQL(xPrvMonthDate) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA <=" & Date_SQL(xPrvMonthDate) & ") END) - "
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'CT' AND AD_DOA <=" & Date_SQL(xPrvMonthDate) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'CT' AND AD_DOA <=" & Date_SQL(xPrvMonthDate) & ") END)) + "
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA >=" & Date_SQL(xMonthBeg) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA >=" & Date_SQL(xMonthBeg) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") END) - "
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'CT' AND AD_DOA >=" & Date_SQL(xMonthBeg) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'CT' AND AD_DOA >=" & Date_SQL(xMonthBeg) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") END) AS OT_OUTST, "
    SQLQ = SQLQ & " JH_JOB, JB_DESCR, JH_REPTAU, SH_SALARY, JH_DHRS,JH_WHRS,SH_SALCD,SUPER.ED_SURNAME AS SSURNAME,SUPER.ED_FNAME AS SFNAME,"
    SQLQ = SQLQ & " ((CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA <=" & Date_SQL(xPrvMonthDate) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA <=" & Date_SQL(xPrvMonthDate) & ") END) - "
    SQLQ = SQLQ & "  (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'CT' AND AD_DOA <=" & Date_SQL(xPrvMonthDate) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'CT' AND AD_DOA <=" & Date_SQL(xPrvMonthDate) & ") END)) AS OT_PREV, "
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA >=" & Date_SQL(xMonthBeg) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA >=" & Date_SQL(xMonthBeg) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") END) AS OT_CURR,"
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'CT' AND AD_DOA >=" & Date_SQL(xMonthBeg) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'CT' AND AD_DOA >=" & Date_SQL(xMonthBeg) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & ") END) AS OT_TAKEN,"
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA >=" & Date_SQL(DateAdd("m", -12, xMonthBeg)) & " AND AD_DOA <=" & Date_SQL(xPrvMonthDate) & ") IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND LEFT(AD_REASON,2) = 'OT' AND AD_DOA >=" & Date_SQL(DateAdd("m", -12, xMonthBeg)) & " AND AD_DOA <=" & Date_SQL(xPrvMonthDate) & ") END) AS OT_12MNTS"
    SQLQ = SQLQ & " FROM ((((HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR AND JH_CURRENT <>0)"
    SQLQ = SQLQ & " INNER JOIN HR_SALARY_HISTORY ON HR_SALARY_HISTORY.SH_EMPNBR = ED_EMPNBR AND SH_CURRENT <>0)"
    SQLQ = SQLQ & " INNER JOIN HRJOB ON HRJOB.JB_CODE = HR_JOB_HISTORY.JH_JOB)"
    SQLQ = SQLQ & " INNER JOIN HREMP SUPER ON SUPER.ED_EMPNBR = HR_JOB_HISTORY.JH_REPTAU)"
    
    SQLQ = SQLQ & " WHERE " & sSQLQ

    SQLQ = SQLQ & " GROUP BY HREMP.ED_EMPNBR,HREMP.ED_FNAME,HREMP.ED_SURNAME,JH_JOB,JB_DESCR,JH_REPTAU,SH_SALARY,JH_DHRS,JH_WHRS,SH_SALCD,SUPER.ED_FNAME,SUPER.ED_SURNAME"
    SQLQ = SQLQ & " ORDER BY SSURNAME, SFNAME, JB_DESCR, OT_OUTST ASC"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: i = 0
        rsHREmp.MoveFirst

        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "OvertimeRptTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "OvertimeRpt" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
    
        FileCopy xlsFileTmp, xlsFileMat
    
        Screen.MousePointer = HOURGLASS
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
    
        exSheet.Cells(1, 2) = Format(Now, "mmm dd, yyyy")
        exSheet.Cells(2, 2) = Time$
        exSheet.Cells(5, 1) = "Report for the Month Ending: " & Format(dlpAsOf.Text, "dd-mmm-yy")
        
        xRow = 11
        'Columns: 1 - Name, 3 - Job Title, 4 - Supervisor, 5 - Previous Vac, 6 - Current Vac, 7 - Taken Vac, 8 - Outstanding Vac, 9 - Salary, 10 - Cost of Oustanding Vacation
        Do While Not rsHREmp.EOF
            If (i / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (i / totNum) * 100
                i = i + 1
            End If
            DoEvents
            
            xHourlyRate = 0
            If rsHREmp("SH_SALCD") = "H" Then
                xHourlyRate = rsHREmp("SH_SALARY")
            ElseIf rsHREmp("SH_SALCD") = "A" And rsHREmp("JH_WHRS") <> 0 Then
                xHourlyRate = Round((rsHREmp("SH_SALARY") / 52) / rsHREmp("JH_WHRS"), 2)
            End If
            
            exSheet.Cells(xRow, 1) = rsHREmp("ESURNAME") & ", " & rsHREmp("EFNAME")
            exSheet.Cells(xRow, 3) = rsHREmp("JB_DESCR")
            'If rsHREmp("JH_REPTAU") <> "" And Not IsNull(rsHREmp("JH_REPTAU")) Then
                exSheet.Cells(xRow, 4) = rsHREmp("SSURNAME") & ", " & rsHREmp("SFNAME")
            'End If
            'If rsHREmp("JH_DHRS") <> 0 And Not IsNull(rsHREmp("JH_DHRS")) Then
                exSheet.Cells(xRow, 5) = Round(rsHREmp("OT_12MNTS"), 2) ' / rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 6) = Round(rsHREmp("OT_PREV"), 2) ' / rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 7) = Round(rsHREmp("OT_CURR"), 2) '/ rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 8) = Round(rsHREmp("OT_TAKEN"), 2) ' / rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 9) = Round(rsHREmp("OT_OUTST"), 2) ' / rsHREmp("JH_DHRS"), 2)
                exSheet.Cells(xRow, 10) = Format(Round(rsHREmp("OT_OUTST"), 2) * xHourlyRate, "#,##0.00")
            'End If
            'exSheet.Cells(xRow, 10) = Format(xHourlyRate, "#,##0.00")
            exSheet.Cells(xRow, 11) = Format(rsHREmp("SH_SALARY"), "#,##0.00")
            
            
            rsHREmp.MoveNext
            xRow = xRow + 1
        Loop
        
                
        exSheet.Cells(xRow + 2, 1) = "Total Number of Employees Reported: " & totNum
        exSheet.Rows(xRow + 2).Font.Bold = True

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing

    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    rsHREmp.Close
    
Exit Sub

Overtime_Report_XLS_HCAS_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
Resume Next

End Sub

Private Sub Attendance_Report_XLS_HCAS()
    On Error GoTo Attendance_Report_XLS_HCAS_Err

    Dim rsHREmp As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim i, totNum, totEmpDaysCurr, totEmpAbsent12
    Dim xHourlyRate
    Dim xAbsentDays, xCostAbsences, xTotAbsent12
    Dim xTotSal
    Dim xPrv12Date As Date

    'Check the Month End Date is within the Current Vacation Year
    If Len(Trim(dlpAsOf.Text)) = 0 Then
        MsgBox "Month Ending Date cannot be blank"
        dlpAsOf.SetFocus
        Exit Sub
    ElseIf Not IsDate(dlpAsOf.Text) Then
        MsgBox "Invalid Month Ending Date"
        dlpAsOf.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(dlpDateRange(0).Text)) = 0 Then
        MsgBox "For the Period From Date cannot be blank"
        dlpDateRange(0).SetFocus
        Exit Sub
    ElseIf Not IsDate(dlpDateRange(0).Text) Then
        MsgBox "Invalid For the Period From Date"
        dlpDateRange(0).SetFocus
        Exit Sub
    End If
    If Len(Trim(dlpDateRange(1).Text)) = 0 Then
        MsgBox "For the Period To Date cannot be blank"
        dlpDateRange(1).SetFocus
        Exit Sub
    ElseIf Not IsDate(dlpDateRange(1).Text) Then
        MsgBox "Invalid For the Period To Date"
        dlpDateRange(1).SetFocus
        Exit Sub
    End If
    If CVDate(dlpDateRange(0).Text) > CVDate(dlpDateRange(1).Text) Then
        MsgBox "For the Period From Date cannot be greater than To Date"
        dlpDateRange(0).SetFocus
        Exit Sub
    End If
    
    xPrv12Date = DateAdd("m", -12, dlpAsOf.Text)
    xPrv12Date = DateAdd("d", 1, xPrv12Date)
    sSQLQ = Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")")
    SQLQ = "SELECT HREMP.ED_EMPNBR AS EEMPNBR, HREMP.ED_FNAME AS EFNAME, HREMP.ED_SURNAME AS ESURNAME, "
    SQLQ = SQLQ & " JH_JOB, JB_DESCR, JH_REPTAU, SH_SALARY, JH_DHRS,JH_WHRS,SH_SALCD,SUPER.ED_SURNAME AS SSURNAME,SUPER.ED_FNAME AS SFNAME,"
    'SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND AD_DOA >=" & Date_SQL(CVDate(Month(dlpAsOf.Text) & "/01/" & Year(dlpAsOf.Text))) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='ADRE' AND TB_ABSENCE=1)) IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND AD_DOA >=" & Date_SQL(CVDate(Month(dlpAsOf.Text) & "/01/" & Year(dlpAsOf.Text))) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='ADRE' AND TB_ABSENCE=1)) END) AS CURR_ABSENT,"
    'SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND AD_DOA >=" & Date_SQL(xPrv12Date) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='ADRE' AND TB_ABSENCE=1)) IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND AD_DOA >=" & Date_SQL(xPrv12Date) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='ADRE' AND TB_ABSENCE=1)) END) AS ABSENT_12"
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND AD_DOA >=" & Date_SQL(CVDate(month(dlpAsOf.Text) & "/01/" & Year(dlpAsOf.Text))) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & " AND LEFT(AD_REASON,3) = 'SIC') IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND AD_DOA >=" & Date_SQL(CVDate(month(dlpAsOf.Text) & "/01/" & Year(dlpAsOf.Text))) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & " AND LEFT(AD_REASON,3) = 'SIC') END) AS CURR_ABSENT,"
    SQLQ = SQLQ & " (CASE WHEN (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND AD_DOA >=" & Date_SQL(xPrv12Date) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & " AND LEFT(AD_REASON,3) = 'SIC') IS NULL THEN 0 ELSE (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HREMP.ED_EMPNBR AND AD_DOA >=" & Date_SQL(xPrv12Date) & " AND AD_DOA <=" & Date_SQL(dlpAsOf.Text) & " AND LEFT(AD_REASON,3) = 'SIC') END) AS ABSENT_12"
    SQLQ = SQLQ & " FROM ((((HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR AND JH_CURRENT <>0)"
    SQLQ = SQLQ & " INNER JOIN HR_SALARY_HISTORY ON HR_SALARY_HISTORY.SH_EMPNBR = ED_EMPNBR AND SH_CURRENT <>0)"
    SQLQ = SQLQ & " INNER JOIN HRJOB ON HRJOB.JB_CODE = HR_JOB_HISTORY.JH_JOB)"
    SQLQ = SQLQ & " INNER JOIN HREMP SUPER ON SUPER.ED_EMPNBR = HR_JOB_HISTORY.JH_REPTAU)"
    
    SQLQ = SQLQ & " WHERE " & sSQLQ

    SQLQ = SQLQ & " GROUP BY HREMP.ED_EMPNBR,HREMP.ED_FNAME,HREMP.ED_SURNAME,JH_JOB,JB_DESCR,JH_REPTAU,SH_SALARY,JH_DHRS,JH_WHRS,SH_SALCD,SUPER.ED_FNAME,SUPER.ED_SURNAME"
    SQLQ = SQLQ & " ORDER BY SSURNAME, SFNAME, JB_DESCR"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: i = 0
        rsHREmp.MoveFirst

        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "AttendanceRptTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "AttendanceRpt" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
    
        FileCopy xlsFileTmp, xlsFileMat
    
        Screen.MousePointer = HOURGLASS
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
    
        exSheet.Cells(1, 2) = Format(Now, "mmm dd, yyyy")
        exSheet.Cells(2, 2) = Time$
        exSheet.Cells(5, 1) = "Report for the Month Ending: " & Format(dlpAsOf.Text, "dd-mmm-yy")
        exSheet.Cells(6, 1) = "Report for the Period: " & Format(dlpDateRange(0).Text, "mmmm dd, yyyy") & " to " & Format(dlpDateRange(1).Text, "mmmm dd, yyyy")
        
        xAbsentDays = 0
        xCostAbsences = 0
        xTotSal = 0
        xTotAbsent12 = 0
        totEmpDaysCurr = 0
        xRow = 12
        totEmpAbsent12 = 0
        'Columns: 1 - Name, 3 - Job Title, 4 - Supervisor, 5 - Previous Vac, 6 - Current Vac, 7 - Taken Vac, 8 - Outstanding Vac, 9 - Salary, 10 - Cost of Oustanding Vacation
        Do While Not rsHREmp.EOF
            If (i / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (i / totNum) * 100
                i = i + 1
            End If
            DoEvents
            
            xHourlyRate = 0
            If rsHREmp("SH_SALCD") = "H" Then
                xHourlyRate = rsHREmp("SH_SALARY")
            ElseIf rsHREmp("SH_SALCD") = "A" And rsHREmp("JH_WHRS") <> 0 Then
                xHourlyRate = Round((rsHREmp("SH_SALARY") / 52) / rsHREmp("JH_WHRS"), 2)
            End If
            
            exSheet.Cells(xRow, 1) = rsHREmp("ESURNAME") & ", " & rsHREmp("EFNAME")
            exSheet.Cells(xRow, 3) = rsHREmp("JB_DESCR")
            exSheet.Cells(xRow, 4) = rsHREmp("SSURNAME") & ", " & rsHREmp("SFNAME")
            
            If rsHREmp("JH_DHRS") <> 0 And Not IsNull(rsHREmp("JH_DHRS")) Then
                '*** Days Absent in Current months ***
                exSheet.Cells(xRow, 5) = Round(rsHREmp("CURR_ABSENT") / rsHREmp("JH_DHRS"), 2)
                
                'Total # of days Absent in the Current Month
                xAbsentDays = xAbsentDays + Round(rsHREmp("CURR_ABSENT") / rsHREmp("JH_DHRS"), 2)
                
                'Total # of Employees absent in the current month
                If Round(rsHREmp("CURR_ABSENT") / rsHREmp("JH_DHRS"), 2) > 0 Then
                    totEmpDaysCurr = totEmpDaysCurr + 1
                End If
                
                '*** Days Absent in 12 months ***
                exSheet.Cells(xRow, 6) = Round(rsHREmp("ABSENT_12") / rsHREmp("JH_DHRS"), 2)
                                
                'Total # of Days Absent in 12 months
                xTotAbsent12 = xTotAbsent12 + Round(rsHREmp("ABSENT_12") / rsHREmp("JH_DHRS"), 2)
                
                'Total # of Employee absent in 12 months
                If Round(rsHREmp("ABSENT_12") / rsHREmp("JH_DHRS"), 2) > 0 Then
                    totEmpAbsent12 = totEmpAbsent12 + 1
                End If

                '*** Cost of Absence in 12 months ***
                exSheet.Cells(xRow, 8) = Format(Round((rsHREmp("ABSENT_12") / rsHREmp("JH_DHRS")), 2) * (xHourlyRate * rsHREmp("JH_DHRS")), "#,##0.00")
                
                'Total Cost of Absense in 12 Months
                xCostAbsences = xCostAbsences + Format(Round((rsHREmp("ABSENT_12") / rsHREmp("JH_DHRS")), 2) * (xHourlyRate * rsHREmp("JH_DHRS")), "#,##0.00")
                
                '*** Employee's Salary per Day ***
                'exSheet.Cells(xRow, 7) = Format(xHourlyRate * rsHREmp("JH_DHRS"), "#,##0.00")
                exSheet.Cells(xRow, 7) = Format(rsHREmp("SH_SALARY"), "#,##0.00")
                
                'Total Salary
                xTotSal = xTotSal + Format(xHourlyRate * rsHREmp("JH_DHRS"), "#,##0.00")
            End If
            
            rsHREmp.MoveNext
            xRow = xRow + 1
        Loop
        
                
        exSheet.Cells(xRow + 2, 1) = "Total Number of Employees Reported: " & totEmpDaysCurr
        exSheet.Rows(xRow + 2).Font.Bold = True
        exSheet.Cells(xRow + 3, 1) = "Total Number of Days Absent: " & xAbsentDays
        exSheet.Rows(xRow + 3).Font.Bold = True
        If totEmpDaysCurr <> 0 Then
            exSheet.Cells(xRow + 4, 1) = "Average Number of Days Absent: " & Round(xAbsentDays / totEmpDaysCurr, 2)
        Else
            exSheet.Cells(xRow + 4, 1) = "Average Number of Days Absent: 0"
        End If
        exSheet.Rows(xRow + 4).Font.Bold = True
        exSheet.Cells(xRow + 5, 1) = "Cost of Total Absences: " & Format(Round(xCostAbsences, 2), "#,##0.00")
        exSheet.Rows(xRow + 5).Font.Bold = True
        exSheet.Cells(xRow + 6, 1) = "% Cost of Absences to Total Payroll: " '& Round(xCostAbsences / xTotSal * 100, 2)
        exSheet.Rows(xRow + 6).Font.Bold = True
        
        exSheet.Cells(xRow + 7, 1) = "Total Number of Employee with Recorded Absences (12 month period): " & totEmpAbsent12
        exSheet.Rows(xRow + 7).Font.Bold = True
        If totEmpAbsent12 <> 0 Then
            exSheet.Cells(xRow + 8, 1) = "Average # of days absent of Employees with Recorded Absences (12 month period): " & Round(xTotAbsent12 / totEmpAbsent12, 2)
        Else
            exSheet.Cells(xRow + 8, 1) = "Average # of days absent of Employees with Recorded Absences (12 month period): 0"
        End If
        exSheet.Rows(xRow + 8).Font.Bold = True
        
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing

    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    rsHREmp.Close
    
Exit Sub

Attendance_Report_XLS_HCAS_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
Resume Next

End Sub

Private Sub Request_Report_XLS_HCAS()
    On Error GoTo Request_Report_XLS_HCAS_Err
    
    Dim rsHREmp As New ADODB.Recordset
    Dim rsAttOT As New ADODB.Recordset
    Dim rsAttCT As New ADODB.Recordset
    Dim rsAccVAC As New ADODB.Recordset
    Dim rsAttVACT As New ADODB.Recordset
    
    Dim rsAccFrwOT As New ADODB.Recordset
    Dim rsAccFrwCT As New ADODB.Recordset
    Dim rsAccFrwVAC As New ADODB.Recordset
    Dim rsAccFrwVACT As New ADODB.Recordset
    
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim i, totNum, x, z
    Dim xDate As Date
    Dim xHourlyRate
    Dim xReptAuth As String
    Dim xAccFrwDate As Date
    Dim xTotOTBal, xTotVacBal
    Dim xSumOTBal, xSumVacBal
    
       
    If Len(Trim(dlpDateRange(0).Text)) = 0 Then
        MsgBox "For the Period From Date cannot be blank"
        dlpDateRange(0).SetFocus
        Exit Sub
    ElseIf Not IsDate(dlpDateRange(0).Text) Then
        MsgBox "Invalid For the Period From Date"
        dlpDateRange(0).SetFocus
        Exit Sub
    End If
    If Len(Trim(dlpDateRange(1).Text)) = 0 Then
        MsgBox "For the Period To Date cannot be blank"
        dlpDateRange(1).SetFocus
        Exit Sub
    ElseIf Not IsDate(dlpDateRange(1).Text) Then
        MsgBox "Invalid For the Period To Date"
        dlpDateRange(1).SetFocus
        Exit Sub
    End If
    If CVDate(dlpDateRange(0).Text) > CVDate(dlpDateRange(1).Text) Then
        MsgBox "For the Period From Date cannot be greater than To Date"
        dlpDateRange(0).SetFocus
        Exit Sub
    End If
    xReptAuth = ""
    If Len(elpRept.Text) > 0 Then
        xReptAuth = " JH_REPTAU= " & Trim(elpRept.Text) & " "
    End If
    
    sSQLQ = Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")")
    
'    'Accrued Forward Hours for Overtime
'    SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS ACCFRW_OT"
'    SQLQ = SQLQ & " FROM HR_ATTENDANCE "
'    SQLQ = SQLQ & " WHERE LEFT(AD_REASON,2) = 'OT' AND AD_DOA <" & Date_SQL(dlpDateRange(0).Text)
'    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & sSQLQ & ")"
'    SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
'    SQLQ = SQLQ & " ORDER BY AD_EMPNBR"
'    rsAccFrwOT.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
'
'    'Accrued Forward Hours for Comp. Time
'    SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS ACCFRW_CT"
'    SQLQ = SQLQ & " FROM HR_ATTENDANCE "
'    SQLQ = SQLQ & " WHERE LEFT(AD_REASON,2) = 'CT' AND AD_DOA <" & Date_SQL(dlpDateRange(0).Text)
'    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & sSQLQ & ")"
'    SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
'    SQLQ = SQLQ & " ORDER BY AD_EMPNBR"
'    rsAccFrwCT.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
'    'Accrued Forward Hours for Vacation
'    SQLQ = "SELECT AC_EMPNBR, SUM(AC_HRS) AS ACCFRW_VAC"
'    SQLQ = SQLQ & " FROM HR_ACCRUAL"
'    SQLQ = SQLQ & " WHERE AC_TYPE = 'VAC' AND AC_ACTION IN ('U','C') AND AC_COMMENTS LIKE 'Current Vac.%' AND"
'    SQLQ = SQLQ & " AC_EDATE <" & Date_SQL(dlpDateRange(0).Text)
'    SQLQ = SQLQ & " AND AC_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & sSQLQ & ")"
'    SQLQ = SQLQ & " GROUP BY AC_EMPNBR"
'    SQLQ = SQLQ & " ORDER BY AC_EMPNBR"
'    rsAccFrwVAC.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
'
'    'Accrued Forward Hours for Vacation Taken
'    SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS ACCFRW_VACT"
'    SQLQ = SQLQ & " FROM HR_ATTENDANCE"
'    SQLQ = SQLQ & " WHERE LEFT(AD_REASON,3) = 'VAC' AND AD_DOA <" & Date_SQL(dlpDateRange(0).Text)
'    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & sSQLQ & ")"
'    SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
'    SQLQ = SQLQ & " ORDER BY AD_EMPNBR"
'    rsAccFrwVACT.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
'    'Overtime Earned
'    SQLQ = "SELECT AD_EMPNBR, MONTH(AD_DOA) AS MONTH, SUM(AD_HRS) AS OT"
'    SQLQ = SQLQ & " FROM HR_ATTENDANCE "
'    SQLQ = SQLQ & " WHERE LEFT(AD_REASON,2) = 'OT' AND AD_DOA >=" & Date_SQL(dlpDateRange(0).Text) & " AND AD_DOA <=" & Date_SQL(dlpDateRange(1).Text)
'    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & sSQLQ & ")"
'    SQLQ = SQLQ & " GROUP BY MONTH(AD_DOA), AD_EMPNBR"
'    SQLQ = SQLQ & " ORDER BY AD_EMPNBR, MONTH(AD_DOA)"
'    rsAttOT.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
'    'Comp. Time Taken
'    SQLQ = "SELECT AD_EMPNBR, MONTH(AD_DOA) AS MONTH, SUM(AD_HRS) AS CT"
'    SQLQ = SQLQ & " FROM HR_ATTENDANCE "
'    SQLQ = SQLQ & " WHERE LEFT(AD_REASON,2) = 'CT' AND AD_DOA >=" & Date_SQL(dlpDateRange(0).Text) & " AND AD_DOA <=" & Date_SQL(dlpDateRange(1).Text)
'    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & sSQLQ & ")"
'    SQLQ = SQLQ & " GROUP BY MONTH(AD_DOA), AD_EMPNBR"
'    SQLQ = SQLQ & " ORDER BY AD_EMPNBR, MONTH(AD_DOA)"
'    rsAttCT.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
'    'Vacation Accrued
'    SQLQ = "SELECT AC_EMPNBR, MONTH(AC_EDATE) AS MONTH, SUM(AC_HRS) AS VAC_ACCRUE"
'    SQLQ = SQLQ & " FROM HR_ACCRUAL"
'    SQLQ = SQLQ & " WHERE AC_TYPE = 'VAC' AND AC_ACTION IN ('U','C') AND AC_COMMENTS LIKE 'Current Vac.%' AND"
'    SQLQ = SQLQ & " AC_EDATE >=" & Date_SQL(dlpDateRange(0).Text) & " AND AC_EDATE <=" & Date_SQL(dlpDateRange(1).Text)
'    SQLQ = SQLQ & " AND AC_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & sSQLQ & ")"
'    SQLQ = SQLQ & " GROUP BY MONTH(AC_EDATE), AC_EMPNBR"
'    SQLQ = SQLQ & " ORDER BY AC_EMPNBR, MONTH(AC_EDATE)"
'    rsAccVAC.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
'    'Vacation Taken
'    SQLQ = "SELECT AD_EMPNBR, MONTH(AD_DOA) AS MONTH, SUM(AD_HRS) AS VAC_TAKEN"
'    SQLQ = SQLQ & " FROM HR_ATTENDANCE"
'    SQLQ = SQLQ & " WHERE LEFT(AD_REASON,3) = 'VAC' AND AD_DOA >=" & Date_SQL(dlpDateRange(0).Text) & " AND AD_DOA <=" & Date_SQL(dlpDateRange(1).Text)
'    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & sSQLQ & ")"
'    SQLQ = SQLQ & " GROUP BY MONTH(AD_DOA), AD_EMPNBR"
'    SQLQ = SQLQ & " ORDER BY AD_EMPNBR, MONTH(AD_DOA)"
'    rsAttVACT.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
    
    xAccFrwDate = DateAdd("d", -1, dlpDateRange(0).Text)
    'Employees
    SQLQ = "SELECT HREMP.ED_EMPNBR AS EEMPNBR, HREMP.ED_FNAME AS EFNAME, HREMP.ED_SURNAME AS ESURNAME, "
    SQLQ = SQLQ & " SUPER.ED_SURNAME AS SSURNAME,SUPER.ED_FNAME AS SFNAME,SH_SALARY, JH_DHRS,JH_WHRS,SH_SALCD,JH_REPTAU,HREMP.ED_PVAC,HREMP.ED_VAC"
    SQLQ = SQLQ & " FROM (((HREMP INNER JOIN HR_JOB_HISTORY ON ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR AND JH_CURRENT <>0)"
    SQLQ = SQLQ & " INNER JOIN HREMP SUPER ON SUPER.ED_EMPNBR = HR_JOB_HISTORY.JH_REPTAU)"
    SQLQ = SQLQ & " INNER JOIN HR_SALARY_HISTORY ON HR_SALARY_HISTORY.SH_EMPNBR = HREMP.ED_EMPNBR AND SH_CURRENT <>0)"
    SQLQ = SQLQ & " WHERE " & sSQLQ
    If Len(xReptAuth) > 0 Then
        SQLQ = SQLQ & " AND " & xReptAuth
    End If
    SQLQ = SQLQ & " GROUP BY HREMP.ED_EMPNBR,HREMP.ED_FNAME,HREMP.ED_SURNAME,SUPER.ED_FNAME,SUPER.ED_SURNAME,SH_SALARY, JH_DHRS,JH_WHRS,SH_SALCD,JH_REPTAU,HREMP.ED_PVAC,HREMP.ED_VAC"
    SQLQ = SQLQ & " ORDER BY ESURNAME, EFNAME ASC" 'SSURNAME, SFNAME ASC"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
    If Not rsHREmp.EOF Then
        totNum = rsHREmp.RecordCount: i = 0
        rsHREmp.MoveFirst

        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "RequestRptTmp.xls"
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "RequestRpt" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
    
        FileCopy xlsFileTmp, xlsFileMat
    
        Screen.MousePointer = HOURGLASS
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
    
        exSheet.Cells(1, 2) = Format(Now, "mmm dd, yyyy")
        exSheet.Cells(2, 2) = Time$
        exSheet.Cells(3, 3) = Format(dlpDateRange(0).Text, "mmmm dd, yyyy") & " to " & Format(dlpDateRange(1).Text, "mmmm dd, yyyy")
        exSheet.Cells(9, 3) = "As of " & Format(xAccFrwDate, "mmm dd, yyyy")
        
        'Display Months Names across the columns
        xDate = dlpDateRange(0).Text
        For x = 1 To 12
            If xDate <= dlpDateRange(1).Text Then
                exSheet.Cells(7, 3 + x) = Format(xDate, "mmm")
                exSheet.Cells(8, 3 + x) = Format(xDate, "m")
                xDate = DateAdd("m", 1, xDate)
            Else
                Exit For
            End If
        Next x
                
        xRow = 10
        Do While Not rsHREmp.EOF
            If (i / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (i / totNum) * 100
                i = i + 1
            End If
            DoEvents
    
            xHourlyRate = 0
            If rsHREmp("SH_SALCD") = "H" Then
                xHourlyRate = rsHREmp("SH_SALARY")
            ElseIf rsHREmp("SH_SALCD") = "A" And rsHREmp("JH_WHRS") <> 0 Then
                xHourlyRate = Round((rsHREmp("SH_SALARY") / 52) / rsHREmp("JH_WHRS"), 2)
            End If
    
            'Employee Name
            exSheet.Cells(xRow, 1) = rsHREmp("ESURNAME") & ", " & rsHREmp("EFNAME")
            exSheet.Range("A" & xRow).Font.Bold = True
            exSheet.Cells(xRow, 18) = Format(rsHREmp("SH_SALARY"), "#,##0.00")
            
            'Overtime ---------------------------------------------------------------------
            If chkOvertime Then
                xRow = xRow + 1
                exSheet.Cells(xRow, 1) = "OT Accrued"
                
                'Get Accrued Forward
                'Accrued Forward Hours for Overtime
                SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS ACCFRW_OT"
                SQLQ = SQLQ & " FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE LEFT(AD_REASON,2) = 'OT' AND AD_DOA <" & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & sSQLQ & ")"
                SQLQ = SQLQ & " AND AD_EMPNBR = " & rsHREmp("EEMPNBR")
                SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
                'SQLQ = SQLQ & " ORDER BY AD_EMPNBR"
                rsAccFrwOT.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
                
                'Accrued Forward Hours for Comp. Time
                SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS ACCFRW_CT"
                SQLQ = SQLQ & " FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE LEFT(AD_REASON,2) = 'CT' AND AD_DOA <" & Date_SQL(dlpDateRange(0).Text)
                SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & sSQLQ & ")"
                SQLQ = SQLQ & " AND AD_EMPNBR = " & rsHREmp("EEMPNBR")
                SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
                'SQLQ = SQLQ & " ORDER BY AD_EMPNBR"
                rsAccFrwCT.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
                
                If Not rsAccFrwOT.EOF Then
                    If Not rsAccFrwCT.EOF Then
                        'Accrued Forward
                        exSheet.Cells(xRow, 3) = Format(IIf(Not IsNull(rsAccFrwOT("ACCFRW_OT")) And rsAccFrwOT("ACCFRW_OT") <> "", rsAccFrwOT("ACCFRW_OT"), 0) - IIf(Not IsNull(rsAccFrwCT("ACCFRW_CT")) And rsAccFrwCT("ACCFRW_CT") <> "", rsAccFrwCT("ACCFRW_CT"), 0), "#,##0.00")
                    Else
                        'Accrued Forward
                        exSheet.Cells(xRow, 3) = Format(IIf(Not IsNull(rsAccFrwOT("ACCFRW_OT")) And rsAccFrwOT("ACCFRW_OT") <> "", rsAccFrwOT("ACCFRW_OT"), 0) - 0, "#,##0.00")
                    End If
                Else
                    If Not rsAccFrwCT.EOF Then
                        'Accrued Forward
                        exSheet.Cells(xRow, 3) = Format(0 - IIf(Not IsNull(rsAccFrwCT("ACCFRW_CT")) And rsAccFrwCT("ACCFRW_CT") <> "", rsAccFrwCT("ACCFRW_CT"), 0), "#,##0.00")
                    Else
                        exSheet.Cells(xRow, 3) = ""
                    End If
                End If
                rsAccFrwOT.Close
                rsAccFrwCT.Close
                Set rsAccFrwOT = Nothing
                Set rsAccFrwCT = Nothing
                
                '12 Months Data
                'Overtime Accrued
                SQLQ = "SELECT AD_EMPNBR, MONTH(AD_DOA) AS MONTH, SUM(AD_HRS) AS OT"
                SQLQ = SQLQ & " FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE LEFT(AD_REASON,2) = 'OT' AND AD_DOA >=" & Date_SQL(dlpDateRange(0).Text) & " AND AD_DOA <=" & Date_SQL(dlpDateRange(1).Text)
                SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & sSQLQ & ")"
                SQLQ = SQLQ & " AND AD_EMPNBR = " & rsHREmp("EEMPNBR")
                SQLQ = SQLQ & " GROUP BY MONTH(AD_DOA), AD_EMPNBR"
                SQLQ = SQLQ & " ORDER BY AD_EMPNBR, MONTH(AD_DOA)"
                rsAttOT.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
                
                If Not rsAttOT.EOF Then
                    For x = 1 To 12
                        exSheet.Cells(xRow, 3 + x) = Format(0, "#,##0.00")
                    Next

                    rsAttOT.MoveFirst
                    Do While Not rsAttOT.EOF
                        For x = 1 To 12
                            If rsAttOT("MONTH") = exSheet.Cells(8, 3 + x) Then
                                exSheet.Cells(xRow, 3 + x) = Format(rsAttOT("OT"), "#,##0.00")
                                Exit For
                            End If
                        Next
                        rsAttOT.MoveNext
                    Loop
                Else
                    For x = 1 To 12
                        exSheet.Cells(xRow, 3 + x) = Format(0, "#,##0.00")
                    Next
                End If
                rsAttOT.Close
                Set rsAttOT = Nothing
                
                'Sum Monthly OT Accrued
                exSheet.Range("P" & xRow).Formula = "=Sum(R" & xRow & "C4:R" & xRow & "C15)"
                exSheet.Cells(xRow, 16) = Format(exSheet.Cells(xRow, 16), "Fixed")
                
                'OT Taken
                xRow = xRow + 1
                exSheet.Cells(xRow, 1) = "OT Taken"
                
                'Get CT - OT Taken
                SQLQ = "SELECT AD_EMPNBR, MONTH(AD_DOA) AS MONTH, SUM(AD_HRS) AS CT"
                SQLQ = SQLQ & " FROM HR_ATTENDANCE "
                SQLQ = SQLQ & " WHERE LEFT(AD_REASON,2) = 'CT' AND AD_DOA >=" & Date_SQL(dlpDateRange(0).Text) & " AND AD_DOA <=" & Date_SQL(dlpDateRange(1).Text)
                SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & sSQLQ & ")"
                SQLQ = SQLQ & " AND AD_EMPNBR = " & rsHREmp("EEMPNBR")
                SQLQ = SQLQ & " GROUP BY MONTH(AD_DOA), AD_EMPNBR"
                SQLQ = SQLQ & " ORDER BY AD_EMPNBR, MONTH(AD_DOA)"
                rsAttCT.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
                
                If Not rsAttCT.EOF Then
                    For x = 1 To 12
                        exSheet.Cells(xRow, 3 + x) = Format(0, "#,##0.00")
                    Next
                    
                    rsAttCT.MoveFirst
                    Do While Not rsAttCT.EOF
                        For x = 1 To 12
                            If rsAttCT("MONTH") = exSheet.Cells(8, 3 + x) Then
                                exSheet.Cells(xRow, 3 + x) = Format(rsAttCT("CT"), "#,##0.00")
                                Exit For
                            End If
                        Next
                        rsAttCT.MoveNext
                    Loop
                Else
                    For x = 1 To 12
                        exSheet.Cells(xRow, 3 + x) = Format(0, "#,##0.00")
                    Next
                End If
                rsAttCT.Close
                Set rsAttCT = Nothing
                
                'Sum OT Taken
                exSheet.Range("P" & xRow).Formula = "=Sum(R" & xRow & "C4:R" & xRow & "C15)"
                exSheet.Cells(xRow, 16) = Format(exSheet.Cells(xRow, 16), "Fixed")
                
                'OT Balance
                xRow = xRow + 1
                exSheet.Cells(xRow, 1) = "OT Balance"
                For x = 1 To 12
                    If x = 1 Then
                        exSheet.Cells(xRow, 3 + x) = Format((exSheet.Cells(xRow - 2, 3) + exSheet.Cells(xRow - 2, 3 + x)) - exSheet.Cells(xRow - 1, 3 + x), "#,##0.00")
                    Else
                        exSheet.Cells(xRow, 3 + x) = Format((exSheet.Cells(xRow, (3 + x) - 1) + exSheet.Cells(xRow - 2, 3 + x)) - exSheet.Cells(xRow - 1, 3 + x), "#,##0.00")
                    End If
                Next
                
                'Sum OT Balance
                'exSheet.Range("P" & xRow).Formula = "=Sum(R" & xRow & "C4:R" & xRow & "C15)"
                'exSheet.Cells(xRow, 16) = Format(exSheet.Cells(xRow, 16), "Fixed")
                
                'OT Accrued Cost ****
                'Get Total of Monthly Balance
                xTotOTBal = 0
                'Ticket #14841 - Not the Monthly Balance but the last month's balance
                'For x = 4 To 15
                '    xTotOTBal = xTotOTBal + exSheet.Cells(xRow, x)
                'Next
                xTotOTBal = exSheet.Cells(xRow, 15) 'Ticket #14841
                exSheet.Cells(xRow, 17) = Format(xTotOTBal * xHourlyRate, "Currency")
                exSheet.Range("Q" & xRow).Font.Bold = True
            End If
            
            
            'Vacation ---------------------------------------------------------------------
            If chkVacation Then
                If chkOvertime Then xRow = xRow + 1     'Create a line Gap
                xRow = xRow + 1
                exSheet.Cells(xRow, 1) = "Vacation Accrued"
                
'                'Accrued Forward Hours for Vacation
'                SQLQ = "SELECT AC_EMPNBR, SUM(AC_HRS) AS ACCFRW_VAC"
'                SQLQ = SQLQ & " FROM HR_ACCRUAL"
'                SQLQ = SQLQ & " WHERE AC_TYPE = 'VAC' AND AC_ACTION IN ('U','C') AND AC_COMMENTS LIKE 'Current Vac.%' AND"
'                SQLQ = SQLQ & " AC_EDATE <" & Date_SQL(dlpDateRange(0).Text)
'                SQLQ = SQLQ & " AND AC_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & sSQLQ & ")"
'                SQLQ = SQLQ & " AND AC_EMPNBR = " & rsHREmp("EEMPNBR")
'                SQLQ = SQLQ & " GROUP BY AC_EMPNBR"
'                rsAccFrwVAC.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
'
'                'Accrued Forward Hours for Vacation Taken
'                SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS ACCFRW_VACT"
'                SQLQ = SQLQ & " FROM HR_ATTENDANCE"
'                SQLQ = SQLQ & " WHERE LEFT(AD_REASON,3) = 'VAC' AND AD_DOA <" & Date_SQL(dlpDateRange(0).Text)
'                SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & sSQLQ & ")"
'                SQLQ = SQLQ & " AND AD_EMPNBR = " & rsHREmp("EEMPNBR")
'                SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
'                rsAccFrwVACT.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
                
'                If Not rsAccFrwVAC.EOF Then
'                    If Not rsAccFrwVACT.EOF Then
'                        'Accrued Forward
'                        exSheet.Cells(xRow, 3) = IIf(Not IsNull(rsAccFrwVAC("ACCFRW_VAC")) And rsAccFrwVAC("ACCFRW_VAC") <> "", rsAccFrwVAC("ACCFRW_VAC"), 0) - IIf(Not IsNull(rsAccFrwVACT("ACCFRW_VACT")) And rsAccFrwVACT("ACCFRW_VACT") <> "", rsAccFrwVACT("ACCFRW_VACT"), 0)
'                    Else
'                        'Accrued Forward
'                        exSheet.Cells(xRow, 3) = IIf(Not IsNull(rsAccFrwVAC("ACCFRW_VAC")) And rsAccFrwVAC("ACCFRW_VAC") <> "", rsAccFrwVAC("ACCFRW_VAC"), 0) - 0
'                    End If
'                Else
'                    If Not rsAccFrwVACT.EOF Then
'                        'Accrued Forward
'                        exSheet.Cells(xRow, 3) = 0 - IIf(Not IsNull(rsAccFrwVACT("ACCFRW_VACT")) And rsAccFrwVACT("ACCFRW_VACT") <> "", rsAccFrwVACT("ACCFRW_VACT"), 0)
'                    Else
'                        exSheet.Cells(xRow, 3) = ""
'                    End If
'                End If
'                rsAccFrwVAC.Close
'                rsAccFrwVACT.Close
'                Set rsAccFrwVAC = Nothing
'                Set rsAccFrwVACT = Nothing
                If rsHREmp("JH_DHRS") <> 0 And Not IsNull(rsHREmp("JH_DHRS")) Then
                    exSheet.Cells(xRow, 3) = Format(Round(IIf(Not IsNull(rsHREmp("ED_PVAC")) And rsHREmp("ED_PVAC") <> "", rsHREmp("ED_PVAC") / rsHREmp("JH_DHRS"), 0), 2), "#,##0.00")
                End If

               
                '12 Months Data
                'Vacation Accrued ****
'                SQLQ = "SELECT AC_EMPNBR, MONTH(AC_EDATE) AS MONTH, SUM(AC_HRS) AS VAC_ACCRUE"
'                SQLQ = SQLQ & " FROM HR_ACCRUAL"
'                SQLQ = SQLQ & " WHERE AC_TYPE = 'VAC' AND AC_ACTION IN ('U','C') AND AC_COMMENTS LIKE 'Current Vac.%' AND"
'                SQLQ = SQLQ & " AC_EDATE >=" & Date_SQL(dlpDateRange(0).Text) & " AND AC_EDATE <=" & Date_SQL(dlpDateRange(1).Text)
'                SQLQ = SQLQ & " AND AC_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & sSQLQ & ")"
'                SQLQ = SQLQ & " AND AC_EMPNBR = " & rsHREmp("EEMPNBR")
'                SQLQ = SQLQ & " GROUP BY MONTH(AC_EDATE), AC_EMPNBR"
'                SQLQ = SQLQ & " ORDER BY AC_EMPNBR, MONTH(AC_EDATE)"
'                rsAccVAC.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
'
'                If Not rsAccVAC.EOF Then
'                    For x = 1 To 12
'                        exSheet.Cells(xRow, 3 + x) = Format(0, "#,##0.00")
'                    Next
'
'                    rsAccVAC.MoveFirst
'                    Do While Not rsAccVAC.EOF
'                        For x = 1 To 12
'                            If rsAccVAC("MONTH") = exSheet.Cells(8, 3 + x) Then
'                                exSheet.Cells(xRow, 3 + x) = Format(rsAccVAC("VAC_ACCRUE"), "#,##0.00")
'                                Exit For
'                            End If
'                        Next
'                        rsAccVAC.MoveNext
'                    Loop
'                Else
'                    For x = 1 To 12
'                        exSheet.Cells(xRow, 3 + x) = Format(0, "#,##0.00")
'                    Next
'                End If
'                rsAccVAC.Close
'                Set rsAccVAC = Nothing
                'Accumulated Vacation for the First Month.
                If rsHREmp("JH_DHRS") <> 0 And Not IsNull(rsHREmp("JH_DHRS")) Then
                    exSheet.Cells(xRow, 4) = Format(Round(IIf(Not IsNull(rsHREmp("ED_VAC")) And rsHREmp("ED_VAC") <> "", rsHREmp("ED_VAC") / rsHREmp("JH_DHRS"), 0), 2), "#,##0.00")
                End If
                                                
                'Sum Monthly Vacation Accrues
                exSheet.Range("P" & xRow).Formula = "=Sum(R" & xRow & "C4:R" & xRow & "C15)"
                exSheet.Cells(xRow, 16) = Format(exSheet.Cells(xRow, 16), "Fixed")

                'Vacation Taken ****
                xRow = xRow + 1
                exSheet.Cells(xRow, 1) = "Vacation Taken"
                
                'Retrieve Vacation Taken
                If rsHREmp("JH_DHRS") <> 0 And Not IsNull(rsHREmp("JH_DHRS")) Then
                    SQLQ = "SELECT AD_EMPNBR, MONTH(AD_DOA) AS MONTH, SUM(AD_HRS) AS VAC_TAKEN"
                    SQLQ = SQLQ & " FROM HR_ATTENDANCE"
                    SQLQ = SQLQ & " WHERE LEFT(AD_REASON,3) = 'VAC' AND AD_DOA >=" & Date_SQL(dlpDateRange(0).Text) & " AND AD_DOA <=" & Date_SQL(dlpDateRange(1).Text)
                    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & sSQLQ & ")"
                    SQLQ = SQLQ & " AND AD_EMPNBR = " & rsHREmp("EEMPNBR")
                    SQLQ = SQLQ & " GROUP BY MONTH(AD_DOA), AD_EMPNBR"
                    SQLQ = SQLQ & " ORDER BY AD_EMPNBR, MONTH(AD_DOA)"
                    rsAttVACT.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
                    
                    If Not rsAttVACT.EOF Then
                        For x = 1 To 12
                            exSheet.Cells(xRow, 3 + x) = Format(0, "#,##0.00")
                        Next
                        rsAttVACT.MoveFirst
                        Do While Not rsAttVACT.EOF
                            For x = 1 To 12
                                If rsAttVACT("MONTH") = exSheet.Cells(8, 3 + x) Then
                                    exSheet.Cells(xRow, 3 + x) = Format(rsAttVACT("VAC_TAKEN") / rsHREmp("JH_DHRS"), "#,##0.00")
                                    Exit For
                                End If
                            Next
                            rsAttVACT.MoveNext
                        Loop
                    Else
                        For x = 1 To 12
                            exSheet.Cells(xRow, 3 + x) = Format(0, "#,##0.00")
                        Next
                    End If
                    rsAttVACT.Close
                    Set rsAttVACT = Nothing
                End If
                'Sum Vacation Taken
                exSheet.Range("P" & xRow).Formula = "=Sum(R" & xRow & "C4:R" & xRow & "C15)"
                exSheet.Cells(xRow, 16) = Format(exSheet.Cells(xRow, 16), "Fixed")
                
                'Vacation Balance ****
                xRow = xRow + 1
                exSheet.Cells(xRow, 1) = "Vacation Balance"
                For x = 1 To 12
                    If x = 1 Then
                        exSheet.Cells(xRow, 3 + x) = Format((exSheet.Cells(xRow - 2, 3) + exSheet.Cells(xRow - 2, 3 + x)) - exSheet.Cells(xRow - 1, 3 + x), "#,##0.00")
                    Else
                        exSheet.Cells(xRow, 3 + x) = Format((exSheet.Cells(xRow, (3 + x) - 1) + exSheet.Cells(xRow - 2, 3 + x)) - exSheet.Cells(xRow - 1, 3 + x), "#,##0.00")
                    End If
                Next
                
                'Sum Vacation Balance
                'exSheet.Range("P" & xRow).Formula = "=Sum(R" & xRow & "C4:R" & xRow & "C15)"
                'exSheet.Cells(xRow, 16) = Format(exSheet.Cells(xRow, 16), "Fixed")
                
                'Vacation Accrued Cost ****
                'Get Total of Monthly Balance
                xTotVacBal = 0
                'Ticket #14841 - Not the Monthly Balance but the last month's balance
                'For x = 4 To 15
                '    xTotVacBal = xTotVacBal + exSheet.Cells(xRow, x)
                'Next
                xTotVacBal = exSheet.Cells(xRow, 15)     'Ticket #14841
                If rsHREmp("JH_DHRS") <> 0 And Not IsNull(rsHREmp("JH_DHRS")) Then
                    exSheet.Cells(xRow, 17) = Format((xTotVacBal * rsHREmp("JH_DHRS")) * xHourlyRate, "Currency")
                End If
                exSheet.Range("Q" & xRow).Font.Bold = True
                
                xRow = xRow + 1
            End If
    
            rsHREmp.MoveNext
            xRow = xRow + 1
            
            If chkOvertime And Not chkVacation Then
                xRow = xRow + 1
            End If
        Loop
    
        'Overall OT and Vacation Monthly Balance Totals
        xRow = xRow + 1
        If chkOvertime Then
            exSheet.Cells(xRow, 1) = "OT Balance Total : "
            exSheet.Range("A" & xRow).Font.Bold = True
        End If
        If chkVacation Then
            exSheet.Cells(xRow + 1, 1) = "Vac Balance Total :"
            exSheet.Range("A" & xRow + 1).Font.Bold = True
        End If
        
            '=SUM(D1664,D1669,D1673)
            z = 68  '=> CHR(68) = D
            xSumOTBal = 0
            xSumVacBal = 0
            For x = 4 To 15  'Move across Columns
                If chkOvertime Then
                    For i = 13 To xRow - 3 Step 9    'Move down the Row
                        'xSumOTBal = xSumOTBal & Chr(z) & i & ","
                        xSumOTBal = xSumOTBal + exSheet.Cells(i, x)
                    Next
                    'xSumOTBal = Left(xSumOTBal, Len(xSumOTBal) - 1)
                    'exSheet.Cells(xRow, x) = "=SUM(" & xSumOTBal & ")"
                    exSheet.Cells(xRow, x) = xSumOTBal
                    exSheet.Range(Chr(z) & xRow).Font.Bold = True
                    
                    If chkVacation Then
                        For i = 17 To xRow - 3 Step 9    'Move down the Row
                            'xSumVacBal = xSumVacBal & Chr(z) & i & ","
                            xSumVacBal = xSumVacBal + exSheet.Cells(i, x)
                        Next
                        'xSumVacBal = Left(xSumVacBal, Len(xSumVacBal) - 1)
                        'exSheet.Cells(xRow + 1, x) = "=SUM(" & xSumVacBal & ")"
                        exSheet.Cells(xRow + 1, x) = xSumVacBal
                        exSheet.Range(Chr(z) & xRow + 1).Font.Bold = True
                    End If
                End If
                
                If Not chkOvertime And chkVacation Then
                    For i = 13 To xRow - 3 Step 9    'Move down the Row
                        'xSumVacBal = xSumVacBal & Chr(z) & i & ","
                        xSumVacBal = xSumVacBal + exSheet.Cells(i, x)
                    Next
                    'xSumVacBal = Left(xSumVacBal, Len(xSumVacBal) - 1)
                    'exSheet.Cells(xRow + 1, x) = "=SUM(" & xSumVacBal & ")"
                    exSheet.Cells(xRow + 1, x) = xSumVacBal
                    exSheet.Range(Chr(z) & xRow + 1).Font.Bold = True
                End If
                
                xSumOTBal = 0
                xSumVacBal = 0
                z = z + 1
            Next
    
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing

    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    
    rsHREmp.Close

Exit Sub

Request_Report_XLS_HCAS_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", fglbEmpTable, "SELECT")
Resume Next

End Sub
