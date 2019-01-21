VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmERetirement 
   Caption         =   "Retirment"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   12075
   WindowState     =   2  'Maximized
   Begin VB.Frame frmDeathProcess 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   5280
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   6615
      Begin VB.OptionButton optDth 
         Caption         =   "Retiree"
         Height          =   195
         Index           =   0
         Left            =   4275
         TabIndex        =   36
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optDth 
         Caption         =   "Spouse"
         Height          =   195
         Index           =   1
         Left            =   5280
         TabIndex        =   35
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optDth 
         Caption         =   "Employee"
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   34
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optDth 
         Caption         =   "Terminated"
         Height          =   195
         Index           =   3
         Left            =   3000
         TabIndex        =   33
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdDStart 
         Appearance      =   0  'Flat
         Caption         =   "Update"
         Height          =   330
         Left            =   1800
         TabIndex        =   27
         Tag             =   "Terminate the Employee Selected"
         Top             =   2880
         Width           =   2220
      End
      Begin VB.CheckBox chkToSpouse 
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   960
         Width           =   375
      End
      Begin INFOHR_Controls.DateLookup dlpDeath 
         Height          =   285
         Left            =   1800
         TabIndex        =   22
         Tag             =   "41-Date of Death"
         Top             =   0
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin MSMask.MaskEdBox medAmtDeh 
         Height          =   285
         Left            =   2115
         TabIndex        =   25
         Tag             =   "20-Deemed PE "
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
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
         Format          =   "##0.00;(##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Deemed PE "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   29
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Who died?"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   28
         Tag             =   "41-Benefit End Date"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Is Pension Transferring to Spouse?"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   26
         Tag             =   "41-Benefit End Date"
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Death"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   23
         Tag             =   "41-Date of Death"
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Frame frmRetirement 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdRetireRet 
         Appearance      =   0  'Flat
         Caption         =   "Working Retiree Retirement"
         Height          =   330
         Left            =   1800
         TabIndex        =   32
         Tag             =   "Terminate the Employee Selected"
         Top             =   3600
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.CommandButton cmdRetireWorking 
         Appearance      =   0  'Flat
         Caption         =   "Retired - Working"
         Height          =   330
         Left            =   1800
         TabIndex        =   31
         Tag             =   "Terminate the Employee Selected"
         Top             =   3240
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.CommandButton cmdRetire 
         Appearance      =   0  'Flat
         Caption         =   "Retire the Employee"
         Height          =   330
         Left            =   1800
         TabIndex        =   11
         Tag             =   "Terminate the Employee Selected"
         Top             =   2880
         Width           =   2220
      End
      Begin INFOHR_Controls.DateLookup dlpLastWorkDate 
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Tag             =   "41-Last Day Worked"
         Top             =   0
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpRetireDate 
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Tag             =   "41-Retirement Date"
         Top             =   480
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin MSMask.MaskEdBox medAmount 
         Height          =   285
         Left            =   2115
         TabIndex        =   14
         Tag             =   "20-Deemed PE "
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
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
         Format          =   "##0.00;(##0.00)"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.CodeLookup clpBGroup 
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Tag             =   "01-Benefit - Group Code"
         Top             =   1440
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "BGMF"
         MaxLength       =   10
         SecurityMaintainable=   0
      End
      Begin INFOHR_Controls.DateLookup dlpDOther2 
         DataSource      =   " "
         Height          =   285
         Left            =   1800
         TabIndex        =   16
         Tag             =   "40-Other Date 2"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1045
      End
      Begin VB.Label lbOtherDate2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Other Date 2"
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
         Height          =   195
         Left            =   0
         TabIndex        =   30
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Retirement Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   20
         Tag             =   "41-Retirement Date"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Day Worked"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   19
         Tag             =   "41-Last Day Worked"
         Top             =   0
         Width           =   1470
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Deemed PE "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   18
         Top             =   960
         Width           =   900
      End
      Begin VB.Label lblBen 
         BackStyle       =   0  'Transparent
         Caption         =   "Benefit Group"
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   1440
         Width           =   1545
      End
   End
   Begin VB.TextBox txtEmpNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9120
      MaxLength       =   12
      TabIndex        =   7
      Tag             =   "Enter New Employee Number"
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtMain 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2235
      TabIndex        =   0
      Tag             =   "10-Enter Employee Number"
      Top             =   360
      Width           =   1500
   End
   Begin VB.Frame frmAT 
      Caption         =   "Employee Lookup"
      Height          =   615
      Left            =   2235
      TabIndex        =   4
      Top             =   750
      Width           =   2295
      Begin VB.OptionButton optActTerm 
         Caption         =   "Term"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optActTerm 
         Caption         =   "Active"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   9960
      TabIndex        =   1
      Tag             =   "10-Enter Employee Number"
      Top             =   6360
      Visible         =   0   'False
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin VB.Label lblEmpExist 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number Already Exist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   6720
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   3645
   End
   Begin VB.Label lblEmpno 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New Employee Number"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6720
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Label lblDesc 
      Caption         =   "*** NOT ATTACHED ***"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3960
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image imgIcon 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1800
      Picture         =   "frmERetirement.frx":0000
      Top             =   360
      Width           =   240
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
      TabIndex        =   5
      Top             =   360
      Width           =   1290
   End
End
Attribute VB_Name = "frmERetirement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsEmp As New ADODB.Recordset
Dim rsHRTRMEMP As New ADODB.Recordset
Dim rsDAT_Other As New ADODB.Recordset
Dim locWFCPenEligible As Boolean
Dim locUnion As String 'SEQID&
Dim xlocTERM_SEQ&, xlocEmpnbr, xlocPayrollID, xlocDOB
Dim IsSalaried As Boolean
Dim SaveBGroup As String
Dim NewBGroup As String
Dim AbortTerm As Boolean
Dim xPenStatus As String
Dim xCovClass As String, xBenAccount As String
Dim xTermDate
Dim isNGS As Boolean

Private Sub chkToSpouse_Click()
    'Spouse cannot be selected if "Pension Transferring to Spouse" is Yes.
    If chkToSpouse.Value Then
        If optDth(1).Value Then 'Ticket #23491 Franks 04/02/2013
            optDth(2).Value = True
        End If
        'optDth(0).Value = True
        optDth(1).Enabled = False
    Else
        optDth(1).Enabled = True
    End If
End Sub

Private Function getPenStatus()
Dim rsPen As New ADODB.Recordset
Dim SQLQ As String
Dim xSection, xPenType, xDBStatus, xSalHly
Dim xDOB, xEarlyRet, xNorRet, xLateRet, xDate, xYear, xSIN, xUnion
Dim rsTB As New ADODB.Recordset
Dim xstrTmp As String
Dim xEmpNo
Dim retval As String
    retval = ""

    If Not rsEmp.EOF Then
        If IsNull(rsEmp("ED_SIN")) Then GoTo end_line
        If IsNull(rsEmp("ED_SECTION")) Then GoTo end_line
        If IsNull(rsEmp("ED_DOH")) Then GoTo end_line
        If IsNull(rsEmp("ED_DOB")) Then GoTo end_line
        If IsNull(rsEmp("ED_ORG")) Then GoTo end_line
        If IsNull(rsEmp("ED_ELIGIBLE")) Then GoTo end_line
        
        xSIN = rsEmp("ED_SIN")
        xSection = rsEmp("ED_SECTION")
        xUnion = rsEmp("ED_ORG")
        'xYear = Year(dlpDeath.Text)
        xPenType = getDBType(xSection, xUnion, "PenType", rsEmp("ED_DOH")) 'Ticket #26707 Franks 02/25/2015
        xSalHly = getDBType(xSection, xUnion, "HlySal")
        If Len(xPenType) = 0 Then
             GoTo end_line
        End If
        If Len(xSalHly) = 0 Then
             GoTo end_line
        End If
    
        
        SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
        'SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
        SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
        SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
        If rsPen.State <> 0 Then rsPen.Close
        rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsPen.EOF Then
            If Not IsNull(rsPen("PE_DB_STATUS")) Then
                retval = rsPen("PE_DB_STATUS")
            End If
        End If
        rsPen.Close
    End If
end_line:
    getPenStatus = retval
End Function

Private Sub cmdDStart_Click()
'Dim rsEMP As New ADODB.Recordset
Dim rsPen As New ADODB.Recordset
'Dim rsPenPre As New ADODB.Recordset
Dim SQLQ As String
Dim xSection, xPenType, xDBStatus, xSalHly
Dim xDOB, xEarlyRet, xNorRet, xLateRet, xDate, xYear, xSIN, xUnion
Dim rsTB As New ADODB.Recordset
Dim rsT_PARCO As New ADODB.Recordset
Dim Msg$, DgDef As Variant, Response%
Dim Title$, EID&, TermDate$
Dim DelayProcess As Boolean
Dim xTimminsBenefits As Boolean
Dim xEarnedPen, xPAData
Dim xLastDay
Dim xstrTmp As String
Dim xEmpNo
    
    Msg$ = ""
    If Not chkDeathProc() Then Exit Sub

    Msg$ = Msg$ & "Are you sure you want to do the update?"
    
    Title$ = ("Death of an Employee/Spouse")
    DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
    
    If Response% = IDNO Then    ' Evaluate response
        Exit Sub
    End If
    
    If rsEmp.EOF Then
        Exit Sub
    End If


    
    xlocEmpnbr = getEmpnbr(txtMain.Text)
    xEmpNo = xlocEmpnbr
    
    If glbWFC Then 'Ticket #29624 Franks 01/09/2017
        If optDth(2).Value Then 'Employee only
            glbWFC_CancelTransaction = False
            Call CheckWFCReptAuthExistNew(xlocEmpnbr, dlpDeath)
            If glbWFC_CancelTransaction Then
                Exit Sub
            End If
        End If
    End If
    
    If optDth(0).Value Then ' Retiree
        rsEmp("ED_FNAME") = Left((rsEmp("ED_FNAME") & " (Deceased)"), 40)
    End If
    If optDth(2).Value Then ' Employee 'Ticket #23491 Franks 04/02/2013
        rsEmp("ED_FNAME") = Left((rsEmp("ED_FNAME") & " (Deceased)"), 40)
    End If
    If optDth(3).Value Then ' Terminated 'Ticket #23491 Franks 04/02/2013
        rsEmp("ED_FNAME") = Left((rsEmp("ED_FNAME") & " (Deceased)"), 40)
    End If
    If chkToSpouse.Value Then
        'Update Middle Name with beneficiary name in HR Demo screen,
        'this action only work if DB Beneficiary Relationship = Spouse
        xstrTmp = SpouseName4DB
        If Len(xstrTmp) > 0 Then
            rsEmp("ED_MIDNAME") = Left(xstrTmp, 30)
        End If
    End If
    
    rsEmp.Update
     
    'open IHREMP_OTHER
    Call OpenEMP_OTHER
    If Not rsDAT_Other.EOF Then
        rsDAT_Other("ER_PENSIONDATE5") = dlpDeath.Text
        rsDAT_Other("ER_LDATE") = Date
        rsDAT_Other("ER_LTIME") = Time$
        rsDAT_Other("ER_LUSER") = glbUserID
        rsDAT_Other.Update
    End If
        
    
    'Pension Eligible employees - begin
    If Not locWFCPenEligible Then
        GoTo ToSpouse_line
    End If
    If Not rsEmp.EOF Then
        If IsNull(rsEmp("ED_ELIGIBLE")) Then GoTo ToSpouse_line
        If IsNull(rsEmp("ED_SIN")) Then GoTo ToSpouse_line
        If IsNull(rsEmp("ED_SECTION")) Then GoTo ToSpouse_line
        If IsNull(rsEmp("ED_DOH")) Then GoTo ToSpouse_line
        If IsNull(rsEmp("ED_DOB")) Then GoTo ToSpouse_line
        If IsNull(rsEmp("ED_ORG")) Then GoTo ToSpouse_line
        
        
        xSIN = rsEmp("ED_SIN")
        xSection = rsEmp("ED_SECTION")
        xUnion = rsEmp("ED_ORG")
        xYear = Year(dlpDeath.Text)
        xPenType = getDBType(xSection, xUnion, "PenType", rsEmp("ED_DOH")) 'Ticket #26707 Franks 02/25/2015
        xSalHly = getDBType(xSection, xUnion, "HlySal")
        If Len(xPenType) = 0 Then
            GoTo ToSpouse_line
        End If
        If Len(xSalHly) = 0 Then
            GoTo ToSpouse_line
        End If
    
        
        If chkToSpouse.Value Then
            xDBStatus = "B"
        Else
            xDBStatus = "D"
        End If
        
        SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
        SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
        SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
        SQLQ = SQLQ & "AND PE_DB_STATUS = '" & xDBStatus & "'  "
        SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
        If rsPen.State <> 0 Then rsPen.Close
        rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsPen.EOF Then
            'Pen Master Audit -  Ticket #19954 - begin
            UpdPenAudDirect = False
            UpdPenAudit = True
            Call PenMasterAuditOldValSetup("Blank")
            toTYPE = "Add"
            'Pen Master Audit -  Ticket #19954 - end
            rsPen.Close
            'Call WFCPensionMaster(xEmpNo, "N", xDBStatus, , xYear, "UsePenStatus", "Y", xPenType)
            'Ticket #23459 Franks 03/26/2013 - change xCopyPreRecord from "Y" to "YY"
            Call WFCPensionMaster(xEmpNo, "N", xDBStatus, , xYear, "UsePenStatus", "YY", xPenType, , , xlocTERM_SEQ&)
            'reopen this recordset again
            SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
            SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
            SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
            SQLQ = SQLQ & "AND PE_DB_STATUS = '" & xDBStatus & "'  "
            SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
            If rsPen.State <> 0 Then rsPen.Close
            rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            'Pen Master Audit -  Ticket #19954 - begin
            Call PenMasterAuditOldValSetup("CurValues", rsPen)
            toTYPE = "Change"
            'Pen Master Audit -  Ticket #19954 - end
        End If
        If optDth(0).Value Then
            If InStr(1, UCase(rsPen("PE_FNAME")), UCase("Deceased")) = 0 Then
                rsPen("PE_FNAME") = Left((rsPen("PE_FNAME") & " (Deceased)"), 40)
            End If
            rsPen("PE_EXIT_DATE") = dlpDeath.Text
        End If
        If optDth(2).Value Then ' Employee 'Ticket #23491 Franks 04/02/2013
            If InStr(1, UCase(rsPen("PE_FNAME")), UCase("Deceased")) = 0 Then
                rsPen("PE_FNAME") = Left((rsPen("PE_FNAME") & " (Deceased)"), 40)
            End If
            rsPen("PE_EXIT_DATE") = dlpDeath.Text
            If xSalHly = "Salaried" Then 'Ticket #23505 Franks 04/03/2013
                xPAData = 0
                If IsNumeric(medAmtDeh.Text) Then xPAData = Val(medAmtDeh.Text)
                If xPAData > 0 Then 'xPAData -  Pensionable Earnings
                    xEarnedPen = xPAData * 0.009
                    rsPen("PE_YEAR_AMOUNT") = xEarnedPen
                End If
            End If
        End If
        If optDth(3).Value Then ' Terminated 'Ticket #23491 Franks 04/02/2013
            If InStr(1, UCase(rsPen("PE_FNAME")), UCase("Deceased")) = 0 Then
                rsPen("PE_FNAME") = Left((rsPen("PE_FNAME") & " (Deceased)"), 40)
            End If
            rsPen("PE_EXIT_DATE") = dlpDeath.Text
        End If
        rsPen("PE_DB_STATUS") = xDBStatus
        rsPen("PE_DB_STATUS_DATE") = dlpDeath.Text
        rsPen("PE_LUSER") = glbUserID
        rsPen("PE_LTIME") = Time$
        rsPen("PE_LDATE") = Date
        rsPen.Update
        'Pen Master Audit Ticket #19954
        toSOURCE = "IHR Death of Employee/Spouse"
        Call AUDIT_PENSION_MASTER(rsPen)
        rsPen.Close
            
        'Death of an Employee/Spouse on Other DB Pensions
        'One employee can have one DBS plus other DB pensions, such as DBKIPL
        'Employee Dan Dubblestyne had DBS and DBKIPL pensions
        'Pen Master Audit Ticket #19954
        toSOURCE = "IHR Death of Employee/Spouse"
        Call WFCOtherPenUpt(xEmpNo, xSIN, xYear, xSalHly, xPenType, "B", dlpDeath.Text)
            
    End If
    'Pension Eligible employees - end
    
    Call mod_Upd_Pos_Budget_WFC("", "", xEmpNo) 'Ticket #25911 Franks 03/2/2015
    
ToSpouse_line:
    If chkToSpouse.Value = 0 Then
        If optDth(1).Value Then 'Spouse
            'set date of death
            If optActTerm(0).Value Then 'Active
                SQLQ = "UPDATE HRBENS SET BD_DEATHDATE = " & Date_SQL(dlpDeath.Text) & " "
                SQLQ = SQLQ & "WHERE BD_EMPNBR = " & xEmpNo & " "
                SQLQ = SQLQ & "AND BD_RELATE = 'Spouse' "
                gdbAdoIhr001.Execute SQLQ
            End If
            If optActTerm(1).Value Then 'Term
                SQLQ = "UPDATE Term_HRBENS SET BD_DEATHDATE = " & Date_SQL(dlpDeath.Text) & " "
                SQLQ = SQLQ & "WHERE BD_EMPNBR = " & xEmpNo & " "
                SQLQ = SQLQ & "AND TERM_SEQ = " & glbTERM_Seq & " "
                SQLQ = SQLQ & "AND BD_RELATE = 'Spouse' "
                gdbAdoIhr001.Execute SQLQ
            End If
            'Spouse Pension Beneficiary
            SQLQ = "UPDATE HRP_PENSION_BENEFICIARY SET PE_DEATHDATE = " & Date_SQL(dlpDeath.Text) & " "
            SQLQ = SQLQ & "WHERE PE_SIN = '" & xSIN & "' " '
            SQLQ = SQLQ & "AND PE_BEN_RELATE = 'Spouse' "
            gdbAdoIhr001.Execute SQLQ
        End If
        
        If optDth(2).Value Then 'Ticket #23247 Franks 04/22/2014
            Call WFC_NGS_Trans(xEmpNo)
        End If
        
        'If optActTerm(0).Value Then
        'Ticket #23491 Franks 04/02/2013
        If optActTerm(0).Value Then 'Retiree or Employee
            'In info:HR, the employee record is automatically terminated
            'with Reason Code = "DECD" and the Date of Termination= Pension Date 5
            Call TerminateEmp(xEmpNo, dlpDeath.Text, "DECD")
        End If
    End If

    
    MDIMain.panHelp(0).FloodPercent = 0
    
    'MsgBox "   Done!   "
    Unload Me

End Sub

Private Sub cmdRetire_Click()
Dim rsTB As New ADODB.Recordset
Dim rsT_PARCO As New ADODB.Recordset
Dim Msg$, DgDef As Variant, Response%
Dim Title$, EID&, TermDate$
Dim SQLQ
Dim DelayProcess As Boolean
Dim xTimminsBenefits As Boolean
Dim xPAData
Dim xLastDay
Dim xPenType
Dim xOldEmp
    
    Msg$ = ""
    If Not chkRetire() Then Exit Sub

    If optActTerm(1).Value Then
        Msg$ = Msg$ & "This employee will be rehired from terminated employees first and then to be retired. " & Chr(10)
    End If
    Msg$ = Msg$ & "Are you sure you want to retire this employee ?"

    
    Title$ = ("Retire Employee")
    DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
    
    If Response% = IDNO Then    ' Evaluate response
        Exit Sub
    End If
    
    If rsEmp.EOF Then
        Exit Sub
    End If
    
    
    
    If optActTerm(1).Value Then
        If EmpRehire Then
        Else
            Exit Sub
        End If
        xlocEmpnbr = getEmpnbr(txtEmpNo.Text)
    Else 'active
        xlocEmpnbr = getEmpnbr(txtMain.Text)
    End If
    
    If glbWFC Then 'Ticket #29624 Franks 01/06/2017
        glbWFC_CancelTransaction = False
        Call CheckWFCReptAuthExistNew(xlocEmpnbr, dlpRetireDate.Text)
        If glbWFC_CancelTransaction Then
            Exit Sub
        End If
    End If

    
    'Ticket #21491 Franks 01/25/2012
    'add status change to Employee History table
    Call EmpHisCalc(1, xlocEmpnbr, "", "", "RET", "", "", "", "", dlpRetireDate.Text, , , dlpRetireDate.Text)
        
    'open IHREMP_OTHER
    Call OpenEMP_OTHER
    
    If IsSalaried Then
        rsEmp("ED_DIV") = "9001"
    Else
        rsEmp("ED_DIV") = "9005"
    End If
    rsEmp("ED_DIVEDATE") = dlpRetireDate.Text
    
    NewBGroup = clpBGroup.Text

    If SaveBGroup <> NewBGroup Then
        Msg = "Do you want add/update the Employee's Benefits "
        Msg = Msg & " with the Benefit Codes defined for the Benefit Group? "
        If MsgBox(Msg, 36, "info:HR") = 6 Then
            'Call UpdateBenefitGroup
            Call glbUpdateBenefitGroup(xlocEmpnbr, SaveBGroup, NewBGroup, dlpRetireDate.Text)
            DoEvents
            frmBENGRLIST.Show 1
        Else 'Frank 10/04/2003 Delete Benefit Group on Employee Benefit screen if wipe off the Benefit Group
            If Len(clpBGroup.Text) = 0 Then
                SQLQ = "UPDATE HRBENFT SET BF_GROUP = NULL WHERE NOT (BF_GROUP IS NULL) AND BF_EMPNBR =" & xlocEmpnbr
                gdbAdoIhr001.Execute SQLQ
            End If
        End If
        'SaveBGroup = clpBGroup.Text
        
        'If the Benefit Group changes, go to the Benefit Group Matrix to update
        'the HREMP 's Coverage Class and Benefit Account on the status/dates screen.
        If Len(NewBGroup) > 0 Then
            Call getValsFromBenGrpMatrix(NewBGroup, rsEmp("ED_DIV"))
            If Len(xBenAccount) = 0 Then
                rsEmp("ED_USER_NUM1") = Null
            Else
                rsEmp("ED_USER_NUM1") = xBenAccount
            End If
            If Len(xCovClass) = 0 Then
                rsEmp("ED_USER_TEXT2") = Null
            Else
                rsEmp("ED_USER_TEXT2") = xCovClass
            End If
        End If
    End If

    'Pension Status are Term Deferred & Active(T or A): Change ed_admin to "CPEN".
    If xPenStatus = "A" Or xPenStatus = "T" Then
        rsEmp("ED_ADMINBY") = "CPEN"
    End If
    If IsNull(rsEmp("ED_EMP")) Then xOldEmp = "" Else xOldEmp = rsEmp("ED_EMP")
    rsEmp("ED_EMP") = "RET"
    If IsNull(rsEmp("ED_LDAY")) Then
        rsEmp("ED_LDAY") = dlpLastWorkDate.Text
    End If
    If optActTerm(0).Value Then
        rsEmp("ED_SFDATE") = dlpRetireDate.Text
    Else
        If IsDate(xTermDate) Then
            rsEmp("ED_SFDATE") = xTermDate
        End If
    End If
    If Len(NewBGroup) = 0 Then
        rsEmp("ED_BENEFIT_GROUP") = Null
    Else
        rsEmp("ED_BENEFIT_GROUP") = NewBGroup
    End If
    'Ticket #20384 Franks 05/30/2011 - begin
    rsEmp("ED_BADGEID") = Null
    If IsNull(rsEmp("ED_BONUSDEPT")) Then
        rsEmp("ED_BONUSDEPT") = "000000"
    Else
        If Len(Trim(rsEmp("ED_BONUSDEPT"))) = 0 Then
            rsEmp("ED_BONUSDEPT") = "000000"
        Else
            'do not update it if there is value there
        End If
    End If
    'Ticket #20384 Franks 05/30/2011 - end
    rsEmp.Update
    

    If Not rsDAT_Other.EOF Then
        rsDAT_Other("ER_PENSIONDATE6") = dlpRetireDate.Text
        If optActTerm(1).Value Then 'term
            If IsDate(xTermDate) Then
                'If IsNull(rsDAT_Other("ER_PENSIONDATE4")) Then
                    rsDAT_Other("ER_PENSIONDATE4") = xTermDate
                'End If
            End If
        End If
        rsDAT_Other("ER_LDATE") = Date
        rsDAT_Other("ER_LTIME") = Time$
        rsDAT_Other("ER_LUSER") = glbUserID
        rsDAT_Other.Update
    End If
    
    'create Pension Master, PA Master and PA Details
    'Pension Status = "R"
    'Pension Effective Date = Date of Retirement
    If IsDate(dlpLastWorkDate.Text) Then
        xLastDay = dlpLastWorkDate.Text
    Else
        xLastDay = dlpRetireDate.Text
    End If
    xPAData = 0
    If Len(medAmount.Text) > 0 Then
        If IsNumeric(medAmount.Text) Then
            If Val(medAmount.Text) > 0 Then
                xPAData = medAmount.Text
            End If
        End If
    End If
    toSOURCE = "IHR Retirement" 'Ticket #19954
    If xPAData = 0 Then
        'No Pensionable Earnings, it can be blank for Terminated employees
        Call WFCPensionMasUpt(xlocEmpnbr, "Retirement", dlpRetireDate.Text, xLastDay, Year(CVDate(dlpRetireDate.Text)))
    Else
        Call WFCPensionMasUpt(xlocEmpnbr, "Retirement", dlpRetireDate.Text, xLastDay, Year(CVDate(dlpRetireDate.Text)), medAmount.Text)
    End If
    
    'retire other DB Pension
    'One employee can have one DBS plus other DB pensions, such as DBKIPL
    'Employee Dan Dubblestyne had DBS and DBKIPL pensions
    toSOURCE = "IHR Retirement" 'Ticket #19954
    xPenType = getDBType(locSection, locUnion, "PenType", GetEmpData(xlocEmpnbr, "ED_DOH")) 'Ticket #26707 Franks 02/25/2015
    Call WFCOtherPenUpt(xlocEmpnbr, glbSIN, Year(dlpRetireDate.Text), "", xPenType, "R", dlpRetireDate.Text, xLastDay, "DB")
    'retire other DC Pension
    Call WFCOtherPenUpt(xlocEmpnbr, glbSIN, Year(dlpRetireDate.Text), "", "", "R", dlpRetireDate.Text, xLastDay, "DC")

    Call WFCPensionAlerts(xlocEmpnbr, dlpRetireDate.Text, "Termination  - RET")

    If dlpDOther2.Visible Then 'Ticket #19266 Franks 12/03/10
        Call WFC_NGS_Trans(xlocEmpnbr)
    End If
    
    Call mod_Upd_Pos_Budget_WFC("", "", xlocEmpnbr) 'Ticket #25911 Franks 12/18/2014
    
    If gsEMAIL_ONTERM Then
        Call cmdEmailWFCPension
        
        If AbortTerm = True Then
            Screen.MousePointer = vbDefault
            MDIMain.panHelp(0).FloodType = 1
            MDIMain.panHelp(0).Caption = "Retirement Email Aborted"
            MsgBox "Error sending email.  Retirement Email aborted.", vbCritical + vbOKOnly, "Error"
            'Exit Sub
        End If
    End If
    MDIMain.panHelp(0).FloodPercent = 0
    
    'MsgBox "   Done!   "
    Unload Me

End Sub

Private Sub getValsFromBenGrpMatrix(NewBGroup, xDiv)
Dim rsBenGrpMrx As New ADODB.Recordset
Dim SQLQ As String
    xCovClass = ""
    xBenAccount = ""
    SQLQ = "SELECT * FROM HR_BENEFITS_GROUP_MATRIX WHERE BM_BENEFIT_GROUP = '" & NewBGroup & "' "
    If Len(xDiv) > 0 Then
        SQLQ = SQLQ & "AND BM_DIV = '" & xDiv & "' "
    End If
    rsBenGrpMrx.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsBenGrpMrx.EOF Then
        xCovClass = rsBenGrpMrx("BM_BENEFIT_CLASS")
        xBenAccount = rsBenGrpMrx("BM_BENEFIT_ACCOUNT")
    End If
    rsBenGrpMrx.Close
End Sub

Sub cmdEmailWFCPension()
    Dim rsPen As New ADODB.Recordset
    Dim MailBody As String
    Dim SecCode As String, SecDesc As String
    Dim PenType As String
    Dim UnionCode As String
    Dim SalHrl As String
    Dim xEmpNo As Double
    Dim SQLQ As String
    Dim xStr As String
    Dim xTmpVal As Double
    Dim xCredSer As Double
    Dim xContSer As Double
    Dim DBEarns, DBCR, DBCS, DBCalDB, DBCashout
    Dim DBEarnsHly, DBCRHly, DBCSHly, DBCalDBHly, DBCashoutHly
    Dim DCEarns, DCER, DCEE, DCCashout
    Dim xSalFlag As Boolean
    Dim xHlyFlag As Boolean
    Dim xDBList As String
    
    On Error GoTo ErrorHandler
    
    xDBList = "AND (LEFT(PE_PENSIONTYPE,2) = 'DB' OR LEFT(PE_PENSIONTYPE,3) = 'IDL' OR LEFT(PE_PENSIONTYPE,3) = 'UPG' OR LEFT(PE_PENSIONTYPE,3) = 'PRE' OR PE_PENSIONTYPE = 'DBSUP' OR PE_PENSIONTYPE = 'MON' ) "

    'Exit Sub
    Load frmSendEmail

    frmSendEmail.txtSubject.Text = "info:HR Retirement Notice - " & lblDesc.Caption
    MailBody = "The employee below has been retired." & vbCrLf & vbCrLf
    MailBody = MailBody & "Employee #: " & txtMain.Text & vbCrLf
    MailBody = MailBody & "Name: " & lblDesc.Caption & vbCrLf
    frmSendEmail.txtTo.Text = "pension@woodbridgegroup.com"

    xEmpNo = txtMain.Text
    SecCode = GetEmpData(xEmpNo, "ED_SECTION")
    UnionCode = GetEmpData(xEmpNo, "ED_ORG")
    SecDesc = GetTABLDesc("EDSE", SecCode)
    SalHrl = GetSalHourly(SecCode, UnionCode)
    PenType = GetPensionType(SecCode, UnionCode)

    MailBody = MailBody & lStr("Section: ") & " - " & SecDesc & vbCrLf
    MailBody = MailBody & "Salaried/Hourly: " & SalHrl & vbCrLf

    MailBody = MailBody & "Date: " & dlpRetireDate.Text & vbCrLf & vbCrLf

    'DB Pensions - Salaried
    xSalFlag = False
    DBEarns = 0: DBCR = 0: DBCS = 0: DBCalDB = 0: DBCashout = 0
    SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & glbSIN & "' "
    SQLQ = SQLQ & xDBList
    SQLQ = SQLQ & "AND PE_HRLYSAL = 'Salaried' "
    SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC, PE_PENSIONTYPE "
    rsPen.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Do While Not rsPen.EOF
        xSalFlag = True
        If Not IsNull(rsPen("PE_YEAR_AMOUNT")) Then
            DBEarns = DBEarns + rsPen("PE_YEAR_AMOUNT")
        End If
        If Not IsNull(rsPen("PE_CREDITED_SERV")) Then
            DBCR = DBCR + rsPen("PE_CREDITED_SERV")
        End If
        If Not IsNull(rsPen("PE_CONT_SERV")) Then
            DBCS = DBCS + rsPen("PE_CONT_SERV")
        End If
        If Not IsNull(rsPen("PE_ANNDEFERRED")) Then
            DBCalDB = DBCalDB + rsPen("PE_ANNDEFERRED")
        End If
        If Not IsNull(rsPen("PE_PAYOUT_VALUE")) Then
            DBCashout = DBCashout + rsPen("PE_PAYOUT_VALUE")
        End If
        rsPen.MoveNext
    Loop
    rsPen.Close
    'If DBEarns > 0 Then
    If xSalFlag Then
        xStr = "Salaried DB Pensions:"
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Earned Pension: " & "$" & DBEarns
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Credited Service: " & Format((DBCR / 12), "#0.0000") & " (in yrs)"
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Cont. Service: " & Format((DBCS / 12), "#0.0000") & " (in yrs)"
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Calculated Pension: " & "$" & DBCalDB
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Cashed Out: " & "$" & DBCashout
        MailBody = MailBody & xStr & vbCrLf & vbCrLf
    End If
    
    'DB Pensions - Hourly
    xHlyFlag = False
    DBEarnsHly = 0: DBCRHly = 0: DBCSHly = 0: DBCalDBHly = 0: DBCashoutHly = 0
    SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & glbSIN & "' "
    'SQLQ = SQLQ & "AND (LEFT(PE_PENSIONTYPE,2) = 'DB' OR LEFT(PE_PENSIONTYPE,3) = 'IDL' OR LEFT(PE_PENSIONTYPE,3) = 'UPG' OR PE_PENSIONTYPE = 'DBSUP' OR PE_PENSIONTYPE = 'MON' ) "
    SQLQ = SQLQ & xDBList
    SQLQ = SQLQ & "AND PE_HRLYSAL = 'Hourly' "
    SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC, PE_PENSIONTYPE "
    rsPen.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Do While Not rsPen.EOF
        xHlyFlag = True
        If Not IsNull(rsPen("PE_YEAR_AMOUNT")) Then
            DBEarnsHly = DBEarnsHly + rsPen("PE_YEAR_AMOUNT")
        End If
        If Not IsNull(rsPen("PE_CREDITED_SERV")) Then
            DBCRHly = DBCRHly + rsPen("PE_CREDITED_SERV")
        End If
        If Not IsNull(rsPen("PE_CONT_SERV")) Then
            DBCSHly = DBCSHly + rsPen("PE_CONT_SERV")
        End If
        If Not IsNull(rsPen("PE_ANNDEFERRED")) Then
            DBCalDBHly = DBCalDBHly + rsPen("PE_ANNDEFERRED")
        End If
        If Not IsNull(rsPen("PE_PAYOUT_VALUE")) Then
            DBCashoutHly = DBCashoutHly + rsPen("PE_PAYOUT_VALUE")
        End If
        rsPen.MoveNext
    Loop
    rsPen.Close
    'If DBEarnsHly > 0 Then
    If xHlyFlag Then
        xStr = "Hourly DB Pensions:"
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Earned Pension: " & "$" & DBEarnsHly
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Credited Service: " & Format((DBCRHly / 12), "#0.0000") & " (in yrs)"
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Cont. Service: " & Format((DBCSHly / 12), "#0.0000") & " (in yrs)"
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Calculated Pension: " & "$" & DBCalDBHly
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Cashed Out: " & "$" & DBCashoutHly
        MailBody = MailBody & xStr & vbCrLf & vbCrLf
    End If
    
    'DC Pension
    DCEarns = 0: DCER = 0: DCEE = 0: DCCashout = 0
    SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_EMPNBR = " & xEmpNo & " "
    If Len(SecCode) > 0 Then
        SQLQ = SQLQ & "AND PE_SECTION = '" & SecCode & "' "
    End If
    SQLQ = SQLQ & "AND PE_PENSIONTYPE = 'DC' "
    SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC, PE_PENSIONTYPE "
    rsPen.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Do While Not rsPen.EOF
        If Not IsNull(rsPen("PE_YEAR_AMOUNT")) Then
            DCER = DCER + rsPen("PE_YEAR_AMOUNT")
        End If
        If Not IsNull(rsPen("PE_MEM_DOLLAR")) Then
            DCEE = DCEE + rsPen("PE_MEM_DOLLAR")
        End If
        If Not IsNull(rsPen("PE_PAYOUT_VALUE")) Then
            DCCashout = DCCashout + rsPen("PE_PAYOUT_VALUE")
        End If
        rsPen.MoveNext
    Loop
    rsPen.Close
    If DCER > 0 Then
        xStr = "DC Pensions:"
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Employer Portion: " & "$" & DCER
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Employee Portion: " & "$" & DCEE
        MailBody = MailBody & xStr & vbCrLf
        xStr = "    Cashed Out: " & "$" & DCCashout
        MailBody = MailBody & xStr & vbCrLf
    End If
    
    frmSendEmail.txtBody.Text = MailBody

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = "Sending email..."
    frmSendEmail.Tag = ""
    frmSendEmail.cmdSend_Click
    Do
        DoEvents
    Loop Until frmSendEmail.Tag <> ""   ' MC - dkostka - 05/03/01 - Changed from = "DONE" to <> ""
    ' AC - dkostka - 05/03/01 - Added checking to make sure the email went through,
    '   otherwise refuse to terminate the employee.
    If frmSendEmail.Tag = "DONE" Then
        Unload frmSendEmail
        AbortTerm = False
    Else
        Unload frmSendEmail
        AbortTerm = True
    End If
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(0).FloodType = 1

exH:
    Exit Sub
    
ErrorHandler:
    If Err.Number = 364 Then Exit Sub
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Resume exH

End Sub

Private Function chkDeathProc()
Dim rsTB As New ADODB.Recordset
Dim dd As Integer
Dim xStr As String
chkDeathProc = False

If Len(txtMain.Text) < 1 Then
    MsgBox ("Employee Number is a required field")
    txtMain.SetFocus
    Exit Function
Else
    If lblDesc.Caption = "Unassigned" Then
        MsgBox ("Invalid Employee Number.")
        txtMain.SetFocus
        Exit Function
    End If
End If
If rsEmp.EOF Then
    MsgBox ("Invalid Employee Number.")
    txtMain.SetFocus
    Exit Function
End If

'Ticket #23505 Franks 04/03/2013 - This function works for all employees in all countries
'If Not locWFCPenEligible Then
'    MsgBox lStr("Employment Type") & (" is not Yes. " & Chr(10) & "Cannot do Death of an Employee/Spouse for this employee")
'    txtMain.SetFocus
'    Exit Function
'End If

If optDth(0).Value Then ''Ticket #23491 Franks 04/02/2013
    If Not (rsEmp("ED_EMP") = "RET") Then
        MsgBox "Employment Status is not 'RET'" & Chr(10) & "Cannot do Death of an Employee/Spouse for this employee"
        txtMain.SetFocus
        Exit Function
    End If
End If

If optDth(2).Value Then 'Employee -Ticket #23491 Franks 04/02/2013
    If (rsEmp("ED_EMP") = "RET") Then
        MsgBox "Employment Status is 'RET'" & Chr(10) & "Cannot do Death of an Employee/Spouse for this employee"
        txtMain.SetFocus
        Exit Function
    End If
    'Ticket #23565 Franks 04/10/2013
    'check if this employee can be found in active list
    If Not IsEmpExist(txtMain.Text, "A") Then
        MsgBox "Cannot find this employee from Active Employee List"
        txtMain.SetFocus
        Exit Function
    End If
    If locWFCPenEligible Then
        If IsSalaried Then
            If Len(medAmtDeh.Text) = 0 Then
                MsgBox "Deemed PE is required for Salaried Pension Employees"
                medAmtDeh.SetFocus
                Exit Function
            End If
            If Not IsNumeric(medAmtDeh.Text) Then
                MsgBox "Invalid Deemed PE"
                medAmtDeh.SetFocus
                Exit Function
            End If
        End If
    End If
End If
If optDth(3).Value Then 'Employee -Ticket #23491 Franks 04/02/2013
    If (rsEmp("ED_EMP") = "RET") Then
        MsgBox "Employment Status is 'RET'" & Chr(10) & "Cannot do Death of an Employee/Spouse for this employee"
        txtMain.SetFocus
        Exit Function
    End If
    'Ticket #23565 Franks 04/10/2013
    'check if this employee can be found from term list
    If Not IsEmpExist(txtMain.Text, "T") Then
        MsgBox "Cannot find this employee from Terminated Employee List"
        txtMain.SetFocus
        Exit Function
    End If
End If
If Len(dlpDeath.Text) < 1 Then
    MsgBox ("Date of Death is a required field")
    dlpDeath.SetFocus
    Exit Function
End If
If Not IsDate(dlpDeath.Text) Then
    MsgBox ("Date of Death is not valid.")
    dlpDeath.SetFocus
    Exit Function
End If

'If optDth(0).Value Then 'Retiree only
'Ticket #23491 Franks 04/02/2013
If optDth(0).Value Or optDth(2).Value Or optDth(3).Value Then 'Retiree, Employee and Terminated
'check First Name
    If InStr(1, UCase(glbLEE_FName), UCase("Deceased")) > 0 Then
        MsgBox "Please see the employee first name: " & glbLEE_FName
        txtMain.SetFocus
        Exit Function
    End If
End If

chkDeathProc = True


End Function
Private Function chkRetWorking()
Dim rsTB As New ADODB.Recordset
Dim dd As Integer
Dim xStr As String
Dim xYear

chkRetWorking = False

If Len(txtMain.Text) < 1 Then
    MsgBox ("Employee Number is a required field")
    txtMain.SetFocus
    Exit Function
Else
    If lblDesc.Caption = "Unassigned" Then
        MsgBox ("Invalid Employee Number.")
        txtMain.SetFocus
        Exit Function
    End If
End If
If rsEmp.EOF Then
    MsgBox ("Invalid Employee Number.")
    txtMain.SetFocus
    Exit Function
End If
If Not locWFCPenEligible Then
    MsgBox lStr("Employment Type") & (" is not Yes. " & Chr(10) & "Cannot do Retirement for this employee")
    txtMain.SetFocus
    Exit Function
End If
If rsEmp("ED_EMP") = "ACP" Or rsEmp("ED_EMP") = "RET" Then
    If rsEmp("ED_EMP") = "ACP" Then
        MsgBox "Employment Status is already 'ACP'" '& Chr(10) & "Cannot do Retirement for this employee"
    End If
    If rsEmp("ED_EMP") = "RET" Then
        MsgBox "Employment Status is already 'RET'" '& Chr(10) & "Cannot do Retirement for this employee"
    End If
    txtMain.SetFocus
    Exit Function
End If

'Benefit Group should be left blank but made mandatory for Eligible for Pension = "Y".
If isNGS Then  ' dlpDOther2.Visible Then 'Ticket #19266 Franks 12/03/10
'NGS employees will not need to enter "Deemed PE" or "Benefit Group".
Else
    ''If locWFCPenEligible Then
    ''    If Len(clpBGroup.Text) = 0 Then
    ''        MsgBox lStr("Employment Type") & (" is Yes. " & Chr(10) & "Benefit Group is required.")
    ''        clpBGroup.SetFocus
    ''        Exit Function
    ''    End If
    ''End If
End If

If optActTerm(0).Value Then 'active
    If Len(dlpLastWorkDate.Text) < 1 Then
        MsgBox ("Pension Exit Date is a required field for Active employee")
        dlpLastWorkDate.SetFocus
        Exit Function
    End If
    If Not IsDate(dlpLastWorkDate.Text) Then
        MsgBox ("Pension Exit Date is not a valid date.")
        dlpLastWorkDate.SetFocus
        Exit Function
    End If
End If
If Len(dlpRetireDate.Text) < 1 Then
    MsgBox ("Retirement Date is a required field")
    dlpRetireDate.SetFocus
    Exit Function
End If
If Not IsDate(dlpRetireDate.Text) Then
    MsgBox ("Retirement Date is not a valid date.")
    dlpRetireDate.SetFocus
    Exit Function
End If

'Calculated employee's age using the Pension Exit Date.
'--If employee's age is less than 65 as of the Pension Exit Date, display a message saying "An employee cannot retire and continue to work unless their age is 65 or greater."

'Calculated employee's age using the Retirement Date.
'If they are less than 55, display a message saying "Employee has not reached age 55 yet. They cannot retire.".
If IsDate(xlocDOB) Then
    'Ticket #22285 Franks 07/17/2012
    xYear = DateDiff("d", CVDate(xlocDOB), CVDate(dlpLastWorkDate.Text))
    xYear = Round(xYear / 365, 1)
    If xYear < 65 Then
        MsgBox "An employee cannot retire and continue to work unless their age is 65 or greater."
        dlpLastWorkDate.SetFocus
        Exit Function
    End If

    xYear = DateDiff("d", CVDate(xlocDOB), CVDate(dlpRetireDate.Text))
    xYear = Round(xYear / 365, 1)
    If xYear < 55 Then
        MsgBox "Employee has not reached age 55 yet. They cannot retire."
        dlpRetireDate.SetFocus
        Exit Function
    End If
End If

'Check Pension Status. If active employee, only employees which Pension Status
'of A, L, I, S, W, X. If terminated, only status available to retire is T.
'Pension Changes - June 1-2010(Jul009).docx
xPenStatus = getPenStatus
If optActTerm(0).Value Then 'Active
    'Ticket #21544 Franks 02/13/2012 Add M to the list
    If Not (xPenStatus = "A" Or xPenStatus = "L" Or xPenStatus = "I" Or xPenStatus = "S" Or xPenStatus = "W" Or xPenStatus = "X" Or xPenStatus = "M") Then
        MsgBox "Pension Status is not A, L,M, I, S, W, X." & Chr(10) & "They cannot retire."
        Exit Function
    End If
End If
''If optActTerm(1).Value Then 'Term
''    If Not (xPenStatus = "T") Then
''        MsgBox "Pension Status is not T." & Chr(10) & "They cannot retire."
''        Exit Function
''    End If
''End If

'If salaried, user must enter "Deemed PE " ("Pensionable Earnings")
'By default, sum the total Payroll Transaction records for Pay Code DN49 of the current year.
'If none are found, the user must enter a value greater than zero.
'--more Pension Changes - June 1-2010(Jul009).docx
'--Retirement Function - Term Deferred:
'--Pensionable Earnings should be optional. If entered, the Pension Master, PA Detail and Master record needs to be updated using the same logic as the Retirement from Active does.
If IsSalaried Then
    If optActTerm(0).Value Then
        If dlpDOther2.Visible Then 'Ticket #19266 Franks 12/03/10
        'NGS employees will not need to enter "Deemed PE" or "Benefit Group".
        Else
            If Not IsNumeric(medAmount.Text) Then
                If xPenStatus = "T" Then
                Else
                    'MsgBox ("Pensionable Earnings is required since the Union is " & locUnion)
                    MsgBox ("Deemed PE is required since the Union is " & locUnion)
                    medAmount.SetFocus
                    Exit Function
                End If
            Else
                If Val(medAmount.Text) <= 0 Then
                    'MsgBox ("Pensionable Earnings must be greater than zero since the Union is " & locUnion)
                    MsgBox ("Deemed PE must be greater than zero since the Union is " & locUnion)
                    medAmount.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
End If



'If Not chkBenefit.Value Then
'    MsgBox "'Benefits Correct?' is not checked."
'    chkBenefit.SetFocus
'    Exit Function
'End If
'If Not chkBenefic.Value Then
'    MsgBox "'Beneficiaries Correct?' is not checked."
'    chkBenefic.SetFocus
'    Exit Function
'End If

''terminated
'If optActTerm(1).Value Then
'End If

chkRetWorking = True


End Function
Private Function chkRetRetiree()
Dim rsTB As New ADODB.Recordset
Dim dd As Integer
Dim xStr As String
Dim xYear

chkRetRetiree = False

If Len(txtMain.Text) < 1 Then
    MsgBox ("Employee Number is a required field")
    txtMain.SetFocus
    Exit Function
Else
    If lblDesc.Caption = "Unassigned" Then
        MsgBox ("Invalid Employee Number.")
        txtMain.SetFocus
        Exit Function
    End If
End If
If rsEmp.EOF Then
    MsgBox ("Invalid Employee Number.")
    txtMain.SetFocus
    Exit Function
End If
''If Not locWFCPenEligible Then
''    MsgBox lStr("Employment Type") & (" is not Yes. " & Chr(10) & "Cannot do Retirement for this employee")
''    txtMain.SetFocus
''    Exit Function
''End If
If rsEmp("ED_EMP") = "RET" Then
    MsgBox "Employment Status is already 'RET'" '& Chr(10) & "Cannot do Retirement for this employee"
    txtMain.SetFocus
    Exit Function
End If

'Ticket #22321 Franks 07/24/2012
If Not rsEmp("ED_EMP") = "ACP" Then
    MsgBox "Employment Status is not 'ACP'" & Chr(10) & "Cannot do Working Retiree Retirement for this employee"
    txtMain.SetFocus
    Exit Function
End If

'Benefit Group should be left blank but made mandatory for Eligible for Pension = "Y".
If dlpDOther2.Visible Then 'Ticket #19266 Franks 12/03/10
'NGS employees will not need to enter "Deemed PE" or "Benefit Group".
Else
    If locWFCPenEligible Then
        'If Len(clpBGroup.Text) = 0 Then
        '    MsgBox lStr("Employment Type") & (" is Yes. " & Chr(10) & "Benefit Group is required.")
        '    clpBGroup.SetFocus
        '    Exit Function
        'End If
    End If
End If

If optActTerm(0).Value Then 'active
    If Len(dlpLastWorkDate.Text) < 1 Then
        MsgBox ("Last Day Worked is a required field for Active employee")
        dlpLastWorkDate.SetFocus
        Exit Function
    End If
    If Not IsDate(dlpLastWorkDate.Text) Then
        MsgBox ("Last Day Worked is not a valid date.")
        dlpLastWorkDate.SetFocus
        Exit Function
    End If
End If


''Calculated employee's age using the Retirement Date.
''If they are less than 55, display a message saying "Employee has not reached age 55 yet. They cannot retire.".
'If IsDate(xlocDOB) Then
'    xYear = DateDiff("d", CVDate(xlocDOB), CVDate(dlpRetireDate.Text))
'    xYear = Round(xYear / 365, 1)
'    If xYear < 55 Then
'        MsgBox "Employee has not reached age 55 yet. They cannot retire."
'        dlpRetireDate.SetFocus
'        Exit Function
'    End If
'End If

If dlpDOther2.Visible Then 'Ticket #19266 Franks 12/03/10
    If Not IsDate(dlpDOther2.Text) Then
        MsgBox lStr("Other Date 2") & " is required field"
        dlpDOther2.SetFocus
        Exit Function
    End If
End If

chkRetRetiree = True

End Function
Private Function chkRetire()
Dim rsTB As New ADODB.Recordset
Dim dd As Integer
Dim xStr As String
Dim xYear

chkRetire = False

If Len(txtMain.Text) < 1 Then
    MsgBox ("Employee Number is a required field")
    txtMain.SetFocus
    Exit Function
Else
    If lblDesc.Caption = "Unassigned" Then
        MsgBox ("Invalid Employee Number.")
        txtMain.SetFocus
        Exit Function
    End If
End If
If rsEmp.EOF Then
    MsgBox ("Invalid Employee Number.")
    txtMain.SetFocus
    Exit Function
End If
If Not locWFCPenEligible Then
    MsgBox lStr("Employment Type") & (" is not Yes. " & Chr(10) & "Cannot do Retirement for this employee")
    txtMain.SetFocus
    Exit Function
End If
If rsEmp("ED_EMP") = "RET" Then
    MsgBox "Employment Status is already 'RET'" '& Chr(10) & "Cannot do Retirement for this employee"
    txtMain.SetFocus
    Exit Function
End If

'Benefit Group should be left blank but made mandatory for Eligible for Pension = "Y".
If dlpDOther2.Visible Then 'Ticket #19266 Franks 12/03/10
'NGS employees will not need to enter "Deemed PE" or "Benefit Group".
Else
    If locWFCPenEligible Then
        If Len(clpBGroup.Text) = 0 Then
            MsgBox lStr("Employment Type") & (" is Yes. " & Chr(10) & "Benefit Group is required.")
            clpBGroup.SetFocus
            Exit Function
        End If
    End If
End If

If optActTerm(0).Value Then 'active
    If Len(dlpLastWorkDate.Text) < 1 Then
        MsgBox ("Last Day Worked is a required field for Active employee")
        dlpLastWorkDate.SetFocus
        Exit Function
    End If
    If Not IsDate(dlpLastWorkDate.Text) Then
        MsgBox ("Last Day Worked is not a valid date.")
        dlpLastWorkDate.SetFocus
        Exit Function
    End If
End If
If Len(dlpRetireDate.Text) < 1 Then
    MsgBox ("Retirement Date is a required field")
    dlpRetireDate.SetFocus
    Exit Function
End If
If Not IsDate(dlpRetireDate.Text) Then
    MsgBox ("Retirement Date is not a valid date.")
    dlpRetireDate.SetFocus
    Exit Function
End If

'Calculated employee's age using the Retirement Date.
'If they are less than 55, display a message saying "Employee has not reached age 55 yet. They cannot retire.".
If IsDate(xlocDOB) Then
    xYear = DateDiff("d", CVDate(xlocDOB), CVDate(dlpRetireDate.Text))
    xYear = Round(xYear / 365, 1)
    If xYear < 55 Then
        MsgBox "Employee has not reached age 55 yet. They cannot retire."
        dlpRetireDate.SetFocus
        Exit Function
    End If
End If

'Check Pension Status. If active employee, only employees which Pension Status
'of A, L, I, S, W, X. If terminated, only status available to retire is T.
'Pension Changes - June 1-2010(Jul009).docx
xPenStatus = getPenStatus
If optActTerm(0).Value Then 'Active
    'Ticket #21544 Franks 02/13/2012 Add M to the list
    If Not (xPenStatus = "A" Or xPenStatus = "L" Or xPenStatus = "I" Or xPenStatus = "S" Or xPenStatus = "W" Or xPenStatus = "X" Or xPenStatus = "M") Then
        MsgBox "Pension Status is not A, L,M, I, S, W, X." & Chr(10) & "They cannot retire."
        Exit Function
    End If
End If
If optActTerm(1).Value Then 'Term
    If Not (xPenStatus = "T") Then
        MsgBox "Pension Status is not T." & Chr(10) & "They cannot retire."
        Exit Function
    End If
End If

'If salaried, user must enter "Deemed PE " ("Pensionable Earnings")
'By default, sum the total Payroll Transaction records for Pay Code DN49 of the current year.
'If none are found, the user must enter a value greater than zero.
'--more Pension Changes - June 1-2010(Jul009).docx
'--Retirement Function - Term Deferred:
'--Pensionable Earnings should be optional. If entered, the Pension Master, PA Detail and Master record needs to be updated using the same logic as the Retirement from Active does.
If IsSalaried Then
    If optActTerm(0).Value Then
        If dlpDOther2.Visible Then 'Ticket #19266 Franks 12/03/10
        'NGS employees will not need to enter "Deemed PE" or "Benefit Group".
        Else
            If Not IsNumeric(medAmount.Text) Then
                If xPenStatus = "T" Then
                Else
                    'MsgBox ("Pensionable Earnings is required since the Union is " & locUnion)
                    MsgBox ("Deemed PE is required since the Union is " & locUnion)
                    medAmount.SetFocus
                    Exit Function
                End If
            Else
                If Val(medAmount.Text) <= 0 Then
                    'MsgBox ("Pensionable Earnings must be greater than zero since the Union is " & locUnion)
                    MsgBox ("Deemed PE must be greater than zero since the Union is " & locUnion)
                    medAmount.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
End If



'If Not chkBenefit.Value Then
'    MsgBox "'Benefits Correct?' is not checked."
'    chkBenefit.SetFocus
'    Exit Function
'End If
'If Not chkBenefic.Value Then
'    MsgBox "'Beneficiaries Correct?' is not checked."
'    chkBenefic.SetFocus
'    Exit Function
'End If

'terminated
If optActTerm(1).Value Then
    lblEmpExist.Caption = ""
    lblEmpExist.Visible = False
    'check if it was rehired before
    If Not rsHRTRMEMP.EOF Then
        If Not IsNull(rsHRTRMEMP("Term_DOR")) Then
            If IsDate(rsHRTRMEMP("Term_DOR")) Then
                MsgBox "This record has beed rehired before." & Chr(10) & "You can not retire this employee. "
                txtMain.SetFocus
                Exit Function
            End If
        End If
    End If
    'check the employee number
    If Len(txtEmpNo) = 0 Then
        MsgBox "Employee Number Missing"
        txtEmpNo.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtEmpNo) Then
        MsgBox "Invalid Employee Number"
        txtEmpNo.SetFocus
        Exit Function
    End If
    
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR = " & txtEmpNo.Text & " "
    rsTB.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    If Not rsTB.EOF Then
        'xStr = "Duplicated employee # found in active employee list. " & Chr(10) & "Please go to Rehire function to rehire this employee with a new employee #" & Chr(10)
        'xStr = xStr & "And then do this Employee Retirement"
        'MsgBox xStr
        lblEmpExist.Caption = " Employee # " & txtEmpNo.Text & " already active - A NEW Number is required"
        lblEmpExist.Visible = True
        txtEmpNo.SetFocus
        Exit Function
    End If
    rsTB.Close
    
End If

If dlpDOther2.Visible Then 'Ticket #19266 Franks 12/03/10
    If Not IsDate(dlpDOther2.Text) Then
        MsgBox lStr("Other Date 2") & " is required field"
        dlpDOther2.SetFocus
        Exit Function
    End If
End If

chkRetire = True

End Function


Private Sub cmdRetireRet_Click()
Dim rsTB As New ADODB.Recordset
Dim rsT_PARCO As New ADODB.Recordset
Dim Msg$, DgDef As Variant, Response%
Dim Title$, EID&, TermDate$
Dim SQLQ
Dim DelayProcess As Boolean
Dim xTimminsBenefits As Boolean
Dim xPAData
Dim xLastDay
Dim xPenType
Dim xOldEmp
    
    Msg$ = ""
    If Not chkRetRetiree() Then Exit Sub

    Msg$ = Msg$ & "Are you sure you want to retire this employee from ACP to RET ?"

    Title$ = ("Retire Employee")
    DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
    
    If Response% = IDNO Then    ' Evaluate response
        Exit Sub
    End If
    
    If rsEmp.EOF Then
        Exit Sub
    End If
    
    xLastDay = dlpLastWorkDate.Text
    xLastDay = DateAdd("d", 1, xLastDay)
    
    If optActTerm(1).Value Then
    Else 'active
        xlocEmpnbr = getEmpnbr(txtMain.Text)
    End If
    
    If glbWFC Then 'Ticket #29624 Franks 01/06/2017
        glbWFC_CancelTransaction = False
        Call CheckWFCReptAuthExistNew(xlocEmpnbr, dlpRetireDate.Text)
        If glbWFC_CancelTransaction Then
            Exit Sub
        End If
    End If
    
    'Ticket #21491 Franks 01/25/2012
    'add status change to Employee History table
    Call EmpHisCalc(1, xlocEmpnbr, "", "", "RET", "", "", "", "", xLastDay, , , xLastDay)
        
    'open IHREMP_OTHER
    Call OpenEMP_OTHER
    
    If IsSalaried Then
        rsEmp("ED_DIV") = "9001"
    Else
        rsEmp("ED_DIV") = "9005"
    End If
    'rsEmp("ED_DIVEDATE") = dlpRetireDate.Text
    
    NewBGroup = clpBGroup.Text

    If SaveBGroup <> NewBGroup And Len(NewBGroup) > 0 Then
        Msg = "Do you want add/update the Employee's Benefits "
        Msg = Msg & " with the Benefit Codes defined for the Benefit Group? "
        If MsgBox(Msg, 36, "info:HR") = 6 Then
            'Call UpdateBenefitGroup
            Call glbUpdateBenefitGroup(xlocEmpnbr, SaveBGroup, NewBGroup, dlpRetireDate.Text)
            DoEvents
            frmBENGRLIST.Show 1
        Else 'Frank 10/04/2003 Delete Benefit Group on Employee Benefit screen if wipe off the Benefit Group
            If Len(clpBGroup.Text) = 0 Then
                SQLQ = "UPDATE HRBENFT SET BF_GROUP = NULL WHERE NOT (BF_GROUP IS NULL) AND BF_EMPNBR =" & xlocEmpnbr
                gdbAdoIhr001.Execute SQLQ
            End If
        End If
        'SaveBGroup = clpBGroup.Text
        
        'If the Benefit Group changes, go to the Benefit Group Matrix to update
        'the HREMP 's Coverage Class and Benefit Account on the status/dates screen.
        If Len(NewBGroup) > 0 Then
            Call getValsFromBenGrpMatrix(NewBGroup, rsEmp("ED_DIV"))
            If Len(xBenAccount) = 0 Then
                rsEmp("ED_USER_NUM1") = Null
            Else
                rsEmp("ED_USER_NUM1") = xBenAccount
            End If
            If Len(xCovClass) = 0 Then
                rsEmp("ED_USER_TEXT2") = Null
            Else
                rsEmp("ED_USER_TEXT2") = xCovClass
            End If
        End If
    End If

    'Pension Status are Term Deferred & Active(T or A): Change ed_admin to "CPEN".
    'If xPenStatus = "A" Or xPenStatus = "T" Then
        rsEmp("ED_ADMINBY") = "CPEN"
    'End If
    If IsNull(rsEmp("ED_EMP")) Then xOldEmp = "" Else xOldEmp = rsEmp("ED_EMP")
    rsEmp("ED_EMP") = "RET"
    'If IsNull(rsEMP("ED_LDAY")) Then
    If IsDate(dlpLastWorkDate.Text) Then
        rsEmp("ED_LDAY") = CVDate(dlpLastWorkDate.Text)
    End If
    If optActTerm(0).Value Then
        rsEmp("ED_SFDATE") = xLastDay 'dlpRetireDate.Text
    'Else
    '    If IsDate(xTERMDATE) Then
    '        rsEmp("ED_SFDATE") = xTERMDATE
    '    End If
    End If
    If Len(NewBGroup) = 0 Then
        rsEmp("ED_BENEFIT_GROUP") = Null
    Else
        rsEmp("ED_BENEFIT_GROUP") = NewBGroup
    End If
    'Ticket #20384 Franks 05/30/2011 - begin
    rsEmp("ED_BADGEID") = Null
    If IsNull(rsEmp("ED_BONUSDEPT")) Then
        rsEmp("ED_BONUSDEPT") = "000000"
    Else
        If Len(Trim(rsEmp("ED_BONUSDEPT"))) = 0 Then
            rsEmp("ED_BONUSDEPT") = "000000"
        Else
            'do not update it if there is value there
        End If
    End If
    'Ticket #20384 Franks 05/30/2011 - end
    rsEmp.Update
    
    If dlpDOther2.Visible Then 'Ticket #19266 Franks 12/03/10
        Call WFC_NGS_Trans(xlocEmpnbr)
    End If

    Call mod_Upd_Pos_Budget_WFC("", "", xlocEmpnbr) 'Ticket #25911 Franks 03/02/2015

    MDIMain.panHelp(0).FloodPercent = 0
    
    'MsgBox "   Done!   "
    Unload Me

End Sub

Private Sub cmdRetireWorking_Click()
Dim rsTB As New ADODB.Recordset
Dim rsT_PARCO As New ADODB.Recordset
Dim rsBen As New ADODB.Recordset
Dim Msg$, DgDef As Variant, Response%
Dim Title$, EID&, TermDate$
Dim SQLQ
Dim DelayProcess As Boolean
Dim xTimminsBenefits As Boolean
Dim xPAData
Dim xLastDay
Dim xPenType
Dim xOldEmp
    
    Msg$ = ""
    If Not chkRetWorking() Then Exit Sub

    'Msg$ = Msg$ & "Are you sure you want to do Retired - Working for this employee ?"
    Msg$ = Msg$ & "Are you sure that you want to move this employee from an active status to a retired and working status?"

    Title$ = ("Retire Employee")
    DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
    
    If Response% = IDNO Then    ' Evaluate response
        Exit Sub
    End If
    
    If rsEmp.EOF Then
        Exit Sub
    End If
    
    If optActTerm(1).Value Then
        'If EmpRehire Then
        'Else
        '    Exit Sub
        'End If
        'xlocEmpnbr = getEmpnbr(txtEmpNo.Text)
    Else 'active
        xlocEmpnbr = getEmpnbr(txtMain.Text)
    End If
    
    'Ticket #21491 Franks 01/25/2012
    'add status change to Employee History table RET -> ACP
    Call EmpHisCalc(1, xlocEmpnbr, "", "", "ACP", "", "", "", "", dlpRetireDate.Text, , , dlpRetireDate.Text)
        
    'open IHREMP_OTHER
    Call OpenEMP_OTHER

    rsEmp("ED_DIVEDATE") = dlpRetireDate.Text
    
    NewBGroup = clpBGroup.Text


    'Pension Status are Term Deferred & Active(T or A): Change ed_admin to "CPEN".
    If IsNull(rsEmp("ED_EMP")) Then xOldEmp = "" Else xOldEmp = rsEmp("ED_EMP")
    rsEmp("ED_EMP") = "ACP" '"RET"
    'Ticket #22409 Franks 08/08/2012 - Don't update LAST DAY on the Status/Dates
    'If IsNull(rsEMP("ED_LDAY")) Then
    '    rsEMP("ED_LDAY") = dlpLastWorkDate.Text
    'End If
    If optActTerm(0).Value Then
        rsEmp("ED_SFDATE") = dlpRetireDate.Text
    End If

    rsEmp.Update

    If Not rsDAT_Other.EOF Then
        rsDAT_Other("ER_PENSIONDATE6") = dlpRetireDate.Text
        rsDAT_Other("ER_LDATE") = Date
        rsDAT_Other("ER_LTIME") = Time$
        rsDAT_Other("ER_LUSER") = glbUserID
        rsDAT_Other.Update
    End If
    
    '"   Removed Benefit Group and logic with the exception of the DB Benefit code.
    'An END DATE should be added to that Benefit. The End Date would equal to Pension Exit Date
    SQLQ = "SELECT * FROM HRBENFT "
    SQLQ = SQLQ & " WHERE BF_EMPNBR = " & xlocEmpnbr & " "
    SQLQ = SQLQ & "AND BF_BCODE = 'DB' "
    SQLQ = SQLQ & " ORDER BY BF_BCODE, BF_EDATE "
    If rsBen.State <> 0 Then rsBen.Close
    rsBen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsBen.EOF Then
        If IsDate(dlpLastWorkDate.Text) Then
            rsBen("BF_CEASEDATE") = CVDate(dlpLastWorkDate.Text)
        End If
        rsBen.Update
    End If
    rsBen.Close
    
    'create Pension Master, PA Master and PA Details
    'Pension Status = "R"
    'Pension Effective Date = Date of Retirement
    If IsDate(dlpLastWorkDate.Text) Then
        xLastDay = dlpLastWorkDate.Text
    Else
        xLastDay = dlpRetireDate.Text
    End If
    xPAData = 0
    If Len(medAmount.Text) > 0 Then
        If IsNumeric(medAmount.Text) Then
            If Val(medAmount.Text) > 0 Then
                xPAData = medAmount.Text
            End If
        End If
    End If
    toSOURCE = "IHR Retirement" 'Ticket #19954
    If xPAData = 0 Then
        'No Pensionable Earnings, it can be blank for Terminated employees
        Call WFCPensionMasUpt(xlocEmpnbr, "Retirement", dlpRetireDate.Text, xLastDay, Year(CVDate(dlpRetireDate.Text)))
    Else
        Call WFCPensionMasUpt(xlocEmpnbr, "Retirement", dlpRetireDate.Text, xLastDay, Year(CVDate(dlpRetireDate.Text)), medAmount.Text)
    End If
    
    'retire other DB Pension
    'One employee can have one DBS plus other DB pensions, such as DBKIPL
    'Employee Dan Dubblestyne had DBS and DBKIPL pensions
    toSOURCE = "IHR Retirement" 'Ticket #19954
    xPenType = getDBType(locSection, locUnion, "PenType", GetEmpData(xlocEmpnbr, "ED_DOH")) 'Ticket #26707 Franks 02/25/2015
    Call WFCOtherPenUpt(xlocEmpnbr, glbSIN, Year(dlpRetireDate.Text), "", xPenType, "R", dlpRetireDate.Text, xLastDay, "DB")
    ''retire other DC Pension
    'Ticket #22409 Franks 08/08/2012 - The program should not end the DC pension type or create a pension master record for DC
    'Call WFCOtherPenUpt(xlocEmpnbr, glbSIN, Year(dlpRetireDate.Text), "", "", "R", dlpRetireDate.Text, xLastDay, "DC")

    Call WFCPensionAlerts(xlocEmpnbr, dlpRetireDate.Text, "Termination  - RET")
    
    'Eligible for Pension ="N".
    rsEmp("ED_EMPTYPE") = "N"
    rsEmp.Update
    rsEmp.Close
    
    'If dlpDOther2.Visible Then 'Ticket #19266 Franks 12/03/10
    '    Call WFC_NGS_Trans(xlocEmpnbr)
    'End If
    
    If gsEMAIL_ONTERM Then
        Call cmdEmailWFCPension
        
        If AbortTerm = True Then
            Screen.MousePointer = vbDefault
            MDIMain.panHelp(0).FloodType = 1
            MDIMain.panHelp(0).Caption = "Retirement Email Aborted"
            MsgBox "Error sending email.  Retirement Email aborted.", vbCritical + vbOKOnly, "Error"
            'Exit Sub
        End If
    End If
    MDIMain.panHelp(0).FloodPercent = 0
    
    'MsgBox "   Done!   "
    Unload Me

End Sub

Private Sub dlpLastWorkDate_LostFocus()
    If glbFrmCaption$ = "Retired – Working Process" Then
        'don't call this funcion
    Else
        Call setRetireDate
    End If
End Sub
Private Sub setRetireDate()
Dim xstrTemp
    If optActTerm(0).Value Then 'Active
        If IsDate(dlpLastWorkDate.Text) Then
            'If Len(dlpRetireDate.Text) = 0 Then
                xstrTemp = MonthName(month(dlpLastWorkDate.Text)) & " 1, " & Year(dlpLastWorkDate.Text)
                xstrTemp = CVDate(xstrTemp)
                xstrTemp = DateAdd("M", 1, xstrTemp)
                dlpRetireDate.Text = xstrTemp
            'End If
        End If
    End If
End Sub
Private Sub Form_Activate()
glbOnTop = "frmERetirement"
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim x%

glbOnTop = "frmERetirement"
Me.Caption = glbFrmCaption$

If glbFrmCaption$ = "Retirement Process" Then
    frmRetirement.Visible = True
    cmdRetire.Enabled = gSec_Upd_RetirementProc
End If
'Ticket #22285 Franks 07/17/2012 - begin
If glbFrmCaption$ = "Retired – Working Process" Then
    frmRetirement.Visible = True
    cmdRetire.Enabled = gSec_Upd_RetirementProc
    optActTerm(1).Visible = False
    lblBen.Visible = False: clpBGroup.Visible = False
    cmdRetireWorking.Top = cmdRetire.Top: cmdRetire.Visible = False: cmdRetireWorking.Visible = True
    lbOtherDate2.Top = lblBen.Top: dlpDOther2.Top = clpBGroup.Top
    lblTitle(0).Caption = "Pension Exit Date"
End If
If glbFrmCaption$ = "Working Retiree Retirement Process" Then
    frmRetirement.Visible = True
    cmdRetire.Enabled = gSec_Upd_RetirementProc
    optActTerm(1).Visible = False
    
    lblTitle(1).Visible = False
    dlpRetireDate.Visible = False
    lblTitle(2).Visible = False
    medAmount.Visible = False
    
    lblBen.Visible = True: clpBGroup.Visible = True
    cmdRetireRet.Caption = "Retire the Employee"
    cmdRetireRet.Top = cmdRetire.Top: cmdRetire.Visible = False: cmdRetireRet.Visible = True
    lbOtherDate2.Top = lblTitle(2).Top: dlpDOther2.Top = medAmount.Top
    lblBen.Top = lblTitle(1).Top:  clpBGroup.Top = dlpRetireDate.Top
End If
'Ticket #22285 Franks 07/17/2012 - end
If glbFrmCaption$ = "Death of an Employee/Spouse" Then
    frmDeathProcess.Top = 750
    frmDeathProcess.Left = frmRetirement.Left
    frmDeathProcess.Visible = True
    cmdDStart.Enabled = gSec_Upd_DeathProc
    frmAT.Visible = False
End If

Call INI_Controls(Me)

If glbtermopen Then
    'term
    If glbTERM_ID > 0 Then
        optActTerm(1).Value = True
        txtMain.Text = glbTERM_ID
    End If
Else
    'active
    If glbLEE_ID > 0 Then
        txtMain.Text = glbLEE_ID
    End If
End If
End Sub
Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateEMP
End Property
Public Property Get UpdateRight() As Boolean
UpdateRight = False   'gSec_Upd_Terminations
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
Printable = False
End Property
Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
TF = True
UpdateState = OPENING
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
End Sub

Private Sub imgIcon_Click()
Call txtMain_DblClick
End Sub


Private Sub medAmount_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medAmtDeh_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optActTerm_Click(Index As Integer)
    If glbFrmCaption$ = "Retirement Process" Then
        txtMain.Text = ""
        txtEmpNo.Text = ""
        dlpLastWorkDate.Text = ""
        dlpRetireDate.Text = ""
        medAmount.Text = ""
        lblEmpExist.Caption = ""
        lblEmpExist.Visible = False
        If Index = 0 Then
            lblEmpno.Visible = False
            txtEmpNo.Visible = False
        Else
            lblEmpno.Visible = True
            txtEmpNo.Visible = True
        End If
    End If
    'If glbFrmCaption$ = "Death of an Employee/Spouse" Then 'Ticket #23491 Franks 04/02/2013
    '    txtMain.Text = ""
    'End If
End Sub

Private Sub optDth_Click(Index As Integer)
    'Ticket #23565 Franks 04/10/2013
    If optDth(0).Value Or optDth(2).Value Or optDth(3).Value Then
        chkToSpouse.Enabled = True
        lblTitle(4).Enabled = True
    End If
    If optDth(1).Value Then
        chkToSpouse.Value = False
        chkToSpouse.Enabled = False
        lblTitle(4).Enabled = False
    End If
    'Ticket #23491 Franks 04/02/2013 - begin
    If optDth(3).Value Then 'terminated
        optActTerm(1).Value = True
        Call optActTerm_Click(1)
        'optActTerm(0).Value = False
    Else
        optActTerm(0).Value = True
        Call optActTerm_Click(0)
    End If
    If optDth(2).Value Then 'Employee
        medAmtDeh.Text = ""
        lblTitle(6).Enabled = True
        medAmtDeh.Enabled = True
    Else
        medAmtDeh.Text = ""
        lblTitle(6).Enabled = False
        medAmtDeh.Enabled = False
    End If
    'Ticket #23491 Franks 04/02/2013 - end
End Sub

Private Sub txtMain_Change()
 Call locRefreshDescription(optActTerm(0).Value)
 If glbFrmCaption$ = "Death of an Employee/Spouse" Then  'Ticket #23505 Franks 04/03/2013
    If locWFCPenEligible Then 'only Pension Employees need these two fields
        lblTitle(4).Enabled = True
        chkToSpouse.Enabled = True
        If optDth(2).Value Then 'Employee only 'Ticket #23565 Franks 04/10/2013
            lblTitle(6).Enabled = True
            medAmtDeh.Enabled = True
        End If
    Else
        lblTitle(4).Enabled = False
        chkToSpouse.Value = False
        chkToSpouse.Enabled = False
        lblTitle(6).Enabled = False
        medAmtDeh.Text = ""
        medAmtDeh.Enabled = False
    End If
 End If
End Sub

Private Sub txtMain_DblClick()
Dim locTermOK As Boolean
Dim locLEE_ID As Double
Dim locTERM_Seq As Double
Dim locTERM_ID As Long
Dim locTermDate$
Dim locLEE_FName As String
Dim locLEE_SName As String
Dim locUNIONTe As String
Dim locTerm_FName As String
Dim locTerm_SName As String
Dim locRehireDt As String
Dim locBand As String

    
    locTermOK = glbTermOK
    locLEE_ID = glbLEE_ID
    locTERM_Seq = glbTERM_Seq
    locTERM_ID = glbTERM_ID
    locTermDate$ = glbTermDate$
    locLEE_FName = glbLEE_FName
    locLEE_SName = glbLEE_SName
    locUNIONTe = glbUNIONTe
    locTerm_FName = glbTerm_FName
    locTerm_SName = glbTerm_SName
    locRehireDt = glbRehireDt
    locBand = glbBand
    
    
    If optActTerm(0).Value Then 'Active glbLEE_ID
        xlocTERM_SEQ& = 0
        'glbCode = txtMain.Text
        frmEEFIND.Show 1
    
        'If glbCode = "" Then Exit Sub
        If glbEEOK Then
            txtMain.Text = glbLEE_ID
        End If
        'glbCode = ""
    End If
    If optActTerm(1).Value Then 'Term
        xlocTERM_SEQ& = 0
        frmTERMEMPL.Show 1
        If glbTermOK Then
            txtMain.Text = glbTERM_ID
            xlocTERM_SEQ& = glbTERM_Seq
        End If
    End If
    'glbTermOK = locTermOK
    'glbLEE_ID = locLEE_ID
    'glbTERM_Seq = locTERM_Seq
    'glbTERM_ID = locTERM_ID
    'glbTermDate$ = locTermDate$
    'glbLEE_FName = locLEE_FName
    'glbLEE_SName = locLEE_SName
    'glbUNIONTe = locUNIONTe
    'glbTerm_FName = locTerm_FName
    'glbTerm_SName = locTerm_SName
    'glbRehireDt = locRehireDt
    'glbBand = locBand
    
End Sub

Private Sub locRefreshDescription(xACT)
Dim xField
Dim SQLQ As String
Dim xstrTemp
    
    If rsEmp.State <> 0 Then rsEmp.Close
    If txtMain.Text = "" Then
        lblDesc.Visible = False
    Else
        lblDesc.Visible = True
    End If
    
    locWFCPenEligible = False
    locUnion = ""
    xlocPayrollID = ""
    xlocDOB = ""
    xTermDate = ""
    If txtMain.Text <> "" And IsNumeric(txtMain.Text) Then
        lblDesc.Caption = "Unassigned"
        IsSalaried = False
        xField = getEmpnbr(txtMain.Text)
        If xACT Then 'Active
            If gdbAdoIhr001 Is Nothing Then Exit Sub
            xlocTERM_SEQ& = 0
            'xField = getEmpnbr(txtMain.Text)
            'rsEmp.Open "SELECT ED_SURNAME,ED_FNAME,ED_SIN,ED_ORG,ED_EMPTYPE,ED_EMP FROM HREMP WHERE ED_EMPNBR=" & xField, gdbAdoIhr001, adOpenForwardOnly
            rsEmp.Open "SELECT * FROM HREMP WHERE ED_EMPNBR=" & xField, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsEmp.EOF Then
                lblDesc.Caption = rsEmp("ED_SURNAME") & ", " & rsEmp("ED_FNAME")
                If Not IsNull(rsEmp("ED_SIN")) Then
                    glbSIN = rsEmp("ED_SIN")
                End If
                If Not IsNull(rsEmp("ED_EMPTYPE")) Then
                    If UCase(rsEmp("ED_EMPTYPE")) = "Y" Then
                        locWFCPenEligible = True
                    End If
                End If
                If Not IsNull(rsEmp("ED_ORG")) Then
                    locUnion = rsEmp("ED_ORG")
                    'If locWFCPenEligible Then
                        If locUnion = "NONE" Then 'Or locUnion = "EXEC" Then
                            IsSalaried = True
                            medAmount.Text = getPayrollTransaction(xField, "DN49", Year(Date))
                        End If
                    'End If
                End If
                If Not IsNull(rsEmp("ED_BENEFIT_GROUP")) Then
                    'clpBGroup.Text = rsEMP("ED_BENEFIT_GROUP")
                    SaveBGroup = rsEmp("ED_BENEFIT_GROUP")
                Else
                    'clpBGroup.Text = ""
                    SaveBGroup = ""
                End If
                'SaveBGroup = clpBGroup.Text
                clpBGroup.Text = ""
                If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then
                    xlocPayrollID = rsEmp("ED_PAYROLL_ID")
                End If
                If Not IsNull(rsEmp("ED_DOB")) Then
                    xlocDOB = rsEmp("ED_DOB")
                End If
            End If
            
            'Ticket #19266 Franks 12/03/10
            Call WFCOther2Screen(xField)
            'rsEmp.Close
        Else 'TERM
            If gdbAdoIhr001X Is Nothing Then Exit Sub
            'xField = getEmpnbr(txtMain.Text)
            xlocTERM_SEQ& = 0
            'SQLQ = "SELECT ED_SURNAME,ED_FNAME,ED_SIN,ED_ORG,ED_EMPTYPE,TERM_SEQ FROM TERM_HREMP WHERE ED_EMPNBR=" & xField
            SQLQ = "SELECT * FROM TERM_HREMP WHERE ED_EMPNBR=" & xField
            SQLQ = SQLQ & " ORDER BY ED_EMPNBR, TERM_SEQ DESC"
            rsEmp.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
            If Not rsEmp.EOF Then
                If rsHRTRMEMP.State Then rsHRTRMEMP.Close
                SQLQ = "SELECT * FROM Term_HRTRMEMP WHERE TERM_SEQ = " & rsEmp("TERM_SEQ")
                rsHRTRMEMP.Open SQLQ, gdbAdoIhr001X, adOpenStatic
                If Not rsHRTRMEMP.EOF Then
                    dlpLastWorkDate.Text = rsHRTRMEMP("Term_DOT")
                    xTermDate = rsHRTRMEMP("Term_DOT")
                End If
                xstrTemp = MonthName(month(Date)) & " 1, " & Year(Date)
                xstrTemp = CVDate(xstrTemp)
                xstrTemp = DateAdd("M", 1, xstrTemp)
                dlpRetireDate.Text = xstrTemp
                
                lblDesc.Caption = rsEmp("ED_SURNAME") & ", " & rsEmp("ED_FNAME")
                xlocTERM_SEQ& = rsEmp("TERM_SEQ")
                If Not IsNull(rsEmp("ED_SIN")) Then
                    glbSIN = rsEmp("ED_SIN")
                End If
                If Not IsNull(rsEmp("ED_EMPTYPE")) Then
                    If UCase(rsEmp("ED_EMPTYPE")) = "Y" Then
                        locWFCPenEligible = True
                    End If
                End If
                If Not IsNull(rsEmp("ED_ORG")) Then
                    locUnion = rsEmp("ED_ORG")
                    'If locWFCPenEligible Then
                        If locUnion = "NONE" Then 'Or locUnion = "EXEC" Then
                            IsSalaried = True
                            medAmount.Text = getPayrollTransaction(xField, "DN49", Year(Date), xlocTERM_SEQ&)
                        End If
                    'End If
                End If
                If Not IsNull(rsEmp("ED_BENEFIT_GROUP")) Then
                    'clpBGroup.Text = rsEMP("ED_BENEFIT_GROUP")
                    SaveBGroup = rsEmp("ED_BENEFIT_GROUP")
                Else
                    'clpBGroup.Text = ""
                    SaveBGroup = ""
                End If
                'SaveBGroup = clpBGroup.Text
                clpBGroup.Text = ""
                
                If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then
                    xlocPayrollID = rsEmp("ED_PAYROLL_ID")
                End If
                If Not IsNull(rsEmp("ED_DOB")) Then
                    xlocDOB = rsEmp("ED_DOB")
                End If
                txtEmpNo.Text = txtMain.Text
                lblEmpExist.Caption = ""
                lblEmpExist.Visible = False
            End If
            'Ticket #19266 Franks 12/03/10
            Call WFCOther2Screen(xField, xlocTERM_SEQ&)
            'rsEmp.Close
        End If
        lblDesc.Caption = Replace(lblDesc.Caption, "&", "&&")
    Else
        'If propShowUnassigned = Always Then lblDesc.Caption = "Unassigned" Else lblDesc.Caption = ""
        lblDesc.Caption = ""
    End If
    
End Sub

Private Function getPayrollTransaction(xEmpNo, xCode, xYear, Optional xTermSEQ)
Dim rsPayTran As New ADODB.Recordset
Dim SQLQ As String
Dim retval As Double
    retval = 0
    If IsMissing(xTermSEQ) Then
        SQLQ = "SELECT SUM(PT_DOLLARAMT) AS TOTAMT FROM HR_PAYROLL_TRANSACTION WHERE PT_EMPNBR = " & xEmpNo & " "
    Else
        SQLQ = "SELECT SUM(PT_DOLLARAMT) AS TOTAMT FROM TERM_PAYROLL_TRANSACTION WHERE PT_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
    End If
    SQLQ = SQLQ & "AND PT_PAYCODE = '" & xCode & "' "
    SQLQ = SQLQ & "AND YEAR(PT_PAYSTART) = " & xYear & " "
    rsPayTran.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsPayTran.EOF Then
        If Not IsNull(rsPayTran("TOTAMT")) Then
            retval = rsPayTran("TOTAMT")
        End If
    End If
    rsPayTran.Close
    getPayrollTransaction = retval
End Function

Private Sub OpenEMP_OTHER()
Dim SQLQ As String
    
    ''SQLQ = "SELECT * "
    ''SQLQ = SQLQ & " from HREMP_OTHER"
    ''SQLQ = SQLQ & " where ER_EMPNBR = " & xlocEmpnbr
        
    If optActTerm(1).Value Then 'TERM
        SQLQ = "SELECT * "
        SQLQ = SQLQ & " from Term_HREMP_OTHER"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & xlocTERM_SEQ
        SQLQ = SQLQ & " AND ER_EMPNBR = " & xlocEmpnbr 'glbLEE_ID
    Else
        SQLQ = "SELECT * "
        SQLQ = SQLQ & " from HREMP_OTHER"
        SQLQ = SQLQ & " where ER_EMPNBR = " & xlocEmpnbr 'glbLEE_ID
    End If
    
    If rsDAT_Other.State <> 0 Then rsDAT_Other.Close
    rsDAT_Other.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    'DataOther.RecordSource = SQLQ
        
    If rsDAT_Other.EOF Then
        rsDAT_Other.AddNew
        rsDAT_Other("ER_COMPNO") = "001"
        rsDAT_Other("ER_EMPNBR") = xlocEmpnbr 'glbLEE_ID
        'If optActTerm(1).Value And xlocTERM_SEQ > 0 Then 'TERM
        '    rsDAT_Other("TERM_SEQ") = xlocTERM_SEQ
        'End If
        rsDAT_Other("ER_LDATE") = Date
        rsDAT_Other("ER_LTIME") = Time$
        rsDAT_Other("ER_LUSER") = glbUserID
        rsDAT_Other.Update
    End If

End Sub

Function EmpRehire()
Dim xCommText As String
    EmpRehire = False
    If rsEmp.State <> 0 Then rsEmp.Close
    If rsDAT_Other.State <> 0 Then rsDAT_Other.Close
    
    'Prior to deleting the termination record, create a Comment record with the termination comments in it.
    'Comment Type = RET and date = date of retirement. Make sure the comments include reason, date and any comments.
    'Pension Changes - June 1-2010(Jul009).docx
    
    'xCommText = "This employee was retired on " & dlpRetireDate.Text
    'Call CreateRetireComment(txtEmpNo.Text, xlocTERM_SEQ, dlpRetireDate.Text, "RET", xCommText)
    Call CreateRetireComment(txtEmpNo.Text, xlocTERM_SEQ, "TERM") ', dlpRetireDate.Text, "RET", xCommText)
    
    Call modReinMove(txtEmpNo.Text, xlocTERM_SEQ, Date)
    
    Call modNukeEETerm(xlocTERM_SEQ)
    
    'reopen HREMP and HREMP_OTHER

    rsEmp.Open "SELECT * FROM HREMP WHERE ED_EMPNBR=" & txtEmpNo.Text, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    glbEEOK = True
    glbTermOK = False
    glbLEE_ID = txtEmpNo.Text
    
    EmpRehire = True
    
End Function

Private Sub CreateRetireComment(EID&, EESEQ&, xType)  ', xCommText)
Dim rsComm As New ADODB.Recordset
Dim rsETerm As New ADODB.Recordset
Dim xDate, xCode, xDesc, xComm, xMess
Dim SQLQ As String
    SQLQ = "SELECT * FROM Term_HRTRMEMP WHERE TERM_SEQ = " & EESEQ& & " "
    rsETerm.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsETerm.EOF Then
        xDate = rsETerm("Term_DOT")
        xCode = rsETerm("Term_Reason")
        xDesc = GetTABLDesc("TERM", xCode)
        xMess = "This employee was terminated on " & xDate & " and reason was " & xDesc & ". "
        If Not IsNull(rsETerm("Term_Comments")) Then
            xComm = rsETerm("Term_Comments")
        Else
            xComm = ""
        End If
        xComm = xMess & Chr(10) & xComm
        SQLQ = "Select * from Term_COMMENTS "
        SQLQ = SQLQ & " WHERE Term_COMMENTS.TERM_SEQ = " & EESEQ& & " AND Term_COMMENTS.CO_TYPE = '" & xType & "' "
        SQLQ = SQLQ & "AND CO_EDATE = " & Date_SQL(RetireDate$) & " "
        rsComm.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsComm.EOF Then
            rsComm.AddNew
            rsComm("CO_COMPNO") = "001"
            rsComm("CO_EMPNBR") = EID&
            rsComm("CO_EDATE") = xDate 'RetireDate$
            rsComm("CO_TYPE") = xType
            If Len(xComm) > 0 Then
                rsComm("CO_COMMENTS") = xComm 'xCommText
            End If
            rsComm("CO_LDATE") = Date
            rsComm("CO_LTIME") = Time$
            rsComm("CO_LUSER") = glbUserID
            rsComm("TERM_SEQ") = EESEQ&
            rsComm.Update
        End If
        rsComm.Close
    End If
    rsETerm.Close
End Sub

Private Function modReinMove(EID&, EESEQ&, TermDate$)
Dim x%, DtTm   As Variant, TRDesc$

Screen.MousePointer = HOURGLASS
modReinMove = False
DtTm = Now

MDIMain.panHelp(0).FloodPercent = 5

If glbCompSerial = "S/N - 2192W" Or glbWFC Or glbCompSerial = "S/N - 2370W" Then  'County of Essex
    x% = REIN_BASIC(EID&, EESEQ&, TermDate$, xlocPayrollID)
Else
    x% = REIN_BASIC(EID&, EESEQ&, TermDate$)
End If
If Not x% Then
    Exit Function
End If

'x% = Reset_BASIC(EID&)
'x% = RehHREMPAudit(EID&, EESEQ&)
x% = REIN_DEPEND(EID&, EESEQ&)
x% = REIN_COBRA(EID&, EESEQ&)

MDIMain.panHelp(0).FloodPercent = 10

x% = REIN_ATTENDANCE(EID&, EESEQ&)

x% = REIN_BENEFITS(EID&, EESEQ&)
'x% = RehBENEFITSAudit(EID&, EESEQ&)   'Laura jan 13, 1998

MDIMain.panHelp(0).FloodPercent = 20

x% = REIN_HealthCost(EID&, EESEQ&)

MDIMain.panHelp(0).FloodPercent = 25

x% = REIN_HealthSafety(EID&, EESEQ&)
x% = REIN_OHS_CONTACT(EID&, EESEQ&)
x% = REIN_OHS_CORRECTIVE(EID&, EESEQ&)
x% = REIN_OHS_ROOT_CAUSES(EID&, EESEQ&)
x% = REIN_OHS_CLAIM_MEDICAL(EID&, EESEQ&)

MDIMain.panHelp(0).FloodPercent = 30


x% = REIN_JOB(EID&, EESEQ&)
'x% = RehJOBAudit(EID&, EESEQ&)     'laura jan 13, 1998
MDIMain.panHelp(0).FloodPercent = 40

x% = REIN_PERFORM(EID&, EESEQ&)
MDIMain.panHelp(0).FloodPercent = 60

x% = REIN_SALARY(EID&, EESEQ&)
'x% = RehSALARYAudit(EID&, EESEQ&)   'laura jan 13, 1998
MDIMain.panHelp(0).FloodPercent = 75

x% = REIN_EDUCSEM(EID&, EESEQ&)

x% = REIN_COMMENTS(EID&, EESEQ&)

x% = REIN_EARN(EID&, EESEQ&)

x% = REIN_EDU(EID&, EESEQ&)

x% = REIN_EMPSKL(EID&, EESEQ&)

x% = REIN_TRADE(EID&, EESEQ&)

x% = REIN_DOLENT(EID&, EESEQ&)

'Ticket #28789 - Actual Amounts Details
x% = REIN_DOLENT_ACTDTL(EID&, EESEQ&)

x% = REIN_ENTHRS(EID&, EESEQ&)

'x% = REIN_RSP(EID&, EESEQ&)

x% = REIN_COUNSEL(EID&, EESEQ&)

x% = REIN_USERDEFINED(EID&, EESEQ&)     'Hemu - User Defined Table


If gsAttachment_DB Then
    x% = REIN_HRDOC_EMP(EID&, EESEQ&)
    x% = REIN_HRDOC_JOB_HISTORY(EID&, EESEQ&)
    x% = REIN_HRDOC_COMMENTS(EID&, EESEQ&)
    x% = REIN_HRDOC_HEALTH_SAFETY(EID&, EESEQ&)
    x% = REIN_HRDOC_HEALTH_SAFETY_2(EID&, EESEQ&)
    x% = REIN_HRDOC_COUNSEL(EID&, EESEQ&)
    x% = REIN_HRDOC_PERFORM_HISTORY(EID&, EESEQ&)
    'add missed tables by Frank 01/21/10 Ticket #17894 - begin
    x% = REIN_HRDOC_EDSEM(EID&, EESEQ&)
    x% = REIN_HRDOC_EDSEM_RETEST(EID&, EESEQ&)
    x% = REIN_HRDOC_HREDU(EID&, EESEQ&)
    x% = REIN_HRDOC_DOLENT(EID&, EESEQ&)
    'add missed tables by Frank 01/21/10 Ticket #17894 - end
End If


x% = REIN_SUCCESSION(EID&, EESEQ&) 'George Apr 4,2006 #10595
x% = REIN_LANGUAGE(EID&, EESEQ&) 'George Apr 4,2006 #10595
x% = REIN_HREMPHIS(EID&, EESEQ&)


x% = RemoveHREMPEQU_DOT(EID&)
'
'Dim HRChanges As New Collection
'Call isChanged_Field(HRChanges, oDOH, dlpDOH)
'Call isChanged_Field(HRChanges, oFday, dlpLTHire)
'Call Passing_Changes(HRChanges, Rehire, "M", Date, EID&)
'
'Call AddNewPayrollEmp(Rehire, Date, EID&, "")

modReinMove = True

Screen.MousePointer = DEFAULT

Exit Function

modReinMoveErr_Msg:
Screen.MousePointer = DEFAULT
'MsgBox "Problem Creating Audit record - Termination Aborted"
Resume Next

End Function

Private Function RemoveHREMPEQU_DOT(EmpN As Long)
Dim SQLQ As String
Dim dynEmp As New ADODB.Recordset

SQLQ = "SELECT * FROM HREMPEQU WHERE HREMPEQU.EQ_EMPNBR = "
SQLQ = SQLQ & EmpN

dynEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If dynEmp.RecordCount > 0 Then
    'Release 8.0 - Ticket #24309: Addition option to enter Terminated Employees on hte Employment Equity Survey screen
    'SQLQ = "UPDATE HREMPEQU SET HREMPEQU.EQ_DOT = Null "
    SQLQ = "UPDATE HREMPEQU SET HREMPEQU.EQ_DOT = Null, EQ_TYPE = 'A' "
    SQLQ = SQLQ & "WHERE HREMPEQU.EQ_EMPNBR = " & EmpN
    gdbAdoIhr001.Execute SQLQ
End If

End Function

Private Function SpouseName4DB()
Dim SQLQ As String
Dim rsBenFic As New ADODB.Recordset
Dim retval As String
    retval = ""
    If optActTerm(0).Value Then
        SQLQ = "SELECT * FROM HRBENS WHERE BD_BCODE = 'DB' "
        SQLQ = SQLQ & "AND BD_EMPNBR = " & xlocEmpnbr & " "
    End If
    If optActTerm(1).Value Then
        SQLQ = "SELECT * FROM Term_HRBENS WHERE BD_BCODE = 'DB' "
        SQLQ = SQLQ & "AND BD_EMPNBR = " & xlocEmpnbr & " "
        SQLQ = SQLQ & "AND TERM_SEQ = " & glbTERM_Seq & " "
    End If
    rsBenFic.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsBenFic.EOF Then
        If Not IsNull(rsBenFic("BD_RELATE")) Then
            If rsBenFic("BD_RELATE") = "Spouse" Then
                retval = rsBenFic("BD_BNAME")
            End If
        End If
    End If
    rsBenFic.Close
    SpouseName4DB = retval
End Function

Private Sub TerminateEmp(xEmpNo, xDOT, xTermReason)
Dim rsTB As New ADODB.Recordset
Dim rsT_PARCO As New ADODB.Recordset
Dim Title$, EID&, TermDate$
Dim SQLQ

rsTB.Open "Term_HRSEQ", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic, adCmdTableDirect
If rsTB.EOF And rsTB.BOF Then
    glbTERM_Seq = 1
    rsTB.AddNew
Else
    rsTB.MoveFirst 'Jaddy 10/28/99
    glbTERM_Seq = rsTB("TERM_SEQ_NEXT")
End If
rsTB("TERM_SEQ_NEXT") = glbTERM_Seq + 1
rsTB.Update

rsTB.Close

If Not AUDITTERM(xEmpNo, xDOT, xTermReason) Then MsgBox "ERROR - AUDIT FILE"

If Not modTermMove(xEmpNo, xDOT, xTermReason) Then Exit Sub

EID& = xEmpNo 'CLng(lblEEID)
TermDate$ = xDOT

If Not Term_Superv(xEmpNo) Then Exit Sub  'laura

MDIMain.panHelp(0).FloodPercent = 100

gdbAdoIhr001.Execute "delete from HR_PHOTO where PT_EMPNBR=" & EID&

Call NukeEE2(EID&)
MDIMain.panHelp(0).FloodPercent = 0

rsT_PARCO.Open "HRPARCO", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
rsT_PARCO("PC_NUMBER_EMPLOYEES") = rsT_PARCO("PC_NUMBER_EMPLOYEES") - 1 'UPDATE FIELD WITH ACTUAL COUNT
rsT_PARCO.Update
rsT_PARCO.Close
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

SQLQ = "UPDATE Term_HREMP SET ED_OMERS=" & Date_SQL(xDOT)
SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
gdbAdoIhr001X.Execute SQLQ

glbLEE_ID = 0

End Sub

Private Function AUDITTERM(xEmpNo, xDOT, xTermReason)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsTACheck As New ADODB.Recordset
Dim xPT As String, xDiv As String, XSNAME As String, XFNAME As String, xEmpType As String, xDOH As String, xSENDTE As String
Dim SQLQ As String, strFields As String

On Error GoTo AUDIT_ERR

AUDITTERM = False
 
rsTB.Open "SELECT ED_PT,ED_DIV,ED_SURNAME,ED_FNAME,ED_EMPTYPE,ED_DOH,ED_SENDTE FROM HREMP WHERE ED_EMPNBR=" & xEmpNo, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    If Not IsNull(rsTB("ED_PT")) Then   'Hemu - Gives an error when it's Null and this checking is not done
        xPT = rsTB("ED_PT")
    Else
        xPT = ""
    End If
    
    If Not IsNull(rsTB("ED_DIV")) Then 'George Apr 4,2006
        'xDiv = rsTB("ED_DIV")
        If IsNull(rsTB("ED_DIV")) Then xDiv = "" Else xDiv = rsTB("ED_DIV")
    Else
        xDiv = ""
    End If
    XSNAME = rsTB("ED_SURNAME")
    XFNAME = rsTB("ED_FNAME")
    If IsNull(rsTB("ED_EMPTYPE")) Then
        xEmpType = ""
    Else
        xEmpType = rsTB("ED_EMPTYPE")
    End If
    xDOH = rsTB("ED_DOH")
    If IsNull(rsTB("ED_SENDTE")) Then
        xSENDTE = ""
    Else
        xSENDTE = rsTB("ED_SENDTE")
    End If
Else
    xPT = ""
    xDiv = ""
    XSNAME = ""
    XFNAME = ""
    xEmpType = ""
    xDOH = ""
    xSENDTE = ""
End If
rsTB.Close
'Linamar doesn't need Audit records when Transfer Out
'WFC need Audit records when Transfer Out
'Ticket# 7337 For Linamar Interface
'If glbTermTran Or Not glbLinamar Then
    'strFields added by Bryan 02/Dec/05 Ticket#9899
    strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, "
    strFields = strFields & "AU_TREAS_TABL, AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_EMPTYPE, AU_SURNAME, "
    strFields = strFields & "AU_FNAME, AU_DOT, AU_TREAS, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID,AU_VADIM2,AU_SIN,AU_SSN "
    rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
    rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
    rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_NEWEMP") = "N"
    rsTA("AU_PTUPL") = xPT
    rsTA("AU_DIVUPL") = xDiv
    rsTA("AU_EMPTYPE") = xEmpType
    rsTA("AU_SURNAME") = XSNAME
    rsTA("AU_FNAME") = XFNAME
    rsTA("AU_DOT") = xDOT
    rsTA("AU_TREAS") = xTermReason
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = xEmpNo
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "T"
    'If glbSoroc Or glbSyndesis Then
    Dim rsEmp As New ADODB.Recordset
    'Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_SIN,ED_SSN FROM HREMP WHERE ED_EMPNBR =" & xEmpNo
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
        'Ticket #16749
        If glbWFC Then
            rsTA("AU_SIN") = rsEmp("ED_SIN")
            rsTA("AU_SSN") = rsEmp("ED_SSN")
        End If
    End If
    rsEmp.Close
    'End If
    rsTA.Update
    rsTA.Close
'End If


AUDITTERM = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '29July99 js
Resume Next
End Function

Private Function modTermMove(xEmpNo, xDOT, xTermReason)
Dim x%
Dim EEID&, TReason$, DtTm  As Variant, TRDesc$
Dim TComment$
Dim TRehire$
Dim TCause

Screen.MousePointer = HOURGLASS
modTermMove = False
DtTm = xDOT
EEID& = xEmpNo
TReason$ = xTermReason
TComment$ = ""
TRehire$ = "No"
TRDesc$ = ""
TCause = ""

gdbAdoIhr001.BeginTrans
'gdbAdoIhr001X.BeginTrans

x% = TERM_LIST(EEID&, DtTm, TReason$, TRDesc$, TComment$, TRehire$)
MDIMain.panHelp(0).FloodPercent = 5
x% = TERM_BASIC(EEID&)
MDIMain.panHelp(0).FloodPercent = 10
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_EDUCSEM(EEID&)                  'laura nov 5, 1997
MDIMain.panHelp(0).FloodPercent = 13      '
If Not x Then GoTo modTermMoveErr_Msg    '
x% = TERM_ATTENDANCE(EEID&, DtTm)
MDIMain.panHelp(0).FloodPercent = 15
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_ATTENDANCE_HISTORY(EEID&, DtTm)
MDIMain.panHelp(0).FloodPercent = 18
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_JOB(EEID&)
MDIMain.panHelp(0).FloodPercent = 20
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_PERFORM(EEID&)
MDIMain.panHelp(0).FloodPercent = 22
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_SALARY(EEID&)
MDIMain.panHelp(0).FloodPercent = 25
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_HealthSafety(EEID&)
MDIMain.panHelp(0).FloodPercent = 28
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_BENEFITS(EEID&)
MDIMain.panHelp(0).FloodPercent = 30
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_DEPEND(EEID&)
MDIMain.panHelp(0).FloodPercent = 31
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_HealthCost(EEID&)
MDIMain.panHelp(0).FloodPercent = 32
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_OHS_Contact(EEID&)
MDIMain.panHelp(0).FloodPercent = 35
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_COMMENTS(EEID&)               'FRANK 4/5/2000
MDIMain.panHelp(0).FloodPercent = 38
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_COBRA(EEID&)
MDIMain.panHelp(0).FloodPercent = 39
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_OHS_Corrective(EEID&)
MDIMain.panHelp(0).FloodPercent = 40
If Not x Then GoTo modTermMoveErr_Msg
x% = Term_OHS_ROOT_CAUSES(EEID&)

If Not x Then GoTo modTermMoveErr_Msg
x% = Term_OHS_CLAIM_MEDICAL(EEID&)

MDIMain.panHelp(0).FloodPercent = 43
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_DOLENT(EEID&)

'Ticket #28789 - Actual Amounts Details
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_DOLENT_ACTDTL(EEID&)

MDIMain.panHelp(0).FloodPercent = 45
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_ENTHRS(EEID&)                 'FRANK 4/5/2000
MDIMain.panHelp(0).FloodPercent = 46
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_EARN(EEID&)                   'FRANK 4/5/2000
MDIMain.panHelp(0).FloodPercent = 48
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_EDU(EEID&)                    'FRANK 4/5/2000
MDIMain.panHelp(0).FloodPercent = 50
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_EMPSKL(EEID&)                 'FRANK 4/5/2000
MDIMain.panHelp(0).FloodPercent = 52
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_TRADE(EEID&)                  'FRANK 4/5/2000
If Not x Then GoTo modTermMoveErr_Msg
MDIMain.panHelp(0).FloodPercent = 53
x% = TERM_COUNSEL(EEID&)                ' dkostka - 10/02/2001
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_HREMPHIS(EEID&)                ' Hemu - 06/30/2004
If Not x Then GoTo modTermMoveErr_Msg
'If glbWFC Then
x% = TERM_EMPOTHER(EEID&)                  'FRANK 11/05/2004
If Not x Then GoTo modTermMoveErr_Msg
'End If
x% = TERM_USERDEFINE_TABLE(EEID&)          'Hemu - 02/28/2008
If Not x Then GoTo modTermMoveErr_Msg

x% = TERM_SUCCESSION(EEID&)          'George 04/04/2006 #10595
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_LANGUAGE(EEID&)          'George 04/04/2006 #10595
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_EMP_FLAGS(EEID&)          'Bryan 05/04/2006
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_GLDIST(EEID&)             'Bryan 05/04/2006
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_EMPADP(EEID&)                  'FRANK 06/08/2006
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_EMPPAYROLL_TRANSACTION(EEID&)  'FRANK 03/18/2010 Ticket #18232
If Not x Then GoTo modTermMoveErr_Msg

If gsAttachment_DB Then
    x% = TERM_HRDOC_EMP(EEID&)                  'FRANK 01/10/2006
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_JOB_HISTORY(EEID&)          'George 01/19/2006 #10266
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_COMMENTS(EEID&)          'George 01/26/2006 #10266
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_HEALTH_SAFETY(EEID&)          'George 02/17/2006 #10266
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_HEALTH_SAFETY_2(EEID&)
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_COUNSEL(EEID&)          'George 01/26/2006 #10266
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_PERFORM_HISTORY(EEID&)          'George 01/26/2006 #10266
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_EDSEM(EEID&)
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_EDSEM_RETEST(EEID&)
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_HREDU(EEID&)
    If Not x Then GoTo modTermMoveErr_Msg '
    x% = TERM_HRDOC_HRDOLENT(EEID&)
    If Not x Then GoTo modTermMoveErr_Msg
End If '

MDIMain.panHelp(0).FloodPercent = 55
If Not x Then GoTo modTermMoveErr_Msg

gdbAdoIhr001.CommitTrans
'gdbAdoIhr001X.CommitTrans

modTermMove = True

Screen.MousePointer = DEFAULT
Exit Function

modTermMoveErr_Msg:
Screen.MousePointer = DEFAULT

'MsgBox TranStr("Problem Creating Audit record - Termination Aborted")

End Function


Private Function Term_Superv(xEmpNo)
'Laura
Term_Superv = False
Dim SQLQDel As String, SQLQCom As String, strTable As String
Dim dynHRAT As New ADODB.Recordset
Dim strComm

On Error GoTo Database_Err
'Set Superv_DB = OpenDatabase(glbIHRDB, False, False)

Screen.MousePointer = HOURGLASS

SQLQCom = "UPDATE HR_ATTENDANCE SET AD_SUPER = 0 WHERE AD_SUPER = " & xEmpNo
gdbAdoIhr001.Execute SQLQCom

SQLQCom = "UPDATE HR_ATTENDANCE_HISTORY SET AH_SUPER = 0 WHERE AH_SUPER = " & xEmpNo
gdbAdoIhr001.Execute SQLQCom

SQLQCom = "UPDATE HR_PERFORM_HISTORY SET PH_REPTAU = 0 WHERE PH_REPTAU = " & xEmpNo
gdbAdoIhr001.Execute SQLQCom

SQLQCom = "UPDATE HR_JOB_HISTORY SET JH_REPTAU = 0 WHERE JH_REPTAU = " & xEmpNo
gdbAdoIhr001.Execute SQLQCom

SQLQCom = "UPDATE HR_OCC_HEALTH_SAFETY SET EC_EMPNOT = 0 WHERE EC_EMPNOT = " & xEmpNo
gdbAdoIhr001.Execute SQLQCom


SQLQCom = "UPDATE HR_OCC_HEALTH_SAFETY SET EC_SUPERVISOR = 0 WHERE EC_SUPERVISOR = " & xEmpNo
gdbAdoIhr001.Execute SQLQCom


Screen.MousePointer = DEFAULT

Term_Superv = True
Exit Function

Database_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_Superv", strTable, "TERMINATE")

End Function

Private Sub NukeEE2(EEID As Long)
Dim snapEETables As New ADODB.Recordset
Dim SQLQ As String, TabName$
Dim EEIDAlias$

On Error GoTo NukeEE2_Err
Dim rsSE As New ADODB.Recordset
Dim xUserID As String
rsSE.Open "SELECT USERID FROM HR_SECURE_BASIC WHERE EMPNBR=" & EEID&, gdbAdoIhr001, adOpenStatic
If Not rsSE.EOF Then
    xUserID = rsSE("USERID")
    Call NukeUSERID(xUserID)
End If
rsSE.Close

SQLQ = "SELECT * FROM INFO_HR_TABLES "
SQLQ = SQLQ & " WHERE Employee_Keyed <>0"
SQLQ = SQLQ & " AND TERMINATION_TABLE=0"
'Ticket #20415 - Add Serial # to the select statement so custom tables also gets employee # changed.
'Serial 9999 is by default for all standard info:HR table.
'SQLQ = SQLQ & " AND (SERIAL = 'S/N - 9999W' OR SERIAL = '" & glbCompSerial & "')"
'Ticket #20893 Franks 09/02/2011 - only remove data for the standard INFO:HR tables
SQLQ = SQLQ & " AND (SERIAL = 'S/N - 9999W' OR SERIAL IS NULL) "

snapEETables.Open SQLQ, gdbAdoIhr001, adOpenStatic

If snapEETables.RecordCount < 1 Then Exit Sub
snapEETables.MoveFirst

While Not snapEETables.EOF
    TabName$ = snapEETables("Table_Name")
    If UCase(Right(TabName$, 3)) <> "WRK" Then
      EEIDAlias$ = snapEETables("EMPNBR_Alias")
        If glbVadim And TabName$ = "HRBENFT" Then ' special process for vadim integration
            gdbAdoIhr001.BeginTrans
            gdbAdoIhr001.Execute "UPDATE HRBENFT SET BF_LUSER='VADIM_INTEGRATION' WHERE BF_EMPNBR=" & EEID&
            gdbAdoIhr001.CommitTrans
        End If
      Call NukeEERows2(TabName$, EEIDAlias$, EEID&)
    End If
    snapEETables.MoveNext
Wend

snapEETables.Close
Call UpdVacTimeRequest(EEID&, "D")
If glbCompSerial = "S/N - 2362W" Then 'CITY OF SARNIA
    SQLQ = "DELETE FROM HR_OHS_REOCCURENCE WHERE CC_EMPNBR =" & EEID & " "
    gdbAdoIhr001.Execute SQLQ
End If

If glbCompSerial = "S/N - 2279W" Then  'Friesens Corporation
    SQLQ = "DELETE FROM HR_PERFORM_FRIESEN WHERE PH_EMPNBR =" & EEID & " "
    gdbAdoIhr001.Execute SQLQ
End If

SQLQ = "DELETE FROM HR_PAYROLL_TRANSACTION WHERE PT_EMPNBR =" & EEID & " "
gdbAdoIhr001.Execute SQLQ
    
Exit Sub

NukeEE2_Err:
glbFrmCaption$ = "Delete Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HR_TABLES Error", "TabName$", "Search")
Call RollBack '29July99 js

End Sub

Private Sub NukeEERows2(TabName As String, EEIDAlias As String, EEID As Long)
' returns number of records found for ee in table
Dim Rows%, SQLQ As String
Dim gdbESS As New ADODB.Connection

On Error GoTo NukeEERows2_Err

If TabName$ = "HREMPEQU" Then
    Exit Sub
End If

SQLQ = "DELETE FROM " & TabName
SQLQ = SQLQ & " WHERE " & EEIDAlias & " = " & EEID
gdbAdoIhr001.Execute SQLQ

Exit Sub

NukeEERows2_Err:
glbFrmCaption$ = "Nuke Rows"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Delete EE Rows", TabName$, "Delete")
Call RollBack '29July99 js

End Sub

Private Sub WFCOther2Screen(xEmpNo, Optional xTermSEQ)
Dim rsEmpee As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim SQLQ As String
Dim xUnion As String
Dim xSalHly As String
Dim xInSubGrp As String
Dim xLDate
Dim xNGSStart
Dim xTermFlag As Boolean
    
    isNGS = False
    If Not glbNGS_OnFlag Then
        Exit Sub
    End If
    
    If IsMissing(xTermSEQ) Then
        xTermFlag = False
    Else
        xTermFlag = True
    End If
    
    lbOtherDate2.Visible = False
    dlpDOther2.Visible = False
    
    If xTermFlag Then
        SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2 FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
    Else
        SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2 FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    End If
    rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmpee.EOF Then
        Exit Sub
    Else
        If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
        If IsNull(rsEmpee("ED_ORG")) Then glbUNION = "" Else glbUNION = rsEmpee("ED_ORG")
        If IsNull(rsEmpee("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = rsEmpee("ED_VADIM1")
        If IsNull(rsEmpee("ED_VADIM2")) Then glbWFCPayGroup = "" Else glbWFCPayGroup = rsEmpee("ED_VADIM2")
    End If
    rsEmpee.Close
    
    'No NGS Sub Group, skip
    If Len(glbWFCNGSSubGroup) = 0 Then Exit Sub

    xNGSStart = ""
    If xTermFlag Then
        SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1 FROM Term_HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
    Else
        SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1 FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo & " "
    End If
    rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpOther.EOF Then
        If IsDate(rsEmpOther("ER_OTHERDATE1")) Then
            xNGSStart = rsEmpOther("ER_OTHERDATE1")
        End If
    End If
    rsEmpOther.Close
    'No NGS Effective Date, skip
    If Len(xNGSStart) = 0 Then Exit Sub
    isNGS = True
    
    If glbFrmCaption$ = "Retired – Working Process" Then
        'Ticket #22285 Franks 07/17/2012 - begin
        'do not enter NGS End Date for "Retired – Working Process"
    Else
        lbOtherDate2.Caption = lStr("Other Date 2")
        lbOtherDate2.Visible = True
        dlpDOther2.Visible = True
    End If
    
End Sub

Private Sub WFC_NGS_Trans(xEmpNo) '#19266
Dim rsEmpee As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim SQLQ As String
Dim xUnion As String
Dim xSalHly As String
Dim xInSubGrp As String
Dim xLDate
Dim xNGSStart
Dim xNGSEndDate

    If Not glbNGS_OnFlag Then
        Exit Sub
    End If
    
    SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2 FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmpee.EOF Then
        Exit Sub
    Else
        If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
        If IsNull(rsEmpee("ED_ORG")) Then glbUNION = "" Else glbUNION = rsEmpee("ED_ORG")
        If IsNull(rsEmpee("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = rsEmpee("ED_VADIM1")
        If IsNull(rsEmpee("ED_VADIM2")) Then glbWFCPayGroup = "" Else glbWFCPayGroup = rsEmpee("ED_VADIM2")
    End If
    rsEmpee.Close
    
    'No NGS Sub Group, skip
    If Len(glbWFCNGSSubGroup) = 0 Then Exit Sub

    xLDate = dlpRetireDate.Text 'Date
    
    xNGSStart = ""
    SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1 FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo & ""
    rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpOther.EOF Then
        If IsDate(rsEmpOther("ER_OTHERDATE1")) Then
            xNGSStart = rsEmpOther("ER_OTHERDATE1")
        End If
    End If
    rsEmpOther.Close
    'No NGS Effective Date, skip
    If Len(xNGSStart) = 0 Then Exit Sub

    If glbUNION = "NONE" Or glbUNION = "EXEC" Then
        xSalHly = "Y"
    Else
        xSalHly = "N"
    End If
    
    If glbFrmCaption$ = "Death of an Employee/Spouse" Then 'Ticket #23247 Franks 04/22/2014
        xNGSEndDate = DateAdd("D", 1, CVDate(dlpDeath.Text))
        Call Upt_EmpOtherByField(xEmpNo, "ER_OTHERDATE2", xNGSEndDate)
        Call NGSAuditAdd(xEmpNo, "M", "Employee Death", lStr("Other Date 2"), "", xNGSEndDate, xNGSEndDate)
        glbMsgCustomVal = 0
        Call WFCUpdBenefitEndDatePublic(xEmpNo, xNGSEndDate, "ALL")
    Else 'retirement
        If IsDate(dlpDOther2.Text) Then
            Call Upt_EmpOtherByField(xEmpNo, "ER_OTHERDATE2", CVDate(dlpDOther2.Text))
            Call NGSAuditAdd(xEmpNo, "M", "Retirment", lStr("Other Date 2"), "", CVDate(dlpDOther2.Text), xLDate)
        End If
    End If
    
End Sub

Private Function IsEmpExist(xEmpNo, xType) 'Ticket #23565 Franks 04/10/2013
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim retval As Boolean
    retval = False
    If xType = "A" Then
        SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    Else 'T
        SQLQ = "SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    End If
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        retval = True
    End If
    rsTemp.Close
    IsEmpExist = retval
End Function

Private Sub CheckWFCReptAuthExistNew(xEmpNo, xDate) 'Ticket #29624 Franks 01/06/2017
    If IsWFCReptAuth(xEmpNo, "") Then
        glbWFC_IncePlanID = xEmpNo
        'If glbTermTran Then
             glbWFC_IPPopFormName = "WFCEmpListWithRepTerm"
        'Else
        '    glbWFC_IPPopFormName = "WFCEmpListWithRepTran"
        'End If
        frmCheckListView.lblStDate = xDate ' dlpRetireDate.Text
        frmCheckListView.Show 1
    End If
End Sub

