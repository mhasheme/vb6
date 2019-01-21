VERSION 5.00
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmUEmpNum 
   Appearance      =   0  'Flat
   Caption         =   "Active Employee Number Mass Change"
   ClientHeight    =   5550
   ClientLeft      =   930
   ClientTop       =   1560
   ClientWidth     =   8565
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5550
   ScaleWidth      =   8565
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.EmployeeLookup elpEmpNum 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Tag             =   "11-Old Employee Number"
      Top             =   480
      Width           =   5000
      _ExtentX        =   8811
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin VB.Frame frmlinamar 
      Caption         =   "NEW Employee #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   600
      TabIndex        =   8
      Top             =   2280
      Width           =   7875
      Begin INFOHR_Controls.CodeLookup clpDIV 
         Height          =   285
         Left            =   1410
         TabIndex        =   4
         Top             =   300
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin VB.TextBox txtEmpID 
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
         Height          =   285
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "01-Employee ID in the Division"
         Top             =   810
         Width           =   825
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Message to user"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   4560
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label lblEEID 
         Caption         =   "lblEEID"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   2640
         TabIndex        =   12
         Top             =   1140
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblTitle 
         Caption         =   "Employee ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   570
         TabIndex        =   11
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Facility"
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
         Index           =   0
         Left            =   570
         TabIndex        =   10
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label lblEENum 
         AutoSize        =   -1  'True
         Caption         =   "lblEENum"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2790
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   825
      End
   End
   Begin VB.Frame frmGeneral 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   525
      Left            =   450
      TabIndex        =   13
      Top             =   840
      Width           =   7875
      Begin VB.TextBox txtEmpNum 
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
         Height          =   285
         Index           =   1
         Left            =   1890
         MaxLength       =   9
         TabIndex        =   1
         Tag             =   "11-New Employee Number"
         Top             =   90
         Width           =   1185
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Message to user"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   3330
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label lblEmpNum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NEW Employee #"
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
         Index           =   1
         Left            =   150
         TabIndex        =   14
         Top             =   135
         Width           =   1275
      End
   End
   Begin INFOHR_Controls.DateLookup dlpTermDate 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Tag             =   "41-Date Terminated"
      Top             =   1380
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpNewHireDate 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Tag             =   "41-Date New Hired"
      Top             =   1800
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin VB.Label lblNewHireDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New Hire Date"
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
      Left            =   600
      TabIndex        =   18
      Tag             =   "41-Date Terminated"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label lblTermDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Termination Date"
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
      Left            =   600
      TabIndex        =   16
      Tag             =   "41-Date Terminated"
      Top             =   1380
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label lblEmpNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OLD Employee #"
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
      Left            =   585
      TabIndex        =   7
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label lblMessage 
      Caption         =   "Ensure that no other user is accessing this employee's information during this procedure."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   3990
      Width           =   7995
   End
End
Attribute VB_Name = "frmUEmpNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ODIV, ODivD, xGlbDiv, xGlbDivDesc
Dim tmpUNION

Public Sub cmdClose_Click()
Unload Me
If glbOnTop = "FRMUEMPNUM" Then glbOnTop = ""
End Sub

Public Sub cmdModify_Click()
Dim Title$, Msg$, DgDef As Variant, Response%
Dim SQLQ
Dim dyn_Table As New ADODB.Recordset
On Error GoTo Mod_Err


If Not chkEMPID Then Exit Sub

SQLQ = "Select ED_EMPNBR,ED_ORG_TABL,ED_ORG from HREMP WHERE ED_EMPNBR = " & elpEmpNum(0).Text
dyn_Table.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not dyn_Table.EOF Then
    tmpUNION = dyn_Table("ED_ORG")
End If

If glbNoNONE Then
    If tmpUNION = "NONE" Then
        MsgBox "You Do Not Have Authority For This Transaction"
        glbOnTop = Empty
        'Unload Me
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
End If
If glbNoEXEC Then    'Hemu -EXE
    If tmpUNION = "EXEC" Then  'Hemu -EXE
        MsgBox "You Do Not Have Authority For This Transaction"
        glbOnTop = Empty
        'Unload Me
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
End If
Set dyn_Table = Nothing

If glbLinamar Then
    If lblEEName(0) = "This number already exists" Then
        MsgBox "This number already exists"
        glbOnTop = Empty
        'Unload Me
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
Else
    If lblEEName(1) = "This number already exists" Then
        MsgBox "This number already exists"
        glbOnTop = Empty
        'Unload Me
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
End If
Title$ = "Employee Number Change"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to change the Employee number for this Employee?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

If Not modUpdate() Then
    MsgBox "All the tables could not be updated"
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

Screen.MousePointer = DEFAULT
MsgBox "Records Updated Successfully"
glbLEE_ID = getEmpnbr(txtEmpNum(1)) 'CHANGE THE ID NUMBER OF THE GLOBAL SELECTED EMPLOYEE
DoEvents
elpEmpNum(0).Text = ""
txtEmpNum(1) = ""
clpDIV = ""
txtEmpID = ""
MDIMain.panHelp(0).FloodType = 0
Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).FloodType = 0
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub Form_Activate()
glbOnTop = "FRMUEMPNUM"
Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
glbOnTop = "FRMUEMPNUM"
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim x%
Screen.MousePointer = DEFAULT

glbOnTop = "FRMUEMPNUM"

frmlinamar.Visible = glbLinamar
If glbLinamar Then
    frmlinamar.Left = lblEmpNum(0).Left
    frmlinamar.Top = 1500
    'Ticket #21805 Franks 03/27/2012
    txtEmpNum(1).MaxLength = 10
End If

frmGeneral.Visible = Not glbLinamar

If glbCompSerial = "S/N - 2296W" Then
    lblTermDate.Visible = True
    dlpTermDate.Visible = True
    'Ticket #10580
    lblNewHireDate.Visible = True
    dlpNewHireDate.Visible = True
End If

If glbLEE_ID = 0 Then frmEEFIND.Show 1
If glbLEE_ID > 0 Then
    frmUEmpNum.Show
Else
    frmEEFIND.Show 1  ' find EE if not present
End If

If Len(glbLEE_SName) < 1 Then Exit Sub

Screen.MousePointer = HOURGLASS


If Len(glbLEE_SName) > 0 Then
    Me.elpEmpNum(0).Caption = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If

elpEmpNum(0).Text = ShowEmpnbr(glbLEE_ID)

If glbCompSerial = "S/N - 2296W" Then
    If Not modECountChk() Then
        MsgBox "You have reached the maximum number of employees for your license"
    Else
        If glbSysGen = True Then
            txtEmpNum(1) = glbNextEmpl
        End If
    End If
End If

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

End Sub

Private Function modUpdate()
Dim dyn_Table As New ADODB.Recordset
Dim xCount, xx
Dim SQLQ, x%, xFldTitle, xFld As String, xTable As String
modUpdate = False
On Error GoTo modUpdate_cmdUpdErr
Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 0
'Ticket #20415 - Add Serial # to the select statement so custom tables also gets employee # changed.
'Serial 9999 is by default for all standard info:HR table.
SQLQ = "SELECT * FROM INFO_HR_TABLES WHERE TERMINATION_TABLE=0 AND (SERIAL = 'S/N - 9999W' OR SERIAL = '" & glbCompSerial & "')"

dyn_Table.Open SQLQ, gdbAdoIhr001, adOpenStatic
MDIMain.panHelp(0).FloodPercent = 10
xCount = dyn_Table.RecordCount
xx = 0
Do Until dyn_Table.EOF
    MDIMain.panHelp(0).FloodPercent = (xx / xCount) * 60 + 10
    xTable = dyn_Table("Table_Name")
    If IsNull(dyn_Table("EMPNBR_Alias")) Then xFld = "" Else xFld = dyn_Table("EMPNBR_Alias")
    If InStr(xFld, "_") = 0 Then xFldTitle = "" Else xFldTitle = Left(xFld, 3)
    If dyn_Table("Employee_Keyed") Then
        If xTable = "HREMP_OTHER" Then 'Ticket #16462
            SQLQ = "DELETE FROM " & xTable & " WHERE ER_EMPNBR = " & txtEmpNum(1)
            gdbAdoIhr001.Execute SQLQ
        End If
        Call UpdateEMPNBR(xTable, xFld, xFldTitle, txtEmpNum(1), elpEmpNum(0), clpDIV)
        Select Case xTable
        Case "HREMP"
            'added by Bryan 24/Oct/05 Ticket#9607
            If glbCompSerial = "S/N - 2378W" Then
                Call UpdateEMPNBR(xTable, "ED_PAYROLL_ID", xFldTitle, "'" & txtEmpNum(1) & "'", "'" & elpEmpNum(0) & "'", clpDIV)
            End If
            'added by Frank 17/Nov/08 Ticket#15793
            If glbCompSerial = "S/N - 2390W" Then
                Call UpdateEMPNBR(xTable, "ED_BADGEID", xFldTitle, "'" & txtEmpNum(1) & "'", "'" & elpEmpNum(0) & "'", clpDIV)
            End If
        Case "HR_ATTENDANCE", "HR_ATTENDANCE_HISTORY"
            Call UpdateEMPNBR(xTable, xFldTitle & "SUPER", xFldTitle, txtEmpNum(1), elpEmpNum(0), clpDIV)
        Case "HR_JOB_HISTORY"
            Call UpdateEMPNBR(xTable, xFldTitle & "REPTAU", xFldTitle, txtEmpNum(1), elpEmpNum(0), clpDIV)
            Call UpdateEMPNBR(xTable, xFldTitle & "REPTAU2", xFldTitle, txtEmpNum(1), elpEmpNum(0), clpDIV)
            Call UpdateEMPNBR(xTable, xFldTitle & "REPTAU3", xFldTitle, txtEmpNum(1), elpEmpNum(0), clpDIV)
            'added by Bryan 24/Oct/05 Ticket#9607
            If glbCompSerial = "S/N - 2378W" Then
                Call UpdateEMPNBR(xTable, "JH_PAYROLL_ID", xFldTitle, "'" & txtEmpNum(1) & "'", "'" & elpEmpNum(0) & "'", clpDIV)
            End If
        Case "HR_SALARY_HISTORY"
            'added by Bryan 24/Oct/05 Ticket#9607
            If glbCompSerial = "S/N - 2378W" Then
                Call UpdateEMPNBR(xTable, "SH_PAYROLL_ID", xFldTitle, "'" & txtEmpNum(1) & "'", "'" & elpEmpNum(0) & "'", clpDIV)
            End If
        Case "HR_PERFORM_HISTORY"
            Call UpdateEMPNBR(xTable, xFldTitle & "REPTAU", xFldTitle, txtEmpNum(1), elpEmpNum(0), clpDIV)
            Call UpdateEMPNBR(xTable, xFldTitle & "REPTAU2", xFldTitle, txtEmpNum(1), elpEmpNum(0), clpDIV)
            Call UpdateEMPNBR(xTable, xFldTitle & "REPTAU3", xFldTitle, txtEmpNum(1), elpEmpNum(0), clpDIV)
            
            If glbCompSerial = "S/N - 2279W" Then  'Friesens Corporation
                Call UpdateEMPNBR("HR_PERFORM_FRIESEN", xFld, xFldTitle, txtEmpNum(1), elpEmpNum(0), clpDIV)
                Call UpdateEMPNBR("HR_PERFORM_FRIESEN", xFldTitle & "REPTAU", xFldTitle, txtEmpNum(1), elpEmpNum(0), clpDIV)
            End If
        Case "HR_OCC_HEALTH_SAFETY"
            Call UpdateEMPNBR(xTable, xFldTitle & "EMPNOT", xFldTitle, txtEmpNum(1), elpEmpNum(0), clpDIV)
            Call UpdateEMPNBR(xTable, xFldTitle & "SUPERVISOR", xFldTitle, txtEmpNum(1), elpEmpNum(0), clpDIV)
        End Select
    End If
    dyn_Table.MoveNext
    xx = xx + 1
Loop
MDIMain.panHelp(0).FloodPercent = 70

'Call modStatus
Call modHRAUDIT

If glbLinamar Then Call modTRALOG

'Franks May 09,2002 for Essex County Library
If glbCompSerial = "S/N - 2296W" Then
    If glbSysGen = True And txtEmpNum(1) >= glbNextEmpl Then
        SQLQ = "UPDATE HRPARCO SET PC_NEXT_AVAILABLE_NBR = " & (Val(txtEmpNum(1)) + 1) & " "
        gdbAdoIhr001.Execute SQLQ
    End If
    Call ChgNumInSN2296
End If

If glbCompSerial = "S/N - 2296W" Or glbCompSerial = "S/N - 2380W" Then 'Essex County Library or VitalAire
    'Ticket #12616
    Call AddHRAUDIT
End If

Call Employee_Master_Integration(Val(elpEmpNum(0)), Val(txtEmpNum(1)))

Call UpdVacTimeRequest(Val(txtEmpNum(1)), "C", Val(elpEmpNum(0)))

'George added on May 11,2005 for THE CUSTOMER WHO HAS TIMESHEET
Call UpdTimesheetEMPNum(Val(txtEmpNum(1)), "C", Val(elpEmpNum(0)))

'Franks May 09,2002 for Essex County Library

'''Ticket #25678 franks 06/30/2014 - the pension table mass update uses INFO_HR_TABLES
'''Pension tables Ticket #15537
''If glbSQL Then
''    If glbWFC Then
''        Call UpdHRPensionTables(Val(txtEmpNum(1)), Val(elpEmpNum(0)))
''    End If
''End If

'Ticket #18349 - Change the Employee # in document attachment as well
If gsAttachment_DB Then
    Call UpdateEMPNBR("HRDOC_EMP", "RE_EMPNBR", "RE_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    Call UpdateEMPNBR("HRDOC_JOB_HISTORY", "DJ_EMPNBR", "DJ_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    Call UpdateEMPNBR("HRDOC_COMMENTS", "DO_EMPNBR", "DO_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    Call UpdateEMPNBR("HRDOC_HEALTH_SAFETY", "DE_EMPNBR", "DE_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    Call UpdateEMPNBR("HRDOC_HEALTH_SAFETY_2", "DE_EMPNBR", "DE_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    Call UpdateEMPNBR("HRDOC_COUNSEL", "DC_EMPNBR", "DC_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    Call UpdateEMPNBR("HRDOC_PERFORM_HISTORY", "DH_EMPNBR", "DH_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    Call UpdateEMPNBR("HRDOC_EDSEM", "ES_EMPNBR", "ES_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    Call UpdateEMPNBR("HRDOC_EDSEM_RETEST", "ES_EMPNBR", "ES_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    Call UpdateEMPNBR("HRDOC_HREDU", "EU_EMPNBR", "EU_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    Call UpdateEMPNBR("HRDOC_HRDOLENT", "DE_EMPNBR", "DE_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    
    Call UpdateEMPNBR("HRDOC_EDSEM_SUBMIT", "ES_EMPNBR", "ES_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    
    'New attachment tables
    Call UpdateEMPNBR("HRDOC_ATTENDANCE", "AD_EMPNBR", "AD_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    Call UpdateEMPNBR("HRDOC_EMP_FLAGS", "EF_EMPNBR", "EF_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    Call UpdateEMPNBR("HRDOC_TRADE", "TD_EMPNBR", "TD_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    
    'Applicant Tracker table
    Call UpdateEMPNBR("HRDOC_HRA_EMP", "RE_EMPNBR", "RE_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    Call UpdateEMPNBR("HRDOC_HRA_REFRENCES", "ED_EMPNBR", "ED_", txtEmpNum(1), elpEmpNum(0), clpDIV)
        
    If glbWSIBModule Then
        Call UpdateEMPNBR("HRDOC_HEALTH_SAFETY_CONCERNSWF7", "W7_EMPNBR", "W7_", txtEmpNum(1), elpEmpNum(0), clpDIV)
    End If
End If

MDIMain.panHelp(0).FloodPercent = 100
Screen.MousePointer = DEFAULT
modUpdate = True

Exit Function
modUpdate_cmdUpdErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MDIMain.panHelp(0).FloodType = 0
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modUpdate Error", xTable, "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub ChgNumInSN2296()
Dim gdbAdoSN2296 As New ADODB.Connection
Dim glbAdoSN2296 As String
Dim glbSN2296 As String, SQLQ

    'Ticket #18984 - these tables are in the main database now after the conversion.
    'If gdbAdoSN2296.State = adStateOpen Then gdbAdoSN2296.Close
    'glbSN2296 = Replace(UCase(glbIHRDB), "IHR001.MDB", "SN2296.MDB")
    'glbAdoSN2296 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & glbSN2296
    'gdbAdoSN2296.Mode = adModeReadWrite
    'gdbAdoSN2296.Open glbAdoSN2296
    
    SQLQ = "UPDATE ANN_EMP SET ED_EMPNBR = " & txtEmpNum(1) & " WHERE ED_EMPNBR = " & elpEmpNum(0).Text & " "
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE ANN_EMP_TEMP SET ED_EMPNBR = " & txtEmpNum(1) & " WHERE ED_EMPNBR = " & elpEmpNum(0).Text & " "
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE MTH_SICK SET EMPNBR = " & txtEmpNum(1) & " WHERE EMPNBR = " & elpEmpNum(0).Text & " "
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE Status_Change SET EMPNBR = " & txtEmpNum(1) & " WHERE EMPNBR = " & elpEmpNum(0).Text & " "
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE TIMETABLE SET TI_EMPNBR = " & txtEmpNum(1) & " WHERE TI_EMPNBR = " & elpEmpNum(0).Text & " "
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE TOTALHRS SET EMPNBR = " & txtEmpNum(1) & " WHERE EMPNBR = " & elpEmpNum(0).Text & " "
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE TOTALHRS_02 SET EMPNBR = " & txtEmpNum(1) & " WHERE EMPNBR = " & elpEmpNum(0).Text & " "
    gdbAdoIhr001.Execute SQLQ
    
    'gdbAdoSN2296.Close
End Sub

Private Sub AddHRAUDIT()
Dim SQLQ As String, xEMPNBR As Long
Dim rsAU As New ADODB.Recordset
Dim rsPerform As New ADODB.Recordset
Dim rsEMP As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim rsSAL As New ADODB.Recordset
Dim strFields As String

On Error GoTo HRAUDIT_cmdUpdErr
Screen.MousePointer = HOURGLASS
If rsAU.State <> 0 Then rsAU.Close

'strfields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_COMPNO, AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL,  "
strFields = strFields & "AU_TREAS_TABL, AU_DOLENT_TABL, AU_EARN_TABL, AU_ADMINBY_TABL, AU_LANG1_TABL, AU_LANG2_TABL, "
strFields = strFields & "AU_EMPNBR, AU_TYPE, AU_NEWEMP, AU_SURNAME, AU_FNAME, AU_DOT, AU_LDAY, AU_SALARY, AU_SALCD, AU_PAYP, AU_SEDATE, "
strFields = strFields & "AU_PHONE, AU_SIN, AU_VACPC, AU_DOB, AU_SEX, AU_MSTAT, AU_DEPTNO, AU_LOC, AU_SENDTE, AU_PROV, "
strFields = strFields & "AU_SMOKER, AU_ADDR1, AU_CITY, AU_PCODE, AU_BANK, AU_BRANCH, AU_ACCOUNT, AU_EMP, AU_ORG, "
strFields = strFields & "AU_DOH, AU_LDAY, AU_TD1DOL, AU_TD3, AU_WCB, AU_OMDAY, AU_USRDAT1, AU_WHRS, AU_PHRS, " 'JH_PHRS
strFields = strFields & "AU_UPLOAD, AU_LUSER, AU_LDATE, AU_LTIME, "
strFields = strFields & "AU_DIVUPL"
rsAU.Open "SELECT " & strFields & " FROM HRAUDIT ", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

'For old Number
rsAU.AddNew
rsAU("AU_COMPNO") = "001"
rsAU("AU_LOC_TABL") = "EDLC": rsAU("AU_SECTION_TABL") = "EDSE": rsAU("AU_EMP_TABL") = "EDEM"
rsAU("AU_SUPCODE_TABL") = "EDSP": rsAU("AU_ORG_TABL") = "EDOR": rsAU("AU_PAYP_TABL") = "SDPP"
rsAU("AU_BCODE_TABL") = "BNCD": rsAU("AU_TREAS_TABL") = "TERM": rsAU("AU_DOLENT_TABL") = "EDOL"
rsAU("AU_EARN_TABL") = "EARN": rsAU("AU_ADMINBY_TABL") = "EDAB": rsAU("AU_LANG1_TABL") = "EDL1":: rsAU("AU_LANG2_TABL") = "EDL1"
rsAU("AU_EMPNBR") = elpEmpNum(0).Text
rsAU("AU_TYPE") = "T"
rsAU("AU_NEWEMP") = "N"
rsAU("AU_SURNAME") = glbLEE_SName
rsAU("AU_FNAME") = glbLEE_FName
'SQLQ = "SELECT * FROM HR_PERFORM_HISTORY WHERE PH_CURRENT <> 0 AND PH_EMPNBR = " & txtEmpNum(1) & " "
'rsPerform.Open SQLQ, gdbAdoIhr001, adOpenStatic
'If Not rsPerform.EOF Then
'    If IsDate(rsPerform("PH_PNEXT")) Then
'        rsAU("AU_DOT") = rsPerform("PH_PNEXT")
'        rsAU("AU_LDAY") = rsPerform("PH_PNEXT")
'    Else
'        rsAU("AU_DOT") = Format(Now, "SHORT DATE")
'        rsAU("AU_LDAY") = Format(Now, "SHORT DATE")
'    End If
'Else
'    rsAU("AU_DOT") = Format(Now, "SHORT DATE")
'    rsAU("AU_LDAY") = Format(Now, "SHORT DATE")
If glbCompSerial = "S/N - 2380W" Then 'VitalAire
    rsAU("AU_DOT") = Format(Now, "SHORT DATE")
    rsAU("AU_LDAY") = Format(Now, "SHORT DATE")
Else
    rsAU("AU_DOT") = dlpTermDate.Text
    rsAU("AU_LDAY") = dlpTermDate.Text
End If
'End If
'rsPerform.Close
xEMPNBR = txtEmpNum(1)
SQLQ = "SELECT ED_EMPNBR, ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEMPNBR
rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEMP.EOF Then
    'So that you can use in the Import/Export module to filter the export by Division
    If Not IsNull(rsEMP("ED_DIV")) Then rsAU("AU_DIVUPL") = rsEMP("ED_DIV")
End If
rsEMP.Close

rsAU("AU_UPLOAD") = "N"
rsAU("AU_LUSER") = glbUserID
rsAU("AU_LDATE") = Date
rsAU("AU_LTIME") = Time$
rsAU.Update

xEMPNBR = txtEmpNum(1)
'For New Number

SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEMPNBR
rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
rsAU.AddNew
rsAU("AU_COMPNO") = "001"
rsAU("AU_LOC_TABL") = "EDLC": rsAU("AU_SECTION_TABL") = "EDSE": rsAU("AU_EMP_TABL") = "EDEM"
rsAU("AU_SUPCODE_TABL") = "EDSP": rsAU("AU_ORG_TABL") = "EDOR": rsAU("AU_PAYP_TABL") = "SDPP"
rsAU("AU_BCODE_TABL") = "BNCD": rsAU("AU_TREAS_TABL") = "TERM": rsAU("AU_DOLENT_TABL") = "EDOL"
rsAU("AU_EARN_TABL") = "EARN": rsAU("AU_ADMINBY_TABL") = "EDAB": rsAU("AU_LANG1_TABL") = "EDL1"
rsAU("AU_LANG2_TABL") = "EDL1"
rsAU("AU_EMPNBR") = txtEmpNum(1)
rsAU("AU_TYPE") = "A"
rsAU("AU_NEWEMP") = "Y"
rsAU("AU_SURNAME") = glbLEE_SName
rsAU("AU_FNAME") = glbLEE_FName

If Not IsNull(rsEMP("ED_PHONE")) Then rsAU("AU_PHONE") = rsEMP("ED_PHONE")
If Not IsNull(rsEMP("ED_SIN")) Then rsAU("AU_SIN") = rsEMP("ED_SIN")
If Not IsNull(rsEMP("ED_VACPC")) Then rsAU("AU_VACPC") = rsEMP("ED_VACPC")
If IsDate(rsEMP("ED_DOB")) Then rsAU("AU_DOB") = rsEMP("ED_DOB")
If Not IsNull(rsEMP("ED_SEX")) Then rsAU("AU_SEX") = rsEMP("ED_SEX")
If Not IsNull(rsEMP("ED_MSTAT")) Then rsAU("AU_MSTAT") = rsEMP("ED_MSTAT")
If Not IsNull(rsEMP("ED_DEPTNO")) Then rsAU("AU_DEPTNO") = rsEMP("ED_DEPTNO")
If Not IsNull(rsEMP("ED_LOC")) Then rsAU("AU_LOC") = rsEMP("ED_LOC")
If Not IsNull(rsEMP("ED_SENDTE")) Then rsAU("AU_SENDTE") = rsEMP("ED_SENDTE")
If Not IsNull(rsEMP("ED_PROV")) Then rsAU("AU_PROV") = rsEMP("ED_PROV")
If Not IsNull(rsEMP("ED_DIV")) Then rsAU("AU_DIVUPL") = rsEMP("ED_DIV")

'Hemu - 07/02/2003 Begin - The commented line was giving error, HRAUDIT table has AU_SMOKER
'                          as Text(3) type of field and HREMP table has Yes/No type of field
'                          Change was made in 7.0 first
If Not IsNull(rsEMP("ED_SMOKER")) Then
    If rsEMP("ED_SMOKER") <> 0 Then
        rsAU("AU_SMOKER") = "Yes"
    ElseIf rsEMP("ED_SMOKER") = 0 Then
        rsAU("AU_SMOKER") = "No"
    End If
End If
'If Not IsNull(rsEMP("ED_SMOKER")) Then rsAU("AU_SMOKER") = rsEMP("ED_SMOKER")
'Hemu - 07/02/2003 End


If Not IsNull(rsEMP("ED_ADDR1")) Then rsAU("AU_ADDR1") = rsEMP("ED_ADDR1")
If Not IsNull(rsEMP("ED_CITY")) Then rsAU("AU_CITY") = rsEMP("ED_CITY")
If Not IsNull(rsEMP("ED_PCODE")) Then rsAU("AU_PCODE") = rsEMP("ED_PCODE")
If Not IsNull(rsEMP("ED_BANK")) Then rsAU("AU_BANK") = rsEMP("ED_BANK")
If Not IsNull(rsEMP("ED_BRANCH")) Then rsAU("AU_BRANCH") = rsEMP("ED_BRANCH")
If Not IsNull(rsEMP("ED_ACCOUNT")) Then rsAU("AU_ACCOUNT") = rsEMP("ED_ACCOUNT")
If Not IsNull(rsEMP("ED_EMP")) Then rsAU("AU_EMP") = rsEMP("ED_EMP")
If Not IsNull(rsEMP("ED_ORG")) Then rsAU("AU_ORG") = rsEMP("ED_ORG")
If IsDate(dlpNewHireDate) Then
    rsAU("AU_DOH") = dlpNewHireDate
Else
    If IsDate(rsEMP("ED_DOH")) Then rsAU("AU_DOH") = rsEMP("ED_DOH")
End If
If IsDate(rsEMP("ED_LDAY")) Then rsAU("AU_LDAY") = rsEMP("ED_LDAY")
If Not IsNull(rsEMP("ED_TD1DOL")) Then rsAU("AU_TD1DOL") = rsEMP("ED_TD1DOL")
If Not IsNull(rsEMP("ED_TD3")) Then rsAU("AU_TD3") = rsEMP("ED_TD3")
If Not IsNull(rsEMP("ED_WCB")) Then rsAU("AU_WCB") = rsEMP("ED_WCB")
If Not IsNull(rsEMP("ED_OMERS")) Then rsAU("AU_OMDAY") = rsEMP("ED_OMERS")
If IsDate(rsEMP("ED_USRDAT1")) Then rsAU("AU_USRDAT1") = rsEMP("ED_USRDAT1")

'For Job History
SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & xEMPNBR & " "
If rsJOB.State <> 0 Then rsJOB.Close
rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsJOB.EOF Then
    If Not IsNull(rsJOB("JH_WHRS")) Then rsAU("AU_WHRS") = rsJOB("JH_WHRS")
    If Not IsNull(rsJOB("JH_PHRS")) Then rsAU("AU_PHRS") = rsJOB("JH_PHRS")
End If
rsJOB.Close
'For Salary History
SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & xEMPNBR & " " 'rsSal
If rsSAL.State <> 0 Then rsSAL.Close
rsSAL.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsSAL.EOF Then
    If Not IsNull(rsSAL("SH_SALARY")) Then rsAU("AU_SALARY") = rsSAL("SH_SALARY")
    If Not IsNull(rsSAL("SH_SALCD")) Then rsAU("AU_SALCD") = rsSAL("SH_SALCD")
    If Not IsNull(rsSAL("SH_PAYP")) Then rsAU("AU_PAYP") = rsSAL("SH_PAYP")
    If IsDate(rsSAL("SH_EDATE")) Then rsAU("AU_SEDATE") = rsSAL("SH_EDATE") '
End If
rsSAL.Close
rsAU("AU_UPLOAD") = "N"
rsAU("AU_LUSER") = glbUserID
rsAU("AU_LDATE") = Date
rsAU("AU_LTIME") = Time$
rsAU.Update

rsEMP.Close
rsAU.Close
Screen.MousePointer = DEFAULT

Exit Sub
HRAUDIT_cmdUpdErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "AddHRAUDIT Error", "HRAUDIT", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If
End Sub
'Private Sub modStatus()
'Dim SQLQ, xCount, xx
'Dim rsST As New ADODB.Recordset
'On Error GoTo HRASTATUS_cmdUpdErr
'Screen.MousePointer = HOURGLASS
'rsST.Open "SELECT SC_EMPNBR,SC_LTIME,SC_LUSER,SC_LDATE FROM HRSTATUS WHERE SC_EMPNBR= " & getEmpnbr(elpEmpNum(0)), gdbadoihr001, adOpenStatic, adLockPessimistic
'xCount = rsST.RecordCount
'xx = 0
'Do Until rsST.EOF
'    MDIMain.panHelp(0).FloodPercent = (xx / xCount) * 10 + 80
'    rsST("SC_EMPNBR") = getEmpnbr(txtEmpNum(1))
'    rsST("SC_LTIME") = Time$
'    rsST("SC_LUSER") = glbUserID
'    rsST("SC_LDATE") = Date
'    rsST.MoveNext
'    xx = xx + 1
'Loop
'
'
'Screen.MousePointer = DEFAULT
'
'Exit Sub
'HRASTATUS_cmdUpdErr:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modHRASTATUS Error", "HRASTATUS", "Update")
'Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
'    RollBack
'    Resume Next
'Else
'    Unload Me
'End If
'End Sub

Private Sub modTRALOG()
Dim SQLQ, xCount, xx
Dim rsTL As New ADODB.Recordset
Dim xNewEmpnbr
Dim xOldEmpnbr
On Error GoTo TRALOG_cmdUpdErr
Screen.MousePointer = HOURGLASS
xOldEmpnbr = Val(getEmpnbr(elpEmpNum(0)))
xNewEmpnbr = Val(getEmpnbr(txtEmpNum(1)))

rsTL.Open "SELECT TL_EMPNBR,TL_NEWEMPNBR,TL_NEWDIV,TL_CURRENTDIV,TL_KEY,TL_LTIME,TL_LUSER,TL_LDATE FROM LN_TRALOG WHERE TL_KEY= 'E" & xOldEmpnbr & "'", gdbAdoIhr001, adOpenStatic, adLockPessimistic
Do Until rsTL.EOF
    If rsTL("TL_EMPNBR") = xOldEmpnbr Then
        rsTL("TL_EMPNBR") = xNewEmpnbr
    End If
    If rsTL("TL_NEWEMPNBR") = xOldEmpnbr Then
        rsTL("TL_NEWEMPNBR") = xNewEmpnbr
        rsTL("TL_NEWDIV") = Right(xNewEmpnbr, 3)
        rsTL("TL_CURRENTDIV") = Right(xNewEmpnbr, 3)
    End If
    rsTL("TL_KEY") = "E" & xNewEmpnbr
    rsTL("TL_LTIME") = Time$
    rsTL("TL_LUSER") = glbUserID
    rsTL("TL_LDATE") = Date
    rsTL.MoveNext
    xx = xx + 1
Loop
rsTL.Close

Screen.MousePointer = DEFAULT

Exit Sub
TRALOG_cmdUpdErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modHRASTATUS Error", "HRASTATUS", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub modHRAUDIT()
Dim SQLQ, xCount, xx
Dim rsAU As New ADODB.Recordset

On Error GoTo HRAUDIT_cmdUpdErr

Screen.MousePointer = HOURGLASS

rsAU.Open "SELECT AU_ID,AU_EMPNBR,AU_LTIME,AU_LUSER,AU_LDATE FROM HRAUDIT WHERE AU_EMPNBR= " & getEmpnbr(elpEmpNum(0).Text), gdbAdoIhr001X, adOpenStatic, adLockOptimistic
xCount = rsAU.RecordCount
xx = 0
Do Until rsAU.EOF
    MDIMain.panHelp(0).FloodPercent = (xx / xCount) * 20 + 70
    rsAU("AU_EMPNBR") = getEmpnbr(txtEmpNum(1))
    rsAU("AU_LTIME") = Time$
    rsAU("AU_LUSER") = glbUserID
    rsAU("AU_LDATE") = Date
    rsAU.Update
    rsAU.MoveNext
    xx = xx + 1
Loop
rsAU.Close
Set rsAU = Nothing

rsAU.Open "SELECT AU_ID,AU_EMPNBR,AU_LTIME,AU_LUSER,AU_LDATE FROM HRAUDIT2 WHERE AU_EMPNBR= " & getEmpnbr(elpEmpNum(0).Text), gdbAdoIhr001X, adOpenStatic, adLockOptimistic
xCount = rsAU.RecordCount
xx = 0
Do Until rsAU.EOF
    MDIMain.panHelp(0).FloodPercent = (xx / xCount) * 20 + 70
    rsAU("AU_EMPNBR") = getEmpnbr(txtEmpNum(1))
    rsAU("AU_LTIME") = Time$
    rsAU("AU_LUSER") = glbUserID
    rsAU("AU_LDATE") = Date
    rsAU.Update
    rsAU.MoveNext
    xx = xx + 1
Loop
rsAU.Close
Set rsAU = Nothing

If glbWFC Then
    rsAU.Open "SELECT MT_ID,MT_EMPNBR,MT_LTIME,MT_LUSER,MT_LDATE FROM HR_MANULIFE_TRAN_AUDIT WHERE MT_EMPNBR= " & getEmpnbr(elpEmpNum(0).Text), gdbAdoIhr001X, adOpenStatic, adLockOptimistic
    xCount = rsAU.RecordCount
    xx = 0
    Do Until rsAU.EOF
        MDIMain.panHelp(0).FloodPercent = (xx / xCount) * 20 + 70
        rsAU("MT_EMPNBR") = getEmpnbr(txtEmpNum(1))
        rsAU("MT_LTIME") = Time$
        rsAU("MT_LUSER") = glbUserID
        rsAU("MT_LDATE") = Date
        rsAU.Update
        rsAU.MoveNext
        xx = xx + 1
    Loop
    rsAU.Close
    Set rsAU = Nothing

    rsAU.Open "SELECT NG_ID,NG_EMPNBR,NG_LTIME,NG_LUSER,NG_LDATE FROM WFC_NGS_AUDIT WHERE NG_EMPNBR= " & getEmpnbr(elpEmpNum(0).Text), gdbAdoIhr001X, adOpenStatic, adLockOptimistic
    xCount = rsAU.RecordCount
    xx = 0
    Do Until rsAU.EOF
        MDIMain.panHelp(0).FloodPercent = (xx / xCount) * 20 + 70
        rsAU("NG_EMPNBR") = getEmpnbr(txtEmpNum(1))
        rsAU("NG_LTIME") = Time$
        rsAU("NG_LUSER") = glbUserID
        rsAU("NG_LDATE") = Date
        rsAU.Update
        rsAU.MoveNext
        xx = xx + 1
    Loop
    rsAU.Close
    Set rsAU = Nothing

End If

Screen.MousePointer = DEFAULT

Exit Sub
HRAUDIT_cmdUpdErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modHRAUDIT Error", "HRAUDIT", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmUEmpNum = Nothing 'carmen apr 2000
End Sub

Private Sub lblEENum_Change()
txtEmpNum(1) = lblEENum
Dim rsEMP As New ADODB.Recordset
lblEEName(0) = ""
lblEEName(0).Visible = False
If Len(clpDIV) = 3 And Val(txtEmpID) > 0 Then
    rsEMP.Open "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR=" & txtEmpID & clpDIV, gdbAdoIhr001, adOpenForwardOnly
    If Not rsEMP.EOF Then
        lblEEName(0) = "This number already exists"
        lblEEName(0).Visible = True
    End If
End If
End Sub

Private Sub txtEmpID_Change()
Call CountEmpNbr
End Sub
Private Sub txtEmpID_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub




Private Sub txtEmpNum_Change(Index As Integer)

If Index = 1 Then
    Dim rsEMP As New ADODB.Recordset
    lblEEName(Index) = ""
    lblEEName(Index).Visible = False
    If Len(txtEmpNum(Index)) > 0 Then
        rsEMP.Open "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR=" & txtEmpNum(1), gdbAdoIhr001, adOpenForwardOnly
        If Not rsEMP.EOF Then
            lblEEName(Index) = "This number already exists"
            lblEEName(Index).Visible = True
        End If
    End If
End If
End Sub

Private Sub txtEmpNum_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub CountEmpNbr()
lblEENum.Visible = True
If Len(clpDIV) = 3 And Val(txtEmpID) > 0 Then
    lblEENum = Format(clpDIV, "000") & "-" & Val(txtEmpID)
    lblEEID = Val(txtEmpID) & Format(clpDIV, "000")
Else
    lblEENum = ""
End If
End Sub

Private Function chkEMPID()
Dim rsHREmp As New ADODB.Recordset
Dim SQLQ As String

chkEMPID = False

If Len(elpEmpNum(0).Text) = 0 Then
    MsgBox "The Employee number must be entered."
    elpEmpNum(0).SetFocus
    Exit Function
End If

If elpEmpNum(0).Caption = "Enter Valid Employee #" Then  'the OLD employee number not valid
    MsgBox "The OLD Employee number is not valid."
    elpEmpNum(0).SetFocus
    Exit Function
End If

If getEmpnbr(elpEmpNum(0).Text) = glbEmpNbr Then
    MsgBox "You cannot change your own Employee number."
    elpEmpNum(0).SetFocus
    Exit Function
End If

If Not elpEmpNum(0).ListChecker Then
    Exit Function
End If

'Check if the user has access to change employee # of this employee
SQLQ = "SELECT ED_EMPNBR FROM HREMP "
SQLQ = SQLQ & " WHERE " & glbSeleDeptUn
'Ticket #20711 Franks 08/02/2011
'SQLQ = SQLQ & " AND ED_EMPNBR = " & elpEmpNum(0)
SQLQ = SQLQ & " AND ED_EMPNBR = " & getEmpnbr(elpEmpNum(0).Text)
rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If rsHREmp.EOF Then
    MsgBox "You do not have access to change this OLD Employee number."
    elpEmpNum(0).SetFocus
    Exit Function
End If
rsHREmp.Close
Set rsHREmp = Nothing


If glbLinamar Then
    If clpDIV.Caption = "Unassigned" Or Len(clpDIV) <> 3 Or Not IsNumeric(clpDIV) Then
        MsgBox lStr("Invalid Division")
        clpDIV.SetFocus
        Exit Function
    End If
    If Len(txtEmpID) = 0 Then
        MsgBox "Employee ID is a required field"
        txtEmpID.SetFocus
        Exit Function
    Else
        If Not IsNumeric(txtEmpID) Then
            MsgBox "Invalid Employee ID"
            txtEmpID.SetFocus
            Exit Function
        Else
            If Val(txtEmpID) = 0 Then
                MsgBox "Invalid Employee ID"
                txtEmpID.SetFocus
                Exit Function
            End If
        End If
    End If
    If lblEEName(1).Visible = True Then 'the employee number to change to already exists
      MsgBox "The NEW Employee number already exists."
      txtEmpID.SetFocus
      Exit Function
    End If
Else
    If Len(txtEmpNum(1)) = 0 Then
      MsgBox "The NEW Employee number must be entered."
      txtEmpNum(1).SetFocus
      Exit Function
    End If
    If lblEEName(1).Visible = True Then 'the employee number to change to already exists
      MsgBox "The NEW Employee number already exists."
      txtEmpNum(1).SetFocus
      Exit Function
    End If
    
    If glbCompSerial = "S/N - 2296W" Then
        If Len(dlpTermDate.Text) < 1 Then
            MsgBox ("Termination Date is a required field.")
            dlpTermDate.SetFocus
            Exit Function
        End If
        
        If Not IsDate(dlpTermDate.Text) Then
            MsgBox ("Termination Date is not a valid date.")
            dlpTermDate.SetFocus
            Exit Function
        End If
        
        If Len(dlpNewHireDate.Text) < 1 Then
            MsgBox ("New Hire Date is a required field.")
            dlpNewHireDate.SetFocus
            Exit Function
        End If
        
        If Not IsDate(dlpNewHireDate.Text) Then
            MsgBox ("New Hire Date is not a valid date.")
            dlpNewHireDate.SetFocus
            Exit Function
        End If
    End If
    
End If

chkEMPID = True

End Function

Private Sub clpDIV_Change()
    Call CountEmpNbr
End Sub




Private Sub txtEmpNum_KeyPress(Index As Integer, KeyAscii As Integer)
    If glbLinamar And KeyAscii = 45 Then Exit Sub
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0

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
UpdateRight = True
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property
Public Property Get Updateble() As Boolean
Updateble = True
End Property
Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = False
End Property


Private Sub UpdHRPensionTables(zEMPNBR, Optional zOldEmpnbr)
On Error GoTo EndSub
Dim SQLQ
    SQLQ = "UPDATE HRP_PA_DETAILS SET PE_EMPNBR=" & zEMPNBR & " WHERE PE_EMPNBR = " & zOldEmpnbr
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PA_MASTER SET PE_EMPNBR=" & zEMPNBR & " WHERE PE_EMPNBR = " & zOldEmpnbr
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PENSION_BENEFICIARY SET PE_EMPNBR=" & zEMPNBR & " WHERE PE_EMPNBR = " & zOldEmpnbr
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PENSION_MASTER SET PE_EMPNBR=" & zEMPNBR & " WHERE PE_EMPNBR = " & zOldEmpnbr
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PENSION_MEMBERSHIP SET PE_EMPNBR=" & zEMPNBR & " WHERE PE_EMPNBR = " & zOldEmpnbr
    gdbAdoIhr001.Execute SQLQ
Exit Sub
EndSub:
End Sub
