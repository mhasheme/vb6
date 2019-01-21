VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmURepAuth 
   Caption         =   "Reporting Authority Mass Change"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14310
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   14310
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.EmployeeLookup elpRept 
      Height          =   285
      Index           =   2
      Left            =   9240
      TabIndex        =   5
      Tag             =   "00-New Reporting Authority 3"
      Top             =   1395
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.EmployeeLookup elpRept 
      Height          =   285
      Index           =   1
      Left            =   9240
      TabIndex        =   3
      Tag             =   "00-New Reporting Authority 2"
      Top             =   915
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.EmployeeLookup elpRept 
      Height          =   285
      Index           =   0
      Left            =   9240
      TabIndex        =   1
      Tag             =   "00-New Reporting Authority 1"
      Top             =   435
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin Threed.SSFrame Frame3D1 
      Height          =   1425
      Left            =   120
      TabIndex        =   21
      Top             =   6120
      Width           =   10695
      _Version        =   65536
      _ExtentX        =   18865
      _ExtentY        =   2514
      _StockProps     =   14
      Caption         =   "Update Employees"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin INFOHR_Controls.EmployeeLookup elpEEID 
         Height          =   285
         Left            =   1650
         TabIndex        =   17
         ToolTipText     =   "If Employee Number is not entered then only employee(s) with the matching Old Reporting Authority (1/2/3/4) will be updated."
         Top             =   450
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   8395
         RefreshDescriptionWhen=   2
         MultiSelect     =   -1  'True
      End
      Begin VB.Label lblNotes 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " - All Reporting Authority # affect the same list of employees."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   0
         Left            =   2400
         TabIndex        =   39
         Top             =   1080
         Width           =   5190
      End
      Begin VB.Label lblNotes 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Note: - Employee Number is mandatory when no other Selection Criteria is entered."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   28
         Top             =   840
         Width           =   7050
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
         TabIndex        =   27
         Top             =   495
         Width           =   1290
      End
   End
   Begin INFOHR_Controls.EmployeeLookup elpRept 
      Height          =   285
      Index           =   3
      Left            =   9240
      TabIndex        =   7
      Tag             =   "00-New Reporting Authority 3"
      Top             =   1875
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.EmployeeLookup elpORept 
      Height          =   285
      Index           =   2
      Left            =   2160
      TabIndex        =   4
      Tag             =   "00-Old Reporting Authority 3"
      Top             =   1395
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.EmployeeLookup elpORept 
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Tag             =   "00-Old Reporting Authority 2"
      Top             =   915
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.EmployeeLookup elpORept 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Tag             =   "00-Old Reporting Authority 1"
      Top             =   435
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.EmployeeLookup elpORept 
      Height          =   285
      Index           =   3
      Left            =   2160
      TabIndex        =   6
      Tag             =   "00-Old Reporting Authority 3"
      Top             =   1875
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   1800
      TabIndex        =   16
      Top             =   5640
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   15
      Top             =   5295
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   13
      Top             =   4605
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   11
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   3915
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Tag             =   "00-Specific Department Desired"
      Top             =   3225
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Tag             =   "00-Specific Division Desired"
      Top             =   2880
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   20
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   10
      Tag             =   "00-Enter Union Code"
      Top             =   3570
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Tag             =   "EDPT-Category"
      Top             =   4260
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   14
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   4950
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   38
      Top             =   4995
      Width           =   540
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   37
      Top             =   2925
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
      Left            =   240
      TabIndex        =   36
      Top             =   3270
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
      Left            =   240
      TabIndex        =   35
      Top             =   3615
      Width           =   420
   End
   Begin VB.Label lblEStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   34
      Top             =   3960
      Width           =   1350
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
      Left            =   120
      TabIndex        =   33
      Top             =   2640
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
      Left            =   240
      TabIndex        =   32
      Top             =   4650
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
      Left            =   240
      TabIndex        =   31
      Top             =   5340
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
      Left            =   240
      TabIndex        =   30
      Top             =   5685
      Width           =   1125
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   29
      Top             =   4305
      Width           =   630
   End
   Begin VB.Label lblORep1 
      AutoSize        =   -1  'True
      Caption         =   "Old Rept. Authority 1:"
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   480
      Width           =   1515
   End
   Begin VB.Label lblORep2 
      AutoSize        =   -1  'True
      Caption         =   "Old Rept. Authority 2:"
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   960
      Width           =   1515
   End
   Begin VB.Label lblORep3 
      AutoSize        =   -1  'True
      Caption         =   "Old Rept. Authority 3:"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   1440
      Width           =   1515
   End
   Begin VB.Label lblORep4 
      AutoSize        =   -1  'True
      Caption         =   "Old Rept. Authority 4:"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   1515
   End
   Begin VB.Label lblRep4 
      AutoSize        =   -1  'True
      Caption         =   "New Rept. Authority 4:"
      Height          =   195
      Left            =   7200
      TabIndex        =   22
      Top             =   1920
      Width           =   1605
   End
   Begin VB.Label lblRep3 
      AutoSize        =   -1  'True
      Caption         =   "New Rept. Authority 3:"
      Height          =   195
      Left            =   7200
      TabIndex        =   20
      Top             =   1440
      Width           =   1605
   End
   Begin VB.Label lblRep2 
      AutoSize        =   -1  'True
      Caption         =   "New Rept. Authority 2:"
      Height          =   195
      Left            =   7200
      TabIndex        =   19
      Top             =   960
      Width           =   1605
   End
   Begin VB.Label lblRep1 
      AutoSize        =   -1  'True
      Caption         =   "New Rept. Authority 1:"
      Height          =   195
      Left            =   7200
      TabIndex        =   18
      Top             =   480
      Width           =   1605
   End
End
Attribute VB_Name = "frmURepAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Add Reporting Authority to the Employee's Current Position and
'Current Performance History records
'HR_JOB_HISTORY  & HR_PERFORM_HISTORY
'Both Access Version and SQL Server Version
Option Explicit
Dim fglbESQLQ As String

Private Function chkFURep()
Dim oCode As String, OCodeD As String
Dim xSelEntered  As Boolean
Dim I
Dim x%

chkFURep = False
xSelEntered = False

If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
    clpDept.SetFocus
    Exit Function
ElseIf Len(clpDept.Text) > 0 Then
    xSelEntered = True
End If

If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    MsgBox lStr("If Division Entered - it must be known")
    clpDiv.SetFocus
    Exit Function
ElseIf Len(clpDiv.Text) > 0 Then
    xSelEntered = True
End If

If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    MsgBox lStr("Category code must be valid")
    clpPT.SetFocus
    Exit Function
ElseIf Len(clpPT.Text) > 0 Then
    xSelEntered = True
End If

For x% = 1 To 6
    If Len(clpCode(x%).Text) > 0 And clpCode(x%).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(x%).SetFocus
        Exit Function
    ElseIf Len(clpCode(x%).Text) > 0 Then
        xSelEntered = True
    End If
Next x%


For I = 0 To 3 '2
    If elpRept(I).Caption = "Enter Valid Employee #" Then
        MsgBox "If Employee Entered - they must exist"
        elpRept(I).SetFocus
        Exit Function
    End If
Next
    
If Not elpEEID.ListChecker Then
    Exit Function
'Jerry asked to make the selection criteria and Employee # optional
'ElseIf Len(Trim(elpEEID)) = 0 Then
'    'if any of the selection criteria is entered then Employee # is not mandatory
'    If xSelEntered = False Then
'        MsgBox "Employee Number is mandatory when no selection criteria is entered. You have to enter at least 1 valid Employee Number."
'        elpEEID.SetFocus
'        Exit Function
'    End If
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
chkFURep = True

End Function

Public Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdModify_Click()
Dim Title$, DgDef As Variant, Response%, Msg$
Dim SQLQ As String
Dim x, count As Integer, xRep As Integer, xORep As Integer
Dim z As Integer, duplicate As Integer
Dim recCount As Integer
Dim xOREPTAUSQL As String
Dim yOREPTAUSQL As String


On Error GoTo AddN_Err

'check for all controls
If Not chkFURep() Then Exit Sub

'check if New reporting authority entered
For x = 0 To 3 '2
    If elpRept(x).Text <> "" Then
        xRep = True
    End If
Next x

'Check if any one Old Reporting Authority Entered
xORep = False
For x = 0 To 3
    If elpORept(x).Text <> "" Then
        xORep = True
    End If
Next x

If xRep = False Then
    elpRept(0).SetFocus
    MsgBox "You have to enter at least one valid New Reporting Authority number!", vbExclamation, "info:HR: Missing New Reporting Authority"
    Exit Sub
End If

'check if the user enter 2 times the same reporting authority
For z = 0 To 2 '1
    For x = (z + 1) To 3 '2
        If elpRept(z).Text = elpRept(x).Text And elpRept(z).Text <> "" And elpRept(x).Text <> "" Then
            duplicate = True
            elpRept(z).SetFocus
            'Ticket #29143 - Jerry does not want to stop users from entering duplicate. So suggested to just warn for the duplicates
            'MsgBox "You entered duplicate New Reporting Authority!", vbExclamation, "info:HR: Duplicate New Reporting Authority"
            Msg = "You have entered duplicate New Reporting Authority." & vbCrLf & vbCrLf & "Do you wish to proceed with the duplicate New Reporting Authority?"
            DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
            Response = MsgBox(Msg, DgDef, "info:HR: Duplicate New Reporting Authority")
            If Response = IDNO Then Exit Sub
         End If
    Next x
Next z


'Ticket #22682 - Release 8.0: Added Old Reporting Authority and display the # of employees to update accordingly
If Len(elpEEID) = 0 Then
    ''Check if any one Old Reporting Authority Entered
    'xORep = False
    'For x = 0 To 3
    '    If elpORept(x).Text <> "" Then
    '        xORep = True
    '    End If
    'Next x
    
    'If none of the Old Reporting Authorities entered, then it's okay not to enter the Employee # because then only
    'employees matching the Old Reporting Authorities will be updated.
    'If xORep = False Then
    '    If Len(Trim(elpEEID)) = 0 Then
    '        elpEEID.SetFocus
    '        MsgBox "You have to enter at least one valid Employee Number to update employee(s) with New Reporting Authority!", vbExclamation, "info:HR: Missing Employee Number"
    '        Exit Sub
    '    End If
    'End If
End If

recCount = getRecordCount_Modify
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " employee Position Record " Else Msg$ = Msg$ & " employee Position Records "
    Msg$ = Msg$ & "will be Updated. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No employee Position record found to update."
    Exit Sub
End If

'=========== Begin main function ===============
Dim xREPTAU, yREPTAU
Dim xEmpList, rsJob As New ADODB.Recordset ', rsPERM As New ADODB.Recordset

Screen.MousePointer = HOURGLASS

xEmpList = getEmpnbr(elpEEID)
For z = 0 To 3 '2
    If xORep = True Then
        'Ticket #22682 - Release 8.0: Added Old Reporting Authority and display the # of employees to update accordingly
        'Old Reporting Authority entered, build SQL query
        xOREPTAUSQL = ""
        yOREPTAUSQL = ""
        If Len(elpORept(z).Text) > 0 Then
            Select Case z
            Case 0
                xOREPTAUSQL = "JH_REPTAU = " & getEmpnbr(elpORept(z).Text)
                yOREPTAUSQL = "PH_REPTAU = " & getEmpnbr(elpORept(z).Text)
            Case 1
                xOREPTAUSQL = "JH_REPTAU2 = " & getEmpnbr(elpORept(z).Text)
                yOREPTAUSQL = "PH_REPTAU2 = " & getEmpnbr(elpORept(z).Text)
            Case 2
                xOREPTAUSQL = "JH_REPTAU3 = " & getEmpnbr(elpORept(z).Text)
                yOREPTAUSQL = "PH_REPTAU3 = " & getEmpnbr(elpORept(z).Text)
            Case Else
                xOREPTAUSQL = "JH_REPTAU4 = " & getEmpnbr(elpORept(z).Text)
                yOREPTAUSQL = ""
            End Select
        End If
        
        xREPTAU = ""
        yREPTAU = ""
        If Len(elpRept(z).Text) > 0 Then
            Select Case z
            Case 0
                xREPTAU = "JH_REPTAU"
                yREPTAU = "PH_REPTAU"
            Case 1
                xREPTAU = "JH_REPTAU2"
                yREPTAU = "PH_REPTAU2"
            Case 2
                xREPTAU = "JH_REPTAU3"
                yREPTAU = "PH_REPTAU3"
            Case Else
                xREPTAU = "JH_REPTAU4" 'Ticket #21947 Franks 04/27/2012
                yREPTAU = "" 'Ticket #21947 Franks 04/27/2012
            End Select
        End If
    
        '  update HR_JOB_HISTORY
        If Len(xREPTAU) > 0 Then
            SQLQ = "UPDATE HR_JOB_HISTORY SET " & xREPTAU & " =  " & getEmpnbr(elpRept(z).Text) & " "
            SQLQ = SQLQ & " WHERE JH_CURRENT <> 0"
            If Len(xEmpList) > 0 Then
                SQLQ = SQLQ & " AND JH_EMPNBR IN (" & xEmpList & ") "
            End If
            If Len(fglbESQLQ) > 0 Then
                SQLQ = SQLQ & " AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
            End If
            If Len(xOREPTAUSQL) > 0 Then
                SQLQ = SQLQ & " AND " & xOREPTAUSQL
            End If
            gdbAdoIhr001.Execute SQLQ
        End If
        
        '  update HR_PERFORM_HISTORY
        If Len(yREPTAU) > 0 Then 'Ticket #21947 Franks 04/27/2012, no PH_REPTAU4
            SQLQ = "UPDATE HR_PERFORM_HISTORY SET " & yREPTAU & " =  " & getEmpnbr(elpRept(z).Text) & " "
            SQLQ = SQLQ & " WHERE PH_CURRENT <> 0"
            If Len(xEmpList) > 0 Then
                SQLQ = SQLQ & " AND PH_EMPNBR IN (" & xEmpList & ") "
            End If
            If Len(fglbESQLQ) > 0 Then
                SQLQ = SQLQ & " AND PH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
            End If
            If Len(yOREPTAUSQL) > 0 Then
                SQLQ = SQLQ & " AND " & yOREPTAUSQL
            End If
            gdbAdoIhr001.Execute SQLQ
        End If
    Else
        xREPTAU = ""
        yREPTAU = ""
        If Len(elpRept(z).Text) > 0 Then
            Select Case z
            Case 0
                xREPTAU = "JH_REPTAU"
                yREPTAU = "PH_REPTAU"
            Case 1
                xREPTAU = "JH_REPTAU2"
                yREPTAU = "PH_REPTAU2"
            Case 2
                xREPTAU = "JH_REPTAU3"
                yREPTAU = "PH_REPTAU3"
            Case Else
                xREPTAU = "JH_REPTAU4" 'Ticket #21947 Franks 04/27/2012
                yREPTAU = "" 'Ticket #21947 Franks 04/27/2012
            End Select
           
            '  update HR_JOB_HISTORY
            If Len(xREPTAU) > 0 Then
                SQLQ = "UPDATE HR_JOB_HISTORY SET " & xREPTAU & " =  " & getEmpnbr(elpRept(z).Text) & " "
                SQLQ = SQLQ & " WHERE JH_CURRENT <> 0"
                If Len(xEmpList) > 0 Then
                    SQLQ = SQLQ & " AND JH_EMPNBR IN (" & xEmpList & ") "
                End If
                If Len(fglbESQLQ) > 0 Then
                    SQLQ = SQLQ & " AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
                End If
                gdbAdoIhr001.Execute SQLQ
            End If
            
            '  update HR_PERFORM_HISTORY
            If Len(yREPTAU) > 0 Then 'Ticket #21947 Franks 04/27/2012, no PH_REPTAU4
                SQLQ = "UPDATE HR_PERFORM_HISTORY SET " & yREPTAU & " =  " & getEmpnbr(elpRept(z).Text) & " "
                SQLQ = SQLQ & " WHERE PH_CURRENT <> 0"
                If Len(xEmpList) > 0 Then
                    SQLQ = SQLQ & " AND PH_EMPNBR IN (" & xEmpList & ") "
                End If
                If Len(fglbESQLQ) > 0 Then
                    SQLQ = SQLQ & " AND PH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
                End If
                
                gdbAdoIhr001.Execute SQLQ
            End If
        End If
    End If
Next z

If glbGP Then 'Ticket #21918 Franks 04/30/2012
    Call Employee_Integration_Emplist(xEmpList)
    Call Employee_Master_Integration(getEmpnbr(elpRept(0).Text)) 'supervisor
End If
        

Screen.MousePointer = DEFAULT

SQLQ = "SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0"
If Len(xEmpList) > 0 Then
    SQLQ = SQLQ & " AND JH_EMPNBR IN (" & xEmpList & ") "
End If
If Len(fglbESQLQ) > 0 Then
    SQLQ = SQLQ & " AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
End If
rsJob.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
'rsJob.Open "SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_EMPNBR IN (" & xEmpList & ")", gdbAdoIhr001, adOpenForwardOnly
If rsJob.EOF Then
    MsgBox "No position(s) was setup."
Else
    MsgBox "Update completed."
End If


Exit Sub

AddN_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_JOB_HISTORY", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Employee_Integration_Emplist(xInList)   'Ticket #21918 Franks 04/30/2012
Dim I As Integer
Dim xCnt As Integer
Dim xEList As String
Dim xlEmpNo
    xEList = xInList
    If Len(xEList) > 0 Then
        xCnt = HRSSCharCount(xEList, ",")
        xCnt = xCnt + 1
        For I = 1 To xCnt
            xlEmpNo = CSVGet(xEList, I)
            Call Employee_Master_Integration(xlEmpNo)
        Next
    End If
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE

glbOnTop = "FRMUREPAUTH"

End Sub

Private Sub Form_Load()

glbOnTop = "FRMUREPAUTH"

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Call setCaption(lblDiv)
Call setCaption(lblDept)
Call setCaption(lblLocation)
Call setCaption(lblRegion)
Call setCaption(lblAdmin)
Call setCaption(lblSection)
Call setCaption(lblUnion)
Call setCaption(lblPT)

Call setCaption(lblRep1)
Call setCaption(lblORep1)

Call setCaption(lblRep2)
Call setCaption(lblORep2)

Call setCaption(lblRep3)
Call setCaption(lblORep3)

Call setCaption(lblRep4)
Call setCaption(lblORep4)

Call INI_Controls(Me)
Screen.MousePointer = DEFAULT
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

Private Function getRecordCount_Modify()
    Dim SQLQ As String
    Dim rsJob As New ADODB.Recordset
    Dim recCount As Integer
    Dim xEmpList As String
    Dim xORep As Integer
    Dim xREPTAU As String
    Dim x, z As Integer
    
    'Ticket #22682 - Release 8.0: Added Old Reporting Authority and display the # of employees to update accordingly
    'Check if any one Old Reporting Authority Entered
    xORep = False
    For x = 0 To 3
        If elpORept(x).Text <> "" Then
            xORep = True
        End If
    Next x

    'Ticket #22682 - Release 8.0: Added Old Reporting Authority and display the # of employees to update accordingly
    'Old Reporting Authority entered, build SQL query
    xREPTAU = ""
    If xORep = True Then
        For z = 0 To 3
            If Len(elpORept(z).Text) > 0 Then
                Select Case z
                Case 0
                    If Len(xREPTAU) > 0 Then xREPTAU = xREPTAU & " OR "
                    xREPTAU = xREPTAU & "JH_REPTAU = " & getEmpnbr(elpORept(z).Text)
                Case 1
                    If Len(xREPTAU) > 0 Then xREPTAU = xREPTAU & " OR "
                    xREPTAU = xREPTAU & "JH_REPTAU2 = " & getEmpnbr(elpORept(z).Text)
                Case 2
                    If Len(xREPTAU) > 0 Then xREPTAU = xREPTAU & " OR "
                    xREPTAU = xREPTAU & "JH_REPTAU3 = " & getEmpnbr(elpORept(z).Text)
                Case Else
                    If Len(xREPTAU) > 0 Then xREPTAU = xREPTAU & " OR "
                    xREPTAU = xREPTAU & "JH_REPTAU4 = " & getEmpnbr(elpORept(z).Text)
                End Select
            End If
        Next
    End If

    getRecordCount_Modify = 0
    recCount = 0
    xEmpList = getEmpnbr(elpEEID)
    Call getWSQLQ

    
    SQLQ = "SELECT COUNT(JH_EMPNBR) AS TOT_REC FROM HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE JH_CURRENT <> 0"
    If Len(xEmpList) > 0 Then
        SQLQ = SQLQ & " AND JH_EMPNBR in (" & xEmpList & ") "
    End If
    If Len(fglbESQLQ) > 0 Then
        SQLQ = SQLQ & " AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
    End If
    'Ticket #22682 - Release 8.0: Added Old Reporting Authority and display the # of employees to update accordingly
    If Len(xREPTAU) > 0 Then
        SQLQ = SQLQ & " AND (" & xREPTAU & ")"
    End If
    
    rsJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsJob.EOF Then
        recCount = rsJob("TOT_REC")
    Else
        recCount = 0
    End If
    rsJob.Close
    Set rsJob = Nothing
    
    getRecordCount_Modify = recCount

End Function

Private Function getWSQLQ()

fglbESQLQ = glbSeleDeptUn
If Len(clpDept.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DEPTNO = '" & clpDept.Text & "'"
If Len(clpDiv.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & clpDiv.Text & "' "
If Len(clpCode(1).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & clpCode(1).Text & "' "
If Len(clpCode(2).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & clpCode(2).Text & "' "
If Len(clpCode(3).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMP = '" & clpCode(3).Text & "' "
If Len(clpCode(4).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & clpCode(4).Text & "' "
If Len(clpCode(5).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_REGION = '" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(5).Text & "' "
If Len(clpCode(6).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ADMINBY = '" & clpCode(6).Text & "' "
If Len(clpPT.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_PT = '" & clpPT.Text & "' "

'If Len(elpEEID.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "

End Function
