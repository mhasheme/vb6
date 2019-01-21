VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmNewEmployee 
   Caption         =   "Create New Employee"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frPositionCode 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   4440
      TabIndex        =   24
      Top             =   240
      Visible         =   0   'False
      Width           =   5175
      Begin INFOHR_Controls.CodeLookup clpJob 
         DataField       =   "JH_JOB"
         Height          =   285
         Left            =   1530
         TabIndex        =   25
         Tag             =   "01-Position code"
         Top             =   360
         Width           =   5115
         _ExtentX        =   9022
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   405
         Width           =   1185
      End
   End
   Begin VB.Frame frmUnion 
      Height          =   1215
      Left            =   4440
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtPSAC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   21
         Tag             =   "Enter PSAC Rate Level"
         Top             =   720
         Visible         =   0   'False
         Width           =   945
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataSource      =   " "
         Height          =   285
         Index           =   2
         Left            =   1395
         TabIndex        =   17
         Tag             =   "00-Enter Union Code"
         Top             =   240
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOR"
      End
      Begin INFOHR_Controls.CodeLookup clpVadim1 
         Height          =   285
         Left            =   1395
         TabIndex        =   18
         Top             =   720
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDV1"
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
         TabIndex        =   20
         Top             =   285
         Width           =   420
      End
      Begin VB.Label lblVadim1 
         AutoSize        =   -1  'True
         Caption         =   "Vadim Field 1"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   780
         Width           =   945
      End
   End
   Begin INFOHR_Controls.CodeLookup clpDIv 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   660
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin VB.TextBox txtEmpID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2000
      MaxLength       =   6
      TabIndex        =   3
      Tag             =   "Employee ID in the Division"
      Top             =   1035
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3015
      TabIndex        =   9
      Top             =   1890
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1245
      TabIndex        =   8
      Top             =   1890
      Width           =   1125
   End
   Begin VB.Frame frmRPPCode 
      Height          =   1215
      Left            =   4200
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.ComboBox cmbRPPCode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1305
         TabIndex        =   14
         Tag             =   "10-Choose Pension Code"
         Top             =   420
         Width           =   2550
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RPP Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   750
      End
   End
   Begin Threed.SSCheck chkHrs 
      Height          =   255
      Left            =   1680
      TabIndex        =   22
      Tag             =   "If X-Show Attendance Details"
      Top             =   1440
      Visible         =   0   'False
      Width           =   915
      _Version        =   65536
      _ExtentX        =   1614
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Hourly"
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
      Value           =   -1  'True
   End
   Begin Threed.SSCheck chkSal 
      Height          =   255
      Left            =   3000
      TabIndex        =   23
      Tag             =   "If X-Show Attendance Details"
      Top             =   1440
      Visible         =   0   'False
      Width           =   915
      _Version        =   65536
      _ExtentX        =   1614
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Salary"
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   2400
      Visible         =   0   'False
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataSource      =   " "
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Tag             =   "00-Enter Union Code"
      Top             =   2760
      Visible         =   0   'False
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin MSMask.MaskEdBox medHours 
      Height          =   285
      Index           =   1
      Left            =   1995
      TabIndex        =   6
      Tag             =   "10- Number of hours in work week"
      Top             =   3120
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      DataField       =   "ED_DEPTNO"
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Tag             =   "00-Department"
      Top             =   275
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin MSMask.MaskEdBox medSIN 
      DataField       =   "ED_SIN"
      Height          =   285
      Left            =   2000
      TabIndex        =   0
      Tag             =   "00-Social Insurance Number"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "S.I.N."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      Index           =   11
      Left            =   120
      TabIndex        =   30
      Top             =   315
      Width           =   1335
   End
   Begin VB.Label lblEEStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   2445
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lblUnion2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   2760
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblHrsWeek 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hours/Week"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblEEID 
      Caption         =   "lblEEID"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   3480
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblEENum 
      AutoSize        =   -1  'True
      Caption         =   "lblEENum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3600
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblTitle 
      Caption         =   "Division"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   660
      Width           =   1275
   End
   Begin VB.Label lblTitle 
      Caption         =   "Employee ID"
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1035
      Width           =   1035
   End
End
Attribute VB_Name = "frmNewEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oDiv, ODivD, xGlbDiv, xGlbDivDesc
Dim xHRSoftCOUNTRY As String

Private Sub chkHrs_Click(Value As Integer)
If glbWFC Then
    If chkHrs.Value Then
        chkSal.Value = False
    Else
        chkSal.Value = True
    End If
End If
End Sub

Private Sub chkSal_Click(Value As Integer)
If glbWFC Then
    If chkSal.Value Then
        chkHrs.Value = False
    Else
        chkHrs.Value = True
    End If
End If
End Sub

Private Sub clpCode_Change(Index As Integer)
If glbWFC Then 'Ticket #24695 Franks 11/26/2013
    If Index = 0 Then 'Union
        If Len(clpCode(0).Text) > 0 Then ' And Not clpCode(0).Caption = "Unassigned" Then
            If clpCode(0).Text = "NONE" Or clpCode(0).Text = "EXEC" Then
                chkSal.Value = True
                chkHrs.Value = False
            Else
                chkSal.Value = False
                chkHrs.Value = True
            End If
        End If
    End If
End If
End Sub

Private Sub cmdCancel_Click()
'City of Kawartha Lakes
If glbCompSerial = "S/N - 2363W" Then
    If glbOMERS_Date Or glbUnionCode Then   'From Status/Dates screen
        glbTrsDIV = "Cancel"
        glbTrsVadim1 = "Cancel"
    Else
        glbTrsVadim1 = "Cancel"     'From Demographics screen
        glbTrsEE_ID = ""    'New Hire
    End If
'City of Timmins or City of Niagara Falls
ElseIf glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2276W" Then
    If glbUnionCode Then   'From Status/Dates screen
        glbTrsVadim1 = "Cancel"
    Else
        glbTrsEE_ID = ""    'New Hire
    End If
Else
    glbTrsEE_ID = ""    'New Hire
End If

'Ticket #29660 - Contract Employee - User cancelled the Add process so undo the last # allocated.
If glbWFC And glbWFCContractEmployee Then
    Call WFC_CancelContractEmployeeNo
End If

Unload Me

End Sub

Private Sub cmdOK_Click()
Dim Msg As String
Dim a%
Dim DivCountry

glbHRSoftAction = ""

If glbCompSerial = "S/N - 2394W" And glbLEE_ID <> 0 Then  'St. John's Rehab Hospital 'Ticket #14752
    If Len(Trim(clpJob.Text)) = 0 Then
        MsgBox "Position Code cannot be blank"
        clpJob.SetFocus
        Exit Sub
    ElseIf Len(clpJob.Text) > 0 And clpJob.Caption = "Unassigned" Then
        MsgBox "Position Code is invalid. Please enter a correct Position Code."
        clpJob.SetFocus
        Exit Sub
    Else
        glbJob = clpJob.Text  'Position Code
    End If
    GoTo closeform
End If

If glbWFC And Not glbWFCContractEmployee Then  'Ticket #24184 Franks 09/24/2013
    If glbCandidate > 0 And glbHRSoftType = "NewHire" Then
        If xHRSoftCOUNTRY = "U.S.A." Then 'Ticket #24652 Franks 12/02/2013
            If Len(medHours(1).Text) = 0 Then
                MsgBox "Hours/Week is required for US employees."
                medHours(1).SetFocus
                Exit Sub
            Else
                If Not IsNumeric(medHours(1).Text) Then
                    MsgBox "Invalid Hours/Week."
                    medHours(1).SetFocus
                    Exit Sub
                End If
            End If
        End If
        glbHRSoftAction = "NewEmp"
    End If
    If glbCandidate > 0 And glbHRSoftType = "ReHire" Then
        If Len(Trim(medSIN.Text)) = 0 Then
            MsgBox lblTitle(3).Caption & " is required."
            medSIN.SetFocus
            Exit Sub
        Else
            If Left(medSIN.Text, 4) = "9999" Then 'Ticket #24421 Franks 10/08/2013
                MsgBox "Invalid " & lblTitle(3).Caption
                medSIN.SetFocus
                Exit Sub
            End If
            If Not ValidSIN(xHRSoftCOUNTRY) Then
                Exit Sub
            End If
        End If
        'try to get this employee from Term_HREMP and HREMP
        If GetIhrEmpFromSIN(medSIN.Text) > 0 Then
            'Ticket #24421 Franks 10/08/2013
            If WFCNamesMatched(glbCandidate, glbLEE_SName, glbLEE_FName) Then
                If Not txtEmpID.Text = glbTERM_ID Then
                    'If Employee ID does not match, display a message saying "Employee ID in the termination file does not match the Employee ID in HRsoft. Do  you want to use the Employee ID from the termination file? Click on Yes to use the data or No to treat this record as a New Hire."
                    Msg = "Employee ID in the termination file does not match the Employee ID in HRsoft. Do  you want to use the Employee ID from the termination file?"
                    a% = MsgBox(Msg, 36, "Confirm")
                    If a% <> 6 Then 'Exit Sub
                        'No -  No to treat this record as a New Hire
                        glbHRSoftAction = "NewEmp"
                        ''do new hire
                        txtEmpID.SetFocus
                        Exit Sub
                    Else 'Yes
                        txtEmpID.Text = glbTERM_ID
                        'rehire
                        glbHRSoftAction = "ReHireEmp"
                    End If
                Else
                    'rehire
                    glbHRSoftAction = "ReHireEmp"
                End If
            Else
                MsgBox "Employee Name in HRsoft does not match the Employee Name in the Terminated File. Rehire cannot proceed."
                medSIN.SetFocus
                Exit Sub
            End If
        Else
            'new hire
            Msg = "Cannot find the matching employee in info:HR "
            Msg = Msg & Chr(10) & "Do you want to process as a New Hire?"
            a% = MsgBox(Msg, 36, "Confirm")
            If a% <> 6 Then Exit Sub
            
            glbHRSoftAction = "NewEmp"
            'do new hire
        End If
    End If
End If

If Not glbOMERS_Date Then
    If Not chkEMPID Then Exit Sub
End If

'City of Kawartha Lakes
If glbCompSerial = "S/N - 2363W" Then
    If glbOMERS_Date Then   'Status/Dates screen
        glbTrsDIV = Left(cmbRPPCode.Text, 1)
        GoTo closeform
    ElseIf glbUnionCode Then    'Status/Dates screen
        glbTrsVadim1 = clpVadim1.Text 'Vadim 1
        GoTo closeform
    ElseIf glbUnionDemog Then    'Demographics screen
        glbTrsVadim1 = clpVadim1.Text 'Vadim 1
        glbTrsDIV = clpCode(2).Text 'Union
        GoTo closeform
    Else
        glbTrsDIV = clpCode(2).Text     'New Hire
        GoTo closeform
    End If
'City of Timmins or City of Niagara Falls
ElseIf glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2276W" Then
    If glbUnionCode Then    'Status/Dates screen
        'glbTrsVadim1 = txtPSAC.Text 'ED_VADIM2
        glbTrsVadim1 = clpVadim1.Text 'ED_VADIM2
        GoTo closeform
    End If
End If

If glbWFC And Not glbWFCContractEmployee Then  'Ticket #23247 Franks 04/19/2013
    If Len(Trim(clpDIv.Text)) = 0 Then 'Ticket #23915 Franks 06/13/2013
        MsgBox lStr("Division") & " cannot be blank"
        clpDIv.SetFocus
        Exit Sub
    End If
    If Not Len(Trim(txtEmpID.Text)) = 8 Then 'Ticket #24695 Franks 11/26/2013
        MsgBox "Employee ID is not valid format" & Chr(10) & "It should be Division Number + 4 digit Employee Number"
        txtEmpID.SetFocus
        Exit Sub
    End If
    If Len(Trim(clpCode(1).Text)) = 0 Then
        MsgBox lblEEStatus.Caption & " cannot be blank"
        clpCode(1).SetFocus
        Exit Sub
    End If
    If Len(Trim(clpCode(0).Text)) = 0 Then
        MsgBox lblUnion2.Caption & " cannot be blank"
        clpCode(0).SetFocus
        Exit Sub
    End If
    DivCountry = GetCountryFromDiv(clpDIv.Text)
    If Len(Trim(medHours(1).Text)) = 0 Then
        If DivCountry = "U.S.A." Then
            MsgBox lblHrsWeek.Caption & " is required for US Divisions"
            medHours(1).SetFocus
            Exit Sub
        End If
    End If
    If Len(Trim(medHours(1).Text)) > 0 Then
        If Not IsNumeric((medHours(1).Text)) Then
            MsgBox lblHrsWeek.Caption & " must be numeric"
            medHours(1).SetFocus
            Exit Sub
        End If
    End If
    glbWFCHrsSal = chkHrs
    glbTrsStatus = Trim(clpCode(1).Text)
    glbTrsUnion = Trim(clpCode(0).Text)
    If IsNumeric(medHours(1).Text) Then
        glbTrsHourWeek = Trim(medHours(1).Text)
    Else
        glbTrsHourWeek = 0
    End If
    If medSIN.Visible And Len(medSIN.Text) > 0 Then
        glbSIN = medSIN.Text
    End If
    If Len(txtEmpID.Text) > 0 Then
        If isEmpIDExist(txtEmpID.Text) Then
            Msg = "Sorry, Employee # " & txtEmpID.Text & " Already exists."
            'Msg = Msg & Chr(10) & rsEmp("ED_SURNAME")
            'Msg = Msg & Chr(10) & "Already exists."
            MsgBox Msg
            Exit Sub
        End If
    End If
ElseIf glbWFC And glbWFCContractEmployee Then  'Ticket #29660 - Contract Employees - Adding a New Hire
    If Len(Trim(clpDIv.Text)) = 0 Then
        MsgBox lStr("Division") & " cannot be blank"
        clpDIv.SetFocus
        Exit Sub
    End If
    If Not Len(Trim(txtEmpID.Text)) = 8 Then
        MsgBox "Employee ID is not valid format" & Chr(10) & "It should be Division Number + 4 digit Employee Number"
        clpDIv.SetFocus     'Set the focus to Division as the Division has to be right to generate correct Employee #
        Exit Sub
    End If
    If Len(txtEmpID.Text) > 0 Then
        If isEmpIDExist(txtEmpID.Text) Then
            'Generate another Employee #
            txtEmpID.Text = WFC_GenerateContractEmployeeNo
            lblEEID = txtEmpID
            
            'Msg = "Sorry, Employee # " & txtEmpID.Text & " Already exists."
            'Msg = Msg & Chr(10) & rsEmp("ED_SURNAME")
            'Msg = Msg & Chr(10) & "Already exists."
            'MsgBox Msg
            Exit Sub
        End If
    End If
End If

glbTrsEE_ID = lblEEID
glbTrsDIV = Trim(clpDIv)
glbTrsDept = Trim(clpDept)

closeform:
Unload Me
End Sub

Private Function ValidSIN(xCountry)
Dim retval As Boolean
    ValidSIN = False
    If xCountry = "BAHAMAS" Then
        If Len(medSIN) <> 8 Then
            MsgBox "Invalid National Ins"
            medSIN.SetFocus
            Exit Function
        End If
    Else
        If Len(medSIN) <> 9 Then
            If xCountry = "CANADA" Then
                MsgBox "Invalid SIN"
                medSIN.SetFocus
                Exit Function
            ElseIf xCountry = "U.S.A." Then
                MsgBox "Invalid SSN"
                medSIN.SetFocus
                Exit Function
            ElseIf xCountry = "MEXICO" Then
                If Len(medSIN) <> 11 Then
                    MsgBox "Invalid SSN"
                    medSIN.SetFocus
                    Exit Function
                End If
            Else
                'MsgBox "Invalid National Ins - if Unassigned set to 999999999"
            End If
            'MedSIN.SetFocus
            'Exit Function
        Else
            If xCountry = "CANADA" And (medSIN <> "999999999" Or glbLinamar) Then
                If Not SIN_chk(medSIN.Text) Then
                    MsgBox "Invalid SIN"
                    medSIN.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    retval = True
    
    ValidSIN = retval
End Function

Private Sub CountEmpNbr()
If glbLinamar Then
    lblEENum.Visible = True
    If Len(clpDIv) = 3 And Val(txtEmpID) > 0 Then
        lblEENum = Format(clpDIv, "000") & "-" & Val(txtEmpID)
        lblEEID = Val(txtEmpID) & Format(clpDIv, "000")
    Else
        lblEENum = ""
    End If
Else
    lblEEID = txtEmpID
End If
End Sub

Private Function chkEMPID()
chkEMPID = False

    'City of Kawartha Lakes
    If glbCompSerial = "S/N - 2363W" Then
        If glbUnionDemog Then
            If Not clpCode(2).ListChecker Then  'Union Code
                Exit Function
            ElseIf Not clpVadim1.ListChecker Then    'Vadim Code
                Exit Function
            ElseIf (Len(clpCode(2).Text) = 0 Or Len(clpVadim1.Text) = 0) Then
                MsgBox lblUnion.Caption & " and " & lblVadim1.Caption & " cannot be blank"
                If Len(clpCode(2).Text) = 0 Then
                    clpCode(2).SetFocus
                Else
                    clpVadim1.SetFocus
                End If
                Exit Function
            Else
                chkEMPID = True
                Exit Function
            End If
        ElseIf glbUnionCode Then
            If Not clpVadim1.ListChecker Then    'Vadim Code
                Exit Function
            ElseIf Len(clpVadim1.Text) = 0 Then
                MsgBox lblVadim1.Caption & " cannot be blank"
                clpVadim1.SetFocus
                Exit Function
            Else
                chkEMPID = True
                Exit Function
            End If
        End If
    'City of Timmins or City of Niagara Falls
    ElseIf glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2276W" Then
        If glbUnionCode Then
            'If Len(txtPSAC.Text) = 0 Then
            '    MsgBox lblVadim1.Caption & " cannot be blank"
            '    txtPSAC.SetFocus
            If Not clpVadim1.ListChecker Then    'Vadim Code
                Exit Function
            Else
                chkEMPID = True
                Exit Function
            End If
        End If
    End If
    
    'Ticket #24259 - Dept on New Hire window
    If Len(clpDept.Text) < 1 Then
        MsgBox lStr("Department is a required field")
        clpDept.SetFocus
        Exit Function
    Else
        If clpDept.Caption = "Unassigned" Then
            MsgBox "Department Code must be valid"
            clpDept.SetFocus
            Exit Function
        End If
    End If

    If glbLinamar Then
        If clpDIv.Caption = "Unassigned" Or Len(clpDIv) <> 3 Or Not IsNumeric(clpDIv) Then
            MsgBox lStr("Invalid Division")
            clpDIv.SetFocus
            Exit Function
        End If
    Else
        If clpDIv.Caption = "Unassigned" Then
            MsgBox lStr("Invalid Division")
            clpDIv.SetFocus
            Exit Function
        End If
    End If
    
    'Ticket #29660 - Contract Employees - No validation on Employee # as it's generated by the system
    If glbWFC And Not glbWFCContractEmployee Then
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
    End If
    
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    If glbLinamar Then
        Msg = "You do not have Authority for that Facility"
    Else
        Msg = lStr("You do not have Authority for that Division")
    End If
    rs.Open "SELECT PD_DEPT, PD_DIV FROM HRPASDEP WHERE PD_USERID='" & Replace(glbUserID, "'", "''") & "'", gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF = False And rs.BOF = False Then
        Do
            If Not IsNull(rs("PD_DIV")) Then
                If clpDIv = rs("PD_DIV") Then
                    chkEMPID = True
                End If
            Else
                chkEMPID = True
            End If
            rs.MoveNext
        Loop Until rs.EOF
        If chkEMPID = False Then
            MsgBox Msg
            clpDIv.SetFocus
            Exit Function
        End If
    Else
        MsgBox Msg
        clpDIv.SetFocus
        Exit Function
    End If
    rs.Close
    Set rs = Nothing

chkEMPID = True
End Function

Private Sub clpDiv_Change()
    'Ticket #29660 - Contract Employee - Generate the employee # - do not prompt to enter
    If glbWFC And glbWFCContractEmployee Then
        If Len(clpDIv) = 4 Then
            'Generate Contract Employee #
            txtEmpID = WFC_GenerateContractEmployeeNo
            lblEEID = txtEmpID
        End If
    Else
        Call CountEmpNbr
    End If
End Sub

Private Sub Form_Load()

'Call setCaption(lblTitle(11))
'Call setCaption(lblTitle(0))
lblTitle(11).Caption = lStr("Department")
lblTitle(0).Caption = lStr("Division")

If glbWFC And Not glbWFCContractEmployee Then  'Ticket #23247 Franks 04/19/2013
    'need large form which is current one
    frmNewEmployee.Height = 4095
Else
    frmNewEmployee.Height = 3000    '2730   'Ticket #24259 - Adding Dept on New Hire Window
End If

'City of Kawartha Lakes
If glbCompSerial = "S/N - 2363W" Then
    If glbOMERS_Date Then       'Status/Dates screen
        frmNewEmployee.Caption = "RPP Code"
        lblTitle(0).Visible = False
        lblTitle(1).Visible = False
        clpDIv.Visible = False
        txtEmpID.Visible = False
        frmUnion.Visible = False
    
        'Ticket #24259 - Adding Dept on New Hire Window
        lblTitle(11).Visible = False
        clpDept.Visible = False
        
        frmRPPCode.Visible = True
        frmRPPCode.Left = 585
        cmbRPPCode.Clear
        cmbRPPCode.AddItem "1 - Retirement Age 65"
        cmbRPPCode.AddItem "2 - Retirement Age 60"
    ElseIf glbUnionCode Then    'Status/Dates screen
        frmUnion.Visible = True
        frmUnion.Left = 240
        lblVadim1.Top = lblVadim1.Top - 180
        clpVadim1.Top = clpVadim1.Top - 180
        cmdOK.Left = 2130
        Call setCaption(lblVadim1)
        frmNewEmployee.Caption = lblVadim1.Caption
        
        cmdCancel.Visible = False
        lblUnion.Visible = False
        clpCode(2).Visible = False
        
        frmRPPCode.Visible = False
        lblTitle(0).Visible = False
        lblTitle(1).Visible = False
        clpDIv.Visible = False
        txtEmpID.Visible = False
        
        'Ticket #24259 - Adding Dept on New Hire Window
        lblTitle(11).Visible = False
        clpDept.Visible = False
        
    ElseIf glbUnionDemog Then   'Demographics screen
        frmNewEmployee.Caption = "Union Fields"
        
        cmdCancel.Visible = False
        frmRPPCode.Visible = False
        lblTitle(0).Visible = False
        lblTitle(1).Visible = False
        clpDIv.Visible = False
        txtEmpID.Visible = False
        
        'Ticket #24259 - Adding Dept on New Hire Window
        lblTitle(11).Visible = False
        clpDept.Visible = False
        
        frmUnion.Visible = True
        frmUnion.Left = 240
        Call setCaption(lblUnion)
        Call setCaption(lblVadim1)
        cmdOK.Left = 2130
    End If

'City of Timmins or City of Niagara Falls
ElseIf glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2276W" Then
    If glbUnionCode Then    'Status/Dates screen
        frmUnion.Visible = True
        frmUnion.Left = 240
        lblVadim1.Top = lblVadim1.Top - 180
        clpVadim1.Top = clpVadim1.Top - 180
        clpVadim1.TablName = "EDV2"
        'txtPSAC.Top = clpVadim1.Top - 180
        cmdOK.Left = 2130
        
        If glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls
            frmNewEmployee.Caption = "Union Rate Level"
        Else
            frmNewEmployee.Caption = "PSCAC Rate Level"
        End If
        
        cmdCancel.Visible = False
        lblUnion.Visible = False
        clpCode(2).Visible = False
        'clpVadim1.Visible = False
        txtPSAC.Visible = False
        
        frmRPPCode.Visible = False
        lblTitle(0).Visible = False
        lblTitle(1).Visible = False
        clpDIv.Visible = False
        txtEmpID.Visible = False
    
        'Ticket #24259 - Adding Dept on New Hire Window
        lblTitle(11).Visible = False
        clpDept.Visible = False
    
    End If
ElseIf glbCompSerial = "S/N - 2394W" And glbLEE_ID <> 0 Then 'St. John's Rehab Hospital 'Ticket #14752
    frmNewEmployee.Caption = "Position Code"
    frmRPPCode.Visible = False
    frmUnion.Visible = False
    frPositionCode.Top = 240
    frPositionCode.Left = 120
    frPositionCode.Visible = True
    cmdCancel.Visible = False
    cmdOK.Left = 1965
Else
    frmRPPCode.Visible = False
    frmUnion.Visible = False
    frPositionCode.Visible = False
    
    If Not glbLinamar Then
        txtEmpID.MaxLength = 30
        txtEmpID.Tag = "Employee ID"
        'lblTitle(0).Caption = lStr(lblTitle(0).Caption)
        lblTitle(0).Caption = lStr("Division") 'Ticket #29286 Franks 09/30/2016
    Else
        lblTitle(0).Caption = "Facility"
    End If
End If

If glbWFC And Not glbWFCContractEmployee Then
    Call WFCScreenSetup 'Ticket #23247 Franks 04/19/2013
    Call WFCHRSoftData 'Ticket #24184 Franks 09/11/2013
End If

'Ticket #29660 - Contract Employees - Do not allow Employee # entry
If glbWFC And glbWFCContractEmployee Then
    Me.Caption = "Create New Contractor"
    lblTitle(0).FontBold = True
    txtEmpID.Enabled = False
End If

Call INI_Controls(Me)

End Sub

Private Sub WFCScreenSetup() 'Ticket #23247 Franks 04/19/2013
Dim I As Integer
Dim K As Integer
    frmNewEmployee.Height = 4110    '3795   'Ticket #24259 - Adding Dept.
    K = 110
    If glbCandidate > 0 And glbHRSoftType = "ReHire" Then 'Ticket #24184 Franks 09/24/2013
        Me.Caption = "Rehire Employee"
        lblTitle(3).Top = 30 + 45
        medSIN.Top = 30
        lblTitle(3).Visible = True
        medSIN.Visible = True
    End If
    
    lblTitle(11).Top = lblTitle(11).Top + K
    clpDept.Top = clpDept.Top + K
    lblTitle(0).Top = lblTitle(0).Top + K
    clpDIv.Top = clpDIv.Top + K
    lblTitle(1).Top = lblTitle(1).Top + K
    txtEmpID.Top = txtEmpID.Top + K
    
    'move 3 field up
    I = 1000    '1000    'Ticket #24259 - Adding Dept.
    lblEEStatus.Top = lblEEStatus.Top - I + K
    clpCode(1).Top = clpCode(1).Top - I + K
    lblUnion2.Top = lblUnion2.Top - I + K
    clpCode(0).Top = clpCode(0).Top - I + K
    lblHrsWeek.Top = lblHrsWeek.Top - I + K
    medHours(1).Top = medHours(1).Top - I + K
    'move Hourly, Salary and buttons down
    I = 1200
    chkHrs.Top = chkHrs.Top + I
    chkSal.Top = chkSal.Top + I
    cmdOK.Top = cmdOK.Top + I
    cmdCancel.Top = cmdCancel.Top + I
    
    lblEEStatus.Visible = True
    clpCode(1).Visible = True
    lblUnion2.Visible = True
    clpCode(0).Visible = True
    lblHrsWeek.Visible = True
    medHours(1).Visible = True
    chkHrs.Visible = True
    chkSal.Visible = True

End Sub

Private Sub txtEmpID_GotFocus()
Call SetPanHelp(ActiveControl)
If glbWFC And Not glbWFCContractEmployee Then
    If Len(clpDIv) = 4 Then
        If Len(txtEmpID) = 0 Then
            txtEmpID.Text = clpDIv.Text
            txtEmpID.SetFocus
            SendKeys "{end}" '"{right}"
        End If
    End If
End If
End Sub
Private Sub txtEmpID_Change()
Call CountEmpNbr
End Sub

Private Sub WFCHRSoftData()
Dim rsCand As New ADODB.Recordset
Dim SQLQ As String

xHRSoftCOUNTRY = ""
If glbCandidate > 0 Then
    SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE SF_CANDIDATE = " & glbCandidate & " "
    If rsCand.State <> 0 Then rsCand.Close
    rsCand.Open SQLQ, gdbAdoIhr001, adLockReadOnly
    If Not rsCand.EOF Then
        If Not IsNull(rsCand("SF_ORG")) Then clpCode(0).Text = rsCand("SF_ORG")
        If Not IsNull(rsCand("SF_EMPCODE")) Then clpCode(1).Text = rsCand("SF_EMPCODE")
        If Not IsNull(rsCand("SF_DIV")) Then clpDIv.Text = rsCand("SF_DIV")
        If Not IsNull(rsCand("SF_COUNTRY")) Then xHRSoftCOUNTRY = rsCand("SF_COUNTRY")
        If glbHRSoftType = "NewHire" Then 'Ticket #25562 Franks 06/16/2014
            If Not IsNull(rsCand("SF_EMPNBR")) Then txtEmpID.Text = rsCand("SF_EMPNBR")
        End If
        If glbHRSoftType = "ReHire" Then
            lblTitle(3).FontBold = True
            If Not IsNull(rsCand("SF_EMPNBR")) Then txtEmpID.Text = rsCand("SF_EMPNBR")
            If Not IsNull(rsCand("SF_COUNTRY")) Then
                xHRSoftCOUNTRY = rsCand("SF_COUNTRY")
                If UCase(rsCand("SF_COUNTRY")) = "CANADA" Then
                    lblTitle(3).Caption = "S.I.N."
                    medSIN.MaxLength = 11
                    medSIN.Mask = "###-###-###"
                ElseIf UCase(rsCand("SF_COUNTRY")) = "U.S.A." Or UCase(rsCand("SF_COUNTRY")) = "MEXICO" Then
                    lblTitle(3).Caption = "S.S.N."
                    If UCase(rsCand("SF_COUNTRY")) = "U.S.A." Then
                        medSIN.MaxLength = 11
                        medSIN.Mask = "###-##-####"
                    End If
                    If UCase(rsCand("SF_COUNTRY")) = "MEXICO" Then
                        medSIN.MaxLength = 15
                        medSIN.Mask = "##########-#"
                    End If
                Else
                    lblTitle(3).Caption = "National Ins."
                    medSIN.MaxLength = 15
                    medSIN.Mask = "###############"
                End If
            End If
        End If
    End If
    rsCand.Close
End If

End Sub

Private Function GetIhrEmpFromSIN(xSIN)
Dim rsLocEmp As New ADODB.Recordset
Dim SQLQ As String
Dim retval
    retval = 0
    SQLQ = "SELECT Term_HREMP.*,Employee_Number,Term_DOT,Term_DOR FROM Term_HREMP INNER JOIN Term_HRTRMEMP ON Term_HRTRMEMP.TERM_SEQ = Term_HREMP.TERM_SEQ "
    SQLQ = SQLQ & "WHERE ED_SIN = '" & xSIN & "' "
    SQLQ = SQLQ & "AND (Term_DOR IS NULL) "
    SQLQ = SQLQ & "AND NOT (TERM_REASON = 'TOUT') "
    'SQLQ = SQLQ & "ORDER BY Term_HREMP.TERM_SEQ DESC "
    SQLQ = SQLQ & "ORDER BY Term_DOT DESC, Term_HREMP.TERM_SEQ DESC "
    If rsLocEmp.State <> 0 Then rsLocEmp.Close
    rsLocEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    'If rsLocEmP.EOF Then
    '    rsLocEmP.Close
    '    SQLQ = "SELECT * FROM HREMP "
    '    SQLQ = SQLQ & "WHERE ED_SIN = '" & xSIN & "' "
    '    rsLocEmP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    'End If
    If Not rsLocEmp.EOF Then
        'found
        retval = rsLocEmp("ED_EMPNBR")
        'get emp info - begin ------------------
        glbTermOK = True   'Added 98/05/09 by Andy
        glbTERM_Seq = rsLocEmp("TERM_SEQ")
        glbTERM_ID = rsLocEmp("Employee_Number")
        glbTermDate = rsLocEmp("Term_DOT")
        glbEmpCountry = UCase(rsLocEmp("ED_COUNTRY")) 'Ticket #20429 Franks 06/07/2011
        If Not IsNull(rsLocEmp("ED_FNAME")) Then
            glbLEE_FName = rsLocEmp("ED_FNAME")
        Else
            glbLEE_FName = "*ERROR*"
        End If
        If Not IsNull(rsLocEmp("ED_SURNAME")) Then
            glbLEE_SName = rsLocEmp("ED_SURNAME")
        Else
            glbLEE_SName = "*ERROR*"
        End If
        If IsNull(rsLocEmp("ED_ORG")) Then
            glbUNIONTe = ""
        Else
            glbUNIONTe = rsLocEmp("ED_ORG")
        End If
        
        glbTerm_FName = glbLEE_FName
        glbTerm_SName = glbLEE_SName
        
        If IsDate(rsLocEmp("Term_DOR")) Then
            glbRehireDt = rsLocEmp("Term_DOR")
        Else
            glbRehireDt = ""
        End If
        If glbWFC Then 'Get the glbBand
            glbBand = get_band(glbTERM_Seq)
            If IsNull(rsLocEmp("ED_SIN")) Then 'Ticket #18566
                glbSIN = ""
            Else
                glbSIN = rsLocEmp("ED_SIN")
            End If
        End If
        If Len(txtEmpID.Text) = 0 Then
            txtEmpID.Text = rsLocEmp("ED_EMPNBR")
        End If
        'get emp info - end --------------------
    End If
    GetIhrEmpFromSIN = retval
End Function

Private Function get_band(empNo)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ
    get_band = ""
    SQLQ = "SELECT SH_EMPNBR,SH_BAND FROM Term_SALARY_HISTORY WHERE SH_CURRENT <>0 AND TERM_SEQ = " & empNo
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("SH_BAND")) Then
            get_band = rsTemp("SH_BAND")
        End If
    End If
    rsTemp.Close
End Function

Private Function isEmpIDExist(xEmpNo)
Dim rsEmp As New ADODB.Recordset
Dim SQLQ
Dim retval As Boolean
    retval = False
    SQLQ = "Select ED_EMPNBR from HREMP"
    SQLQ = SQLQ & " where ED_EMPNBR = " & xEmpNo & ""
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.BOF Then
        retval = True
    Else
        'Ticket #26009 Franks 09/16/2014 - "   Check both active and term. They cannot use the same employee number
        If glbWFC Then
            If glbHRSoftAction = "ReHireEmp" Then 'Ticket #27109 Franks 05/26/2015
                '- dont check the term employee for hrsoft rehire
            Else
                rsEmp.Close
                SQLQ = "Select ED_EMPNBR from Term_HREMP"
                SQLQ = SQLQ & " where ED_EMPNBR = " & xEmpNo & ""
                rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsEmp.BOF Then
                    retval = True
                End If
            End If
        End If
    End If
    rsEmp.Close
    isEmpIDExist = retval
End Function

Private Function WFC_GenerateContractEmployeeNo()
    Dim rsDivContEmp As New ADODB.Recordset
    Dim SQLQ As String
    Dim xNextContNo As Long
        
    SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & clpDIv.Text & "'"
    rsDivContEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsDivContEmp.EOF Then
        If Not IsNull(rsDivContEmp("DV_LSTNUM")) And rsDivContEmp("DV_LSTNUM") <> "" Then
            'rsDivContEmp("DV_LSTNUM") = rsDivContEmp("DV_LSTNUM") + 1
            xNextContNo = rsDivContEmp("DV_LSTNUM")
            
NextContractNo:
            xNextContNo = xNextContNo + 1
                        
            'Check if Employee # already exists
            If WFC_ContractEmpNoExists(xNextContNo) Then
                GoTo NextContractNo
            Else
                rsDivContEmp("DV_LSTNUM") = xNextContNo
            End If
        Else
            'Generate # and # format
            'rsDivContEmp("DV_LSTNUM") = clpDIv.Text & Format("1", "0000")
            xNextContNo = clpDIv.Text & Format("1", "0000")
            
NextContractNo1:
            'Check if Employee # already exists
            If WFC_ContractEmpNoExists(xNextContNo) Then
                xNextContNo = xNextContNo + 1

                GoTo NextContractNo1
            Else
                rsDivContEmp("DV_LSTNUM") = xNextContNo
            End If
        End If
        rsDivContEmp("LDate") = Date
        rsDivContEmp("LTime") = Time$
        rsDivContEmp("LUser") = glbUserID
        rsDivContEmp.Update
        WFC_GenerateContractEmployeeNo = rsDivContEmp("DV_LSTNUM")
    Else
        WFC_GenerateContractEmployeeNo = ""
    End If
    rsDivContEmp.Close
    Set rsDivContEmp = Nothing
End Function

Private Function WFC_CancelContractEmployeeNo()
    Dim rsDivContEmp As New ADODB.Recordset
    Dim rsHREmp As New ADODB.Recordset
    Dim SQLQ As String
        
    SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & clpDIv.Text & "'"
    rsDivContEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsDivContEmp.EOF Then
        If Not IsNull(rsDivContEmp("DV_LSTNUM")) And rsDivContEmp("DV_LSTNUM") <> "" Then
            'Make sure no employee with this # already exists just in case.
            SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR = " & txtEmpID.Text
            rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsHREmp.EOF Then
                'No employee exists with this #
                rsDivContEmp("DV_LSTNUM") = rsDivContEmp("DV_LSTNUM") - 1
            End If
            rsHREmp.Close
            Set rsHREmp = Nothing
        End If
        rsDivContEmp("LDate") = Date
        rsDivContEmp("LTime") = Time$
        rsDivContEmp("LUser") = glbUserID
        rsDivContEmp.Update
    End If
    rsDivContEmp.Close
    Set rsDivContEmp = Nothing
End Function

Private Function WFC_ContractEmpNoExists(xEmpNo) As Boolean
    Dim rsHREmp As New ADODB.Recordset
    Dim SQLQ As String

    WFC_ContractEmpNoExists = False
    
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME FROM HREMP WHERE ED_EMPNBR=" & xEmpNo
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If Not rsHREmp.EOF Then
        WFC_ContractEmpNoExists = True
    Else
        WFC_ContractEmpNoExists = False
    End If
    rsHREmp.Close
    Set rsHREmp = Nothing
   
End Function

