VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmUATTHIS 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Mass Update Attendance History"
   ClientHeight    =   7365
   ClientLeft      =   15
   ClientTop       =   1365
   ClientWidth     =   10080
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
   ScaleHeight     =   7365
   ScaleWidth      =   10080
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkUndoCompTime 
      Caption         =   "Remove Brought Forward Compensatory Time from Attendance"
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
      Left            =   1800
      TabIndex        =   17
      Top             =   6840
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.CheckBox chkUndoSenHrs 
      Caption         =   "Remove Brought Forward Seniority Hours from Attendance"
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
      Left            =   1800
      TabIndex        =   16
      Top             =   6480
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CheckBox chkUndoAttArchive 
      Caption         =   "Move back Attendance record(s) from Attendance History to Attendance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   5520
      Width           =   8055
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   8
      Tag             =   "01-Attendance/absentee Reason"
      Top             =   3060
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "ADRE"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   1365
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin VB.CheckBox chkSeniorHrs 
      Caption         =   "Bring Forward Seniority Hours"
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
      Left            =   1800
      TabIndex        =   11
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CheckBox chkCompTime 
      Caption         =   "Bring Forward Compensatory Time"
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
      Left            =   1800
      TabIndex        =   12
      Top             =   4200
      Width           =   3255
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   690
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   360
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   20
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Tag             =   "00-Enter Union Code"
      Top             =   1035
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Tag             =   "EDPT-Category"
      Top             =   1695
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.DateLookup dlpTo 
      Height          =   285
      Left            =   3330
      TabIndex        =   10
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3390
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpFrom 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Tag             =   "40-Date from and including this date forward"
      Top             =   3390
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Tag             =   "10-Enter Employee Number"
      Top             =   2730
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7075
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   5
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   2040
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   6
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   2370
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7920
      Top             =   6120
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
   Begin INFOHR_Controls.DateLookup dlpUndoTo 
      Height          =   285
      Left            =   4890
      TabIndex        =   15
      Tag             =   "40-Date upto and including this date forward"
      Top             =   5985
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpUndoFrom 
      Height          =   285
      Left            =   3120
      TabIndex        =   14
      Tag             =   "40-Date from and including this date forward"
      Top             =   5985
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin VB.Label lblUndoDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Range to move back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      TabIndex        =   30
      Top             =   6000
      Visible         =   0   'False
      Width           =   2715
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
      Left            =   120
      TabIndex        =   29
      Top             =   2370
      Width           =   540
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
      Left            =   120
      TabIndex        =   28
      Top             =   2040
      Width           =   975
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
      Left            =   120
      TabIndex        =   27
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
      Left            =   120
      TabIndex        =   26
      Top             =   2730
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
      TabIndex        =   25
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label textMulti 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "The Union Code and Category will be validated from the Employee Basic Data"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   210
      TabIndex        =   24
      Top             =   4680
      Visible         =   0   'False
      Width           =   6630
   End
   Begin VB.Label lblFrom 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Range"
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
      Left            =   120
      TabIndex        =   23
      Top             =   3390
      Width           =   870
   End
   Begin VB.Label lblReason 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason Code"
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
      Left            =   120
      TabIndex        =   22
      Top             =   3060
      Width           =   975
   End
   Begin VB.Label lblCriteria 
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
      Left            =   120
      TabIndex        =   21
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
      Left            =   120
      TabIndex        =   20
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
      Left            =   120
      TabIndex        =   19
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
      Left            =   120
      TabIndex        =   18
      Top             =   360
      Width           =   555
   End
End
Attribute VB_Name = "frmUATTHIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fTablHREMP As New ADODB.Recordset ' table view of HREMP
Dim snapEntitle As New ADODB.Recordset  'user vier
Dim glbFrmCaption$, glbErrNum&, UPDTCNT
Dim strEMPLIST 'George Mar 16,2006
Dim RSEMPLIST As New ADODB.Recordset 'George Mar 16,2006

Private Function chkATTHIS()
chkATTHIS = False

On Error GoTo chkATTHIS_Err

If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    MsgBox lStr("Invalid Division Code")
    clpDiv.SetFocus
    Exit Function
End If
If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "Invalid Department Code"
    clpDept.SetFocus
    Exit Function
End If
If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    MsgBox lStr("Category code must be valid")
    clpPT.SetFocus
    Exit Function
End If
If Len(clpCode(1).Text) > 0 And clpCode(1).Caption = "Unassigned" Then
    MsgBox lStr("Invalid Union Code")
    clpCode(1).SetFocus
    Exit Function
End If
If Len(clpCode(2).Text) > 0 And clpCode(2).Caption = "Unassigned" Then
    MsgBox "Invalid Employment Status"
    clpCode(2).SetFocus
    Exit Function
End If
'If Len(clpCode(3).Text) > 0 And clpCode(3).Caption = "Unassigned" Then
'    MsgBox "Invalid Reason Code"
'    clpCode(3).SetFocus
'    Exit Function
'End If
If Not clpCode(3).CheckList Then
    clpCode(3).SetFocus
    Exit Function
End If



If Len(clpCode(0).Text) > 0 And clpCode(0).Caption = "Unassigned" Then
    MsgBox lStr("Invalid Location Code")
    clpCode(0).SetFocus
    Exit Function
End If
If Len(clpCode(4).Text) > 0 And clpCode(4).Caption = "Unassigned" Then
    MsgBox "Invalid Section Code"
    clpCode(4).SetFocus
    Exit Function
End If

If chkUndoAttArchive.Value = 0 Then
    'Ticket #14407 - Begin
    If Len(dlpFrom.Text) = 0 Or Len(dlpTo.Text) = 0 Then
        MsgBox "Both From Date and To Date cannot be blank"
        dlpFrom.SetFocus
        Exit Function
    End If
    If Len(dlpFrom.Text) > 0 Then
        If Not IsDate(dlpFrom.Text) Then
            MsgBox "From Date is not a valid date."
            dlpFrom.SetFocus
            Exit Function
        End If
    End If
    If Len(dlpTo.Text) > 0 Then
        If Not IsDate(dlpTo.Text) Then
            MsgBox "To Date is not a valid date."
            dlpTo.SetFocus
            Exit Function
        End If
    End If
    'Ticket #14407 - End
    
    If CVDate(dlpFrom.Text) > CVDate(dlpTo.Text) Then
        MsgBox "Invalid Date Range From/To"
        dlpFrom.SetFocus
        Exit Function
    End If
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

If chkUndoAttArchive.Value = 1 Then
    If Len(dlpUndoFrom.Text) = 0 Or Len(dlpUndoTo.Text) = 0 Then
        MsgBox "Date Range to move back the Attendance record(s) cannot be blank"
        dlpUndoFrom.SetFocus
        Exit Function
    End If
    If Len(dlpUndoFrom.Text) > 0 Then
        If Not IsDate(dlpUndoFrom.Text) Then
            MsgBox "From Date is not a valid date."
            dlpUndoFrom.SetFocus
            Exit Function
        End If
    End If
    If Len(dlpUndoTo.Text) > 0 Then
        If Not IsDate(dlpUndoTo.Text) Then
            MsgBox "To Date is not a valid date."
            dlpUndoTo.SetFocus
            Exit Function
        End If
    End If
    If CVDate(dlpUndoFrom.Text) > CVDate(dlpUndoTo.Text) Then
        MsgBox "Invalid Date Range From/To"
        dlpUndoFrom.SetFocus
        Exit Function
    End If
    
    'If Reason Code entered and Undo B/F Comp Time or Undo B/F Seniority Hours is checked then
    'undo cannot happen. Only one code entry is allowed and with single code the OTBF or BFHR cannot be removed.
    'The calculation will end up being incorrect.
    'Ticket #21417 - Allowing multiple code selects for Reason Code. This condition is not valid then.
    'If Len(clpCode(3).Text) > 0 And (chkUndoCompTime Or chkUndoSenHrs) Then
    'If (chkUndoCompTime Or chkUndoSenHrs) Then
        'MsgBox "Reason Code cannot be selected if you want to remove the Bring Forward Hours " & vbCrLf & "for Seniority Hours and/or Compensatory Time Hours."
        'clpCode(3).SetFocus
        'Exit Function
    'End If
End If

chkATTHIS = True

Exit Function

chkATTHIS_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEntitle", "HRBENFT", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Public Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdModify_Click()
Dim SQLQ As String, x
Dim Title$, Msg$, DgDef As Variant, Response%
Dim recCount As Long

Title = "Mass Update"

If Not gSec_Upd_Attendance Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

On Error GoTo Mod_Err

If Not chkATTHIS() Then Exit Sub

If chkUndoAttArchive.Value = 1 Then
    If chkUndoCompTime.Value <> 1 And chkUndoSenHrs.Value <> 1 Then
        Msg$ = "Are you sure you want to MOVE BACK all the Attendance History Records for this criteria?"
    Else
        If chkUndoCompTime Or chkUndoSenHrs Then
            Msg$ = "Are you sure you want to MOVE BACK all the Attendance History Records for this criteria and " & vbCrLf & "REMOVE the 'Brought Forward' hours from Attendance?"
        End If
    End If
Else
    Msg$ = "Are you sure you want to update all Records for this criteria?"
End If
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2  ' Describe dialog.
Response = MsgBox(Msg, DgDef, Title)    ' Get user response.
If Response = IDNO Then
    Exit Sub
End If

recCount = getRecordsCount
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Record " Else Msg$ = Msg$ & " Records "
    If chkUndoAttArchive.Value = 1 Then
        Msg$ = Msg$ & " to MOVE BACK to Attendance. " & vbCrLf & vbCrLf & " Do you want to proceed?"
    Else
        Msg$ = Msg$ & " to transfer to Attendance History. " & vbCrLf & vbCrLf & " Do you want to proceed?"
    End If
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    If chkUndoAttArchive.Value = 1 Then
        Msg$ = "No Attendance History records found to MOVE BACK to Attendance."
    Else
        Msg$ = "No Attendance records found to tranfer to Attendance History."
    End If
    GoTo End_Note
End If

Msg$ = "Do you want to print a list of employees updated?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.

Screen.MousePointer = HOURGLASS

''Ticket #19720 Franks 01/21/2011
gdbAdoIhr001.Execute "DELETE FROM HR_EMPLIST_WRK " & in_SQL(glbIHRDBW) & " WHERE TT_WRKEMP='" & glbUserID & "'"

If chkUndoAttArchive.Value = 1 Then
    If Not Move_AttendanceHistory_BACK() Then GoTo bpMod
Else
    If Not modUpdateSelection() Then GoTo bpMod
End If

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " Update Completed"
MDIMain.panHelp(2).Caption = ""

If UPDTCNT = 0 Then
    If chkUndoAttArchive.Value = 1 Then
        Msg$ = "No records moved back!"
    Else
        Msg$ = "No records archived!"
    End If
Else
    Msg$ = Str(UPDTCNT)
    If UPDTCNT = 1 Then Msg$ = Msg$ & " Record " Else Msg$ = Msg$ & " Records "
    If chkUndoAttArchive.Value = 1 Then
        Msg$ = Msg$ & "MOVED BACK to " & Chr(10) & "   ATTENDANCE"
    Else
        Msg$ = Msg$ & "Transferred to " & Chr(10) & "   ATTENDANCE HISTORY"
    End If
    Msg$ = Msg$ & Chr(10) & "         Successfully "

    If Response% = IDYES Then    ' Yes response
        'Call set_PrintState(False)
        Screen.MousePointer = HOURGLASS
        
        'Call getWSQLQ("U")
        
        'report name
        'Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList.rpt"
        If glbSQL Then 'Ticket #19720 Franks 01/21/2011
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList2.rpt"
            If Len(glbstrSelCri) >= 0 Then
                Me.vbxCrystal.SelectionFormula = " {HR_EMPLIST_WRK.TT_WRKEMP}='" & glbUserID & "'"
            End If
        Else
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList.rpt"
            If Len(glbstrSelCri) >= 0 Then
                Me.vbxCrystal.SelectionFormula = getWSQLQRPT
            End If
        End If
            
        Me.vbxCrystal.Formulas(0) = "rTitle='Mass Update Attendance History - Employee Details'"
        
        'set location for database tables
        'commented out by Bryan Ticket#12305
        'report uses ODBC even in access
'       If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
'       Else
'           Me.vbxCrystal.Connect = "PWD=petman;"
'           Me.vbxCrystal.DataFiles(0) = glbIHRDB
'       End If
        
        'window title if appropriate
        Me.vbxCrystal.WindowTitle = "Employees-updated Report"
        
        Me.vbxCrystal.Destination = 0
        Screen.MousePointer = DEFAULT
        Me.vbxCrystal.Action = 1
        vbxCrystal.Reset
    End If
End If

End_Note:

DgDef = MB_ICONINFORMATION
MsgBox Msg$, DgDef, Title

bpMod:

Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

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


Private Sub chkCompTime_Click()
    If chkCompTime Then 'Ticket #16064
        MsgBox "Please do not select this option if you are using Overtime Master setup to maintain employee's Banked Time.", vbOKOnly, "Overtime Master"
    End If
End Sub

Private Sub chkUndoAttArchive_Click()
    If chkUndoAttArchive.Value = 1 Then
        lblFrom.Enabled = False
        dlpFrom.Enabled = False
        dlpTo.Enabled = False
        chkSeniorHrs.Enabled = False
        chkCompTime.Enabled = False
        
        lblUndoDate.Visible = True
        dlpUndoFrom.Visible = True
        dlpUndoTo.Visible = True
        chkUndoSenHrs.Visible = True
        chkUndoCompTime.Visible = True
    Else
        lblFrom.Enabled = True
        dlpFrom.Enabled = True
        dlpTo.Enabled = True
        chkSeniorHrs.Enabled = True
        chkCompTime.Enabled = True
        
        lblUndoDate.Visible = False
        dlpUndoFrom.Visible = False
        dlpUndoTo.Visible = False
        chkUndoSenHrs.Visible = False
        chkUndoCompTime.Visible = False
    End If
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMUATTHIS"
End Sub

Private Sub Form_Load()

glbOnTop = "FRMUATTHIS"

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Screen.MousePointer = HOURGLASS

Call setRptCaption(Me)

If glbMulti Then textMulti.Visible = True
textMulti.Caption = "The " & lStr("Union") & " and " & lStr("Category") & " will be validated from the Employee Basic Data"

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
MDIMain.panHelp(0).Caption = "Select FROM the menu the appropriate function."

Set frmUATTHIS = Nothing  'Carmen apr 2000
End Sub

Private Function modUpdateSelection()
'Dim TIHR_DBK As Database, TIHR_DB As Database
Dim TB As New ADODB.Recordset
Dim TD As New ADODB.Recordset
Dim TC As New ADODB.Recordset
Dim rsEListWRK As New ADODB.Recordset ''Ticket #19720 Franks 01/21/2011
Dim xxx, xx1
Dim xEmpnbr, SQLQ, WSQLQ, ESQLQ
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim TotCounter As Double, DelCounter As Double
Dim NumRec

On Error GoTo modUpdateSelection_Err

modUpdateSelection = False


ESQLQ = glbSeleDeptUn
If Len(clpDept.Text) > 0 Then ESQLQ = ESQLQ & " AND ED_DEPTNO='" & clpDept.Text & "'"
If Len(clpDiv.Text) > 0 Then ESQLQ = ESQLQ & " AND ED_DIV='" & clpDiv.Text & "'"

If Len(clpCode(0).Text) > 0 Then ESQLQ = ESQLQ & " AND ED_LOC='" & clpCode(0).Text & "'"
If Len(clpCode(4).Text) > 0 Then ESQLQ = ESQLQ & " AND ED_SECTION='" & clpCode(4).Text & "'"

If Len(clpCode(1).Text) > 0 Then ESQLQ = ESQLQ & " AND ED_ORG='" & clpCode(1).Text & "'"
If Len(clpPT.Text) > 0 Then ESQLQ = ESQLQ & " AND ED_PT='" & clpPT.Text & "'"
If Len(clpCode(2).Text) > 0 Then ESQLQ = ESQLQ & " AND ED_EMP='" & clpCode(2).Text & "'"
If Len(elpEEID.Text) > 0 Then ESQLQ = ESQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "

WSQLQ = " AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & ESQLQ & ")"
If Len(dlpFrom.Text) > 0 Then WSQLQ = WSQLQ & " AND AD_DOA>=" & Date_SQL(dlpFrom.Text)
If Len(dlpTo.Text) > 0 Then WSQLQ = WSQLQ & " AND AD_DOA<=" & Date_SQL(dlpTo.Text)
If Len(clpCode(3).Text) > 0 Then WSQLQ = WSQLQ & " AND AD_REASON IN ('" & Replace(clpCode(3).Text, ",", "','") & "')"

TD.Open "SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE " & WSQLQ, gdbAdoIhr001, adOpenStatic
UPDTCNT = TD.RecordCount
If UPDTCNT > 0 Then
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 10
    MDIMain.panHelp(1).Caption = " Please Wait"
    
    SQLQ = "INSERT INTO HR_ATTENDANCE_HISTORY "
    SQLQ = SQLQ & "(AH_COMPNO,"
    SQLQ = SQLQ & "AH_EMPNBR,"
    SQLQ = SQLQ & "AH_DOA,"
    SQLQ = SQLQ & "AH_REASON_TABL,"
    SQLQ = SQLQ & "AH_REASON,"
    SQLQ = SQLQ & "AH_HRS,"
    SQLQ = SQLQ & "AH_COMM,"
    SQLQ = SQLQ & "AH_CHRGCODE,"
    SQLQ = SQLQ & "AH_SHIFT,"
    SQLQ = SQLQ & "AH_SUPER,"
    SQLQ = SQLQ & "AH_INCID,"
    SQLQ = SQLQ & "AH_SEN,"
    SQLQ = SQLQ & "AH_JOB,"
    SQLQ = SQLQ & "AH_SALARY,"
    SQLQ = SQLQ & "AH_SALCD,"
    SQLQ = SQLQ & "AH_DHRS,"
    SQLQ = SQLQ & "AH_WHRS,"
    SQLQ = SQLQ & "AH_ORG_TABL,"
    SQLQ = SQLQ & "AH_ORG,"
    SQLQ = SQLQ & "AH_INDICATOR,"
    SQLQ = SQLQ & "AH_WCBNBR,"
    SQLQ = SQLQ & "AH_FMLA,"
    SQLQ = SQLQ & "AH_POINT,"
    SQLQ = SQLQ & "AH_LDATE,"
    SQLQ = SQLQ & "AH_LTIME,"
    SQLQ = SQLQ & "AH_LUSER,"
    SQLQ = SQLQ & "AH_EMELEA,"
    SQLQ = SQLQ & "AH_UPLOAD,"
    SQLQ = SQLQ & "AH_CALCHRS,"
    SQLQ = SQLQ & "AH_DISCIPLINE_TABL,"
    SQLQ = SQLQ & "AH_DISCIPLINE,"
    SQLQ = SQLQ & "AH_PAYROLL_ID,"
    SQLQ = SQLQ & "AH_GLNO,"
    SQLQ = SQLQ & "AH_PROJECT_CODE,"
    SQLQ = SQLQ & "AH_MACHINE_NUM,"
    SQLQ = SQLQ & "AH_MACHINE_HRS,"
    SQLQ = SQLQ & "AH_MACHINE_RATE,"
    SQLQ = SQLQ & "AH_LEPOINT,"
    SQLQ = SQLQ & "AH_PAYENDDATE, "
    SQLQ = SQLQ & "AH_BANKHRS_EXP,"
    SQLQ = SQLQ & "AH_CONSUMED,"
    SQLQ = SQLQ & "AH_DOCKEY,"
    SQLQ = SQLQ & "AH_SOURCE,"
    SQLQ = SQLQ & "AH_SOURCE_TABL,"
    SQLQ = SQLQ & "AH_REQID,"
    SQLQ = SQLQ & "AH_FTIME,"
    SQLQ = SQLQ & "AH_TTIME,"
    SQLQ = SQLQ & "AH_REGION_TABL,"
    SQLQ = SQLQ & "AH_REGION,"
    SQLQ = SQLQ & "AH_SALDIST )"
    
    SQLQ = SQLQ & "SELECT "
    SQLQ = SQLQ & "AD_COMPNO AS AH_COMPNO,"
    SQLQ = SQLQ & "AD_EMPNBR AS AH_EMPNBR,"
    SQLQ = SQLQ & "AD_DOA AS AH_DOA,"
    SQLQ = SQLQ & "AD_REASON_TABL AS AH_REASON_TABL,"
    SQLQ = SQLQ & "AD_REASON AS AH_REASON,"
    SQLQ = SQLQ & "AD_HRS AS AH_HRS,"
    SQLQ = SQLQ & "AD_COMM AS AH_COMM,"
    SQLQ = SQLQ & "AD_CHRGCODE AS AH_CHRGCODE,"
    SQLQ = SQLQ & "AD_SHIFT AS AH_SHIFT,"
    SQLQ = SQLQ & "AD_SUPER AS AH_SUPER,"
    SQLQ = SQLQ & "AD_INCID AS AH_INCID,"
    SQLQ = SQLQ & "AD_SEN AS AH_SEN,"
    SQLQ = SQLQ & "AD_JOB AS AH_JOB,"
    SQLQ = SQLQ & "AD_SALARY AS AH_SALARY,"
    SQLQ = SQLQ & "AD_SALCD AS AH_SALCD,"
    SQLQ = SQLQ & "AD_DHRS AS AH_DHRS,"
    SQLQ = SQLQ & "AD_WHRS AS AH_WHRS,"
    SQLQ = SQLQ & "AD_ORG_TABL AS AH_ORG_TABL,"
    SQLQ = SQLQ & "AD_ORG AS AH_ORG,"
    SQLQ = SQLQ & "AD_INDICATOR AS AH_INDICATOR,"
    SQLQ = SQLQ & "AD_WCBNBR AS AH_WCBNBR,"
    SQLQ = SQLQ & "AD_FMLA AS AH_FMLA,"
    SQLQ = SQLQ & "AD_POINT AS AH_POINT," 'Ticket #13009
    SQLQ = SQLQ & Date_SQL(Date) & " AS AH_LDATE,"
    SQLQ = SQLQ & "'" & Time$ & "' AS AH_LTIME,"
    SQLQ = SQLQ & "'" & glbUserID & "'AS AH_LUSER,"
    SQLQ = SQLQ & "AD_EMELEA AS AH_EMELEA,"
    SQLQ = SQLQ & "AD_UPLOAD AS AH_UPLOAD,"
    SQLQ = SQLQ & "AD_CALCHRS AS AH_CALCHRS,"
    SQLQ = SQLQ & "AD_DISCIPLINE_TABL AS AH_DISCIPLINE_TABL,"
    SQLQ = SQLQ & "AD_DISCIPLINE AS AH_DISCIPLINE,"
    SQLQ = SQLQ & "AD_PAYROLL_ID AS AH_PAYROLL_ID,"
    SQLQ = SQLQ & "AD_GLNO AS AH_GLNO,"
    SQLQ = SQLQ & "AD_PROJECT_CODE AS AH_PROJECT_CODE,"
    SQLQ = SQLQ & "AD_MACHINE_NUM AS AH_MACHINE_NUM,"
    SQLQ = SQLQ & "AD_MACHINE_HRS AS AH_MACHINE_HRS,"
    SQLQ = SQLQ & "AD_MACHINE_RATE AS AH_MACHINE_RATE,"
    SQLQ = SQLQ & "AD_LEPOINT AS AH_LEPOINT,"
    SQLQ = SQLQ & "AD_PAYENDDATE AS AH_PAYENDDATE,"
    SQLQ = SQLQ & "AD_BANKHRS_EXP AS AH_BANKHRS_EXP,"
    SQLQ = SQLQ & "AD_CONSUMED AS AH_CONSUMED,"
    SQLQ = SQLQ & "AD_DOCKEY AS AH_DOCKEY,"
    SQLQ = SQLQ & "AD_SOURCE AS AH_SOURCE,"
    SQLQ = SQLQ & "AD_SOURCE_TABL AS AH_SOURCE_TABL,"
    SQLQ = SQLQ & "AD_REQID AS AH_REQID,"
    SQLQ = SQLQ & "AD_FTIME AS AH_FTIME,"
    SQLQ = SQLQ & "AD_TTIME AS AH_TTIME,"
    SQLQ = SQLQ & "AD_REGION_TABL AS AH_REGION_TABL,"
    SQLQ = SQLQ & "AD_REGION AS AH_REGION,"
    SQLQ = SQLQ & "AD_SALDIST AS AH_SALDIST"
    
    SQLQ = SQLQ & " FROM HR_ATTENDANCE"
    SQLQ = SQLQ & " WHERE " & WSQLQ
    gdbAdoIhr001.Execute SQLQ
    
    If chkSeniorHrs Or chkCompTime Then
        Call DealBringForward(WSQLQ)
    End If
    
    SQLQ = "SELECT distinct AD_EMPNBR FROM HR_ATTENDANCE WHERE " & WSQLQ
    
    RSEMPLIST.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly
    If glbSQL Then ''Ticket #19720 Franks 01/21/2011
        SQLQ = "SELECT * FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
        rsEListWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    Do While Not RSEMPLIST.EOF
        If Len(strEMPLIST) > 0 Then
            strEMPLIST = strEMPLIST & "," & RSEMPLIST("AD_EMPNBR")
        Else
            strEMPLIST = strEMPLIST & RSEMPLIST("AD_EMPNBR")
        End If
        If glbSQL Then ''Ticket #19720 Franks 01/21/2011
            rsEListWRK.AddNew
            rsEListWRK("TT_COMPNO") = "001"
            rsEListWRK("TT_EMPNBR") = RSEMPLIST("AD_EMPNBR")
            rsEListWRK("TT_WRKEMP") = glbUserID
            rsEListWRK.Update
        End If
        RSEMPLIST.MoveNext
    Loop
    RSEMPLIST.Close
    If glbSQL Then ''Ticket #19720 Franks 01/21/2011
        rsEListWRK.Close
    End If
    
    MDIMain.panHelp(0).FloodPercent = 30
    SQLQ = "DELETE FROM HR_ATTENDANCE WHERE " & WSQLQ
    gdbAdoIhr001.Execute SQLQ
    MDIMain.panHelp(0).FloodPercent = 50
    Call EntReCalc(ESQLQ)
    MDIMain.panHelp(0).FloodPercent = 70
    Call EntReCalcHr
    MDIMain.panHelp(0).FloodPercent = 80
    
    'Town of Aurora
    'If glbCompSerial = "S/N - 2378W" Then
        Call ReCalcOvt("")
    'End If
    
    'Hemu - Begin -  Town of Ajax - OT Bank
    If glbCompSerial = "S/N - 2173W" Then 'And chkCompTime Then
        Call Recalculate_OTBANK
    End If
    'Hemu
    MDIMain.panHelp(0).FloodPercent = 90

    If glbWHSCC Then Call Update_ASL_WHSCC(ESQLQ)
End If

modUpdateSelection = True
Screen.MousePointer = DEFAULT

Exit Function

modUpdateSelection_Err:

If Err = 94 Or Err = 3018 Then
    MsgBox "ERROR " & Err
    Err = 0
    Resume Next
    Screen.MousePointer = DEFAULT
    Exit Function
End If

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateEntitle", "HR_EMP", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Function

Private Sub DealBringForward(xSQLQ)
Dim rsTemp As New ADODB.Recordset
Dim rsTabl As New ADODB.Recordset
Dim RsAttHis As New ADODB.Recordset
Dim RsCT As New ADODB.Recordset
Dim rsCurSal As New ADODB.Recordset
Dim SQLQ As String, xFDate
Dim fTempSum, fTempSumCT

    
    If chkSeniorHrs Then
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_KEY = 'BFHR' "
        rsTabl.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsTabl.EOF Then
            rsTabl.AddNew
            rsTabl("TB_COMPNO") = "001"
            rsTabl("TB_NAME") = "ADRE"
            rsTabl("TB_KEY") = "BFHR"
            rsTabl("TB_DESC") = "BRING FORWARD SENIORITY HOURS"
            rsTabl("TB_LDATE") = Date
            rsTabl("TB_LTIME") = Time$
            rsTabl("TB_LUSER") = glbUserID
            rsTabl.Update
        End If
        rsTabl.Close
        
        SQLQ = xSQLQ & " AND AD_SEN <> 0 "
        SQLQ = "SELECT AD_COMPNO, AD_EMPNBR, SUM(AD_HRS) AS SUMHRS FROM HR_ATTENDANCE WHERE " & SQLQ
        SQLQ = SQLQ & " GROUP BY AD_COMPNO, AD_EMPNBR"
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        xFDate = DateAdd("d", 1, CVDate(dlpTo.Text))
        Do While Not rsTemp.EOF
            SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsTemp("AD_EMPNBR") & " "
            SQLQ = SQLQ & " AND AD_DOA = " & Date_SQL(xFDate)
            SQLQ = SQLQ & " AND AD_REASON = 'BFHR' "
            RsAttHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If RsAttHis.EOF Then
                RsAttHis.AddNew
                
                SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & rsTemp("AD_EMPNBR")
                rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                If Not rsCurSal.BOF Then
                    If rsCurSal("SH_SALARY") > 0 Then
                        RsAttHis("AD_SALARY") = rsCurSal("SH_SALARY")
                        RsAttHis("AD_SALCD") = rsCurSal("SH_SALCD")
                    End If
                End If
                rsCurSal.Close
                Set rsCurSal = Nothing
                
                SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS,JH_REPTAU,JH_PAYROLL_ID,JH_SHIFT,JH_GLNO,JH_ORG FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & rsTemp("AD_EMPNBR")
                rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                If Not rsCurSal.EOF Then
                    RsAttHis("AD_JOB") = rsCurSal("JH_JOB")
                    RsAttHis("AD_DHRS") = rsCurSal("JH_DHRS")
                    RsAttHis("AD_WHRS") = rsCurSal("JH_WHRS")
                    RsAttHis("AD_SUPER") = rsCurSal("JH_REPTAU")
                    RsAttHis("AD_PAYROLL_ID") = rsCurSal("JH_PAYROLL_ID")
                    RsAttHis("AD_SHIFT") = rsCurSal("JH_SHIFT")
                    RsAttHis("AD_GLNO") = rsCurSal("JH_GLNO")
                    RsAttHis("AD_ORG") = rsCurSal("JH_ORG")
                End If
                rsCurSal.Close
                Set rsCurSal = Nothing
                
            End If
            RsAttHis("AD_COMPNO") = rsTemp("AD_COMPNO")
            RsAttHis("AD_EMPNBR") = rsTemp("AD_EMPNBR")
            RsAttHis("AD_DOA") = CVDate(xFDate)
            RsAttHis("AD_REASON") = "BFHR"
            RsAttHis("AD_HRS") = rsTemp("SUMHRS")
            RsAttHis("AD_SEN") = -1
            RsAttHis("AD_LDATE") = Date
            RsAttHis("AD_LUSER") = glbUserID
            RsAttHis("AD_LTIME") = Time$
            RsAttHis.Update
            RsAttHis.Close
            rsTemp.MoveNext
        Loop
        rsTemp.Close
    End If
    
    If chkCompTime Then
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_KEY = 'OTBF' "
        rsTabl.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsTabl.EOF Then
            rsTabl.AddNew
            rsTabl("TB_COMPNO") = "001"
            rsTabl("TB_NAME") = "ADRE"
            rsTabl("TB_KEY") = "OTBF"
            rsTabl("TB_DESC") = "BRING FORWARD COMP HOURS"
            rsTabl("TB_LDATE") = Date
            rsTabl("TB_LTIME") = Time$
            rsTabl("TB_LUSER") = glbUserID
            rsTabl.Update
        End If
        rsTabl.Close
        
        SQLQ = xSQLQ & " AND AD_REASON like 'OT%' "
        SQLQ = "SELECT AD_COMPNO, AD_EMPNBR, SUM(AD_HRS) AS SUMHRS FROM HR_ATTENDANCE WHERE " & SQLQ
        SQLQ = SQLQ & " GROUP BY AD_COMPNO, AD_EMPNBR"
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        xFDate = DateAdd("d", 1, CVDate(dlpTo.Text))
        Do While Not rsTemp.EOF
            SQLQ = "SELECT AD_COMPNO, AD_EMPNBR, SUM(AD_HRS) AS SUMHRS FROM HR_ATTENDANCE "
            SQLQ = SQLQ & "WHERE AD_EMPNBR = " & rsTemp("AD_EMPNBR")
            If glbCompSerial = "S/N - 2188W" Then
                SQLQ = SQLQ & " AND AD_REASON IN ('CT','CTP') and " & xSQLQ
            Else
                SQLQ = SQLQ & " AND AD_REASON like 'CT%' and " & xSQLQ
            End If
            SQLQ = SQLQ & " GROUP BY AD_COMPNO, AD_EMPNBR"
            RsCT.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If RsCT.EOF And RsCT.BOF Then
                fTempSumCT = 0
            Else
                fTempSumCT = RsCT("SUMHRS")
            End If
            SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsTemp("AD_EMPNBR") & " "
            SQLQ = SQLQ & "AND AD_DOA = " & Date_SQL(xFDate)
            SQLQ = SQLQ & "AND AD_REASON = 'OTBF' "
            RsAttHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If RsAttHis.EOF Then
                RsAttHis.AddNew
            
                SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & rsTemp("AD_EMPNBR")
                rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                If Not rsCurSal.BOF Then
                    If rsCurSal("SH_SALARY") > 0 Then
                        RsAttHis("AD_SALARY") = rsCurSal("SH_SALARY")
                        RsAttHis("AD_SALCD") = rsCurSal("SH_SALCD")
                    End If
                End If
                rsCurSal.Close
                Set rsCurSal = Nothing
                
                SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS,JH_REPTAU,JH_PAYROLL_ID,JH_SHIFT,JH_GLNO,JH_ORG FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & rsTemp("AD_EMPNBR")
                rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                If Not rsCurSal.EOF Then
                    RsAttHis("AD_JOB") = rsCurSal("JH_JOB")
                    RsAttHis("AD_DHRS") = rsCurSal("JH_DHRS")
                    RsAttHis("AD_WHRS") = rsCurSal("JH_WHRS")
                    RsAttHis("AD_SUPER") = rsCurSal("JH_REPTAU")
                    RsAttHis("AD_PAYROLL_ID") = rsCurSal("JH_PAYROLL_ID")
                    RsAttHis("AD_SHIFT") = rsCurSal("JH_SHIFT")
                    RsAttHis("AD_GLNO") = rsCurSal("JH_GLNO")
                    RsAttHis("AD_ORG") = rsCurSal("JH_ORG")
                End If
                rsCurSal.Close
                Set rsCurSal = Nothing
            
            End If
            fTempSum = rsTemp("SUMHRS") - fTempSumCT
            RsAttHis("AD_COMPNO") = rsTemp("AD_COMPNO")
            RsAttHis("AD_EMPNBR") = rsTemp("AD_EMPNBR")
            RsAttHis("AD_DOA") = CVDate(xFDate)
            RsAttHis("AD_REASON") = "OTBF"
            RsAttHis("AD_HRS") = fTempSum
            RsAttHis("AD_SEN") = 0 '-1 As Linda request, turn off the seniority flag
            RsAttHis("AD_LDATE") = Date
            RsAttHis("AD_LUSER") = glbUserID
            RsAttHis("AD_LTIME") = Time$
            RsAttHis.Update
            RsAttHis.Close
            RsCT.Close
            rsTemp.MoveNext
        Loop
        rsTemp.Close
    End If
    
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
'UpdateRight = gSec_Upd_Attendance
UpdateRight = GetMassUpdateSecurities("Attendance_His_MassUpdate", glbUserID)
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
'Updateble = True
'Ticket #7500 - Town of Ajax
'tkt310423 Jerry said remove serial#control for Add_Attendance security
'If glbCompSerial = "S/N - 2173W" Then
    Updateble = gSec_Upd_Attendance
'End If
End Property

Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = False
End Property

Private Sub Update_ASL_WHSCC(ESQLQ)

'Hemu - 01/29/2004 Begin - Enhancement Ticket # 5551
Dim rsASL As New ADODB.Recordset
Dim xEmpNbrO, xOutHRS, recCount
Dim lastrec
Dim xHours, xSHIFT, xSuper, xIncID, xSEN, xEMELEA, xINDICATOR
Dim rsJOB As New ADODB.Recordset, rsDup As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsATT As New ADODB.Recordset
Dim rsASL1 As New ADODB.Recordset
Dim rsASL2 As New ADODB.Recordset
Dim rsENT As New ADODB.Recordset
Dim rsEMP As New ADODB.Recordset
Dim xDATE
Dim SQLQ1
Dim SQLQ, WSQLQ
If Not glbWHSCC Then Exit Sub


SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE " & ESQLQ

WSQLQ = " AND AS_DOA>=" & Date_SQL(dlpFrom)
WSQLQ = WSQLQ & " AND AS_DOA<=" & Date_SQL(dlpTo)

If (Len(clpCode(3)) > 0 And clpCode(3) = "ASL") Or (Len(clpCode(3)) = 0) Then

    xEmpNbrO = 0
    
    'Select all employee meeting selection criteria
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Do Until rsEMP.EOF
    
        'Get the first record with outstanding balance
        SQLQ1 = "SELECT * FROM WHSCC_ASL WHERE AS_EMPNBR = " & rsEMP("ED_EMPNBR") & WSQLQ & "ORDER BY AS_DOA DESC"
        rsASL.Open SQLQ1, gdbAdoIhr001, adOpenDynamic, adLockOptimistic

        If Not rsASL.EOF Then
            rsASL.MoveFirst

            xEmpNbrO = rsASL("AS_EMPNBR")
        
            xOutHRS = rsASL("AS_HRSOS")

            'Add Attendance record for the outstanding ASL balance
            xSHIFT = Null
            xSuper = Null
            rsJOB.Open "SELECT JH_DHRS,JH_REPTAU,JH_SHIFT FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & xEmpNbrO, gdbAdoIhr001, adOpenForwardOnly
            If Not rsJOB.EOF Then
                If IsNumeric(rsJOB("JH_DHRS")) Then xHours = rsJOB("JH_DHRS") Else xHours = 0
                xSuper = rsJOB("JH_REPTAU")
                xSHIFT = rsJOB("JH_SHIFT")
            End If
            rsJOB.Close

            rsTB.Open "SELECT * FROM HRTABL WHERE TB_NAME='ADRE' AND TB_KEY='ASL'", gdbAdoIhr001, adOpenForwardOnly
            xIncID = 0
            xSEN = 0
            xEMELEA = 0
            xINDICATOR = 0
            If Not rsTB.EOF Then
                xSEN = rsTB("TB_SEN")
                xEMELEA = rsTB("TB_USR3")
                xINDICATOR = rsTB("TB_INDICATOR")
            End If
            rsTB.Close

            xDATE = DateAdd("d", 1, CVDate(dlpTo))

            SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=0"
            rsATT.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            rsATT.AddNew
            rsATT("AD_EMPNBR") = xEmpNbrO
            rsATT("AD_COMPNO") = "001"
            rsATT("AD_DOA") = xDATE
            rsATT("AD_REASON") = "ASL"
            rsATT("AD_HRS") = xOutHRS
            rsATT("AD_SHIFT") = xSHIFT
            rsATT("AD_SUPER") = xSuper
            rsATT("AD_INCID") = xIncID
            rsATT("AD_SEN") = xSEN
            rsATT("AD_EMELEA") = xEMELEA
            rsATT("AD_INDICATOR") = xINDICATOR
            rsATT("AD_LDATE") = Date
            rsATT("AD_LTIME") = Time$
            rsATT("AD_LUSER") = glbUserID
            rsATT.Update
            rsATT.Close

            SQLQ = "SELECT ED_EMPNBR,ED_EFDATES,ED_ETDATES FROM HREMP WHERE ED_EMPNBR = " & xEmpNbrO
            rsENT.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            If Not rsENT.EOF Then
                If IsDate(rsENT("ED_EFDATES")) And IsDate(rsENT("ED_ETDATES")) Then
                    SQLQ = "SELECT * FROM WHSCC_ASL WHERE AS_EMPNBR = " & xEmpNbrO & " "
                    SQLQ = SQLQ & "AND AS_DOA = ('" & Format(rsENT("ED_EFDATES"), "mmm dd,yyyy") & "') "
                    SQLQ = SQLQ & "AND AS_CODE = 'TAKE' "
                    rsASL1.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If rsASL1.EOF Then
                        rsASL1.AddNew
                        rsASL1("AS_HRSREP") = 0
                    End If
                    rsASL1("AS_COMPNO") = "001"
                    rsASL1("AS_EMPNBR") = xEmpNbrO
                    rsASL1("AS_DOA") = xDATE
                    rsASL1("AS_CODE") = "TAKE"
                    rsASL1("AS_HRSTAK") = xOutHRS
                    rsASL1("AS_LDATE") = Format(Now, "SHORT DATE")
                    rsASL1("AS_LTIME") = Time$
                    rsASL1("AS_LUSER") = glbUserID
                    rsASL1.Update
                    rsASL1.Close
                End If
            End If

        End If
        rsASL.Close
        rsEMP.MoveNext
    Loop
    rsEMP.Close
        
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE " & ESQLQ
    WSQLQ = " AS_EMPNBR IN (" & SQLQ & ")"
    WSQLQ = WSQLQ & " AND AS_DOA>=" & Date_SQL(dlpFrom)
    WSQLQ = WSQLQ & " AND AS_DOA<=" & Date_SQL(dlpTo)
    
    SQLQ = "DELETE FROM WHSCC_ASL WHERE " & WSQLQ
    gdbAdoIhr001.Execute SQLQ
    MDIMain.panHelp(0).FloodPercent = 95


    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE " & ESQLQ
    WSQLQ = " AS_EMPNBR IN (" & SQLQ & ")"
    WSQLQ = WSQLQ & " AND AS_DOA=" & Date_SQL(xDATE)

    SQLQ = "SELECT * FROM WHSCC_ASL WHERE " & WSQLQ & " ORDER BY AS_EMPNBR,AS_DOA "
    rsASL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do Until rsASL.EOF

        Call ReCalcASL(rsASL("AS_EMPNBR"), "")

        rsASL.MoveNext
    Loop

    rsASL.Close
    MDIMain.panHelp(0).FloodPercent = 98
End If
'Hemu - 01/29/2004 End
   
End Sub

Private Sub Recalculate_OTBANK()
Dim rsEMP As New ADODB.Recordset
Dim rsAttend As New ADODB.Recordset
Dim rsAttendCT As New ADODB.Recordset
Dim SQLQ

'Set ED_OTBANK to zero for the first time otherwise Null will be updated if some Value - Null
SQLQ = "UPDATE HREMP SET ED_OTBANK = 0"
gdbAdoIhr001.Execute SQLQ

SQLQ = "SELECT ED_EMPNBR, ED_OTBANK FROM HREMP"
rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic

If Not rsEMP.EOF Then
    rsEMP.MoveFirst
    
    Do While Not rsEMP.EOF
        
        If glbOracle Then
            SQLQ = "SELECT SUM(AD_HRS) AS OT_SUM FROM HR_ATTENDANCE WHERE substr(AD_REASON,1,2) = 'OT' AND AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " GROUP BY AD_EMPNBR"
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            
            SQLQ = "SELECT SUM(AD_HRS) AS CT_SUM FROM HR_ATTENDANCE WHERE substr(AD_REASON,1,2) = 'CT' AND AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " GROUP BY AD_EMPNBR"
            rsAttendCT.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
        Else
            SQLQ = "SELECT SUM(AD_HRS) AS OT_SUM FROM HR_ATTENDANCE WHERE LEFT(AD_REASON,2) = 'OT' AND AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " GROUP BY AD_EMPNBR"
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            
            SQLQ = "SELECT SUM(AD_HRS) AS CT_SUM FROM HR_ATTENDANCE WHERE LEFT(AD_REASON,2) = 'CT' AND AD_EMPNBR = " & rsEMP("ED_EMPNBR") & " GROUP BY AD_EMPNBR"
            rsAttendCT.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
        End If
        If Not rsAttend.EOF Then
            If Not rsAttendCT.EOF Then
                SQLQ = "UPDATE HREMP SET ED_OTBANK = " & rsAttend("OT_SUM") - rsAttendCT("CT_SUM") & " WHERE ED_EMPNBR = " & rsEMP("ED_EMPNBR")
            Else
                SQLQ = "UPDATE HREMP SET ED_OTBANK = " & rsAttend("OT_SUM") & " WHERE ED_EMPNBR = " & rsEMP("ED_EMPNBR")
            End If
            gdbAdoIhr001.Execute SQLQ
        Else
            If Not rsAttendCT.EOF Then
                SQLQ = "UPDATE HREMP SET ED_OTBANK = " & 0 - rsAttendCT("CT_SUM") & " WHERE ED_EMPNBR = " & rsEMP("ED_EMPNBR")
            Else
                SQLQ = "UPDATE HREMP SET ED_OTBANK = 0 WHERE ED_EMPNBR = " & rsEMP("ED_EMPNBR")
            End If
            gdbAdoIhr001.Execute SQLQ
        End If
        rsAttend.Close
        rsAttendCT.Close
        
        rsEMP.MoveNext
    Loop
End If
rsEMP.Close

End Sub


Private Function getWSQLQRPT() As String
'getWSQLQRPT = glbSeleDeptUn    'Department security removed by Bryan, redundant, this is a list of changes, whether they have security is irrelevant at this point
'If Len(clpDept.Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_DEPTNO} = '" & clpDept.Text & "')"
'If Len(clpDiv.Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_DIV} = '" & clpDiv.Text & "') "
'If Len(clpCode(1).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_LOC} = '" & clpCode(1).Text & "') "
'If Len(clpCode(2).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_ORG} = '" & clpCode(2).Text & "') "
'If Len(clpCode(3).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_EMP} = '" & clpCode(3).Text & "') "
'If Len(clpCode(5).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_REGION} = '" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(5).Text & "') "
'If Len(clpCode(6).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_ADMINBY} = '" & clpCode(6).Text & "') "
'If Len(clpCode(7).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_BENEFIT_GROUP} = '" & clpCode(7).Text & "') "
'If Len(clpPT.Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_PT} = '" & clpPT.Text & "') "
If Len(strEMPLIST) > 0 Then getWSQLQRPT = " ({HREMP.ED_EMPNBR} IN [" & strEMPLIST & "]) "

End Function

Private Function getRecordsCount()
    Dim TD As New ADODB.Recordset
    Dim SQLQ, WSQLQ, ESQLQ
    
    ESQLQ = glbSeleDeptUn
    If Len(clpDept.Text) > 0 Then ESQLQ = ESQLQ & " AND ED_DEPTNO='" & clpDept.Text & "'"
    If Len(clpDiv.Text) > 0 Then ESQLQ = ESQLQ & " AND ED_DIV='" & clpDiv.Text & "'"
    
    If Len(clpCode(0).Text) > 0 Then ESQLQ = ESQLQ & " AND ED_LOC='" & clpCode(0).Text & "'"
    If Len(clpCode(4).Text) > 0 Then ESQLQ = ESQLQ & " AND ED_SECTION='" & clpCode(4).Text & "'"
    
    If Len(clpCode(1).Text) > 0 Then ESQLQ = ESQLQ & " AND ED_ORG='" & clpCode(1).Text & "'"
    If Len(clpPT.Text) > 0 Then ESQLQ = ESQLQ & " AND ED_PT='" & clpPT.Text & "'"
    If Len(clpCode(2).Text) > 0 Then ESQLQ = ESQLQ & " AND ED_EMP='" & clpCode(2).Text & "'"
    If Len(elpEEID.Text) > 0 Then ESQLQ = ESQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
    
    If chkUndoAttArchive.Value = 1 Then
        WSQLQ = " AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & ESQLQ & ")"
        If Len(dlpUndoFrom.Text) > 0 Then WSQLQ = WSQLQ & " AND AH_DOA>=" & Date_SQL(dlpUndoFrom.Text)
        If Len(dlpUndoTo.Text) > 0 Then WSQLQ = WSQLQ & " AND AH_DOA<=" & Date_SQL(dlpUndoTo.Text)
        If Len(clpCode(3).Text) > 0 Then WSQLQ = WSQLQ & " AND AH_REASON IN ('" & Replace(clpCode(3).Text, ",", "','") & "')"
        
        TD.Open "SELECT AH_EMPNBR FROM HR_ATTENDANCE_HISTORY WHERE " & WSQLQ, gdbAdoIhr001, adOpenStatic
    Else
        WSQLQ = " AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & ESQLQ & ")"
        If Len(dlpFrom.Text) > 0 Then WSQLQ = WSQLQ & " AND AD_DOA>=" & Date_SQL(dlpFrom.Text)
        If Len(dlpTo.Text) > 0 Then WSQLQ = WSQLQ & " AND AD_DOA<=" & Date_SQL(dlpTo.Text)
        If Len(clpCode(3).Text) > 0 Then WSQLQ = WSQLQ & " AND AD_REASON IN ('" & Replace(clpCode(3).Text, ",", "','") & "')"
        
        TD.Open "SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE " & WSQLQ, gdbAdoIhr001, adOpenStatic
    End If
    
    getRecordsCount = TD.RecordCount
    
    TD.Close
    Set TD = Nothing
End Function

Private Function Move_AttendanceHistory_BACK()
Dim SQLQ, ESQLQ, WSQLQ As String
Dim RecsAffected As Long
Dim OTsCTsAffected As Long
Dim xFDate
Dim TD As New ADODB.Recordset
Dim rsEListWRK As New ADODB.Recordset ''Ticket #19720 Franks 01/21/2011

On Error GoTo Move_AttendanceHistory_BACK_Err

Move_AttendanceHistory_BACK = False


ESQLQ = glbSeleDeptUn
If Len(clpDept.Text) > 0 Then ESQLQ = ESQLQ & " AND ED_DEPTNO='" & clpDept.Text & "'"
If Len(clpDiv.Text) > 0 Then ESQLQ = ESQLQ & " AND ED_DIV='" & clpDiv.Text & "'"

If Len(clpCode(0).Text) > 0 Then ESQLQ = ESQLQ & " AND ED_LOC='" & clpCode(0).Text & "'"
If Len(clpCode(4).Text) > 0 Then ESQLQ = ESQLQ & " AND ED_SECTION='" & clpCode(4).Text & "'"

If Len(clpCode(1).Text) > 0 Then ESQLQ = ESQLQ & " AND ED_ORG='" & clpCode(1).Text & "'"
If Len(clpPT.Text) > 0 Then ESQLQ = ESQLQ & " AND ED_PT='" & clpPT.Text & "'"
If Len(clpCode(2).Text) > 0 Then ESQLQ = ESQLQ & " AND ED_EMP='" & clpCode(2).Text & "'"
If Len(elpEEID.Text) > 0 Then ESQLQ = ESQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "

WSQLQ = " AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & ESQLQ & ")"
If Len(dlpUndoFrom.Text) > 0 Then WSQLQ = WSQLQ & " AND AH_DOA>=" & Date_SQL(dlpUndoFrom.Text)
If Len(dlpUndoTo.Text) > 0 Then WSQLQ = WSQLQ & " AND AH_DOA<=" & Date_SQL(dlpUndoTo.Text)
If Len(clpCode(3).Text) > 0 Then WSQLQ = WSQLQ & " AND AH_REASON IN ('" & Replace(clpCode(3).Text, ",", "','") & "')"

TD.Open "SELECT AH_EMPNBR FROM HR_ATTENDANCE_HISTORY WHERE " & WSQLQ, gdbAdoIhr001, adOpenStatic
UPDTCNT = TD.RecordCount
If UPDTCNT > 0 Then
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 10
    MDIMain.panHelp(1).Caption = " Please Wait"
    
    ' Run the query
    SQLQ = "INSERT INTO HR_ATTENDANCE (AD_COMPNO,AD_EMPNBR,AD_DOA,AD_REASON_TABL,AD_REASON,AD_HRS,"
    SQLQ = SQLQ & "AD_COMM,AD_CHRGCODE,AD_SHIFT,AD_SUPER,AD_INCID,AD_SEN,AD_JOB,AD_SALARY,AD_SALCD,AD_DHRS,AD_WHRS,"
    SQLQ = SQLQ & "AD_ORG_TABL,AD_ORG,AD_INDICATOR, AD_WCBNBR, AD_FMLA,AD_LDATE,AD_LTIME,AD_LUSER,AD_EMELEA,"
    SQLQ = SQLQ & "AD_POINT,AD_UPLOAD,AD_CALCHRS,AD_DISCIPLINE_TABL,AD_DISCIPLINE,AD_PAYROLL_ID,AD_GLNO,AD_PROJECT_CODE,AD_MACHINE_NUM,AD_MACHINE_HRS,AD_MACHINE_RATE,AD_LEPOINT,AD_PAYENDDATE,"
    SQLQ = SQLQ & "AD_BANKHRS_EXP,AD_CONSUMED,AD_DOCKEY,AD_SOURCE,AD_SOURCE_TABL,AD_REQID,AD_FTIME,AD_TTIME,AD_REGION_TABL,AD_REGION,AD_SALDIST )"

    SQLQ = SQLQ & "SELECT "
    SQLQ = SQLQ & "AH_COMPNO AS AD_COMPNO,"
    SQLQ = SQLQ & "AH_EMPNBR AS AD_EMPNBR,"
    SQLQ = SQLQ & "AH_DOA AS AD_DOA,"
    SQLQ = SQLQ & "AH_REASON_TABL AS AD_REASON_TABL,"
    SQLQ = SQLQ & "AH_REASON AS AD_REASON,"
    SQLQ = SQLQ & "AH_HRS AS AD_HRS,"
    SQLQ = SQLQ & "AH_COMM AS AD_COMM,"
    SQLQ = SQLQ & "AH_CHRGCODE AS AD_CHRGCODE,"
    SQLQ = SQLQ & "AH_SHIFT AS AD_SHIFT,"
    SQLQ = SQLQ & "AH_SUPER AS AD_SUPER,"
    SQLQ = SQLQ & "AH_INCID AS AD_INCID,"
    SQLQ = SQLQ & "AH_SEN AS AD_SEN,"
    SQLQ = SQLQ & "AH_JOB AS AD_JOB,"
    SQLQ = SQLQ & "AH_SALARY AS AD_SALARY,"
    SQLQ = SQLQ & "AH_SALCD AS AD_SALCD,"
    SQLQ = SQLQ & "AH_DHRS AS AD_DHRS,"
    SQLQ = SQLQ & "AH_WHRS AS AD_WHRS,"
    SQLQ = SQLQ & "AH_ORG_TABL AS AD_ORG_TABL,"
    SQLQ = SQLQ & "AH_ORG AS AD_ORG,"
    SQLQ = SQLQ & "AH_INDICATOR AS AD_INDICATOR,"
    SQLQ = SQLQ & "AH_WCBNBR AS AD_WCBNBR,"
    SQLQ = SQLQ & "AH_FMLA AS AD_FMLA,"
    SQLQ = SQLQ & Date_SQL(Format(Now, "mm/dd/yyyy")) & " AS AD_LDATE,"
    SQLQ = SQLQ & "'" & Left(Time$, 8) & "' AS AD_LTIME,"
    SQLQ = SQLQ & "'999999999'AS AD_LUSER,"
    SQLQ = SQLQ & "AH_EMELEA AS AD_EMELEA,"
    SQLQ = SQLQ & "AH_POINT AS AD_POINT,"
    SQLQ = SQLQ & "AH_UPLOAD AS AD_UPLOAD,"
    SQLQ = SQLQ & "AH_CALCHRS AS AD_CALCHRS,"
    SQLQ = SQLQ & "AH_DISCIPLINE_TABL AS AD_DISCIPLINE_TABL,"
    SQLQ = SQLQ & "AH_DISCIPLINE AS AD_DISCIPLINE,"
    SQLQ = SQLQ & "AH_PAYROLL_ID AS AD_PAYROLL_ID,"
    SQLQ = SQLQ & "AH_GLNO AS AD_GLNO,"
    SQLQ = SQLQ & "AH_PROJECT_CODE AS AD_PROJECT_CODE,"
    SQLQ = SQLQ & "AH_MACHINE_NUM AS AD_MACHINE_NUM,"
    SQLQ = SQLQ & "AH_MACHINE_HRS AS AD_MACHINE_HRS,"
    SQLQ = SQLQ & "AH_MACHINE_RATE AS AD_MACHINE_RATE,"
    SQLQ = SQLQ & "AH_LEPOINT AS AD_LEPOINT,"
    SQLQ = SQLQ & "AH_PAYENDDATE AS AD_PAYENDDATE,"
    SQLQ = SQLQ & "AH_BANKHRS_EXP AS AD_BANKHRS_EXP,"
    SQLQ = SQLQ & "AH_CONSUMED AS AD_CONSUMED,"
    SQLQ = SQLQ & "AH_DOCKEY AS AD_DOCKEY,"
    SQLQ = SQLQ & "AH_SOURCE AS AD_SOURCE,"
    SQLQ = SQLQ & "AH_SOURCE_TABL AS AD_SOURCE_TABL,"
    SQLQ = SQLQ & "AH_REQID AS AD_REQID,"
    SQLQ = SQLQ & "AH_FTIME AS AD_FTIME,"
    SQLQ = SQLQ & "AH_TTIME AS AD_TTIME,"
    SQLQ = SQLQ & "AH_REGION_TABL AS AD_REGION_TABL,"
    SQLQ = SQLQ & "AH_REGION AS AD_REGION,"
    SQLQ = SQLQ & "AH_SALDIST AS AD_SALDIST"
    
    SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY "
        
    SQLQ = SQLQ & " WHERE " & WSQLQ
    gdbAdoIhr001.Execute SQLQ, RecsAffected
    
    
    'Get list of employees affected for the Employee List report
    SQLQ = "SELECT distinct AH_EMPNBR FROM HR_ATTENDANCE_HISTORY WHERE " & WSQLQ
    RSEMPLIST.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly
    If glbSQL Then ''Ticket #19720 Franks 01/21/2011
        SQLQ = "SELECT * FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
        rsEListWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    Do While Not RSEMPLIST.EOF
        If Len(strEMPLIST) > 0 Then
            strEMPLIST = strEMPLIST & "," & RSEMPLIST("AH_EMPNBR")
        Else
            strEMPLIST = strEMPLIST & RSEMPLIST("AH_EMPNBR")
        End If
        If glbSQL Then ''Ticket #19720 Franks 01/21/2011
            rsEListWRK.AddNew
            rsEListWRK("TT_COMPNO") = "001"
            rsEListWRK("TT_EMPNBR") = RSEMPLIST("AH_EMPNBR")
            rsEListWRK("TT_WRKEMP") = glbUserID
            rsEListWRK.Update
        End If
        RSEMPLIST.MoveNext
    Loop
    RSEMPLIST.Close
    If glbSQL Then ''Ticket #19720 Franks 01/21/2011
        rsEListWRK.Close
    End If
    
    MDIMain.panHelp(0).FloodPercent = 50
    
    'Delete the records from Attendance History now
    SQLQ = "DELETE FROM HR_ATTENDANCE_HISTORY "
    'SQLQ = SQLQ & " WHERE (HR_ATTENDANCE_HISTORY.AH_DOA>=" & Date_SQL(Format("05/01/2010", "mm/dd/yyyy")) & " AND HR_ATTENDANCE_HISTORY.AH_DOA<=" & Date_SQL(Format("12/31/2010", "mm/dd/yyyy")) & ") "
    SQLQ = SQLQ & " WHERE " & WSQLQ
    gdbAdoIhr001.Execute SQLQ
    
    'Ticket #21417 - If reason code specified or not is not applicable now because multiple reason codes
    'can be specified now. Also the user has already been given warning for this action.
    'Delete the OTBF and BFHR from Attendance now if no Reason Code was specified because this is only applicable
    'for OTs and CTs, and only one reason code selection option is provided. If OT code is specified only then
    'clearing OTBF will not be correct. Same goes with Seniority tagged code.
    'If Len(clpCode(3).Text) = 0 Then
        'Compute OTBF and BFHR date, based on the date range provided.
        xFDate = DateAdd("d", 1, CVDate(dlpUndoTo.Text))

        If chkUndoCompTime Then
            SQLQ = "DELETE FROM HR_ATTENDANCE "
            SQLQ = SQLQ & " WHERE HR_ATTENDANCE.AD_REASON = 'OTBF' "
            SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(xFDate)
            SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & ESQLQ & ")"
            'SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_DOA =" & Date_SQL(Format("08/01/2010", "mm/dd/yyyy"))
            gdbAdoIhr001.Execute SQLQ
        End If
        
        If chkUndoSenHrs Then
            SQLQ = "DELETE FROM HR_ATTENDANCE "
            SQLQ = SQLQ & " WHERE AD_REASON = 'BFHR' "
            SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(xFDate)
            SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & ESQLQ & ")"
            'SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_DOA =" & Date_SQL(Format("01/01/2006", "mm/dd/yyyy"))
            gdbAdoIhr001.Execute SQLQ
        End If
    'End If
    
    'Recalculate all
    MDIMain.panHelp(0).FloodPercent = 80
    Call EntReCalc(ESQLQ)
    MDIMain.panHelp(0).FloodPercent = 85
    Call EntReCalcHr
    MDIMain.panHelp(0).FloodPercent = 90
    
    Call ReCalcOvt("")
    
    'Hemu - Begin -  Town of Ajax - OT Bank
    If glbCompSerial = "S/N - 2173W" Then 'And chkCompTime Then
        Call Recalculate_OTBANK
    End If
    'Hemu
    MDIMain.panHelp(0).FloodPercent = 100
    
End If
    
Move_AttendanceHistory_BACK = True

Screen.MousePointer = DEFAULT

Exit Function

Move_AttendanceHistory_BACK_Err:

If Err = 94 Or Err = 3018 Then
    MsgBox "ERROR " & Err
    Err = 0
    Resume Next
    Screen.MousePointer = DEFAULT
    Exit Function
End If

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Undo Attendance Archive", "HR_ATTENDANCE_HISTORY", "Move Back")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function
