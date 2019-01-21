VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRJournalEntry 
   Caption         =   "Journal Entry Report"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   9975
   WindowState     =   2  'Maximized
   Begin VB.Frame frmGrouping 
      Height          =   495
      Left            =   2190
      TabIndex        =   37
      Top             =   4590
      Width           =   2925
      Begin Threed.SSOption optGrouping 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Tag             =   "Detailed Report"
         Top             =   150
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "  Detailed"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24.27
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption optGrouping 
         Height          =   255
         Index           =   1
         Left            =   1410
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "Summary Report"
         Top             =   150
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "   Summary"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24.27
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame frmType 
      Height          =   465
      Left            =   2190
      TabIndex        =   36
      Top             =   4050
      Width           =   4005
      Begin Threed.SSOption optType 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Tag             =   "Labour Cost Report"
         Top             =   150
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   " Labour Cost"
         ForeColor       =   16711680
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
      Begin Threed.SSOption optType 
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "Machine Cost Report"
         Top             =   150
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Equipment Cost"
         ForeColor       =   16711680
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
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7080
      MaxLength       =   2
      TabIndex        =   34
      Tag             =   "00-Employee Position Shift"
      Top             =   3330
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Tag             =   "First Level of grouping records"
      Top             =   5505
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Tag             =   "First Level of grouping records"
      Top             =   5820
      Visible         =   0   'False
      Width           =   2325
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1890
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1680
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
      Left            =   1890
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2010
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
      Left            =   1890
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1350
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
      Left            =   1890
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
      Left            =   1890
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   690
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
      Left            =   1890
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
      Index           =   5
      Left            =   1890
      TabIndex        =   9
      Tag             =   "00-Enter Section Code"
      Top             =   3330
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1890
      TabIndex        =   8
      Tag             =   "00-Enter Administered By Code"
      Top             =   3000
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1890
      TabIndex        =   7
      Tag             =   "00-Enter Region Code"
      Top             =   2670
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1890
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2340
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   0
      Top             =   6410
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowWidth     =   480
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   2
      WindowControls  =   -1  'True
      MarginTop       =   720
      MarginBottom    =   720
      PrintFileLinesPerPage=   60
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3600
      TabIndex        =   11
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3690
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1255
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1890
      TabIndex        =   10
      Tag             =   "40-Date from and including this date forward"
      Top             =   3660
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1255
   End
   Begin VB.Label lblShift 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6570
      TabIndex        =   35
      Top             =   3375
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Report"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   33
      Top             =   4200
      Width           =   1290
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      TabIndex        =   32
      Top             =   390
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
      Left            =   270
      TabIndex        =   31
      Top             =   690
      Width           =   825
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   270
      TabIndex        =   30
      Top             =   2340
      Width           =   1290
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      TabIndex        =   29
      Top             =   1350
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
      Left            =   270
      TabIndex        =   28
      Top             =   1695
      Width           =   450
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
      Left            =   210
      TabIndex        =   27
      Top             =   90
      Width           =   1575
   End
   Begin VB.Label lblRepGrp 
      Caption         =   "Report Grouping"
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
      Left            =   210
      TabIndex        =   26
      Top             =   5265
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #1"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   25
      Top             =   5505
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Final Sort"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   24
      Top             =   5820
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      TabIndex        =   23
      Top             =   1020
      Width           =   615
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      TabIndex        =   22
      Top             =   3000
      Width           =   1125
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      TabIndex        =   21
      Top             =   2670
      Width           =   510
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      TabIndex        =   20
      Top             =   3330
      Width           =   540
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      TabIndex        =   19
      Top             =   2010
      Width           =   630
   End
   Begin VB.Label lblFromTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Range"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      TabIndex        =   18
      Top             =   3720
      Width           =   870
   End
End
Attribute VB_Name = "frmRJournalEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jaddy 10/26/99 add a CheckBox and edit lblTitle(5).caption="Display"
Option Explicit

Private Sub cmbHours_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Journal Entry Report Criteria", Me) Then Exit Sub
    Screen.MousePointer = HOURGLASS
     Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
     Call set_PrintState(True)
End If
Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString, , "info:HR"
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdPrint_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
     Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
     Call set_PrintState(True)
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
End If

Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString, , "info:HR"
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ENTITLEMENTS", "VIEW")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdView_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    Select Case intIdx%
    Case 0: strCd$ = "HREMP.ED_LOC"
    Case 1: strCd$ = "HREMP.ED_ORG"
    Case 2: strCd$ = "HREMP.ED_EMP"
    Case 3: strCd$ = "HREMP.ED_REGION"
    Case 4: strCd$ = "HREMP.ED_ADMINBY"
    Case 5: strCd$ = "HREMP.ED_SECTION"
    End Select
        CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
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
    DivCri = "({HREMP.ED_DIV} in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
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
    EECri = "{HREMP.ED_EMPNBR} IN [" & getEmpnbr(elpEEID.Text) & "] "
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

Private Sub Cri_PT()
Dim EECri As String, OneSet%, x%

If Len(clpPT.Text) < 1 Then Exit Sub

EECri = "{HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_FTDates()

Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%, x%
Dim FromDate, ToDate, SQLQ
Dim RsHRPARCO As New ADODB.Recordset

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HR_ATTENDANCE.AD_DOA} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) and "
    
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " ({HR_ATTENDANCE.AD_DOA} <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
End If


Cri_FTDatst:
If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = TempCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Function Cri_SetAll()
Dim x%, xNoFile, xNoWork
Dim WHSCCrpt As Boolean

Cri_SetAll = False

On Error GoTo modSetCriteria_Err

Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_PT
'Call Cri_Shift
Call Cri_EE
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For x% = 0 To 5
    Call Cri_Code(x%)
Next x%

Call Cri_FTDates

If optType(0) Then
    If optGrouping(0) Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2192_JEHOURS_DETAIL.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2192_JEHOURS_SUM.rpt"
    End If
Else
    If optGrouping(0) Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2192_JEMACHINE_DETAIL.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2192_JEMACHINE_SUM.rpt"
    End If
End If

xNoFile = xNoFile + 1
x% = Cri_Sorts()   ' returns number of sections formated

'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

Me.vbxCrystal.Connect = RptODBC_SQL


' window title if appropriate
Me.vbxCrystal.WindowTitle = "Entitlements Report"

Cri_SetAll = True

Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "ENTITLEMENT Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

'Private Sub Cri_Shift()
'Dim EECri As String, OneSet%, x%
'
'If Len(txtShift.Text) < 1 Then Exit Sub
'EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"
'
'If glbiOneWhere Then
'    glbstrSelCri = glbstrSelCri & " AND " & EECri
'Else
'    glbstrSelCri = EECri
'End If
'glbiOneWhere = True
'End Sub

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

Cri_Sorts = 0
If dlpDateRange(0).Text <> "" And dlpDateRange(1).Text <> "" Then
    strSFormat$ = "As of " & dlpDateRange(0).Text & " through " & dlpDateRange(1).Text
    Me.vbxCrystal.Formulas(2) = "Daterange = '" & strSFormat$ & "'"

Else
    strSFormat$ = "No date entered"
    Me.vbxCrystal.Formulas(2) = "Daterange = '" & strSFormat$ & "'"

End If
'grpField$ = getEGroup(comGroup(0).Text)
'If comGroup(0).Text = "Employee Name" Then grpField$ = "{@EEFullName}"
  
'grpField$ = Replace(grpField$, "HRTABL.", "HRTABL1.")
'If comGroup(0) <> "(none)" Then
'    ' <====
'    dscGroup$ = comGroup(0).Text
'    dscGroup$ = "descGroup1 = '" & dscGroup$ & "'"
'    Me.vbxCrystal.Formulas(0) = dscGroup$
'    dscGroup$ = "descName1 = " & grpField$
'    Me.vbxCrystal.Formulas(1) = dscGroup$
'
'    '<====
'
'    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
'    Me.vbxCrystal.GroupCondition(0) = grpCond$
'    grpCond$ = "GROUP" & CStr(2) & ";" & "{@EEFullName}" & ";ANYCHANGE;A"
'    Me.vbxCrystal.GroupCondition(1) = grpCond$
'
'    Me.vbxCrystal.SectionFormat(0) = "GH1;T;T;X;X;X;X;X"
'    Me.vbxCrystal.SectionFormat(1) = "GF1;F,F;F;X;X;X;X"
'Else
'    GrpIdx% = comGroup(1).ListIndex
'    Select Case GrpIdx%
'        Case 0: grpField$ = "{@EEFullName}"
'    End Select
'    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
'    Me.vbxCrystal.GroupCondition(0) = grpCond$
'    Me.vbxCrystal.SectionFormat(0) = "GH1;F;F;X;X;X;X;X"
'    Me.vbxCrystal.SectionFormat(1) = "GF1;F,F;F;X;X;X;X"
'    dscGroup$ = "descGroup1 = ''"
'    Me.vbxCrystal.Formulas(0) = dscGroup$
'    dscGroup$ = "descName1 = ''"
'    Me.vbxCrystal.Formulas(1) = dscGroup$
'End If
''If chkAtt = False Then  'Jaddy 10/26/99 Jaddy 11/5/99
' Me.vbxCrystal.Formulas(2) = "DESCSICK =if not isnull({HREMP.ED_EFDATES}) or not isnull({HREMP.ED_ETDATES}) then totext({HREMP.ED_EFDATES}) +' TO ' + totext({HREMP.ED_ETDATES}) "
' Me.vbxCrystal.Formulas(3) = "DESCVAC = if not isnull({HREMP.ED_EFDATE}) or not isnull({HREMP.ED_ETDATE}) then totext({HREMP.ED_EFDATE}) +' TO ' + totext({HREMP.ED_ETDATE}) "
''End If  'Jaddy 10/26/99 Jaddy 11/5/99

Cri_Sorts = z% ' next section number to format

End Function

Private Function CriCheck()
Dim x%

CriCheck = False

If Not clpDiv.ListChecker Then
'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    'MsgBox lStr("If Division Entered - it must be known"), , "info:HR"
    'clpDiv.SetFocus
    Exit Function
End If

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known", , "info:HR"
    'clpDept.SetFocus
    Exit Function
End If

For x% = 0 To 5
    If Not clpCode(x).ListChecker Then Exit Function
Next x%

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

CriCheck = True
End Function

Private Sub dlpDateRange_LostFocus(Index As Integer)
'dlpDateRange(1).Text = ""
'If IsDate(dlpDateRange(0).Text) Then
'    dlpDateRange(1).Text = DateAdd("d", 13, dlpDateRange(0).Text)
'End If
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()

glbOnTop = "FRMRJOURNALENTRY"

Screen.MousePointer = HOURGLASS

'If Not glbMulti Then
'    lblShift.Visible = True
'    txtShift.Visible = True
'End If
Call setRptCaption(Me)
Call comGrpLoad

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

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
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub

Private Sub comGrpLoad()

'comGroup(0).AddItem lStr("Division")
'comGroup(0).AddItem lStr("Department")
'comGroup(0).AddItem lStr("Location")
'comGroup(0).AddItem lStr("Union")
'comGroup(0).AddItem lStr("Administered By")
'comGroup(0).AddItem "Employee Name"
'comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
'
'If glbLinamar Then ' Frank May 2,2001
'    comGroup(0).AddItem "Employment Type"
'    comGroup(0).AddItem lStr("Region")
'    comGroup(0).AddItem ("Home Line")
'End If
'If Not glbMulti Then comGroup(0).AddItem "Shift"
'comGroup(0).AddItem "(none)"
'comGroup(0).ListIndex = 0
'comGroup(1).AddItem "Employee Name"
'comGroup(1).ListIndex = 0
'comGroup(1).Enabled = False

End Sub

Private Sub txtShift_GotFocus()
Call SetPanHelp(ActiveControl)
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
Printable = True
End Property

Public Sub SET_UP_MODE()
Call set_Buttons
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

