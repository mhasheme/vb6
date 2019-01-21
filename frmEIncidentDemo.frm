VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEIncidentDemo 
   Appearance      =   0  'Flat
   Caption         =   "Demographics"
   ClientHeight    =   6330
   ClientLeft      =   285
   ClientTop       =   1320
   ClientWidth     =   7800
   FillStyle       =   0  'Solid
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6330
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Tag             =   "01-Employee ID in the Division"
   Begin VB.TextBox txtJobDesc 
      Appearance      =   0  'Flat
      DataField       =   "EC_JOBDESC"
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
      Left            =   2120
      MaxLength       =   50
      TabIndex        =   35
      Tag             =   "00-Position Description"
      Top             =   4580
      Width           =   4095
   End
   Begin VB.CommandButton cmdPostion 
      Caption         =   "P&ositions"
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Tag             =   "Postions"
      Top             =   4580
      Width           =   1095
   End
   Begin VB.ComboBox comCountryOfEmp 
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
      Height          =   315
      Left            =   2120
      TabIndex        =   28
      Tag             =   "00-Country of Employment"
      Top             =   4215
      Width           =   1320
   End
   Begin VB.TextBox txtCountryOfEmp 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "EC_WORKCOUNTRY"
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
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   27
      Tag             =   "01-Country"
      Top             =   4245
      Width           =   1515
   End
   Begin VB.TextBox txtIncidentNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      MaxLength       =   4
      TabIndex        =   25
      Tag             =   "01-Type of Incident- Code"
      Top             =   5760
      Visible         =   0   'False
      Width           =   870
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EC_DEPTNO"
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Tag             =   "Department"
      Top             =   960
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EC_DIV"
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   0
      Tag             =   "01-Division"
      Top             =   1320
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   10
      Top             =   5670
      Width           =   7800
      _Version        =   65536
      _ExtentX        =   13758
      _ExtentY        =   1164
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Font3D          =   1
      Alignment       =   1
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Tag             =   "Close and exit this screen"
         Top             =   120
         Width           =   735
      End
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7800
      _Version        =   65536
      _ExtentX        =   13758
      _ExtentY        =   926
      _StockProps     =   15
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BevelInner      =   2
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee#"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   150
         Width           =   945
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "lblEEName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2880
         TabIndex        =   7
         Top             =   135
         Width           =   1185
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1200
         TabIndex        =   8
         Top             =   135
         Width           =   1245
      End
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EC_SECTION"
      Height          =   285
      Index           =   9
      Left            =   1800
      TabIndex        =   1
      Tag             =   "00-Section"
      Top             =   3480
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
      MaxLength       =   8
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EC_ADMINBY"
      Height          =   285
      Index           =   10
      Left            =   1800
      TabIndex        =   4
      Tag             =   "00-Administered By"
      Top             =   3840
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EC_REGION"
      Height          =   285
      Index           =   8
      Left            =   1800
      TabIndex        =   5
      Tag             =   "00-Region"
      Top             =   3120
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EC_LOC"
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   3
      Tag             =   "00-Location - Code"
      Top             =   1680
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EC_ORG"
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   19
      Tag             =   "Union"
      Top             =   2040
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
      MaxLength       =   7
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EC_EMP"
      Height          =   285
      Index           =   6
      Left            =   1800
      TabIndex        =   20
      Tag             =   "00-Enter Status Code"
      Top             =   2400
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EC_PT"
      Height          =   285
      Index           =   7
      Left            =   1800
      TabIndex        =   21
      Tag             =   "00-Category Codes"
      Top             =   2760
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.CodeLookup clpHOME 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   29
      Tag             =   "00-Home Operation Number"
      Top             =   4910
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "HMOP"
      MaxLength       =   12
   End
   Begin INFOHR_Controls.CodeLookup clpHOME 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   30
      Tag             =   "00-Home Line"
      Top             =   5250
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "HMLN"
      MaxLength       =   12
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Left            =   1800
      TabIndex        =   34
      Tag             =   "01-Position code"
      Top             =   4580
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   6
      LookupType      =   5
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Home Line"
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
      Index           =   13
      Left            =   360
      TabIndex        =   32
      Top             =   5250
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Home Operation#"
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
      Index           =   12
      Left            =   360
      TabIndex        =   31
      Top             =   4910
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label lblTopDesp 
      Caption         =   "During the date/time of the incident, the employee worked in the following areas:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   600
      Width           =   7575
   End
   Begin VB.Label lblTitle 
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
      Index           =   5
      Left            =   360
      TabIndex        =   24
      Top             =   2040
      Width           =   420
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   6
      Left            =   360
      TabIndex        =   23
      Top             =   2400
      Width           =   1350
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Index           =   7
      Left            =   360
      TabIndex        =   22
      Top             =   2790
      Width           =   630
   End
   Begin VB.Label lblTitle 
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
      Index           =   4
      Left            =   360
      TabIndex        =   18
      Top             =   1710
      Width           =   615
   End
   Begin VB.Label lblTitle 
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
      Index           =   8
      Left            =   360
      TabIndex        =   17
      Top             =   3120
      Width           =   1410
   End
   Begin VB.Label lblTitle 
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
      Index           =   10
      Left            =   360
      TabIndex        =   16
      Top             =   3860
      Width           =   1125
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   9
      Left            =   360
      TabIndex        =   15
      Top             =   3480
      Width           =   540
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Country of Employment"
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
      Index           =   11
      Left            =   360
      TabIndex        =   13
      Top             =   4200
      Width           =   1620
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   12
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Index           =   2
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   825
   End
End
Attribute VB_Name = "frmEIncidentDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsORG As New ADODB.Recordset
Dim RDept, RGLNum
Dim fglbTERM_Seq
Dim fglbJobID&
Dim JobSnap_PayScale(15) As Double
Dim JobSnap_Salary_Code$
Dim JobSnap_MidPoint!
Dim OHOMELINE, OHOMESHIFT, OHOMEOPRTNBR, OHOMEWRKCNT
Dim xUpdateable
Dim fglbJobList As String

Private Function ChkInput()
Dim x%
Dim Msg$, Response%
ChkInput = False

'If Len(clpCode(1)) = 0 Then
'    MsgBox "Reason for Transfer In is a required field"
'     clpCode(1).SetFocus
'    Exit Function
'Else
'    If clpCode(1).Caption = "Unassigned" Then
'        MsgBox "If code entered it must be known"
'         clpCode(1).SetFocus
'        Exit Function
'    End If
'End If

'If Len(clpCode(2)) = 0 Then
'    MsgBox "Department is a required field"
'     clpCode(2).SetFocus
'    Exit Function
'Else
'    If clpCode(2).Caption = "Unassigned" Then
'        MsgBox "Invalid Department"
'         clpCode(2).SetFocus
'        Exit Function
'    End If
'End If

'For x% = 2 To 10
'If Len(clpCode(x%).Text) > 0 And clpCode(x%).Caption = "Unassigned" Then
'    MsgBox "If code entered it must be known"
'     clpCode(x%).SetFocus
'    Exit Function
'End If
'Next x%

ChkInput = True

End Function
'
'Private Sub clpCode_LostFocus(Index As Integer)
'    'If Index = 2 Then Call Dept_GL
'End Sub

Private Sub clpHOME_Click(Index As Integer)
Call getCodes
End Sub

Private Sub clpHOME_DblClick(Index As Integer)
Call getCodes
End Sub

Private Sub clpJob_Change()
txtJobDesc.Text = GetJobDesc(clpJob.Text)
End Sub

Private Sub clpJob_GotFocus()
txtJobDesc.Text = GetJobDesc(clpJob.Text)
End Sub

Public Sub cmdClose_Click()

    On Error GoTo err_Unload
    
    glbTERM_ID = 0
    glbTran_ID = 0
    glbTran_Seq = 0
    glbOnTop = ""
    
    
    frmEHSINCIDENT.txtDemo(2) = clpCode(2).Text
    frmEHSINCIDENT.txtDemo(3) = clpCode(3).Text
    frmEHSINCIDENT.txtDemo(4) = clpCode(4).Text
    frmEHSINCIDENT.txtDemo(5) = clpCode(5).Text
    frmEHSINCIDENT.txtDemo(6) = clpCode(6).Text
    frmEHSINCIDENT.txtDemo(7) = clpCode(7).Text
    If glbLinamar Then
        If Mid(frmEHSINCIDENT.txtDemo(8), 4) <> clpCode(8).Text Then
            frmEHSINCIDENT.txtDemo(8) = clpCode(3).Text & clpCode(8).Text
        End If
    Else
        frmEHSINCIDENT.txtDemo(8) = clpCode(8).Text
    End If
    
    If glbLinamar Then
        If Mid(frmEHSINCIDENT.txtDemo(9), 4) <> clpCode(9).Text Then
            frmEHSINCIDENT.txtDemo(9) = clpCode(3).Text & clpCode(9).Text
        End If
    Else
        frmEHSINCIDENT.txtDemo(9) = clpCode(9).Text
    End If
    
    frmEHSINCIDENT.txtDemo(10) = clpCode(10).Text
    frmEHSINCIDENT.txtDemo(11) = txtCountryOfEmp
    If glbLinamar Then
        If Len(clpHOME(1).Text) > 0 Then
            frmEHSINCIDENT.txtDemo(12) = clpCode(3).Text & clpHOME(1)
        End If
        If Len(clpHOME(2).Text) > 0 Then
            frmEHSINCIDENT.txtDemo(13) = clpCode(3).Text & clpHOME(2)
        End If
    End If
    frmEHSINCIDENT.txtDemo(14) = txtJobDesc.Text
    Unload Me
    
    Exit Sub
    
err_Unload:
    Unload Me
    Resume Next
    Unload Me

End Sub

Private Function EERetrieve()
Dim SQLQ As String
EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS

'If glbtermopen Then
'    SQLQ = "SELECT * FROM Term_HR_OCC_HEALTH_SAFETY "
'    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
'Else
'    SQLQ = "SELECT *  FROM HR_OCC_HEALTH_SAFETY "
'    SQLQ = SQLQ & " WHERE EC_EMPNBR = " & glbLEE_ID
'End If
'SQLQ = SQLQ & " AND Ec_Case = '" & txtIncidentNo & "'"
'
'If rsORG.State <> 0 Then rsORG.Close
'
'rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic


EERetrieve = True

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REHIRE", "Term_HRTRMEMP", "SELECT")

Resume Next

Exit Function
'
End Function

Private Sub cmdPostion_Click()
Dim OJOB As String, OJobD As String

'OJOB = clpJob.Text
OJobD = txtJobDesc.Text

Load frmJOBS
frmJOBS.Show 1

'If Len(glbJob) < 1 Then
If Len(glbPos) < 1 Then
    'clpJob.Text = OJOB
    txtJobDesc.Text = OJobD
Else
    'clpJob.Text = glbPos
    txtJobDesc.Text = glbPosDesc
End If
End Sub

Private Sub comCountryOfEmp_Click()
    txtCountryOfEmp = comCountryOfEmp.Text
End Sub

Private Sub Form_Activate()
    glbOnTop = "frmEIncidentDemo"
    Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "frmEIncidentDemo"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim rsTerm As New ADODB.Recordset
Dim x%, SQLQ

glbOnTop = "frmEIncidentDemo"

Call setCaption(lblTitle(1))
Call setCaption(lblTitle(2))
Call setCaption(lblTitle(3))
Call setCaption(lblTitle(4))
Call setCaption(lblTitle(5))
Call setCaption(lblTitle(6))
Call setCaption(lblTitle(7))
Call setCaption(lblTitle(8))
Call setCaption(lblTitle(9))
Call setCaption(lblTitle(10))
Call setCaption(lblTitle(11))
If glbLinamar Then
    lblTitle(12).Visible = True
    lblTitle(13).Visible = True
    clpHOME(1).Visible = True
    clpHOME(2).Visible = True
    
    'Ticket #14573
    lblTopDesp.Caption = "During the date/time of the incident, the injury/incident occurred in the following areas:"
End If
Screen.MousePointer = HOURGLASS

If glbLinHS Then
    If Len(glbDivDesc) > 0 Then
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv
    lblTitle(0).Caption = lStr("Division")
Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
        Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    lblEENum.Caption = ShowEmpnbr(glbLEE_ID)
    Call CR_JobHis_Snap
    'clpJob.seleEMPCode = fglbJobList
End If

If EERetrieve() = False Then Exit Sub

'clpDIV = rsORG("TL_NEWDIV")

'fglbTERM_Seq = rsORG("TL_TERM_SEQ")

MDIMain.panHelp(1).Caption = " "
Call addItems

Call INI_Controls(Me)
If Len(txtCountryOfEmp.Text) > 0 Then comCountryOfEmp = txtCountryOfEmp
If Len(fglbJobList) > 0 Then
    clpJob.seleEMPCode = fglbJobList
End If
If glbLinamar Then
    If glbLinHS Then
        Call LinScreenHSDiv
    Else
        Call LinScreenNormal
    End If
End If

End Sub

Private Function RollBack()
On Error GoTo rr
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
rr:
End Function

Private Sub Form_Unload(Cancel As Integer)
    glbOnTop = ""
End Sub
Private Sub UPDMOD()
    'Dim x%
    'For x% = 0 To 2
    '    dlpDate(x%).Enabled = False
    'Next
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
 ChangeAction = OPENING
End Property
Public Property Let ChangeAction(vData As UpdateStateEnum)

End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateTransEmp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Terminations
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property
Public Property Get Updateble() As Boolean
Updateble = xUpdateable
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

UpdateState = OPENING
TF = True
Call set_Buttons(UpdateState)
If Not UpdateRight Then
    TF = False
    Call UPDMOD
End If

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

Private Sub addItems()
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

comCountryOfEmp.ListIndex = 0        '

End Sub

Private Sub txtCountryOfEmp_Change()
    Me.comCountryOfEmp = txtCountryOfEmp
End Sub

Private Sub CR_JobHis_Snap()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String
Dim dynaJobHIS As New ADODB.Recordset
On Error GoTo JobHis_Err
fglbJobList = ""
Screen.MousePointer = HOURGLASS
If glbtermopen Then
    SQLQ = "Select * from Term_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY JH_CURRENT " & IIf(glbSQL, "DESC", "") & ",JH_SDATE DESC"

    If dynaJobHIS.State <> 0 Then dynaJobHIS.Close
    dynaJobHIS.Open SQLQ, gdbAdoIhr001X, adOpenStatic
Else
    SQLQ = "Select * from HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE JH_EMPNBR=" & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY JH_CURRENT " & IIf(glbSQL, "DESC", "") & ",JH_SDATE DESC"

    If dynaJobHIS.State <> 0 Then dynaJobHIS.Close
    dynaJobHIS.Open SQLQ, gdbAdoIhr001, adOpenStatic
End If
If Not dynaJobHIS.EOF Then
    Do Until dynaJobHIS.EOF
        If Not IsNull(dynaJobHIS!JH_JOB) Then
            fglbJobList = fglbJobList & dynaJobHIS!JH_JOB & ","
        End If
        dynaJobHIS.MoveNext
    Loop
    If Right(fglbJobList, 1) = "," Then
        fglbJobList = Left(fglbJobList, Len(fglbJobList) - 1)
    End If
    dynaJobHIS.MoveFirst
End If
Screen.MousePointer = DEFAULT

Exit Sub

JobHis_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Hours per Week", "HR_JOB_History", "SELECT")
Screen.MousePointer = DEFAULT
Resume Next

End Sub

Function GetJobDesc(xCode)
Dim SQLQ As String
Dim xDesc As String
Dim dynaJobHIS As New ADODB.Recordset
    xDesc = ""
    If Len(xCode) > 0 Then
        SQLQ = "SELECT JB_CODE,JB_DESCR FROM HRJOB WHERE JB_CODE = '" & xCode & "' "
        If dynaJobHIS.State <> 0 Then dynaJobHIS.Close
        dynaJobHIS.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not dynaJobHIS.EOF Then
            xDesc = dynaJobHIS("JB_DESCR")
        End If
        dynaJobHIS.Close
    End If
    GetJobDesc = xDesc
End Function

Private Sub LinScreenNormal()
Dim xTmpTop As Integer
Dim xTmpTop2 As Integer
Dim xTmpTop3 As Integer
Dim xTmpTop4 As Integer
Dim xTmpTop5 As Integer
xTmpTop4 = lblTitle(10).Top
xTmpTop5 = lblTitle(11).Top
lblTitle(11).Top = lblTitle(13).Top 'Country
comCountryOfEmp.Top = lblTitle(13).Top
txtCountryOfEmp.Top = lblTitle(13).Top
lblTitle(10).Top = lblTitle(12).Top 'AdminBy
clpCode(10).Top = lblTitle(12).Top
xTmpTop3 = lblTitle(5).Top
lblTitle(5).Top = txtJobDesc.Top 'Union - Operation Sector
clpCode(5).Top = txtJobDesc.Top
xTmpTop = lblTitle(2).Top
lblTitle(2).Top = lblTitle(3).Top
clpCode(2).Top = clpCode(3).Top
lblTitle(3).Top = xTmpTop
clpCode(3).Top = xTmpTop
xTmpTop2 = lblTitle(8).Top
lblTitle(8).Top = lblTitle(4).Top
clpCode(8).Top = clpCode(4).Top
lblTitle(12).Top = xTmpTop3 'Home Operation#
clpHOME(1).Top = xTmpTop3
lblTitle(13).Top = lblTitle(6).Top '
clpHOME(2).Top = lblTitle(6).Top
xTmpTop = lblTitle(9).Top
lblTitle(9).Top = lblTitle(7).Top
clpCode(9).Top = lblTitle(7).Top 'Operation
cmdPostion.Top = xTmpTop2
clpJob.Top = xTmpTop2
txtJobDesc.Top = xTmpTop2
lblTitle(7).Top = xTmpTop
clpCode(7).Top = xTmpTop
lblTitle(6).Top = xTmpTop4
clpCode(6).Top = xTmpTop4
lblTitle(4).Top = xTmpTop5
clpCode(4).Top = xTmpTop5

End Sub
Private Sub LinScreenHSDiv()
Dim xTmpTop As Integer
Dim xTmpTop2 As Integer
Dim xTmpTop3 As Integer
Dim xTmpTop4 As Integer
Dim xTmpTop5 As Integer
Dim xTmpTop6 As Integer
xTmpTop4 = lblTitle(10).Top
xTmpTop5 = lblTitle(11).Top
xTmpTop6 = lblTitle(7).Top
lblTitle(11).Top = lblTitle(13).Top 'Country
comCountryOfEmp.Top = lblTitle(13).Top
txtCountryOfEmp.Top = lblTitle(13).Top
lblTitle(10).Top = lblTitle(12).Top 'AdminBy
clpCode(10).Top = lblTitle(12).Top
xTmpTop3 = lblTitle(5).Top
lblTitle(5).Top = txtJobDesc.Top 'Union - Operation Sector
clpCode(5).Top = txtJobDesc.Top
xTmpTop = lblTitle(2).Top
lblTitle(2).Top = lblTitle(3).Top
clpCode(2).Top = clpCode(3).Top
lblTitle(3).Top = xTmpTop
clpCode(3).Top = xTmpTop
xTmpTop2 = lblTitle(8).Top
lblTitle(8).Top = lblTitle(4).Top
clpCode(8).Top = clpCode(4).Top
lblTitle(12).Top = xTmpTop3 'Home Operation#
clpHOME(1).Top = xTmpTop3
lblTitle(13).Top = lblTitle(6).Top '
clpHOME(2).Top = lblTitle(6).Top
xTmpTop = lblTitle(9).Top
lblTitle(9).Top = lblTitle(7).Top
clpCode(9).Top = lblTitle(7).Top 'Operation
cmdPostion.Top = xTmpTop2
clpJob.Top = xTmpTop2
txtJobDesc.Top = xTmpTop2
lblTitle(7).Top = xTmpTop
clpCode(7).Top = xTmpTop
lblTitle(6).Top = xTmpTop4
clpCode(6).Top = xTmpTop4
lblTitle(4).Top = xTmpTop5
clpCode(4).Top = xTmpTop5
'--------------
lblTitle(12).Visible = False
clpHOME(1).Visible = False
lblTitle(13).Visible = False
clpHOME(2).Visible = False
cmdPostion.Visible = False
clpJob.Visible = False
txtJobDesc.Visible = False
lblTitle(4).Visible = False
clpCode(4).Visible = False
lblTitle(5).Visible = False
clpCode(5).Visible = False
lblTitle(6).Visible = False
clpCode(6).Visible = False
lblTitle(7).Visible = False
clpCode(7).Visible = False
'--------------
lblTitle(9).Top = xTmpTop3
clpCode(9).Top = xTmpTop3
lblTitle(10).Top = lblTitle(13).Top
clpCode(10).Top = lblTitle(13).Top
lblTitle(11).Top = xTmpTop6 'Country
comCountryOfEmp.Top = xTmpTop6
txtCountryOfEmp.Top = xTmpTop6
Me.Height = 4500
End Sub

Sub getCodes()
If glbLinamar Then
    clpHOME(1).TransDiv = clpCode(3).Text
    clpHOME(2).TransDiv = clpCode(3).Text
End If
End Sub

