VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmUEnterLeave 
   Caption         =   "Mass Update Enter Leave"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   10770
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   2
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   1470
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   810
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
      Left            =   1560
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   480
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
      Index           =   6
      Left            =   1560
      TabIndex        =   52
      Tag             =   "00-Enter Union Code"
      Top             =   1140
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Tag             =   "EDPT-Category"
      Top             =   1800
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   4
      Tag             =   "00-Enter Region Code"
      Top             =   2130
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Tag             =   "10-Enter Employee Number"
      Top             =   2460
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
      Left            =   6000
      TabIndex        =   6
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   480
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   6000
      TabIndex        =   7
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   810
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin Threed.SSPanel Panel3D1 
      Height          =   3300
      Left            =   120
      TabIndex        =   23
      Top             =   3600
      Width           =   10035
      _Version        =   65536
      _ExtentX        =   17701
      _ExtentY        =   5821
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Begin VB.CheckBox chkLeave 
         Caption         =   "Leave"
         Height          =   195
         Left            =   0
         TabIndex        =   44
         Top             =   2040
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Frame frmMulti 
         Caption         =   "Position/Salary Information"
         Height          =   2745
         Left            =   5520
         TabIndex        =   25
         Top             =   120
         Visible         =   0   'False
         Width           =   4275
         Begin VB.CommandButton cmdPostion 
            Caption         =   "P&ositions"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Tag             =   "Postions"
            Top             =   300
            Width           =   975
         End
         Begin VB.TextBox txtDHRS 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1935
            TabIndex        =   29
            Tag             =   "00-Usual working hours per day"
            Top             =   1500
            Width           =   855
         End
         Begin VB.TextBox txtWHRS 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1935
            TabIndex        =   28
            Tag             =   "00- Number of hours in work week"
            Top             =   1800
            Width           =   975
         End
         Begin VB.ComboBox comPayPer 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1935
            TabIndex        =   27
            Tag             =   "Choose annum or hour"
            Top             =   1170
            Width           =   1215
         End
         Begin VB.TextBox txtPayrollID 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1930
            MaxLength       =   15
            TabIndex        =   26
            Tag             =   "00-Payroll ID"
            Top             =   2400
            Width           =   1815
         End
         Begin INFOHR_Controls.CodeLookup clpJob 
            Height          =   285
            Left            =   1620
            TabIndex        =   31
            Tag             =   "01-Job Code"
            Top             =   270
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "n/a"
            MaxLength       =   6
            LookupType      =   5
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   4
            Left            =   1620
            TabIndex        =   32
            Tag             =   "00-Union Code"
            Top             =   570
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDOR"
         End
         Begin MSMask.MaskEdBox medSalary 
            Height          =   285
            Left            =   1935
            TabIndex        =   33
            Tag             =   "00-Usual working Salary"
            Top             =   870
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin INFOHR_Controls.CodeLookup clpGLNo 
            Height          =   285
            Left            =   1620
            TabIndex        =   34
            Tag             =   "00-General Ledger - Code"
            Top             =   2100
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   25
            LookupType      =   3
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Per"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   11
            Left            =   180
            TabIndex        =   43
            Top             =   1230
            Width           =   300
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Salary"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   10
            Left            =   180
            TabIndex        =   42
            Top             =   930
            Width           =   540
         End
         Begin VB.Label lblHrsWeek 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hours/Week"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   180
            TabIndex        =   41
            Top             =   1830
            Width           =   1095
         End
         Begin VB.Label lblHrsDay 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hours/Day"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   180
            TabIndex        =   40
            Top             =   1530
            Width           =   855
         End
         Begin VB.Label lblSalCode 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "H/A"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2790
            TabIndex        =   39
            Top             =   1260
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Union"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   9
            Left            =   180
            TabIndex        =   38
            Top             =   630
            Width           =   660
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   37
            Top             =   330
            Width           =   780
         End
         Begin VB.Label lblPayID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payroll ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   180
            TabIndex        =   36
            Top             =   2400
            Width           =   675
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "G/L #"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   20
            Left            =   180
            TabIndex        =   35
            Top             =   2130
            Width           =   435
         End
      End
      Begin VB.CheckBox chkATPaidHours 
         Caption         =   "Paid Hours in AT"
         Height          =   195
         Left            =   2640
         TabIndex        =   24
         Top             =   2040
         Visible         =   0   'False
         Width           =   2235
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   2580
         TabIndex        =   11
         Tag             =   "41-Employement Status Code"
         Top             =   1200
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDEM"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   7
         Left            =   2580
         TabIndex        =   8
         Tag             =   "41-Termination Code "
         Top             =   0
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "TERM"
      End
      Begin INFOHR_Controls.DateLookup dlpTLAYDate 
         Height          =   285
         Index           =   1
         Left            =   2580
         TabIndex        =   10
         Tag             =   "41-To Date"
         Top             =   810
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.DateLookup dlpTLAYDate 
         Height          =   285
         Index           =   0
         Left            =   2580
         TabIndex        =   9
         Tag             =   "40-From Date"
         Top             =   420
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.CodeLookup clpATTCode 
         Height          =   285
         Left            =   2580
         TabIndex        =   12
         Top             =   1590
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ADRE"
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Temporary Lay Off Reason"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   50
         Top             =   -30
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date                        From"
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
         Index           =   1
         Left            =   0
         TabIndex        =   49
         Tag             =   "41-Date Terminated"
         Top             =   420
         Width           =   2265
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "                               To"
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
         Index           =   2
         Left            =   0
         TabIndex        =   48
         Tag             =   "41-Date Terminated"
         Top             =   810
         Width           =   2100
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "New Employment Status"
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
         Index           =   0
         Left            =   0
         TabIndex        =   47
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lblWeeks 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0 Week"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4320
         TabIndex        =   46
         Top             =   840
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Attendance Code"
         Height          =   345
         Left            =   0
         TabIndex        =   45
         Top             =   1620
         Width           =   1995
      End
   End
   Begin VB.Label lblEnterLeave 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Leave"
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
      TabIndex        =   51
      Top             =   3360
      Width           =   1575
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
      TabIndex        =   22
      Top             =   540
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
      TabIndex        =   21
      Top             =   840
      Width           =   825
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   1140
      Width           =   840
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   1470
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
      Left            =   0
      TabIndex        =   18
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2160
      Width           =   510
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
      TabIndex        =   16
      Top             =   2520
      Width           =   1290
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
      TabIndex        =   15
      Top             =   1830
      Width           =   630
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5160
      TabIndex        =   14
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5160
      TabIndex        =   13
      Top             =   840
      Width           =   540
   End
End
Attribute VB_Name = "frmUEnterLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fglbAdd%
Dim fglbDelete%
Dim fglbModify%
Dim fglbSDate As Variant
Dim fglbESQLQ, fglbWSQLQ
Dim xASL As String
Dim xDiscipFlag As Boolean, xOccuAmount
Dim yDiscipFlag As Boolean
Dim fglbRetry, xmedHours, xAnother
Dim SavEML, SavVac, SavSick, AddChg, cntSick, savIncid, SavOutE, SavOutV, SavOutS, SaveHours
Dim Fdate, Tdate, fdateS, tdateS
Dim strEMPLIST
Dim RSEMPLIST As New ADODB.Recordset
Dim xKey
Dim locEEID, locEEName
Dim fglbFollowID
Dim fglbNew
Dim OLDEMP
Dim oFDate, OTDate
Dim fdFdate, fdTdate
Dim fglbJobList As String
Dim AskWeekend, SkipWeekend
Dim xUptEmpAmt
Dim xEJob, xEOrg, xESalary, xESalCD, xEHrsDay, xEHrsWk, xEHrsPay, xEGL, xEPayID

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMUENTERLEAVE"
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Screen.MousePointer = HOURGLASS

glbOnTop = "FRMUENTERLEAVE"

Call setRptCaption(Me)

If glbLinamar Then
    fdFdate = "ED_USRDAT1"
    fdTdate = "ED_UNION"
    Me.Caption = "Temporary Lay Off"
    lblTitle(1) = "Temporary Lay Off Reason"
Else
    fdFdate = "ED_SFDATE"
    fdTdate = "ED_STDATE"
    'lblTitle(1) = "Enter Leave Reason"
    'lblTitle(1).Visible = False
    'clpCode(1).Visible = False
    'clpCode(1).ShowDescription = False
End If

'If glbMulti Then textMulti.Visible = True
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
Set frmUEnterLeave = Nothing
End Sub

Public Sub cmdModify_Click()
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String
Dim Title$, Msg$, DgDef As Variant, Response%
Dim xDays, X, xDATE, xWeekDay

On Error GoTo Mod_Err

strEMPLIST = ""
fglbDelete% = False
fglbAdd% = False
fglbModify% = True
If Not chkMUEnterLeave() Then Exit Sub

Title$ = "Mass Update Enter Leave"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to update all Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

'exclude Saturday/Sunday - Begin
If Len(dlpTLAYDate(1).Text) = 0 Then
    xDays = 0
Else
    xDays = DateDiff("d", dlpTLAYDate(0).Text, dlpTLAYDate(1).Text)
End If
AskWeekend = True
SkipWeekend = False
xDATE = dlpTLAYDate(0).Text
For X = 0 To xDays
   xWeekDay = Weekday(xDATE)
   If xWeekDay = 7 Or xWeekDay = 1 Then
        If AskWeekend Then
            Msg$ = "Do you want exclude Saturday/Sunday for Attendance Records?"
            If MsgBox(Msg$, 36) = 6 Then
                SkipWeekend = True
            End If
        End If
        AskWeekend = False
    End If
    xDATE = DateAdd("d", 1, xDATE)
Next
'exclude Saturday/Sunday - End

If Not modUpdRecs() Then Exit Sub

If xUptEmpAmt > 0 Then
    MsgBox Str(xUptEmpAmt) & " Records Updated Successfully"
Else
    MsgBox "0 Records Updated"
End If

Screen.MousePointer = DEFAULT

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

Private Sub getWSQLQ()
fglbESQLQ = glbSeleDeptUn
If Len(clpDept.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DEPTNO = '" & clpDept.Text & "'"
If Len(clpDiv.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & clpDiv.Text & "' "
If Len(clpCode(6).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & clpCode(2).Text & "' "
If Len(clpCode(3).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMP = '" & clpCode(3).Text & "' "
If Len(clpCode(0).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & clpCode(0).Text & "' "
If Len(clpCode(1).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & clpCode(1).Text & "' "

If Len(clpPT.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_PT = '" & clpPT.Text & "' "
'If glbLinamar Then
'    If Len(clpCode(5).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND (ED_REGION = '" & clpDiv.Text & clpCode(5).Text & "' or  ED_REGION= 'ALL" & clpCode(5).Text & "')"
'End If
If Len(clpCode(5).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_REGION = '" & clpCode(5).Text & "' "
If Len(elpEEID.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
End Sub

Private Function modUpdRecs()

Dim SQLQ As String
Dim rsempt As New ADODB.Recordset
Dim rsCurSal As New ADODB.Recordset
Dim rsAttD As New ADODB.Recordset
Dim fblDoul As Double
Dim xTotalRecs, I
modUpdRecs = False
On Error GoTo modUpdRecs2_Err

Screen.MousePointer = HOURGLASS

Call getWSQLQ

SQLQ = "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME,ED_SECTION,ED_ORG,ED_GLNO,ED_PAYROLL_ID FROM HREMP WHERE (1=1) "
SQLQ = SQLQ & "AND " & fglbESQLQ
rsempt.Open SQLQ, gdbAdoIhr001, adOpenStatic

xUptEmpAmt = 0

If rsempt.EOF Then
    'MsgBox "No employee found in this Selection Criteria."
    GoTo End_Rec
End If

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(2).Caption = ""
'MDIMain.panHelp(0).FloodPercent = 10
MDIMain.panHelp(1).Caption = " Please Wait"
If Not rsempt.EOF Then
    I = 0
    xTotalRecs = rsempt.RecordCount
End If
Do While Not rsempt.EOF
    MDIMain.panHelp(0).FloodPercent = (I / xTotalRecs) * 100
    I = I + 1
    xUptEmpAmt = xUptEmpAmt + 1
    DoEvents
    locEEID = rsempt("ED_EMPNBR")
    locEEName = rsempt("ED_FNAME") & " " & rsempt("ED_SURNAME")
    xEOrg = rsempt("ED_ORG")
    xEGL = rsempt("ED_GLNO")
    xEPayID = rsempt("ED_PAYROLL_ID")
    Call GetEmpJobSalInfo(locEEID)
    
    Call UptEmpEnterLeave
    rsempt.MoveNext
Loop
rsempt.Close

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " Update Completed"
MDIMain.panHelp(2).Caption = ""

End_Rec:


modUpdRecs = True
glbflgFU = False


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

Private Sub GetEmpJobSalInfo(xEmpNo)
Dim TE As New ADODB.Recordset
Dim SQLQ
    xESalary = 0: xESalCD = ""
    SQLQ = "SELECT SH_SALCD,SH_SALARY FROM HR_SALARY_HISTORY WHERE SH_EMPNBR=" & locEEID & " AND SH_CURRENT<>0 "
    TE.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not TE.EOF Then
        xESalCD = TE("SH_SALCD")
        xESalary = TE("SH_SALARY")
    End If
    TE.Close
    
    xEJob = "": xEHrsDay = 0: xEHrsWk = 0: xEHrsPay = 0
    SQLQ = "SELECT JH_ORG,JH_DHRS,JH_WHRS,JH_PAYROLL_ID,JH_GLNO,JH_JOB,JH_PHRS FROM HR_JOB_HISTORY WHERE JH_EMPNBR=" & locEEID & " AND JH_CURRENT<>0 "
    TE.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not TE.EOF Then
        xEJob = TE("JH_JOB")
        xEHrsDay = TE("JH_DHRS")
        xEHrsWk = TE("JH_WHRS")
        xEHrsPay = TE("JH_PHRS")
    End If
    TE.Close
    
End Sub

Private Sub UptEmpEnterLeave()
Dim SQLQ
Dim xEMP

    oFDate = dlpTLAYDate(0).Text
    OTDate = dlpTLAYDate(1).Text

    If Not AUDITTERM() Then MsgBox "ERROR - AUDIT FILE"
    Call updAttendance
    Call updFollow
    Call updStatus
    'If Not glbLambton Then
        If glbCompSerial = "S/N - 2351W" Then 'Ticket #17447 - Burlington Technologies only
            Call updPositionHis
        End If
    'End If
    
    If Len(clpCode(2).Text) > 0 Then xEMP = clpCode(2).Text Else xEMP = "*"
    If Not EmpHisCalc(1, locEEID, "", "", xEMP, "", "", "", "", Date) Then MsgBox "EMPHIS Error"
    
    SQLQ = "UPDATE HREMP  SET "
    SQLQ = SQLQ & fdFdate & "=" & Date_SQL(dlpTLAYDate(0).Text) & ", "
    SQLQ = SQLQ & fdTdate & "=" & Date_SQL(dlpTLAYDate(1).Text) & ", "
    SQLQ = SQLQ & " ED_EMP='" & clpCode(2).Text & "' "
    SQLQ = SQLQ & " WHERE ED_EMPNBR=" & locEEID
    
    gdbAdoIhr001.Execute SQLQ
End Sub

Private Function updPositionHis()
Dim SQLQ As String
    If IsDate(dlpTLAYDate(1).Text) Then
        SQLQ = "UPDATE HR_JOB_HISTORY SET JH_ENDDATE = " & Date_SQL(dlpTLAYDate(1).Text) & " "
        SQLQ = SQLQ & "WHERE NOT (JH_CURRENT = 0) AND JH_EMPNBR = " & glbLEE_ID
        gdbAdoIhr001.Execute SQLQ
    End If
    SQLQ = "UPDATE HR_JOB_HISTORY SET JH_ENDREAS = '" & clpCode(2).Text & "' "
    SQLQ = SQLQ & "WHERE NOT (JH_CURRENT = 0) AND JH_EMPNBR = " & glbLEE_ID
    gdbAdoIhr001.Execute SQLQ
        
End Function

Private Sub updStatus()   'Laura on 11/2/97
Dim SQLQ As String
Dim Msg As String
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
On Error GoTo CrFollow_Err
rsTA.Open "SELECT ED_EMP FROM HREMP WHERE ED_EMPNBR=" & locEEID, gdbAdoIhr001, adOpenKeyset
If rsTA.EOF Then Exit Sub
SQLQ = "SELECT * FROM HRSTATUS "
If IsDate(oFDate) Or IsDate(OTDate) Then
    SQLQ = SQLQ & " WHERE SC_REASON IN ('TLAY', 'LOA') AND SC_EMPNBR=" & locEEID
    If IsDate(oFDate) Then SQLQ = SQLQ & " AND SC_FDATE=" & Date_SQL(oFDate)
    If IsDate(OTDate) Then SQLQ = SQLQ & " AND SC_TDATE=" & Date_SQL(OTDate)
    SQLQ = SQLQ & " AND SC_TYPE='HR'"
End If
rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not rsTB.EOF And (IsDate(oFDate)) Then
    rsTB("SC_TYPE") = Null
    rsTB.Update
End If
rsTB.AddNew
rsTB("SC_COMPNO") = "001"
rsTB("SC_EMPNBR") = locEEID
rsTB("SC_FDATE") = dlpTLAYDate(0).Text
rsTB("SC_TDATE") = dlpTLAYDate(1).Text
rsTB("SC_EMP_TABL") = "EDEM"
rsTB("SC_OLDEMP") = rsTA!ED_EMP
rsTB("SC_NEWEMP") = clpCode(2).Text
rsTB("SC_REASON_TABL") = "SCRE"
If glbLinamar Then
    rsTB("SC_REASON") = clpCode(1).Text
Else
    rsTB("SC_REASON") = "LOA"
End If
rsTB("SC_FOLLOWID") = fglbFollowID

If Len(clpATTCode.Text) > 0 Then rsTB("SC_ATTREASON") = clpATTCode.Text

rsTB("SC_JOB") = ReadJob
rsTB("SC_TYPE") = "HR"
rsTB("SC_LDATE") = Date
rsTB("SC_LTIME") = Time$
rsTB("SC_LUSER") = glbUserID
rsTB.Update

rsTB.Close

 
Exit Sub

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Sub

Private Sub updFollow()   'Laura on 11/2/97
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
On Error GoTo CrFollow_Err
SQLQ = "SELECT * FROM HR_FOLLOW_UP "
If IsDate(OTDate) Then
    SQLQ = SQLQ & " WHERE EF_COMPLETED=0 AND EF_EMPNBR=" & locEEID
    If glbLinamar Then
        SQLQ = SQLQ & " AND EF_FREAS='TLAY' "
    Else
        SQLQ = SQLQ & " AND EF_FREAS='LOA' "
    End If
    SQLQ = SQLQ & " AND EF_FDATE=" & Date_SQL(OTDate)
End If

rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not IsDate(OTDate) Or rsTB.EOF Then
    rsTB.AddNew
    Msg = "A Follow Up Record was created!"
Else
    Msg = "A Follow Up Record was updated!"
End If
rsTB("EF_COMPNO") = "001"
rsTB("EF_EMPNBR") = locEEID
rsTB("EF_FDATE") = CVDate(dlpTLAYDate(1).Text)
rsTB("EF_FREAS_TABL") = "FURE"
'Ticket #24257 - Do not update Admin By for them only
If glbCompSerial <> "S/N - 2262W" Then
    rsTB("EF_ADMINBY_TABL") = "EDAB"
    rsTB("EF_ADMINBY") = GetEmpData(locEEID, "ED_ADMINBY", Null)
End If
If glbLinamar Then
    rsTB("EF_FREAS") = "TLAY"
Else
    rsTB("EF_FREAS") = "LOA"
    Dim rsTT As New ADODB.Recordset
    rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='LOA'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If rsTT.EOF Then
        rsTT.AddNew
        rsTT("TB_COMPNO") = "001"
        rsTT("TB_NAME") = "FURE"
        rsTT("TB_KEY") = "LOA"
        rsTT("TB_DESC") = "Leave of Absence Review"
        rsTT("TB_LUSER") = glbUserID
        rsTT("TB_LDATE") = Date
        rsTT("TB_LTIME") = Time$
        rsTT.Update
    End If
    rsTT.Close
    
    'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
    'follow up record
    Call Grant_FollowUpCode_Security(glbUserID, "LOA", "Leave of Absence Review")
    
End If
rsTB("EF_COMMENTS") = locEEName & " was " & IIf(glbLinamar, "temporarily laid off on ", "on leave from ") & Format(dlpTLAYDate(0).Text, "mmmm dd, yyyy")
rsTB("EF_LDATE") = Date
rsTB("EF_LTIME") = Time$
rsTB("EF_LUSER") = glbUserID
rsTB.Update


fglbFollowID = rsTB("EF_FOLLOWUP_ID")
rsTB.Close
'MsgBox Msg
 
Exit Sub

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Sub

Private Function AUDITTERM()
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String, XSNAME As String, XFNAME As String, xEmpType As String
Dim SQLQ
Dim strFields As String

On Error GoTo AUDIT_ERR

AUDITTERM = False


rsTB.Open "SELECT ED_EMPNBR,ED_PT,ED_DIV,ED_SURNAME,ED_FNAME,ED_EMPTYPE,ED_EMP FROM HREMP WHERE ED_EMPNBR=" & locEEID, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then

    xPT = rsTB("ED_PT")
    If Not IsNull(rsTB("ED_DIV")) Then
        xDiv = rsTB("ED_DIV")
    Else
        xDiv = ""
    End If
    XSNAME = rsTB("ED_SURNAME")
    XFNAME = rsTB("ED_FNAME")
    xEmpType = IIf(IsNull(rsTB("ED_EMPTYPE")), "", rsTB("ED_EMPTYPE"))
    OLDEMP = rsTB("ED_EMP")
Else
    xPT = ""
    xDiv = ""
    XSNAME = ""
    XFNAME = ""
    xEmpType = ""
    OLDEMP = ""
End If

strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_EMPTYPE, AU_EMP, AU_USRDAT1, AU_UNION, AU_SFDATE, "
strFields = strFields & "AU_STDATE, AU_SURNAME, AU_FNAME, AU_DOT, AU_TREAS, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, "
strFields = strFields & "AU_TYPE, AU_PAYROLL_ID "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
xADD = False

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv
rsTA("AU_EMPTYPE") = xEmpType
rsTA("AU_EMP") = clpCode(2).Text

If glbLinamar Then
    rsTA("AU_USRDAT1") = dlpTLAYDate(0).Text
    rsTA("AU_UNION") = dlpTLAYDate(1).Text
Else
    rsTA("AU_SFDATE") = dlpTLAYDate(0).Text
    rsTA("AU_STDATE") = dlpTLAYDate(1).Text
End If
'rsTA("AU_SURNAME") = XSNAME
'rsTA("AU_FNAME") = XFNAME
'rsTA("AU_DOT") = dlpTLAYDate(0)
'rsTA("AU_TREAS") = clpCode(1)

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = locEEID
rsTA("AU_LDATE") = Date
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"

'rsTA("AU_TYPE") = "T"
rsTA("AU_TYPE") = "M"
'If glbSoroc Or glbSyndesis Then
    Dim rsEmp As New ADODB.Recordset
    'Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & locEEID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
    End If
    rsEmp.Close
'End If

rsTA.Update

AUDITTERM = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '29July99 js

End Function
Private Function updAttendance()
Dim SQLQ As String
Dim rsJOB As New ADODB.Recordset, rsDup As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsATT As New ADODB.Recordset
Dim xDays
Dim X, xDATE, xDup
Dim WSQLQ, ESQLQ, Result
Dim TSQLQ
Dim Msg$
Dim xWeekDay
Dim xAskDup
Dim xAddDup
Dim xKey

Dim xHours, xSHIFT, xSuper, xIncID, xSEN, xEMELEA, xIndicator
updAttendance = False
On Error GoTo updAttendance_Err

If Len(clpATTCode.Text) = 0 Then Exit Function
Screen.MousePointer = HOURGLASS

xHours = 0
xSHIFT = Null
xSuper = Null

xHours = xEHrsDay
rsTB.Open "SELECT * FROM HRTABL WHERE TB_NAME='ADRE' AND TB_KEY='" & clpATTCode.Text & "'", gdbAdoIhr001, adOpenForwardOnly
xIncID = 0
xSEN = 0
xEMELEA = 0
xIndicator = 0
If Not rsTB.EOF Then
    xSEN = rsTB("TB_SEN")
    xEMELEA = rsTB("TB_USR3")
    xIndicator = rsTB("TB_INDICATOR")
End If
rsTB.Close

If UCase(clpATTCode.Text) = "OT15" Then xHours = xHours * 1.5
If UCase(clpATTCode.Text) = "OT20" Then xHours = xHours * 2

If Len(dlpTLAYDate(1).Text) = 0 Then
    xDays = 0
Else
    xDays = DateDiff("d", dlpTLAYDate(0).Text, dlpTLAYDate(1).Text)
End If

xDATE = dlpTLAYDate(0).Text
AskWeekend = True
xAskDup = True
xAddDup = True

For X = 0 To xDays
   xWeekDay = Weekday(xDATE)
   If xWeekDay = 7 Or xWeekDay = 1 Then
            If SkipWeekend Then
                xDATE = DateAdd("d", IIf(xWeekDay = 7, 2, 1), xDATE)
                X = X + IIf(xWeekDay = 7, 2, 1)
            End If
    End If
    If Len(dlpTLAYDate(1).Text) > 0 Then
        If CVDate(xDATE) > CVDate(dlpTLAYDate(1).Text) Then Exit For
    Else
        If CVDate(xDATE) > CVDate(dlpTLAYDate(0).Text) Then Exit For
    End If
       
    TSQLQ = "SELECT AD_EMPNBR FROM HR_ATTENDANCE "
    TSQLQ = TSQLQ & " WHERE AD_REASON = '" & clpATTCode.Text & "' "
    TSQLQ = TSQLQ & " AND AD_DOA = " & Date_SQL(xDATE)
    TSQLQ = TSQLQ & " AND AD_EMPNBR =" & locEEID
    rsDup.Open TSQLQ, gdbAdoIhr001, adOpenKeyset
    If Not rsDup.EOF Then
        xDup = True
    Else
        xDup = False
    End If
    rsDup.Close
    
    If Not xDup Then ' Or (xDup And xAddDup) Then
        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=0"
        rsATT.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
        rsATT.AddNew
        rsATT("AD_EMPNBR") = locEEID
        rsATT("AD_COMPNO") = "001"
        rsATT("AD_DOA") = xDATE
        rsATT("AD_REASON") = clpATTCode.Text
        rsATT("AD_HRS") = xHours
        rsATT("AD_SHIFT") = xSHIFT
        rsATT("AD_SUPER") = xSuper
        rsATT("AD_INCID") = xIncID
        rsATT("AD_SEN") = xSEN
        rsATT("AD_EMELEA") = xEMELEA
        rsATT("AD_INDICATOR") = xIndicator
        rsATT("AD_JOB") = xEJob
        rsATT("AD_ORG") = xEOrg
        rsATT("AD_SALARY") = xESalary
        rsATT("AD_DHRS") = xEHrsDay
        rsATT("AD_WHRS") = xEHrsWk
        rsATT("AD_GLNO") = xEGL
        rsATT("AD_PAYROLL_ID") = xEPayID
        rsATT("AD_SALCD") = xESalCD
        rsATT("AD_LDATE") = Date
        rsATT("AD_LTIME") = Time$
        rsATT("AD_LUSER") = glbUserID
        rsATT.Update

      
        rsATT.Close
    End If
    
    xDATE = DateAdd("d", 1, xDATE)
Next

updAttendance = True

Exit Function

updAttendance_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "updAttendance", "Attendance", "Insert")
updAttendance = False
Resume Next

End Function

Private Function chkMUEnterLeave()

Dim SQLQ As String, Msg$, dd&, Response%, X%
Dim DgDef As Variant, Title$, DCurPDate As Variant

chkMUEnterLeave = False

On Error GoTo chkMUEnterLeave_Err

For X% = 0 To 7
If Len(clpCode(X%).Text) > 0 And clpCode(X%).Caption = "Unassigned" Then
    MsgBox "If code entered it must be known"
    clpCode(X%).SetFocus
    Exit Function
End If
Next X%

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

If Len(dlpTLAYDate(0).Text) < 1 Then
    MsgBox "From Date is a required field"
    dlpTLAYDate(0).SetFocus
    Exit Function
End If
If Len(dlpTLAYDate(1).Text) < 1 Then
    MsgBox "To Date is a required field"
    dlpTLAYDate(1).SetFocus
    Exit Function
End If

If Not IsDate(dlpTLAYDate(0).Text) Then
    MsgBox "From Date is not a valid date"
    dlpTLAYDate(0).SetFocus
    Exit Function
End If
If Not IsDate(dlpTLAYDate(0).Text) Then
    MsgBox "From Date is not a valid date"
    dlpTLAYDate(1).SetFocus
    Exit Function
End If
If DateDiff("d", dlpTLAYDate(1), dlpTLAYDate(0)) > 0 Then
    MsgBox "From date must be earlier than To Date"
    dlpTLAYDate(0).SetFocus
    Exit Function
End If
' If statement above should work, but in any case I add If statement under afther test result
If IsDate(dlpTLAYDate(0).Text) And IsDate(dlpTLAYDate(1).Text) Then
    If DaysBetween(dlpTLAYDate(0), dlpTLAYDate(1)) < 0 Then
        MsgBox "From Date can not be prior to To Date"
        dlpTLAYDate(0).SetFocus
        Exit Function
    End If
End If

If Len(clpATTCode.Text) > 0 Then
    If clpATTCode.Caption = "Unassigned" Then
        MsgBox "Attendance code must be valid"
        clpATTCode.SetFocus
        Exit Function
    End If
    If Not clpATTCode.ListChecker Then
        'MsgBox "Attendance code must be valid"
        'clpATTCode.SetFocus
        Exit Function
    End If
End If

If Len(clpCode(2).Text) < 1 Then
    MsgBox "Employment Status is a required field"
    clpCode(2).SetFocus
    Exit Function
Else
    If clpCode(2).Caption = "Unassigned" Then
        MsgBox "Employment Status code must be valid"
        clpCode(2).SetFocus
        Exit Function
    Else
        EMPCode_Desc
        If chkLeave = 0 Then
            MsgBox "This is not code for Leave of Absence"
            clpCode(2).SetFocus
            Exit Function
        End If
    End If
End If

chkMUOK:
chkMUEnterLeave = True

Exit Function

chkMUEnterLeave_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkMUEnterLeave", "HR Attendance", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Function ReadJob()
Dim rsTA As New ADODB.Recordset
Dim IJob
ReadJob = ""

rsTA.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_EMPNBR=" & locEEID, gdbAdoIhr001, adOpenKeyset
If rsTA.EOF Then Exit Function
ReadJob = rsTA("JH_JOB")
rsTA.Close

End Function

Private Sub EMPCode_Desc()
Dim SQLQ As String
Dim rsTA As New ADODB.Recordset
On Error GoTo EMPCode_Desc_Err
chkLeave.Value = 0

If Len(clpCode(2).Text) > 0 Then
    SQLQ = "SELECT TB_USR3 FROM HRTABL WHERE TB_NAME='EDEM' AND TB_KEY = '" & clpCode(2).Text & "'"
    rsTA.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsTA.EOF Then
        chkLeave.Value = IIf(rsTA("TB_USR3"), 1, 0)
    End If
End If

Exit Sub
EMPCode_Desc_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EMP Code Snap", "TABL", "SELECT")
Call RollBack '29July99 js

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
