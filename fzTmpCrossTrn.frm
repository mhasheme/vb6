VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRTmpCrossTrain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Temporary/Cross Training Assignment"
   ClientHeight    =   10950
   ClientLeft      =   435
   ClientTop       =   870
   ClientWidth     =   13215
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10950
   ScaleWidth      =   13215
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   10695
      LargeChange     =   300
      Left            =   11760
      Max             =   4000
      SmallChange     =   300
      TabIndex        =   54
      Top             =   120
      Width           =   255
   End
   Begin Threed.SSPanel panWindow 
      Height          =   10695
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   11415
      _Version        =   65536
      _ExtentX        =   20135
      _ExtentY        =   18865
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.PictureBox panDetails 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7215
         Left            =   0
         ScaleHeight     =   7185
         ScaleWidth      =   11385
         TabIndex        =   31
         Top             =   120
         Width           =   11415
         Begin VB.OptionButton optEmpWork 
            Caption         =   "Employee Work History Report"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Tag             =   "Employee Work History Report"
            Top             =   5040
            Value           =   -1  'True
            Width           =   2715
         End
         Begin VB.OptionButton optCrossTrain 
            Caption         =   "Cross-Training By Position Report"
            Height          =   255
            Left            =   3375
            TabIndex        =   19
            Tag             =   "Cross-Training By Position Report"
            Top             =   5040
            Width           =   2925
         End
         Begin VB.CheckBox chkShowMedical 
            Caption         =   "Show Medical Contacts"
            Height          =   285
            Left            =   7320
            TabIndex        =   27
            Top             =   6360
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.CheckBox chkForAudit 
            Caption         =   "For Data Audit"
            Height          =   285
            Left            =   7320
            TabIndex        =   28
            Top             =   6600
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.CheckBox chkWeeklyEmpList 
            Caption         =   "Show Weekly Employee List"
            Height          =   285
            Left            =   7320
            TabIndex        =   26
            Top             =   6120
            Visible         =   0   'False
            Width           =   2445
         End
         Begin VB.ComboBox comGroup 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   7410
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Tag             =   "Final sorting of records"
            Top             =   6620
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.CheckBox chkLastDay 
            Caption         =   "Show Last Day"
            Height          =   285
            Left            =   7320
            TabIndex        =   25
            Top             =   5880
            Visible         =   0   'False
            Width           =   2565
         End
         Begin VB.TextBox txtShift 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2050
            MaxLength       =   4
            TabIndex        =   17
            Tag             =   "00-Employee Position Shift"
            Top             =   4530
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.ComboBox comGroup 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Tag             =   "Final sorting of records"
            Top             =   6340
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.ComboBox comGroup 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Tag             =   "First Level of grouping records"
            Top             =   6000
            Visible         =   0   'False
            Width           =   2325
         End
         Begin INFOHR_Controls.CodeLookup clpJob 
            Height          =   285
            Left            =   1740
            TabIndex        =   7
            Tag             =   "00-Enter Position Code "
            Top             =   2550
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   6
            LookupType      =   5
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   2
            Left            =   1740
            TabIndex        =   4
            Tag             =   "00-Enter Status Code"
            Top             =   1560
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
            Left            =   1740
            TabIndex        =   5
            Tag             =   "EDPT-Category"
            Top             =   1890
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
            Left            =   1740
            TabIndex        =   3
            Tag             =   "00-Enter Union Code"
            Top             =   1230
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
            Left            =   1740
            TabIndex        =   2
            Tag             =   "00-Enter Location Code"
            Top             =   900
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDLC"
         End
         Begin INFOHR_Controls.CodeLookup clpDept 
            Height          =   285
            Left            =   1740
            TabIndex        =   1
            Tag             =   "00-Specific Department Desired"
            Top             =   570
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
            Left            =   1740
            TabIndex        =   0
            Tag             =   "00-Specific Division Desired"
            Top             =   240
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
            Index           =   8
            Left            =   1740
            TabIndex        =   12
            Tag             =   "00-Enter Administered By Code"
            Top             =   3540
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDAB"
            MaxLength       =   10
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   9
            Left            =   1740
            TabIndex        =   13
            Tag             =   "00-Enter Section Code"
            Top             =   3870
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDSE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   7
            Left            =   1740
            TabIndex        =   11
            Tag             =   "00-Enter Region Code"
            Top             =   3210
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDRG"
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   1
            Left            =   3540
            TabIndex        =   10
            Tag             =   "40-Position Start Date upto and including this date forward"
            Top             =   2880
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   0
            Left            =   1740
            TabIndex        =   9
            Tag             =   "40-Position Start Date from and including this date forward"
            Top             =   2880
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.EmployeeLookup elpEEID 
            Height          =   285
            Left            =   1740
            TabIndex        =   6
            Tag             =   "10-Enter Employee Number"
            Top             =   2220
            Width           =   7515
            _ExtentX        =   13256
            _ExtentY        =   503
            ShowUnassigned  =   1
            TextBoxWidth    =   7195
            RefreshDescriptionWhen=   2
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   3
            Left            =   8970
            TabIndex        =   23
            Tag             =   "40-Date upto and including this date forward"
            Top             =   5265
            Visible         =   0   'False
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin INFOHR_Controls.DateLookup dlpDateRange 
            Height          =   285
            Index           =   2
            Left            =   7140
            TabIndex        =   22
            Tag             =   "40-Date from and including this date forward"
            Top             =   5265
            Visible         =   0   'False
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin INFOHR_Controls.CodeLookup clpGrid 
            Height          =   285
            Left            =   8070
            TabIndex        =   8
            Top             =   2550
            Visible         =   0   'False
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "JBGD"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   10
            Left            =   6600
            TabIndex        =   24
            Tag             =   "00-Benefit - Group Code"
            Top             =   5580
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "BGMF"
         End
         Begin INFOHR_Controls.EmployeeLookup elpRept 
            Height          =   285
            Index           =   0
            Left            =   1740
            TabIndex        =   14
            Tag             =   "10-Reporting Authority 1"
            Top             =   4200
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            ShowDescription =   0   'False
            RefreshDescriptionWhen=   2
         End
         Begin INFOHR_Controls.EmployeeLookup elpRept 
            Height          =   285
            Index           =   1
            Left            =   3660
            TabIndex        =   15
            Tag             =   "10-Reporting Authority 2"
            Top             =   4200
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   503
            RefreshDescriptionWhen=   2
         End
         Begin INFOHR_Controls.EmployeeLookup elpRept 
            Height          =   285
            Index           =   2
            Left            =   5580
            TabIndex        =   16
            Tag             =   "10-Reporting Authority 3"
            Top             =   4200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            ShowDescription =   0   'False
            RefreshDescriptionWhen=   2
         End
         Begin VB.Label lblBenGroup 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Benefit Group"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5520
            TabIndex        =   55
            Top             =   5580
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblGrp 
            BackStyle       =   0  'Transparent
            Caption         =   "Work History Sort"
            Height          =   375
            Index           =   1
            Left            =   5520
            TabIndex        =   53
            Top             =   6645
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label lblGrid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Grid Category"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6840
            TabIndex        =   52
            Top             =   2580
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label lblEmplStFrpmTo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Status From / To Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5520
            TabIndex        =   51
            Top             =   5280
            Visible         =   0   'False
            Width           =   1590
         End
         Begin VB.Label FName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
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
            Left            =   6360
            TabIndex        =   50
            Top             =   6960
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.Label lblShift 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Shift"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   4560
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
            TabIndex        =   48
            Top             =   1890
            Width           =   630
         End
         Begin VB.Label lblRep 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reporting Authority:"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   4200
            Width           =   1395
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
            TabIndex        =   46
            Top             =   3870
            Width           =   540
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
            TabIndex        =   45
            Top             =   3540
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
            Left            =   120
            TabIndex        =   44
            Top             =   3210
            Width           =   510
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
            TabIndex        =   43
            Top             =   900
            Width           =   615
         End
         Begin VB.Label lblGrp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Final Sort"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   42
            Top             =   6340
            Visible         =   0   'False
            Width           =   660
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
            Left            =   120
            TabIndex        =   41
            Top             =   6030
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label lblRepGrp 
            BackStyle       =   0  'Transparent
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
            Left            =   0
            TabIndex        =   40
            Top             =   5760
            Visible         =   0   'False
            Width           =   1575
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
            TabIndex        =   39
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label lblFromTo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "From / To Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label lblPosition 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Position Code"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   2550
            Width           =   975
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
            TabIndex        =   36
            Top             =   2220
            Width           =   1290
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
            TabIndex        =   35
            Top             =   1560
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
            TabIndex        =   34
            Top             =   1230
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
            TabIndex        =   33
            Top             =   570
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
            TabIndex        =   32
            Top             =   240
            Width           =   555
         End
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   12360
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
End
Attribute VB_Name = "frmRTmpCrossTrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportSel, SQLQ
Dim rsTabl As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdPrint_Click()
    Dim X%
    
    On Error GoTo PrntErr
    
    If CriCheck() Then
        If Not PrtForm("Temporary/Cross Training Assignment Criteria", Me) Then Exit Sub
        
        Call set_PrintState(False)
        
        X% = Cri_SetAll()
        
        Me.vbxCrystal.Destination = 1
        MDIMain.Timer1.Enabled = False
        Me.vbxCrystal.Action = 1
        vbxCrystal.Reset
        MDIMain.Timer1.Enabled = True
        
        Call set_PrintState(True)
        Screen.MousePointer = DEFAULT
    End If
    
Exit Sub
PrntErr:
    MsgBox "CRW ERROR : " & Chr(10) & "[" & str(Err) & "] : " & Me.vbxCrystal.LastErrorString
    Resume Next
    Screen.MousePointer = DEFAULT
End Sub

Public Sub cmdView_Click()
    Dim X%
    Dim strWHand As String
    On Error GoTo CRW_Err
    
    If CriCheck() Then
        Screen.MousePointer = HOURGLASS
        Call set_PrintState(False)
    
        'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
        'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
        Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
        
        X% = Cri_SetAll()
        
        Me.vbxCrystal.Destination = 0
        MDIMain.Timer1.Enabled = False
        Screen.MousePointer = DEFAULT
        Me.vbxCrystal.Action = 1
        vbxCrystal.Reset
        MDIMain.Timer1.Enabled = True
        
        Call set_PrintState(True)
    End If
Exit Sub
    
CRW_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    MsgBox "CRW ERROR : " & Chr(10) & "[" & str(Err) & "] : " & Me.vbxCrystal.LastErrorString
    Resume Next
    Screen.MousePointer = DEFAULT
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")  'Jaddy jun 16,1999
    comGroup(0).AddItem lStr("Union")
    comGroup(0).AddItem "Employee Name"
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "Position Code"
    comGroup(0).AddItem lStr("Machine #")
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
    
    comGroup(1).AddItem "Employee Name"
    comGroup(1).ListIndex = 0
    
    comGroup(2).AddItem "Descending"
    comGroup(2).AddItem "Ascending"
    comGroup(2).ListIndex = 0
End Sub

Private Sub Cri_Assoc()
    Dim EECri As String
    
    If Len(clpCode(1).Text) <= 0 Then Exit Sub
    
    If glbMulti Then
        EECri = "{qry_Primary_Temp_TRAIN_rpt.JH_ORG} in  ['" & Replace(clpCode(1).Text, ",", "','") & "']"
    Else
        EECri = "{HREMP.ED_ORG} in  ['" & Replace(clpCode(1).Text, ",", "','") & "']"
    End If
    
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
End Sub

Private Sub Cri_Dept()
    Dim countr   As Integer  ' EEList_Snap is definded at form level
    Dim DeptCri As String
    
    DeptCri = ""
    
    Call glbCri_DeptUN(clpDept.Text)
End Sub

Private Sub Cri_Div()
    Dim DivCri As String
    
    If Len(clpDiv.Text) <= 0 Then Exit Sub
    
    DivCri = "({HREMP.ED_DIV} in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
    
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & DivCri
    Else
        glbstrSelCri = DivCri
    End If
End Sub

Private Sub Cri_EE()
    Dim EECri As String
    
    If Len(elpEEID.Text) <= 0 Then Exit Sub
    
    EECri = "{HREMP.ED_EMPNBR} IN [" & getEmpnbr(elpEEID.Text) & "] "
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
End Sub

Private Sub Cri_RepAuth()
    Dim TempCri As String
    Dim EECri As String, LocCri As String
    Dim I, xTemp As Boolean
    
    xTemp = False
    EECri = ""

    If Len(Trim(elpRept(0).Text)) > 0 Then
        EECri = EECri & "{qry_Primary_Temp_TRAIN_rpt.JH_REPTAU} = " & Trim(elpRept(0).Text) & " "
        xTemp = True
    End If
    If Len(Trim(elpRept(1).Text)) > 0 Then
        If xTemp Then
            EECri = EECri & "AND {qry_Primary_Temp_TRAIN_rpt.JH_REPTAU2} = " & Trim(elpRept(1).Text) & " "
        Else
            EECri = EECri & "{qry_Primary_Temp_TRAIN_rpt.JH_REPTAU2} = " & Trim(elpRept(1).Text) & " "
        End If
        xTemp = True
    End If
    If Len(Trim(elpRept(2).Text)) > 0 Then
        If xTemp Then
            EECri = EECri & "AND {qry_Primary_Temp_TRAIN_rpt.JH_REPTAU3} = " & Trim(elpRept(2).Text) & " "
        Else
            EECri = EECri & "{qry_Primary_Temp_TRAIN_rpt.JH_REPTAU3} = " & Trim(elpRept(2).Text) & " "
        End If
        xTemp = True
    End If
    
    If Len(EECri) > 0 Then
        If Len(glbstrSelCri) > 0 Then
          glbstrSelCri = glbstrSelCri & " AND " & EECri
        Else
          glbstrSelCri = EECri
        End If
    End If
End Sub

Private Sub Cri_FTDates()
    Dim TempCri As String
    Dim dtYYY%, dtMM%, dtDD%
    Dim X%
    Dim EECri As String, LocCri As String
    
    If Len(dlpDateRange(0).Text) = 0 And Len(dlpDateRange(1).Text) = 0 Then Exit Sub
    
    If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
        TempCri = "({qry_Primary_Temp_TRAIN_rpt.JH_SDATE} "
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
        TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
        dtYYY% = Year(dlpDateRange(1).Text)
        dtMM% = month(dlpDateRange(1).Text)
        dtDD% = Day(dlpDateRange(1).Text)
        TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    ElseIf Len(dlpDateRange(0).Text) > 0 Then
        TempCri = "({qry_Primary_Temp_TRAIN_rpt.JH_SDATE} "
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
        TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "    'Hemu - 07/02/2003, Included '='
        GoTo Cri_FTDatst
    ElseIf Len(dlpDateRange(1).Text) > 0 Then
        TempCri = "({qry_Primary_Temp_TRAIN_rpt.JH_SDATE} "
        dtYYY% = Year(dlpDateRange(1).Text)
        dtMM% = month(dlpDateRange(1).Text)
        dtDD% = Day(dlpDateRange(1).Text)
        TempCri = TempCri & " <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "    'Hemu - 07/02/2003, Included '='
        GoTo Cri_FTDatst
    End If
    
    For X% = 0 To 1
        If Len(dlpDateRange(0).Text) > 0 Then
            TempCri = "({qry_Primary_Temp_TRAIN_rpt.JH_SDATE}  "
            If X% = 0 Then
                TempCri = TempCri & " >= "
            Else
                TempCri = TempCri & " <= "
            End If
            dtYYY% = Year(dlpDateRange(0).Text)
            dtMM% = month(dlpDateRange(0).Text)
            dtDD% = Day(dlpDateRange(0).Text)
            TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
            GoTo Cri_FTDatst
        End If
    Next X%

Cri_FTDatst:
    If Len(TempCri) > 0 Then
        If Len(glbstrSelCri) > 0 Then
          glbstrSelCri = glbstrSelCri & " AND " & TempCri
        Else
          glbstrSelCri = TempCri
        End If
    End If
End Sub

Private Sub Cri_Position()
    Dim countr   As Integer  ' EEList_Snap is definded at form level
    Dim PosCri As String
    
    If Len(clpJOB.Text) <= 0 Then Exit Sub
        
    PosCri = "({qry_Primary_Temp_TRAIN_rpt.TW_JOB} = '" & clpJOB.Text & "')"
    
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & PosCri
    Else
        glbstrSelCri = PosCri
    End If
End Sub

Private Sub Cri_Grid()
    Dim countr   As Integer  ' EEList_Snap is definded at form level
    Dim GirdCri As String
    
    If Len(clpGrid.Text) <= 0 Then Exit Sub
        
    GirdCri = "({qry_Primary_Temp_TRAIN_rpt.JH_GRID} = '" & clpGrid.Text & "')"
    
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & GirdCri
    Else
        glbstrSelCri = GirdCri
    End If
End Sub

Private Sub Cri_PT()
    Dim EECri As String
    
    If Len(clpPT.Text) < 1 Then Exit Sub
    
    If glbMulti Then
        EECri = "{qry_Primary_Temp_TRAIN_rpt.JH_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"
    Else
        EECri = "{HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"
    End If
    
    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
End Sub

Private Sub Cri_BenefitGroup()
    Dim EECri As String
    
    If Len(clpCode(10).Text) < 1 Then Exit Sub
    
    EECri = "ED_BENEFIT_GROUP = '" & clpCode(10).Text & "' "
    
    If Len(SQLQ) > 1 Then
        SQLQ = SQLQ & " AND " & EECri
    Else
        SQLQ = EECri
    End If
End Sub

Private Function Cri_SetAll()
    Dim X%, strRName$, I
    
    On Error GoTo modSetCriteria_Err
    
    Cri_SetAll = False
    
    Screen.MousePointer = HOURGLASS
    
    glbiOneWhere = False
    glbstrSelCri = ""
    SQLQ = ""

    ' call cri models set both glbiONeWhere and strSelCri
    Call glbCri_DeptUN(clpDept.Text)
    SQLQ = glbstrSelCri
    Call Cri_Div
    Call Cri_Assoc
    Call Cri_Code(0)
    Call Cri_Code(1)
    Call Cri_Code(2)
    Call Cri_PT
    Call Cri_EE
    Call Cri_Position
    'Call Cri_Grid
    Call Cri_FTDates
    Call Cri_Status
    Call Cri_Code(7)
    Call Cri_Code(8)
    Call Cri_Code(9)
    Call Cri_RepAuth
    Call Cri_Shift
    
    'X% = Cri_Sorts()
    
    'glbstrSelCri = IIf(Len(glbstrSelCri) > 0, glbstrSelCri & " AND ", glbstrSelCri) & " {HREMPWRK.TT_WRKEMP}='" & glbUserID & "'"
    If Len(glbstrSelCri) >= 0 Then
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    End If
        
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    End If
    If optEmpWork Then
        'glbstrSelCri = IIf(Len(glbstrSelCri) > 0, glbstrSelCri & " AND ", glbstrSelCri) & " {HREMPWRK.TT_WRKEMP}='" & glbUserID & "'"
        If Len(glbstrSelCri) >= 0 Then
            Me.vbxCrystal.SelectionFormula = glbstrSelCri
        End If
        
        strRName$ = glbIHRREPORTS & "rzWorkHist.rpt"
        Me.vbxCrystal.WindowTitle = "Employee Work History Report"
    End If
    If optCrossTrain Then
        If Len(glbstrSelCri) Then
            glbstrSelCri = Replace(glbstrSelCri, "qry_Primary_Temp_TRAIN_rpt", "HRCRSTRNWRK")
        End If
        glbstrSelCri = IIf(Len(glbstrSelCri) > 0, glbstrSelCri & " AND ", glbstrSelCri) & " {HRCRSTRNWRK.CT_WRKEMP}='" & glbUserID & "'"
        
        'Call procedure to populate the temporary table with rows to display in the report.
        Call Populate_TempTable_with_Employee_PosCourses
        
        If Len(glbstrSelCri) >= 0 Then
            Me.vbxCrystal.SelectionFormula = glbstrSelCri
        End If
        
        'Call procedure to populate temp. table with the records to populate in the report.
        strRName$ = glbIHRREPORTS & "rzCrsTrainPos.rpt"
        Me.vbxCrystal.WindowTitle = "Cross Training By Position Report"
    End If
    
    Me.vbxCrystal.ReportFileName = strRName$
    
    Cri_SetAll = True

    Screen.MousePointer = DEFAULT

Exit Function
modSetCriteria_Err:
    Screen.MousePointer = DEFAULT
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Select Report Criteria", "Temp/Cross Training", "Select")
    Cri_SetAll = False
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Function

Private Function Cri_Sorts()
    Dim grpCond$, grpField$
    Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
    Dim dscGroup$, GrpIdx%
    
    Cri_Sorts = 0
    
    'first set primary grouping
    z% = 0
    X% = 0
    
    grpField$ = getEGroup(comGroup(0).Text)
    Y% = X% + 1
    
    If comGroup(0) = "(none)" Then grpField$ = "{@EFullName}"
    Call setRptLabel(Me, 0)
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup1 = '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(X%) = dscGroup$
    
    grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(X%) = grpCond$
    
    Cri_Sorts = z% ' next section number to format

End Function

Private Sub Cri_Status()
    Dim EECri As String, LocCri As String
    
    If Len(clpCode(2).Text) <= 0 Then Exit Sub
    
    If Len(clpCode(2).Text) > 0 Then
        EECri = "{HREMP.ED_EMP} in ['" & Replace(clpCode(2).Text, ",", "','") & "'] "
    End If
    
    If Len(EECri) >= 1 Then
        If Len(glbstrSelCri) > 1 Then
            glbstrSelCri = glbstrSelCri & " AND " & EECri
        Else
            glbstrSelCri = EECri
        End If
        glbiOneWhere = True
    End If
End Sub

Private Sub Cri_Code(intIdx%)
    Dim CodeCri As String
    Dim strCd$
    
    If Len(clpCode(intIdx%).Text) > 0 Then
        If intIdx% = 0 Then strCd$ = "HREMP.ED_LOC"
        If intIdx% = 7 Then strCd$ = "HREMP.ED_REGION"
        If intIdx% = 8 Then strCd$ = "HREMP.ED_ADMINBY"
        If intIdx% = 9 Then strCd$ = "HREMP.ED_SECTION"  'Lucy July 4, 2000
    
        If Len(strCd$) > 0 Then
            CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
            If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
                CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
            End If
            If Len(glbstrSelCri) > 1 Then
                glbstrSelCri = glbstrSelCri & " AND " & CodeCri
            Else
                glbstrSelCri = CodeCri
            End If
        End If
    End If
End Sub

Private Function CriCheck()
    Dim X%, I
    
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
        
    For X% = 0 To 2
        If Not clpCode(X).ListChecker Then Exit Function
    Next X%
    
    For X% = 7 To 9
        If Not clpCode(X).ListChecker Then Exit Function
    Next X%
    
    If Len(clpJOB.Text) > 0 And clpJOB.Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
         clpJOB.SetFocus
        Exit Function
    End If
    
    If Not clpPT.ListChecker Then
    'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
        'MsgBox lStr("Category code must be valid")
        'clpPT.SetFocus
        Exit Function
    End If
    
    For X% = 0 To 1
        If Len(dlpDateRange(X%).Text) > 0 Then
            If Not IsDate(dlpDateRange(X%).Text) Then
                MsgBox "Not a valid date"
                dlpDateRange(X%).Text = ""
                dlpDateRange(X%).SetFocus
                Exit Function
            End If
        End If
    Next X%
       
    If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
        If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
            MsgBox "To Date cannot be prior to From Date!"                       '
            Me.dlpDateRange(0).SetFocus                                         '
            Exit Function                                                       '
        End If
    End If
    
    For I = 0 To 2
        If elpRept(I).Caption = "Enter Valid Employee #" Then
            MsgBox "If Reporting Authority Entered - they must exist"
            elpRept(I).SetFocus
            Exit Function
        End If
    Next
    
    If Not elpEEID.ListChecker Then
        Exit Function
    End If
    
    CriCheck = True
    
End Function

Private Sub Form_Activate()
    Call SET_UP_MODE
End Sub

Private Sub Form_Load()
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    glbOnTop = Me.name
    
    If glbMultiGrid Then
        lblGrid.Visible = True
        clpGrid.Visible = True
    End If
    
    If Not glbMulti Then
        lblShift.Visible = True
        txtShift.Visible = True
    End If
    
    Call setRptCaption(Me)
    Call comGrpLoad
    
'    If Me.Caption = "Employee Profile Report" Then
'        lblGrp(1).Visible = True
'        comGroup(2).Visible = True
'    Else
'        lblGrp(1).Visible = False
'        comGroup(2).Visible = False
'    End If
    
    If glbLinamar Then clpCode(7).MaxLength = 8
    If glbCompSerial = "S/N - 2227W" Then clpCode(7).MaxLength = 6
    If glbCompSerial = "S/N - 2381W" Then clpCode(0).MaxLength = 6
    If glbCompSerial = "S/N - 2359W" Then
    End If
    
    Call INI_Controls(Me)
    
    panDetails.BorderStyle = 0 'no border
    panWindow.BevelOuter = 0 ' no bevel

    Screen.MousePointer = DEFAULT
End Sub

Private Sub Form_LostFocus()
    MDIMain.panHelp(0).Caption = " "
    MDIMain.panHelp(1).Caption = " "
    MDIMain.panHelp(2).Caption = " "
    MDIMain.panHelp(3).Caption = " "
End Sub

Private Sub Form_Resize()
On Error GoTo EH
    Dim c As Long
    
    If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
        panWindow.Height = Me.ScaleHeight - 200
        panWindow.Width = Me.ScaleWidth - (scrControl.Width + 200)
        If panWindow.Height >= 7500 Then   '+ 230 Then
            scrControl.Value = 0
            panDetails.Top = 0
            scrControl.Visible = False
        Else
            scrControl.Visible = True
            scrControl.Left = Me.ScaleWidth - scrControl.Width
            scrControl.Height = panWindow.Height
        End If
    End If

exH:
    Exit Sub
EH:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form Resize", "HR_TEMP_WORK", "Form Resize")
    Resume exH
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmRTmpCrossTrain = Nothing  'carmen apr 2000
End Sub

Private Sub scrControl_Change()
    panDetails.Top = 0 - scrControl.Value
End Sub

Private Sub txtShift_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Cri_Shift()
    Dim EECri As String, OneSet%, X%
    
    If Len(txtShift.Text) < 1 Then Exit Sub
        
    EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"

    If Len(glbstrSelCri) > 1 Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    
    glbiOneWhere = True
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

Private Sub Cri_EmpStatFTDates()
    Dim TempCri As String
    Dim dtYYY%, dtMM%, dtDD%, X%
    Dim FromDate, ToDate, SQLQ
    Dim RsHRPARCO As New ADODB.Recordset
    
    If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
        TempCri = "({HREMP.ED_SFDATE} "
        dtYYY% = Year(dlpDateRange(2).Text)
        dtMM% = month(dlpDateRange(2).Text)
        dtDD% = Day(dlpDateRange(2).Text)
        TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) and "
        
        dtYYY% = Year(dlpDateRange(3).Text)
        dtMM% = month(dlpDateRange(3).Text)
        dtDD% = Day(dlpDateRange(3).Text)
        TempCri = TempCri & " ({HREMP.ED_STDATE} <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
    
    If Len(dlpDateRange(2).Text) > 0 Or Len(dlpDateRange(3).Text) > 0 Then
        If Len(dlpDateRange(2).Text) > 0 Then
            TempCri = "({HREMP.ED_SFDATE} "
            TempCri = TempCri & " >= "
            dtYYY% = Year(dlpDateRange(2).Text)
            dtMM% = month(dlpDateRange(2).Text)
            dtDD% = Day(dlpDateRange(2).Text)
            TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
            GoTo Cri_FTDatst
        End If
        If Len(dlpDateRange(3).Text) > 0 Then
            TempCri = TempCri & "({HREMP.ED_STDATE}  "
            TempCri = TempCri & " <= "
            dtYYY% = Year(dlpDateRange(3).Text)
            dtMM% = month(dlpDateRange(3).Text)
            dtDD% = Day(dlpDateRange(3).Text)
            TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
            GoTo Cri_FTDatst
        End If
    Else
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

Private Function GetJobCodeDesc(xKey)
    Dim rsTabl As New ADODB.Recordset
    Dim SQLQ As String, xStr As String
    
    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & xKey & "' "
    rsTabl.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xStr = ""
    If Not rsTabl.EOF Then
        xStr = rsTabl("JB_DESCR")
    End If
    rsTabl.Close
    
    GetJobCodeDesc = xStr
End Function

Private Function GetTABLDesc(xName, xKey)
    Dim rsTabl As New ADODB.Recordset
    Dim SQLQ As String, xStr As String
    
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = '" & xName & "' AND TB_KEY = '" & xKey & "' "
    rsTabl.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xStr = ""
    If Not rsTabl.EOF Then
        xStr = rsTabl("TB_DESC")
    End If
    rsTabl.Close
    GetTABLDesc = xStr
End Function

Private Sub Populate_TempTable_with_Employee_PosCourses()
    Dim rsHRJObs As New ADODB.Recordset
    Dim rsEmpJob As New ADODB.Recordset
    Dim rsCrsTrainPos As New ADODB.Recordset
    Dim SQLQ As String
    Dim xLstEmpNo, xRecNum, X As Integer
    Dim xLstEmpJob As String
    
    X = 0
        
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."
    
    'Delete the existing records for this user in the temp. table
    gdbAdoIhr001.Execute "DELETE FROM HRCRSTRNWRK WHERE CT_WRKEMP ='" & glbUserID & "'"
    
    'Open blank recordset to add records
    SQLQ = "SELECT * FROM HRCRSTRNWRK WHERE 1 = 2"
    rsCrsTrainPos.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    'For reach Job retrieve employees holding this position, current or tracked
    SQLQ = "SELECT * FROM HRJOB"
    rsHRJObs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRJObs.EOF Then
        rsHRJObs.MoveFirst
        
        xRecNum = rsHRJObs.RecordCount
        
        Do While Not rsHRJObs.EOF
        
            MDIMain.panHelp(0).FloodPercent = (X / xRecNum) * 100
            
            'Retrieve employee holding this Position as Current or Tracked
            SQLQ = "SELECT 'Primary' AS Position_Type, JH_EMPNBR AS TW_EMPNBR,JH_JOB AS TW_JOB,JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE, 'TRAIN_COMPLETE' = (CASE WHEN (SELECT TOP 1 ES_DATCOMP FROM HREDSEM WHERE ES_CRSCODE = 'TRAIN' AND ES_EMPNBR = JH_EMPNBR AND ES_JOB = JH_JOB ORDER BY ES_DATCOMP DESC) IS NULL THEN 'No' ELSE 'Yes' END),datediff(Month, jh_enddate, getdate()) AS '# of Months Since End',"
            SQLQ = SQLQ & " JH_REPTAU, JH_REPTAU2, JH_REPTAU3, JH_ORG, JH_PT, JH_GRID"
            SQLQ = SQLQ & " FROM HR_JOB_HISTORY"
            SQLQ = SQLQ & " WHERE ((JH_CURRENT <> 0) Or (JH_TRK_CRS_RENEWAL <> 0))"
            SQLQ = SQLQ & " AND JH_JOB = '" & rsHRJObs("JB_CODE") & "'"
            SQLQ = SQLQ & " UNION"
            SQLQ = SQLQ & " SELECT 'Temporary' AS Position_Type, TW_EMPNBR,TW_JOB,TW_SDATE,TW_ENDDATE,'TRAIN_COMPLETE' = (CASE WHEN (SELECT TOP 1 ES_DATCOMP FROM HREDSEM WHERE ES_CRSCODE = 'TRAIN' AND ES_EMPNBR = TW_EMPNBR AND ES_JOB = TW_JOB ORDER BY ES_DATCOMP DESC) IS NULL THEN 'No' ELSE 'Yes' END),datediff(Month, tw_enddate, getdate()) AS '# of Months Since End',"
            SQLQ = SQLQ & " TW_REPTAU, TW_REPTAU2, TW_REPTAU3, TW_ORG, TW_PT, TW_GRID"
            SQLQ = SQLQ & " FROM HR_TEMP_WORK"
            SQLQ = SQLQ & " WHERE ((TW_CURRENT <> 0) Or (TW_TRK_CRS_RENEWAL <> 0))"
            SQLQ = SQLQ & " AND TW_JOB = '" & rsHRJObs("JB_CODE") & "'"
            SQLQ = SQLQ & " ORDER BY TW_EMPNBR,TW_JOB,TW_SDATE DESC"
            rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsEmpJob.EOF Then
                rsEmpJob.MoveFirst
                
                Do While Not rsEmpJob.EOF
                    If rsEmpJob("TW_EMPNBR") = xLstEmpNo And rsEmpJob("TW_JOB") = xLstEmpJob Then
                        'Skip to next employee
                    Else
                        rsCrsTrainPos.AddNew
                        rsCrsTrainPos("TW_EMPNBR") = rsEmpJob("TW_EMPNBR")
                        rsCrsTrainPos("Position_Type") = rsEmpJob("Position_Type")
                        rsCrsTrainPos("TW_JOB") = rsEmpJob("TW_JOB")
                        rsCrsTrainPos("JH_SDATE") = rsEmpJob("TW_SDATE")
                        rsCrsTrainPos("TW_ENDDATE") = rsEmpJob("TW_ENDDATE")
                        rsCrsTrainPos("TRAIN_COMPLETE") = rsEmpJob("TRAIN_COMPLETE")
                        rsCrsTrainPos("MTH_SINCE_END") = rsEmpJob("# of Months Since End")
                        rsCrsTrainPos("CT_WRKEMP") = glbUserID
                        rsCrsTrainPos("JH_REPTAU") = rsEmpJob("JH_REPTAU")
                        rsCrsTrainPos("JH_REPTAU2") = rsEmpJob("JH_REPTAU2")
                        rsCrsTrainPos("JH_REPTAU3") = rsEmpJob("JH_REPTAU3")
                        rsCrsTrainPos("JH_ORG") = rsEmpJob("JH_ORG")
                        rsCrsTrainPos("JH_PT") = rsEmpJob("JH_PT")
                        rsCrsTrainPos("JH_GRID") = rsEmpJob("JH_GRID")
                        rsCrsTrainPos.Update
                    End If
                                        
                    xLstEmpNo = rsEmpJob("TW_EMPNBR")
                    xLstEmpJob = rsEmpJob("TW_JOB")
                    
                    rsEmpJob.MoveNext
                Loop
            End If
            rsEmpJob.Close
            Set rsEmpJob = Nothing
            
            X = X + 1
            
            rsHRJObs.MoveNext
        Loop
    End If
    rsHRJObs.Close
    Set rsHRJObs = Nothing
    
    rsCrsTrainPos.Close
    Set rsCrsTrainPos = Nothing
    
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    
End Sub
