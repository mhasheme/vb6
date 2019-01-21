VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRUserDef 
   Caption         =   "User Defined Report"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9405
   ScaleWidth      =   10920
   Begin VB.VScrollBar scrControl 
      Height          =   9255
      LargeChange     =   300
      Left            =   10440
      Max             =   4000
      SmallChange     =   300
      TabIndex        =   63
      Top             =   0
      Width           =   255
   End
   Begin Threed.SSPanel panWindow 
      Height          =   9255
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   10455
      _Version        =   65536
      _ExtentX        =   18441
      _ExtentY        =   16325
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
         Height          =   8895
         Left            =   0
         ScaleHeight     =   8865
         ScaleWidth      =   10425
         TabIndex        =   33
         Top             =   120
         Width           =   10455
         Begin VB.CheckBox chkFlag 
            Alignment       =   1  'Right Justify
            Caption         =   "UFlag 5"
            Height          =   255
            Index           =   4
            Left            =   6240
            TabIndex        =   28
            Top             =   6705
            Width           =   1695
         End
         Begin VB.CheckBox chkFlag 
            Alignment       =   1  'Right Justify
            Caption         =   "UFlag 4"
            Height          =   255
            Index           =   3
            Left            =   6240
            TabIndex        =   27
            Top             =   6369
            Width           =   1695
         End
         Begin VB.CheckBox chkFlag 
            Alignment       =   1  'Right Justify
            Caption         =   "UFlag 3"
            Height          =   255
            Index           =   2
            Left            =   6240
            TabIndex        =   26
            Top             =   6036
            Width           =   1695
         End
         Begin VB.CheckBox chkFlag 
            Alignment       =   1  'Right Justify
            Caption         =   "UFlag 2"
            Height          =   255
            Index           =   1
            Left            =   6240
            TabIndex        =   25
            Top             =   5703
            Width           =   1695
         End
         Begin VB.CheckBox chkFlag 
            Alignment       =   1  'Right Justify
            Caption         =   "UFlag 1"
            Height          =   255
            Index           =   0
            Left            =   6240
            TabIndex        =   24
            Top             =   5370
            Width           =   1695
         End
         Begin VB.ComboBox comGroup 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Tag             =   "First Level of grouping records"
            Top             =   7620
            Width           =   2325
         End
         Begin VB.ComboBox comGroup 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Tag             =   "Second level of grouping records"
            Top             =   7995
            Width           =   2325
         End
         Begin VB.ComboBox comGroup 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Tag             =   "Final Sort of Records"
            Top             =   8370
            Width           =   2325
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   2
            Left            =   1800
            TabIndex        =   4
            Tag             =   "00-Enter Status Code"
            Top             =   1692
            Width           =   7515
            _ExtentX        =   13256
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "EDEM"
            MaxLength       =   0
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   3
            Tag             =   "00-Enter Union Code"
            Top             =   1359
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
            Top             =   1026
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDLC"
         End
         Begin INFOHR_Controls.CodeLookup clpDept 
            Height          =   285
            Left            =   1800
            TabIndex        =   1
            Tag             =   "00-Specific Department Desired"
            Top             =   693
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
            Left            =   1800
            TabIndex        =   7
            Tag             =   "00-Enter Administered By Code"
            Top             =   2691
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDAB"
            MaxLength       =   10
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   6
            Left            =   1800
            TabIndex        =   8
            Tag             =   "00-Enter Section Code"
            Top             =   3030
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDSE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   4
            Left            =   1800
            TabIndex        =   6
            Tag             =   "00-Enter Region Code"
            Top             =   2358
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDRG"
         End
         Begin INFOHR_Controls.EmployeeLookup elpEEID 
            Height          =   285
            Left            =   1800
            TabIndex        =   5
            Tag             =   "10-Enter Employee Number"
            Top             =   2025
            Width           =   7515
            _ExtentX        =   13256
            _ExtentY        =   503
            ShowUnassigned  =   1
            TextBoxWidth    =   7195
            RefreshDescriptionWhen=   2
            MultiSelect     =   -1  'True
         End
         Begin INFOHR_Controls.CodeLookup clpUser 
            Height          =   285
            Index           =   0
            Left            =   1800
            TabIndex        =   9
            Tag             =   "00-Enter Section Code"
            Top             =   3480
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "COD1"
         End
         Begin INFOHR_Controls.CodeLookup clpUser 
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   10
            Tag             =   "00-Enter Section Code"
            Top             =   3830
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "COD2"
         End
         Begin INFOHR_Controls.CodeLookup clpUser 
            Height          =   285
            Index           =   2
            Left            =   1800
            TabIndex        =   11
            Tag             =   "00-Enter Section Code"
            Top             =   4180
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "COD3"
         End
         Begin INFOHR_Controls.CodeLookup clpUser 
            Height          =   285
            Index           =   3
            Left            =   1800
            TabIndex        =   12
            Tag             =   "00-Enter Section Code"
            Top             =   4530
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "COD4"
         End
         Begin INFOHR_Controls.CodeLookup clpUser 
            Height          =   285
            Index           =   4
            Left            =   1800
            TabIndex        =   13
            Tag             =   "00-Enter Section Code"
            Top             =   4880
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "COD5"
         End
         Begin INFOHR_Controls.DateLookup dlpTo 
            Height          =   285
            Index           =   0
            Left            =   3840
            TabIndex        =   15
            Tag             =   "40-Date upto and including this date forward"
            Top             =   5370
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpFrom 
            Height          =   285
            Index           =   0
            Left            =   1800
            TabIndex        =   14
            Tag             =   "40-Date from and including this date forward"
            Top             =   5370
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpTo 
            Height          =   285
            Index           =   1
            Left            =   3840
            TabIndex        =   17
            Tag             =   "40-Date upto and including this date forward"
            Top             =   5700
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpFrom 
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   16
            Tag             =   "40-Date from and including this date forward"
            Top             =   5700
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpTo 
            Height          =   285
            Index           =   2
            Left            =   3840
            TabIndex        =   19
            Tag             =   "40-Date upto and including this date forward"
            Top             =   6030
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpFrom 
            Height          =   285
            Index           =   2
            Left            =   1800
            TabIndex        =   18
            Tag             =   "40-Date from and including this date forward"
            Top             =   6030
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpTo 
            Height          =   285
            Index           =   3
            Left            =   3840
            TabIndex        =   21
            Tag             =   "40-Date upto and including this date forward"
            Top             =   6375
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpFrom 
            Height          =   285
            Index           =   3
            Left            =   1800
            TabIndex        =   20
            Tag             =   "40-Date from and including this date forward"
            Top             =   6375
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpTo 
            Height          =   285
            Index           =   4
            Left            =   3840
            TabIndex        =   23
            Tag             =   "40-Date upto and including this date forward"
            Top             =   6705
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.DateLookup dlpFrom 
            Height          =   285
            Index           =   4
            Left            =   1800
            TabIndex        =   22
            Tag             =   "40-Date from and including this date forward"
            Top             =   6705
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin VB.Label lblTo 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Index           =   4
            Left            =   3480
            TabIndex        =   62
            Top             =   6750
            Width           =   375
         End
         Begin VB.Label lblTo 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Index           =   3
            Left            =   3480
            TabIndex        =   61
            Top             =   6420
            Width           =   375
         End
         Begin VB.Label lblTo 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Index           =   2
            Left            =   3480
            TabIndex        =   60
            Top             =   6075
            Width           =   375
         End
         Begin VB.Label lblTo 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Index           =   1
            Left            =   3480
            TabIndex        =   59
            Top             =   5745
            Width           =   375
         End
         Begin VB.Label lblTo 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Index           =   0
            Left            =   3480
            TabIndex        =   58
            Top             =   5415
            Width           =   375
         End
         Begin VB.Label lblDate 
            BackStyle       =   0  'Transparent
            Caption         =   "Date 5"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   57
            Top             =   6750
            Width           =   1590
         End
         Begin VB.Label lblDate 
            BackStyle       =   0  'Transparent
            Caption         =   "Date 4"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   56
            Top             =   6414
            Width           =   1590
         End
         Begin VB.Label lblDate 
            BackStyle       =   0  'Transparent
            Caption         =   "Date 3"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   55
            Top             =   6081
            Width           =   1590
         End
         Begin VB.Label lblDate 
            BackStyle       =   0  'Transparent
            Caption         =   "Date 2"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   54
            Top             =   5748
            Width           =   1590
         End
         Begin VB.Label lblDate 
            BackStyle       =   0  'Transparent
            Caption         =   "Date 1"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   53
            Top             =   5415
            Width           =   1590
         End
         Begin VB.Label lblCode 
            BackStyle       =   0  'Transparent
            Caption         =   "Code 5"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   52
            Top             =   4925
            Width           =   1590
         End
         Begin VB.Label lblCode 
            BackStyle       =   0  'Transparent
            Caption         =   "Code 4"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   51
            Top             =   4575
            Width           =   1590
         End
         Begin VB.Label lblCode 
            BackStyle       =   0  'Transparent
            Caption         =   "Code 3"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   50
            Top             =   4225
            Width           =   1590
         End
         Begin VB.Label lblCode 
            BackStyle       =   0  'Transparent
            Caption         =   "Code 2"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   49
            Top             =   3875
            Width           =   1590
         End
         Begin VB.Label lblCode 
            BackStyle       =   0  'Transparent
            Caption         =   "Code 1"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   48
            Top             =   3525
            Width           =   1590
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
            TabIndex        =   47
            Top             =   2070
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
            Left            =   120
            TabIndex        =   46
            Top             =   405
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
            TabIndex        =   45
            Top             =   738
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
            Left            =   120
            TabIndex        =   44
            Top             =   1404
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
            Left            =   120
            TabIndex        =   43
            Top             =   1737
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
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   1575
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
            Left            =   30
            TabIndex        =   41
            Top             =   7350
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
            Left            =   150
            TabIndex        =   40
            Top             =   7650
            Width           =   885
         End
         Begin VB.Label lblGrp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Grouping #2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   39
            Top             =   8055
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
            Index           =   3
            Left            =   150
            TabIndex        =   38
            Top             =   8430
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
            Left            =   120
            TabIndex        =   37
            Top             =   1071
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
            Left            =   120
            TabIndex        =   36
            Top             =   2403
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
            Left            =   120
            TabIndex        =   35
            Top             =   2736
            Width           =   1125
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
            TabIndex        =   34
            Top             =   3075
            Width           =   540
         End
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   12600
      Top             =   11400
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
Attribute VB_Name = "frmRUserDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Call set_Buttons
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    glbOnTop = Me.name
    Dim c As Long
    
    Me.Caption = lStr("User Defined Table") & " Report"
    
    Call setCaption(lblDiv)
    Call setCaption(lblDept)
    Call setCaption(lblLocation)
    Call setCaption(lblUnion)
    Call setCaption(lblStatus)
    Call setCaption(lblEENum)
    Call setCaption(lblRegion)
    Call setCaption(lblAdmin)
    Call setCaption(lblSection)
    For c = 0 To 4
        Call setCaption(lblCode(c))
        Call setCaption(chkFlag(c))
        Call setCaption(lblDate(c))
    Next c
    
    'Get Grouping
    Call comGrpLoad
    
    'Jerry said no to this.
'    'Ticket #22421 - Donalda Club
'    If glbCompSerial = "S/N - 2438W" Then
'        For c = 0 To 4
'            If lblCode(c).Caption = "." Then
'                lblCode(c).Visible = False
'                clpUser(c).Visible = False
'            Else
'                lblCode(c).Visible = True
'                clpUser(c).Visible = True
'            End If
'            If chkFlag(c).Caption = "." Then
'                chkFlag(c).Visible = False
'            Else
'                chkFlag(c).Visible = True
'            End If
'
'            If lblDate(c).Caption = "." Then
'                lblDate(c).Visible = False
'                dlpFrom(c).Visible = False
'                dlpTo(c).Visible = False
'                lblTo(c).Visible = False
'            Else
'                lblDate(c).Visible = True
'                dlpFrom(c).Visible = True
'                dlpTo(c).Visible = True
'                lblTo(c).Visible = True
'            End If
'        Next c
'    End If
    
    Call INI_Controls(Me)
    panDetails.BorderStyle = 0 'no border
    panWindow.BevelOuter = 0 ' no bevel
    Screen.MousePointer = DEFAULT
    Me.WindowState = vbMaximized
    
exH:
    Exit Sub
EH:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form_Load", "HR_USERDEFINE_TABLE", "SELECT")
    Resume exH
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
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form_Resize", "HR_USERDEFINE_TABLE", "SELECT")
    Resume exH
End Sub

Private Sub comGrpLoad()
    Dim c As Long
    
    'Sort Grouping 1
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Union")
    comGroup(0).AddItem lStr("Section")
    comGroup(0).AddItem lStr("Employee Name")
    comGroup(0).AddItem lStr("(none)")
    comGroup(0).ListIndex = 0
    
    'Group 2
    comGroup(1).AddItem lStr("Employee Name")
    For c = 0 To 4
        comGroup(1).AddItem lStr("Code " & CStr(c + 1))
    Next c
    For c = 0 To 4
        comGroup(1).AddItem lStr("Date " & CStr(c + 1))
    Next c
    comGroup(1).AddItem lStr("(none)")
    comGroup(1).ListIndex = 0
    
    'Final Sort
    comGroup(2).AddItem lStr("Employee Name")
    comGroup(2).AddItem lStr("Employee Number")
    comGroup(2).ListIndex = 0
    
End Sub

Private Function cri_SetAll() As Integer
    On Error GoTo EH
    Dim strRName As String, x As Integer

    cri_SetAll = False
    Screen.MousePointer = HOURGLASS
    glbiOneWhere = False
    glbstrSelCri = ""

       
    Call glbCri_DeptUN(clpDept.Text)
    Call Cri_Div
    Call Cri_Code(0) 'Location
    Call Cri_Code(1) 'Union
    Call Cri_Code(2) 'Status
    Call Cri_Code(4) 'Region
    Call Cri_Code(5) 'Admin By
    Call Cri_Code(6) 'Section
    Call Cri_EE
    Call Cri_User
    Call Cri_Date
    Call Cri_Flag
    
    Call setReportLabels
    
    'Ticket #22421 - Donalda Club
    If glbCompSerial = "S/N - 2438W" Then
        If comGroup(0) = "(none)" And comGroup(1) = "(none)" Then
            strRName = glbIHRREPORTS & "SN2438_SeasStaff1.rpt"
        Else
            strRName = glbIHRREPORTS & "SN2438_SeasStaff.rpt"
        End If
    Else
        strRName = glbIHRREPORTS & "rzUserDef.rpt"
    End If
    Me.vbxCrystal.ReportFileName = strRName$
    
    ' set to sorting/grouping criteria
    x = Cri_Sorts()   ' returns number of sections formated
    Me.vbxCrystal.SelectionFormula = "(" & glbstrSelCri & " )"
    
    Me.vbxCrystal.Connect = RptODBC_SQL
    
    cri_SetAll = True

    Screen.MousePointer = DEFAULT

    
exH:
    Exit Function
EH:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cri_SetAll", "HR_USERDEFINE_TABLE", "SELECT")
    Resume exH
End Function

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

Private Sub Cri_Code(intIdx As Integer)
Dim CodeCri As String
Dim strCd As String

    If Len(clpCode(intIdx).Text) > 0 Then
        If intIdx = 0 Then strCd = "HREMP.ED_LOC"
        If intIdx = 1 Then strCd = "HREMP.ED_ORG"
        If intIdx = 2 Then strCd = "HREMP.ED_EMP"
        If intIdx = 4 Then strCd = "HREMP.ED_REGION"
        If intIdx = 5 Then strCd = "HREMP.ED_ADMINBY"
        If intIdx = 6 Then strCd = "HREMP.ED_SECTION"  'Lucy July 4, 2000
            CodeCri = "({" & strCd & "} in  ['" & Replace(clpCode(intIdx).Text, ",", "','") & "'])"
        If glbLinamar And (strCd = "HREMP.ED_REGION" Or strCd = "HREMP.ED_SECTION") Then
            CodeCri = "(({" & strCd & "} = '" & clpDiv.Text & clpCode(intIdx).Text & "') or ({" & strCd & "} = 'ALL" & clpCode(intIdx).Text & "') )"
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

Private Sub Cri_User()
    Dim StrCri As String
    Dim c As Long
    
    StrCri = ""
    
    For c = 0 To 4
        If Len(clpUser(c).Text) > 0 Then
            If Len(StrCri) > 0 Then
                StrCri = StrCri & " AND "
            End If
            StrCri = StrCri & "{HR_USERDEFINE_TABLE.UD_CODE" & CStr(c + 1) & "}='" & clpUser(c).Text & "'"
        End If
    Next c
    
    If Len(StrCri) > 0 Then
        If Len(glbstrSelCri) > 1 Then
            glbstrSelCri = glbstrSelCri & " AND " & StrCri
        Else
            glbstrSelCri = StrCri
        End If
    End If
End Sub

Private Sub Cri_Date()
    Dim TempCri As String
    Dim c As Long, dtYYY%, dtMM%, dtDD%
    
    TempCri = ""
    
    For c = 0 To 4
        If Len(TempCri) > 0 And (Len(dlpFrom(c).Text) > 0 Or Len(dlpTo(c).Text) > 0) Then
            TempCri = TempCri & " AND "
        End If
        If Len(dlpFrom(c).Text) > 0 And Len(dlpTo(c).Text) > 0 Then
            TempCri = TempCri & "({HR_USERDEFINE_TABLE.UD_DATE" & CStr(c + 1) & "} "
            dtYYY% = Year(dlpFrom(c).Text)
            dtMM% = month(dlpFrom(c).Text)
            dtDD% = Day(dlpFrom(c).Text)
            TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
            dtYYY% = Year(dlpTo(c).Text)
            dtMM% = month(dlpTo(c).Text)
            dtDD% = Day(dlpTo(c).Text)
            TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        ElseIf Len(dlpFrom(c).Text) > 0 Then
            TempCri = TempCri & "({HR_USERDEFINE_TABLE.UD_DATE" & CStr(c + 1) & "} "         ' Added section to enable entering only From date, no To date.
            dtYYY% = Year(dlpFrom(c).Text)
            dtMM% = month(dlpFrom(c).Text)
            dtDD% = Day(dlpFrom(c).Text)
            TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        ElseIf Len(dlpTo(c).Text) > 0 Then
            TempCri = TempCri & "({HR_USERDEFINE_TABLE.UD_DATE" & CStr(c + 1) & "} "         ' Added section to enable entering only To date, no From date.
            dtYYY% = Year(dlpTo(c).Text)
            dtMM% = month(dlpTo(c).Text)
            dtDD% = Day(dlpTo(c).Text)
            TempCri = TempCri & " <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        End If
    Next c
    
    If Len(TempCri) > 0 Then
        If Len(glbstrSelCri) > 1 Then
            glbstrSelCri = glbstrSelCri & " AND " & TempCri
        Else
            glbstrSelCri = TempCri
        End If
    End If
End Sub

Private Sub Cri_Flag()
Dim StrCri As String
    Dim c As Long
    
    StrCri = ""
    
    For c = 0 To 4
        If chkFlag(c).Value <> vbGrayed Then
            If Len(StrCri) > 0 Then
                StrCri = StrCri & " AND "
            End If
            StrCri = StrCri & "{HR_USERDEFINE_TABLE.UD_FLAG" & CStr(c + 1) & "}=" & CBool(chkFlag(c).Value)
        End If
    Next c
    
    If Len(StrCri) > 0 Then
        If Len(glbstrSelCri) > 1 Then
            glbstrSelCri = glbstrSelCri & " AND " & StrCri
        Else
            glbstrSelCri = StrCri
        End If
    End If
End Sub

Private Sub setReportLabels()
    Dim x As Long, c As Long

    x = 1
    For c = 0 To 4
        vbxCrystal.Formulas(x) = "lblCode" & CStr(c + 1) & "='" & lStr("Code " & CStr(c + 1)) & "'"
        x = x + 1
    Next c
    For c = 0 To 4
        vbxCrystal.Formulas(x) = "lblDate" & CStr(c + 1) & "='" & lStr("Date " & CStr(c + 1)) & "'"
        x = x + 1
    Next c
    For c = 0 To 4
        vbxCrystal.Formulas(x) = "lblFlag" & CStr(c + 1) & "='" & lStr("UFlag " & CStr(c + 1)) & "'"
        x = x + 1
    Next c
    For c = 0 To 1
        vbxCrystal.Formulas(x) = "lblUText" & CStr(c + 1) & "='" & lStr("UText " & CStr(c + 1)) & "'"
        x = x + 1
    Next c
    vbxCrystal.Formulas(x) = "lblUComments='" & lStr("UComments") & "'"
    'Ticket #22421 - Donalda Club
    If glbCompSerial = "S/N - 2438W" Then
        vbxCrystal.Formulas(x + 1) = "lblTitle='Summary of Seasonal Staff'"
    Else
        vbxCrystal.Formulas(x + 1) = "lblTitle='" & lStr("User Defined Table") & "'"
    End If
End Sub

Private Sub scrControl_Change()
panDetails.Top = 0 - scrControl.Value
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

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
  If FormEmplPosition% = True Then
    If Not PrtForm("Employee/Position Report Criteria", Me) Then Exit Sub
  ElseIf FormLanguages% = True Then
    If Not PrtForm("Languages Report Criteria", Me) Then Exit Sub    'laura nov 3, 1997
  Else
  End If
    Call set_PrintState(False)
    x% = cri_SetAll()
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
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
    Call set_PrintState(False)
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    x% = cri_SetAll()
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

Private Function CriCheck()
Dim x As Long
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


For x = 0 To 6
    If x <> 3 Then
        If Not clpCode(x).ListChecker Then Exit Function
    End If
Next x

For x = 0 To 4
    If Len(clpUser(x).Text) > 0 And clpUser(x).Caption = "Unassigned" Then
        MsgBox "If " & lStr("Code " & CStr(x + 1)) & " Entered - it must be known"
        clpUser(x).Text = ""
        clpUser(x).SetFocus
        Exit Function
    End If
    If Len(dlpFrom(x).Text) > 0 Then
        If Not IsDate(dlpFrom(x).Text) Then
            MsgBox "Not a valid date"
            dlpFrom(x).Text = ""
            dlpFrom(x).SetFocus
            Exit Function
        End If
    End If
    If Len(dlpTo(x).Text) > 0 Then
        If Not IsDate(dlpTo(x).Text) Then
            MsgBox "Not a valid date"
            dlpTo(x).Text = ""
            dlpTo(x).SetFocus
            Exit Function
        End If
    End If
Next x

If Not elpEEID.ListChecker Then
    Exit Function
End If

CriCheck = True
End Function

Private Function Cri_Sorts() As Integer
    Dim grpCond$, grpField$, strSFormat$
Dim x%, y%, z%
Dim dscGroup$, GrpIdx%

'X% = 0
'y% = 0
'grpField$ = getEGroup(comGroup(0).Text)
'
'If comGroup(0).Text = "None" Then grpField$ = "{HRPARCO.PC_CO}"
'
'y% = X% + 1
'dscGroup$ = comGroup(0).Text
'dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
'Me.vbxCrystal.Formulas(20) = dscGroup$
'
'grpCond$ = "GROUP" & CStr(y%) & ";" & grpField$ & ";ANYCHANGE;A"
'Me.vbxCrystal.GroupCondition(X%) = grpCond$
'
'If comGroup(0) <> "None" Then
'    strSFormat$ = "GH1;T;X;X;X;X;X;X"
'Else
'    strSFormat$ = "GH1;F;X;X;X;X;X;X"
'End If
'Me.vbxCrystal.SectionFormat(0) = strSFormat$
'
'
'dscGroup$ = comGroup(1).Text
'dscGroup$ = "descGroup" & CStr(2) & "= '" & dscGroup$ & "'"
'Me.vbxCrystal.Formulas(21) = dscGroup$
'
'GrpIdx% = comGroup(1).ListIndex
'Select Case GrpIdx%
'    Case 0: grpField$ = "{@EFullName}"
'    Case 1 To 5: grpField$ = "{tblcode" & CStr(GrpIdx%) & ".TB_DESC}"
'    Case 6 To 10: grpField$ = "{HR_USERDEFINE_TABLE.UD_DATE" & CStr(GrpIdx% - 5) & "}"
'    Case 11: grpField$ = "{HRPARCO.PC_CO}"
'End Select
'grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
'Me.vbxCrystal.GroupCondition(1) = grpCond$
'
'If comGroup(1) <> "None" Then
'    strSFormat$ = "GH2;T;X;X;X;X;X;X"
'Else
'    strSFormat$ = "GH2;F;X;X;X;X;X;X"
'End If
'Me.vbxCrystal.SectionFormat(1) = strSFormat$
'
'GrpIdx% = comGroup(2).ListIndex
'Select Case GrpIdx%
'    Case 0: grpField = "{@EFullName}"
'    Case 1: grpField = "{HREMP.ED_EMPNBR}"
'End Select
'grpCond$ = "GROUP3;" & grpField$ & ";ANYCHANGE;A"
'Me.vbxCrystal.GroupCondition(2) = grpCond$
'
'Cri_Sorts = z% ' next section number to format
'


Cri_Sorts = 0
grpField$ = getEGroup(comGroup(0).Text)

If grpField$ <> "(none)" Then
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(20) = dscGroup$

    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$

    If comGroup(0) <> "(none)" Then
        strSFormat$ = "GH1;T;X;X;X;X;X;X"
    Else
        strSFormat$ = "GH1;F;X;X;X;X;X;X"
    End If
    Me.vbxCrystal.SectionFormat(0) = strSFormat$

    dscGroup$ = comGroup(1).Text
    dscGroup$ = "descGroup" & CStr(2) & "= '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(21) = dscGroup$

    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1 To 5: grpField$ = "{tblcode" & CStr(GrpIdx%) & ".TB_DESC}"
        Case 6 To 10: grpField$ = "{HR_USERDEFINE_TABLE.UD_DATE" & CStr(GrpIdx% - 5) & "}"
        Case 11: grpField$ = "(none)"
    End Select
    
    If grpField$ <> "(none)" Then
        grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(1) = grpCond$

        If comGroup(1) <> "(none)" Then
            strSFormat$ = "GH2;T;X;X;X;X;X;X"
        Else
            strSFormat$ = "GH2;F;X;X;X;X;X;X"
        End If
        Me.vbxCrystal.SectionFormat(1) = strSFormat$

        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField = "{@EFullName}"
            Case 1: grpField = "{HREMP.ED_EMPNBR}"
        End Select
        grpCond$ = "GROUP3;" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(2) = grpCond$
    Else
        strSFormat$ = "GH2;F;X;X;X;X;X;X"
        Me.vbxCrystal.SectionFormat(1) = strSFormat$
    End If
Else
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1 To 5: grpField$ = "{tblcode" & CStr(GrpIdx%) & ".TB_DESC}"
        Case 6 To 10: grpField$ = "{HR_USERDEFINE_TABLE.UD_DATE" & CStr(GrpIdx% - 5) & "}"
        Case 11: grpField$ = "(none)"
    End Select
    
    If grpField$ <> "(none)" Then
        dscGroup$ = comGroup(1).Text
        dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
        Me.vbxCrystal.Formulas(20) = dscGroup$
        
        grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(0) = grpCond$
        
        If comGroup(1) <> "(none)" Then
            strSFormat$ = "GH1;T;X;X;X;X;X;X"
        Else
            strSFormat$ = "GH1;F;X;X;X;X;X;X"
        End If
        Me.vbxCrystal.SectionFormat(0) = strSFormat$
        
        'Ticket #22421 - Donalda Club
        If glbCompSerial <> "S/N - 2438W" Then
            dscGroup$ = comGroup(1).Text
            dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
            Me.vbxCrystal.Formulas(20) = dscGroup$
        End If
        
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField = "{@EFullName}"
            Case 1: grpField = "{HREMP.ED_EMPNBR}"
        End Select
        'grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        'Me.vbxCrystal.GroupCondition(1) = grpCond$
    Else
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField = "{@EFullName}"
            Case 1: grpField = "{HREMP.ED_EMPNBR}"
        End Select
        grpCond$ = "GROUP1;" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(0) = grpCond$
    
        'Ticket #22421 - Donalda Club
        If glbCompSerial <> "S/N - 2438W" Then
            dscGroup$ = comGroup(2).Text
            dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
            Me.vbxCrystal.Formulas(20) = dscGroup$
        End If
    End If
End If

Cri_Sorts = z% ' next section number to format

End Function
