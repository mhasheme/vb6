VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRINCOST 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Accident Cost Analysis Report"
   ClientHeight    =   7560
   ClientLeft      =   90
   ClientTop       =   675
   ClientWidth     =   12810
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
   ScaleHeight     =   7560
   ScaleWidth      =   12810
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbHRat 
      Height          =   315
      Left            =   2595
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Tag             =   "00-Hazard Rating - Choose A,B, or C"
      Top             =   5795
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txtHRat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "EC_HAZARD"
      Height          =   315
      Left            =   3300
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   5795
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chkDivStart 
      Caption         =   "Incident Date >= Seniority Date"
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
      Left            =   6120
      TabIndex        =   24
      Top             =   5475
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   5340
      MaxLength       =   4
      TabIndex        =   23
      Tag             =   "00-Shift from H&S Incidents"
      Top             =   5460
      Width           =   450
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   2595
      MaxLength       =   4
      TabIndex        =   22
      Tag             =   "00-00-Employee Position Shift"
      Top             =   5460
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   2280
      TabIndex        =   7
      Tag             =   "00-Rate Group"
      Top             =   2705
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Tag             =   "00-Incident Type"
      Top             =   2370
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
   End
   Begin VB.ComboBox comShift 
      Height          =   315
      Left            =   2600
      TabIndex        =   9
      Tag             =   "00-Incident Number"
      Top             =   3375
      Width           =   1815
   End
   Begin VB.TextBox txtIncNo 
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
      Left            =   4700
      MaxLength       =   30
      TabIndex        =   10
      Tag             =   "01-Type of Incident- Code"
      Top             =   3330
      Visible         =   0   'False
      Width           =   870
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   49
      Top             =   6900
      Width           =   12810
      _Version        =   65536
      _ExtentX        =   22595
      _ExtentY        =   1164
      _StockProps     =   15
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   3540
         Top             =   165
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
   Begin MSMask.MaskEdBox medDOW 
      Height          =   330
      Index           =   0
      Left            =   2580
      TabIndex        =   13
      Tag             =   "10-Day of week 1=Sunday, 2=Monday..."
      Top             =   4075
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      Mask            =   "#"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medDOW 
      Height          =   330
      Index           =   1
      Left            =   3435
      TabIndex        =   14
      Tag             =   "10-Day of week 1=Sunday, 2=Monday..."
      Top             =   4075
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      Mask            =   "#"
      PromptChar      =   "_"
   End
   Begin Threed.SSFrame fraDOW 
      Height          =   2085
      Left            =   10140
      TabIndex        =   37
      Top             =   240
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   3678
      _StockProps     =   14
      Caption         =   "Days of the Week"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "7 - Saturday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   6
         Left            =   450
         TabIndex        =   44
         Top             =   1800
         Width           =   1290
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "6 - Friday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   450
         TabIndex        =   43
         Top             =   1560
         Width           =   1290
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "5 - Thursday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   450
         TabIndex        =   42
         Top             =   1320
         Width           =   1290
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "4 - Wednesday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   450
         TabIndex        =   41
         Top             =   1080
         Width           =   1290
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "3 - Tuesday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   450
         TabIndex        =   40
         Top             =   840
         Width           =   1290
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "2 - Monday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   450
         TabIndex        =   39
         Top             =   600
         Width           =   1290
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1 - Sunday"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   450
         TabIndex        =   38
         Top             =   360
         Width           =   1290
      End
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1030
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   695
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
      Left            =   2280
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
      Index           =   3
      Left            =   2280
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1700
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
      Left            =   2280
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2035
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
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1365
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDOR"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   4380
      TabIndex        =   12
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3690
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   11
      Tag             =   "40-Date from and including this date forward"
      Top             =   3740
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   2280
      TabIndex        =   16
      Tag             =   "00-Enter Administered By Code"
      Top             =   4455
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   7
      Left            =   7320
      TabIndex        =   17
      Tag             =   "00-Enter Section Code"
      Top             =   4455
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   7320
      TabIndex        =   15
      Tag             =   "00-Enter Region Code"
      Top             =   4098
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Tag             =   "10-Enter Employee Number"
      Top             =   3040
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   13
      Left            =   2280
      TabIndex        =   20
      Tag             =   "00-Classification of Incident"
      Top             =   5125
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECCL"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   14
      Left            =   7320
      TabIndex        =   21
      Tag             =   "00-Equipment being used when injured"
      Top             =   5125
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECEQ"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   12
      Left            =   7320
      TabIndex        =   19
      Tag             =   "01-Primary Cause of Injury"
      Top             =   4790
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECCA"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   11
      Left            =   2280
      TabIndex        =   18
      Tag             =   "01-Injury - Code"
      Top             =   4790
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECCD"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   8
      Left            =   2280
      TabIndex        =   26
      Tag             =   "01-Results of Claim - Code"
      Top             =   6165
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECRS"
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
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
      Left            =   300
      TabIndex        =   61
      Top             =   6210
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hazard Rating"
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
      Height          =   255
      Index           =   9
      Left            =   300
      TabIndex        =   60
      Tag             =   "00-Hazard Rationg"
      Top             =   5797
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Primary Injury"
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
      Left            =   300
      TabIndex        =   58
      Top             =   4786
      Width           =   930
   End
   Begin VB.Label lblCause 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Primary Cause"
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
      Left            =   6000
      TabIndex        =   57
      Top             =   4835
      Width           =   1215
   End
   Begin VB.Label lblEquipment 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Equipment"
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
      Left            =   6000
      TabIndex        =   56
      Top             =   5170
      Width           =   900
   End
   Begin VB.Label lblClassification 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Classification"
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
      Left            =   300
      TabIndex        =   55
      Top             =   5123
      Width           =   1155
   End
   Begin VB.Label lblHSShift 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Incident Shift"
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
      Left            =   3780
      TabIndex        =   54
      Top             =   5505
      Width           =   1185
   End
   Begin VB.Label lblShift 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
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
      Left            =   300
      TabIndex        =   53
      Top             =   5460
      Visible         =   0   'False
      Width           =   645
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
      Left            =   300
      TabIndex        =   52
      Top             =   2090
      Width           =   630
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Incident Number"
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
      Left            =   300
      TabIndex        =   51
      Top             =   3438
      Width           =   1515
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Left            =   6000
      TabIndex        =   50
      Top             =   4500
      Width           =   1260
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Left            =   6000
      TabIndex        =   48
      Top             =   4143
      Width           =   1335
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Left            =   300
      TabIndex        =   47
      Top             =   4449
      Width           =   1125
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   300
      TabIndex        =   46
      Top             =   1079
      Width           =   615
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
      TabIndex        =   45
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
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
      Left            =   2985
      TabIndex        =   36
      Top             =   4143
      Width           =   270
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Day of Week"
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
      Left            =   300
      TabIndex        =   35
      Top             =   4112
      Width           =   945
   End
   Begin VB.Label lblFromTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From / To Date"
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
      Left            =   300
      TabIndex        =   34
      Top             =   3775
      Width           =   1095
   End
   Begin VB.Label lblEmpName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   300
      TabIndex        =   33
      Top             =   3101
      Width           =   1290
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Group"
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
      Left            =   300
      TabIndex        =   32
      Top             =   2764
      Width           =   825
   End
   Begin VB.Label lblBCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Incident Type"
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
      Left            =   300
      TabIndex        =   31
      Top             =   2427
      Width           =   975
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Left            =   300
      TabIndex        =   30
      Top             =   1753
      Width           =   450
   End
   Begin VB.Label lblUnion 
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
      Left            =   300
      TabIndex        =   29
      Top             =   1416
      Width           =   420
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
      Left            =   300
      TabIndex        =   28
      Top             =   742
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
      Left            =   300
      TabIndex        =   27
      Top             =   405
      Width           =   555
   End
End
Attribute VB_Name = "frmRINCOST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub cmdClose_Click()
Unload Me
End Sub

Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Analysis of Incidents by Employee Report Criteria", Me) Then Exit Sub
    'cmdPrint.Enabled = False
    'cmdView.Enabled = False
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    'cmdPrint.Enabled = True
    'cmdView.Enabled = True
End If
Exit Sub

PrntErr:
MsgBox "Error Printing - check your Windows Printer setup"
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    'cmdPrint.Enabled = False
    'cmdView.Enabled = False
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    Screen.MousePointer = HOURGLASS
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    'cmdPrint.Enabled = True
    'cmdView.Enabled = True
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "OHS", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    If intIdx% = 0 Then strCd$ = "HREMP.ED_LOC"
    If intIdx% = 1 Then strCd$ = "HR_OCC_HEALTH_SAFETY.EC_TYPE"
    If intIdx% = 2 Then strCd$ = "HREMP.ED_ORG"
    If intIdx% = 3 Then strCd$ = "HREMP.ED_EMP"
    If intIdx% = 4 Then strCd$ = "HR_OCC_HEALTH_SAFETY.EC_RATEGRP"
    If intIdx% = 5 Then strCd$ = "HREMP.ED_REGION"
    If intIdx% = 6 Then strCd$ = "HREMP.ED_ADMINBY"
    If intIdx% = 7 Then strCd$ = "HREMP.ED_SECTION"
    If intIdx% = 11 Then strCd$ = "HR_OCC_HEALTH_SAFETY.EC_CODE"
    If intIdx% = 12 Then strCd$ = "HR_OCC_HEALTH_SAFETY.EC_CAUSECD"
    If intIdx% = 13 Then strCd$ = "HR_OCC_HEALTH_SAFETY.EC_CLASS"
    If intIdx% = 14 Then strCd$ = "HR_OCC_HEALTH_SAFETY.EC_EQUIP"
    If intIdx% = 8 Then strCd$ = "HR_OCC_HEALTH_SAFETY.EC_WCBRES"
    
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

Private Sub Cri_DOW()
Dim EECri As String, OneSet%
Dim x%

For x% = 0 To 1
    If Len(medDOW(x%)) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next x%

If OneSet% = 0 Then Exit Sub

For x% = 0 To 1
    If Len(medDOW(x%)) > 0 Then
        EECri = EECri & medDOW(x%)
        OneSet% = OneSet% - 1
    End If
    If OneSet% > 0 Then
        EECri = EECri & " ,"
    End If
Next x%
EECri = "{@DOW} in [" & EECri & "]"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_IncNo()
Dim EECri As String

If Len(txtIncNo.Text) > 0 Then
    EECri = "{HR_OCC_HEALTH_SAFETY.EC_CASE} = " & Val(txtIncNo) & " "
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

Private Sub Cri_Shift()
Dim EECri As String, OneSet%, x%

If Len(txtShift(0).Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_SHIFT}= '" & txtShift(0).Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Sub Cri_HSShift()
Dim ShftCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim OneSet%, x%

For x% = 1 To 1
    If Len(txtShift(x%)) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next x%
If OneSet% = 0 Then Exit Sub

For x% = 1 To 1
    If Len(txtShift(x%)) > 0 Then
        ShftCri = ShftCri & "'" & txtShift(x%) & "'"
        OneSet% = OneSet% - 1
        If OneSet% > 0 Then
            ShftCri = ShftCri & " ,"
        End If
    End If
Next x%
ShftCri = "{HR_OCC_HEALTH_SAFETY.EC_SHIFT} in [" & ShftCri & "]"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & ShftCri
Else
    glbstrSelCri = ShftCri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_FTDates()

Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HR_OCC_HEALTH_SAFETY.EC_OCCDATE} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
End If

For x% = 0 To 1
    If Len(dlpDateRange(0).Text) > 0 Then
        TempCri = "({HR_OCC_HEALTH_SAFETY.EC_OCCDATE}  "
        If x% = 0 Then
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
Next x%

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

Private Sub Cri_Hazard()
Dim BodyCri As String

If cmbHRat.Text <> "" Then
    BodyCri = "{HR_OCC_HEALTH_SAFETY.EC_HAZARD} = '" & cmbHRat.Text & "'"
Else
    Exit Sub
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & BodyCri
Else
    glbstrSelCri = BodyCri
End If
glbiOneWhere = True

End Sub

Private Function Cri_SetAll()
Dim x%, Rptname, xdbf

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

xdbf = 6
glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
For x% = 0 To 4
    Call Cri_Code(x%)
Next x%
Call Cri_Code(5)
Call Cri_Code(6)
Call Cri_Code(7)
Call Cri_PT
Call Cri_Shift
Call Cri_FTDates
Call Cri_DOW
Call Cri_EE
Call Cri_IncNo  'Add by Frank 04/26/2001 '
Call Cri_HSShift

Call Cri_Code(11)
Call Cri_Code(12)
Call Cri_Code(13)
Call Cri_Code(14)
'Call Cri_Code(8)

'Call Cri_Hazard

'If glbLinamar Then
    If chkDivStart Then
        Call Cri_DivStart
    End If
'End If
Rptname = "rzincost.rpt"

Me.vbxCrystal.ReportFileName = glbIHRREPORTS & Rptname

'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For x% = 0 To 2
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next x%
    Me.vbxCrystal.DataFiles(3) = glbIHRWFC
End If
' window title if appropriate
Me.vbxCrystal.WindowTitle = "Accident Cost Analysis Report"

Cri_SetAll = True

Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "Comp Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Function CriCheck()
Dim x%

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

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

For x% = 0 To 7
    If Not clpCode(x).CheckList Then Exit Function
Next x%

For x% = 0 To 1
 If Len(dlpDateRange(x%).Text) > 0 Then
    If Not IsDate(dlpDateRange(x%).Text) Then
        MsgBox "Not a valid date"
        dlpDateRange(x%).Text = ""
        dlpDateRange(x%).SetFocus
        Exit Function
    End If
 End If
Next x%

For x% = 0 To 1
 If Len(medDOW(x%)) > 0 Then
    If medDOW(x%) < 0 Or medDOW(x%) > 7 Then
        MsgBox "Day of Week must be between 1 and 7"
        medDOW(x%).SetFocus
        Exit Function
    End If
  End If
Next x%

If Not elpEEID.ListChecker Then
    Exit Function
End If

CriCheck = True

End Function

Private Sub clpCode_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbHRat_Change()
'Ticket # 6831 - For Burlington Tech.
If glbCompSerial = "S/N - 2351W" Then
    If cmbHRat = "" Then
        txtHRat = ""
    End If
End If
End Sub

Private Sub cmbHRat_Click()
'Ticket # 6831 - For Burlington Tech.
If glbCompSerial = "S/N - 2351W" Then
    Select Case cmbHRat
        Case "Frequent": txtHRat = "F"
        Case "Occassional": txtHRat = "O"
        Case "Rare": txtHRat = "R"
        Case Else
            txtHRat = ""
    End Select
Else
    txtHRat = cmbHRat
End If
End Sub

Private Sub cmbHRat_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comShift_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comShift_LostFocus()
txtIncNo = comShift
End Sub

Private Sub Form_Activate()
    Call SET_UP_MODE

    chkDivStart.Visible = True
    chkDivStart.Caption = lStr(chkDivStart.Caption)
End Sub

Private Sub Form_Load()

glbOnTop = "FRMRINCOST"

Dim rsTB As New ADODB.Recordset
                     
Screen.MousePointer = HOURGLASS

If Not glbMulti Then
    lblShift.Visible = True
    txtShift(0).Visible = True
End If

rsTB.Open "select * FROM HR_OCC_HEALTH_SAFETY ORDER BY EC_CASE DESC", gdbAdoIhr001, adOpenKeyset
comShift.AddItem ""
Do Until rsTB.EOF
    comShift.AddItem rsTB("EC_CASE")
    rsTB.MoveNext
Loop
rsTB.Close

If glbLinamar Then
    clpCode(5).MaxLength = 8
    'Ticket #14573
    lblClassification.Caption = "Type of Event"
    lblCause.Caption = "Root Cause"
End If

If glbCompSerial = "S/N - 2227W" Then clpCode(5).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

'Ticket # 6831 - For Burlington Tech.
If glbCompSerial = "S/N - 2351W" Then
    lblTitle(9).Caption = "Frequency"
    cmbHRat.Width = 1500
    'cmbHRat.Left = 7560
    cmbHRat.AddItem ""
    cmbHRat.AddItem "Frequent"
    cmbHRat.AddItem "Occassional"
    cmbHRat.AddItem "Rare"
Else
    cmbHRat.AddItem ""
    cmbHRat.AddItem "A"
    cmbHRat.AddItem "B"
    cmbHRat.AddItem "C"
End If

Call setRptCaption(Me)
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

Private Sub medDOW_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtHRat_Change()
'Ticket # 6831 - For Burlington Tech.
If glbCompSerial = "S/N - 2351W" Then
    Select Case txtHRat
        Case "F": cmbHRat.ListIndex = 0
        Case "O": cmbHRat.ListIndex = 1
        Case "R": cmbHRat.ListIndex = 2
        Case Else
            cmbHRat.ListIndex = -1
    End Select
Else
    cmbHRat = txtHRat
End If
End Sub

Private Sub txtShift_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
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

Private Sub Cri_DivStart()
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

CodeCri = "({HR_OCC_HEALTH_SAFETY.EC_OCCDATE} >= {HREMP.ED_SENDTE} )"
If Len(CodeCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CodeCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CodeCri
    End If
    glbiOneWhere = True
End If

End Sub
