VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRINJURY 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Health and Safety Reports"
   ClientHeight    =   9045
   ClientLeft      =   90
   ClientTop       =   675
   ClientWidth     =   11760
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
   ScaleHeight     =   9045
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbHRat 
      Height          =   315
      Left            =   2100
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Tag             =   "00-Hazard Rating - Choose A,B, or C"
      Top             =   5040
      Width           =   645
   End
   Begin VB.TextBox txtHRat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   315
      Left            =   3060
      TabIndex        =   89
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox comGroup 
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
      Index           =   0
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Tag             =   "First Level of grouping records"
      Top             =   6210
      Width           =   2325
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
      Left            =   6000
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox chkComments 
      Caption         =   "Include Comments"
      Height          =   255
      Left            =   6000
      TabIndex        =   24
      Top             =   5435
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   5340
      MaxLength       =   4
      TabIndex        =   13
      Tag             =   "00-Shift from H&S Incidents"
      Top             =   3360
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   6600
      TabIndex        =   7
      Tag             =   "00-Rate Group"
      Top             =   2340
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECGP"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   6
      Tag             =   "00-Incident Type"
      Top             =   2340
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECTY"
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   71
      Top             =   8385
      Visible         =   0   'False
      Width           =   11760
      _Version        =   65536
      _ExtentX        =   20743
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
      Index           =   1
      Left            =   2955
      TabIndex        =   12
      Tag             =   "10-Day of week 1=Sunday, 2=Monday..."
      Top             =   3300
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
      Height          =   2025
      Left            =   9480
      TabIndex        =   50
      Top             =   210
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   3572
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
         Top             =   360
         Width           =   1290
      End
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1800
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
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1020
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1800
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
      Index           =   3
      Left            =   1800
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
      Left            =   1800
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
      Index           =   17
      Left            =   1800
      TabIndex        =   15
      Tag             =   "00-Enter Region Code"
      Top             =   3690
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   18
      Left            =   1800
      TabIndex        =   17
      Tag             =   "00-Enter Administered By Code"
      Top             =   4020
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   19
      Left            =   6600
      TabIndex        =   16
      Tag             =   "00-Enter Section Code"
      Top             =   3690
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   4020
      TabIndex        =   10
      Tag             =   "40-Date upto and including this date forward"
      Top             =   2970
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1255
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   9
      Tag             =   "40-Date from and including this date forward"
      Top             =   2970
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1345
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Tag             =   "10-Enter Employee Number"
      Top             =   2660
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Threed.SSFrame Report8 
      Height          =   1035
      Left            =   240
      TabIndex        =   62
      Top             =   6720
      Visible         =   0   'False
      Width           =   9975
      _Version        =   65536
      _ExtentX        =   17595
      _ExtentY        =   1834
      _StockProps     =   14
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   16
         Left            =   5760
         TabIndex        =   29
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECPA"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   15
         Left            =   1650
         TabIndex        =   27
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECPA"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   14
         Left            =   5760
         TabIndex        =   28
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECPA"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   13
         Left            =   1650
         TabIndex        =   26
         Tag             =   "00-Enter Plant Area Code"
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECPA"
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Plant Area"
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
         Left            =   120
         TabIndex        =   63
         Top             =   270
         Width           =   1020
      End
   End
   Begin Threed.SSFrame Report9 
      Height          =   735
      Left            =   240
      TabIndex        =   58
      Top             =   6720
      Visible         =   0   'False
      Width           =   8775
      _Version        =   65536
      _ExtentX        =   15478
      _ExtentY        =   1296
      _StockProps     =   14
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
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         DataSource      =   "Data1"
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
         Index           =   0
         Left            =   2085
         MaxLength       =   1
         TabIndex        =   34
         Tag             =   "00-Code assigned to the shift that the employee works"
         Top             =   240
         Width           =   330
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         DataSource      =   "Data1"
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
         Left            =   3000
         MaxLength       =   1
         TabIndex        =   35
         Tag             =   "00-Code assigned to the shift that the employee works"
         Top             =   240
         Width           =   330
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         DataSource      =   "Data1"
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
         Index           =   2
         Left            =   3960
         MaxLength       =   1
         TabIndex        =   36
         Tag             =   "00-Code assigned to the shift that the employee works"
         Top             =   240
         Width           =   330
      End
      Begin VB.Label lblShift 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Left            =   330
         TabIndex        =   61
         Top             =   270
         Width           =   315
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
         Index           =   2
         Left            =   2520
         TabIndex        =   60
         Top             =   300
         Width           =   270
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
         Index           =   3
         Left            =   3480
         TabIndex        =   59
         Top             =   300
         Width           =   270
      End
   End
   Begin Threed.SSFrame Report7 
      Height          =   975
      Left            =   240
      TabIndex        =   64
      Top             =   6720
      Visible         =   0   'False
      Width           =   8655
      _Version        =   65536
      _ExtentX        =   15266
      _ExtentY        =   1720
      _StockProps     =   14
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   12
         Left            =   5160
         TabIndex        =   80
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECCD"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   11
         Left            =   1440
         TabIndex        =   79
         Top             =   600
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECCD"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   9
         Left            =   1440
         TabIndex        =   77
         Tag             =   "00-Enter Injury Code"
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECCD"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   10
         Left            =   5160
         TabIndex        =   78
         Tag             =   "00-Enter Injury Code"
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECCD"
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Injury Code"
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
         Left            =   120
         TabIndex        =   65
         Top             =   270
         Width           =   795
      End
   End
   Begin Threed.SSFrame Report1 
      Height          =   1030
      Left            =   240
      TabIndex        =   74
      Top             =   6720
      Visible         =   0   'False
      Width           =   9615
      _Version        =   65536
      _ExtentX        =   16960
      _ExtentY        =   1817
      _StockProps     =   14
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   6
         Left            =   5040
         TabIndex        =   31
         Top             =   240
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECBS"
      End
      Begin VB.TextBox txtCode 
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
         Height          =   645
         Index           =   7
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   75
         Tag             =   "00-Enter Body Site Code"
         Top             =   2640
         Width           =   870
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   7
         Left            =   1320
         TabIndex        =   32
         Top             =   600
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECBS"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   8
         Left            =   5040
         TabIndex        =   33
         Top             =   600
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECBS"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   30
         Top             =   240
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECBS"
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Body Site"
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
         Left            =   330
         TabIndex        =   76
         Top             =   270
         Width           =   825
      End
   End
   Begin Threed.SSFrame Report5 
      Height          =   855
      Left            =   240
      TabIndex        =   37
      Top             =   6720
      Visible         =   0   'False
      Width           =   8655
      _Version        =   65536
      _ExtentX        =   15266
      _ExtentY        =   1508
      _StockProps     =   14
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
      Begin Threed.SSOption optWhatSort 
         Height          =   195
         Index           =   1
         Left            =   3600
         TabIndex        =   39
         TabStop         =   0   'False
         Tag             =   "Based on Position Start Date"
         Top             =   360
         Width           =   3165
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Service by Position Start Date"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   29.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optWhatSort 
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Tag             =   "Based on Date of Hire"
         Top             =   360
         Width           =   3165
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Service by Original Hire Date"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   29.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   23
      Left            =   6615
      TabIndex        =   21
      Tag             =   "00-Equipment being used when injured"
      Top             =   4680
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECEQ"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   21
      Left            =   6615
      TabIndex        =   19
      Tag             =   "01-Primary Cause of Injury"
      Top             =   4350
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECCA"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   20
      Left            =   1800
      TabIndex        =   18
      Tag             =   "01-Injury - Code"
      Top             =   4350
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECCD"
   End
   Begin MSMask.MaskEdBox medDOW 
      Height          =   330
      Index           =   0
      Left            =   2100
      TabIndex        =   11
      Tag             =   "10-Day of week 1=Sunday, 2=Monday..."
      Top             =   3300
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   22
      Left            =   1800
      TabIndex        =   20
      Tag             =   "00-Classification of Incident"
      Top             =   4680
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECCL"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   24
      Left            =   1800
      TabIndex        =   23
      Tag             =   "01-Results of Claim - Code"
      Top             =   5420
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
      Index           =   4
      Left            =   120
      TabIndex        =   91
      Top             =   5445
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
      Left            =   120
      TabIndex        =   90
      Tag             =   "00-Hazard Rationg"
      Top             =   5070
      Width           =   1620
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #1"
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
      Left            =   240
      TabIndex        =   88
      Top             =   6270
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
      Left            =   120
      TabIndex        =   87
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label lblType 
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
      Left            =   120
      TabIndex        =   86
      Top             =   4395
      Width           =   1530
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
      Left            =   5400
      TabIndex        =   85
      Top             =   4350
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
      Left            =   5400
      TabIndex        =   84
      Top             =   4680
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
      Left            =   120
      TabIndex        =   83
      Top             =   4725
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
      TabIndex        =   82
      Top             =   3390
      Width           =   1185
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Report includes Terminated Employees"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   81
      Top             =   7920
      Width           =   3315
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
      TabIndex        =   73
      Top             =   2010
      Width           =   630
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
      Left            =   5400
      TabIndex        =   72
      Top             =   3720
      Width           =   540
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
      Left            =   120
      TabIndex        =   70
      Top             =   3720
      Width           =   510
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
      Left            =   120
      TabIndex        =   69
      Top             =   4065
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
      Left            =   120
      TabIndex        =   68
      Top             =   1020
      Width           =   615
   End
   Begin VB.Label ReportType 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Report#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6960
      TabIndex        =   67
      Top             =   0
      Visible         =   0   'False
      Width           =   705
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
      TabIndex        =   66
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
      Left            =   2505
      TabIndex        =   49
      Top             =   3345
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
      Left            =   120
      TabIndex        =   48
      Top             =   3360
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
      Left            =   120
      TabIndex        =   47
      Top             =   2970
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
      Left            =   120
      TabIndex        =   46
      Top             =   2640
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
      Left            =   5400
      TabIndex        =   45
      Top             =   2340
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
      Left            =   120
      TabIndex        =   44
      Top             =   2340
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
      Left            =   120
      TabIndex        =   43
      Top             =   1680
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
      Left            =   120
      TabIndex        =   42
      Top             =   1350
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
      Left            =   120
      TabIndex        =   41
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
      TabIndex        =   40
      Top             =   360
      Width           =   555
   End
End
Attribute VB_Name = "frmRINJURY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Analysis of Incidents by Employee Report Criteria", Me) Then Exit Sub
'    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
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

Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
'    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

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
    Call set_PrintState(True)
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
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

Private Sub Cri_Body()

Dim BodyCri As String
Dim OneSet%, x%

For x% = 5 To 8
    If Len(clpCode(x%).Text) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next x%
If OneSet% = 0 Then Exit Sub

For x% = 5 To 8
    If Len(clpCode(x%).Text) > 0 Then
        BodyCri = BodyCri & "'" & clpCode(x%).Text & "'"
        OneSet% = OneSet% - 1
    End If
    If OneSet% > 0 Then
        BodyCri = BodyCri & " ,"
    End If
Next x%
BodyCri = "{HR_OCC_HEALTH_SAFETY.EC_PBODY} in [" & BodyCri & "]"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & BodyCri
Else
    glbstrSelCri = BodyCri
End If
glbiOneWhere = True

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
     If intIdx% = 5 Then strCd$ = ("HR_OCC_HEALTH_SAFETY.EC_PBODY")
     If intIdx% = 6 Then strCd$ = "HR_OCC_HEALTH_SAFETY.EC_CODE"
     If intIdx% = 7 Then strCd$ = "HR_OCC_HEALTH_SAFETY.EC_AREA"
     If intIdx% = 17 Then strCd$ = "HREMP.ED_REGION"
     If intIdx% = 18 Then strCd$ = "HREMP.ED_ADMINBY"
     If intIdx% = 19 Then strCd$ = "HREMP.ED_SECTION"  'Lucy July 4, 2000
     If intIdx% = 20 Then strCd$ = "HR_OCC_HEALTH_SAFETY.EC_CODE"
     If intIdx% = 21 Then strCd$ = "HR_OCC_HEALTH_SAFETY.EC_CAUSECD"
     If intIdx% = 22 Then strCd$ = "HR_OCC_HEALTH_SAFETY.EC_CLASS"
     If intIdx% = 23 Then strCd$ = "HR_OCC_HEALTH_SAFETY.EC_EQUIP"
     If intIdx% = 24 Then strCd$ = "HR_OCC_HEALTH_SAFETY.EC_WCBRES"
     
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

Private Sub Cri_InJcode()
Dim BodyCri As String
Dim OneSet%, x%

For x% = 9 To 12
    If Len(clpCode(x%).Text) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next x%
If OneSet% = 0 Then Exit Sub

For x% = 9 To 12
    If Len(clpCode(x%).Text) > 0 Then
        BodyCri = BodyCri & "'" & clpCode(x%).Text & "'"
        OneSet% = OneSet% - 1
    End If
    If OneSet% > 0 Then
        BodyCri = BodyCri & " ,"
    End If
Next x%
BodyCri = "{HR_OCC_HEALTH_SAFETY.EC_CODE} in [" & BodyCri & "]"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & BodyCri
Else
    glbstrSelCri = BodyCri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_Plant()
Dim BodyCri As String
Dim OneSet%, x%

For x% = 13 To 16
    If Len(clpCode(x%).Text) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next x%
If OneSet% = 0 Then Exit Sub

For x% = 13 To 16
    If Len(clpCode(x%).Text) > 0 Then
        BodyCri = BodyCri & "'" & clpCode(x%).Text & "'"
        OneSet% = OneSet% - 1
        If OneSet% > 0 Then
            BodyCri = BodyCri & " ,"
        End If
    End If
Next x%
BodyCri = "{HR_OCC_HEALTH_SAFETY.EC_AREA} in [" & BodyCri & "]"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & BodyCri
Else
    glbstrSelCri = BodyCri
End If
glbiOneWhere = True

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
    BodyCri = "{HR_OCC_HEALTH_SAFETY.EC_HAZARD} ='" & cmbHRat.Text & "'"
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
Dim x%, Rptname, xdbf, xdbfw1, xdbfw2

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

xdbf = 6
glbiOneWhere = False
glbstrSelCri = ""

Call glbCri_DeptUN(clpDept.Text)
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_EE
For x% = 0 To 4
    Call Cri_Code(x%)
Next x%
Call Cri_Code(17)
Call Cri_Code(18)
Call Cri_Code(19)
Call Cri_PT
Call Cri_Shift
Call Cri_FTDates
Call Cri_DOW
Call Cri_Comments

'Ticket #12401
Call Cri_Code(20)
Call Cri_Code(21)
Call Cri_Code(22)
Call Cri_Code(23)
Call Cri_Code(24)

If chkDivStart Then
    Call Cri_DivStart
End If

Call Cri_Hazard

' Create query or view in database for display report
If Not glbSQL And Not glbOracle Then
    Call SETWRK
End If

If ReportType.Caption = "1" Then
    Call Cri_Body
    Rptname = "rzinbody.rpt" 'Body Site Report
    Me.vbxCrystal.WindowTitle = "Analysis of Incidents by Body Site Report"
    xdbf = 11 + 4: xdbfw1 = 4: xdbfw2 = 11
End If
If ReportType.Caption = "2" Then
    Rptname = "rzindow.rpt"  'Day of Week Report
    Me.vbxCrystal.WindowTitle = "Analysis of Incidents by Day of Week Report"
    xdbf = 6 + 6: xdbfw1 = 2: xdbfw2 = 6
End If
If ReportType.Caption = "3" Then
    'Employee
    If glbWFC Then
        Rptname = "rzinempw.rpt"
        xdbf = 7: xdbfw1 = 2: xdbfw2 = 7
    Else
        Rptname = "rzinemp.rpt"
        xdbf = 8 + 5: xdbfw1 = 2: xdbfw2 = 7
    End If
    Me.vbxCrystal.WindowTitle = "Analysis of Incidents by Employee Report"
    'xdbf = 7: xdbfw1 = 2: xdbfw2 = 7
End If
If ReportType.Caption = "4" Then
    Rptname = "rzinet.rpt"  'Employee Trend Analysis Report
    Me.vbxCrystal.WindowTitle = "Analysis of Incidents by Employee Trend Report"
    xdbf = 11 + 4: xdbfw1 = 4: xdbfw2 = 11
End If
If ReportType.Caption = "5" Then
    Call Cri_WhatSort
    Rptname = "rzinexp.rpt"
    Me.vbxCrystal.WindowTitle = "Analysis of Incidents by Experience Report"
    xdbf = 11 + 4: xdbfw1 = 4: xdbfw2 = 11
End If
'If Reporttype.Caption = "6" Then RptName = "rzinbody.rpt"
If ReportType.Caption = "7" Then
    Call Cri_InJcode
    Rptname = "rzincode.rpt" 'Injury Code Report
    Me.vbxCrystal.WindowTitle = "Analysis of Incidents by Injury Code Report"
    xdbf = 11 + 4: xdbfw1 = 4: xdbfw2 = 11
End If
If ReportType.Caption = "8" Then
    Call Cri_Plant
    Rptname = "rzinarea.rpt" 'Plant Area Report
    Me.vbxCrystal.WindowTitle = "Analysis of Incidents by Plant Area Report"
    xdbf = 10 + 4: xdbfw1 = 4: xdbfw2 = 10
End If
If ReportType.Caption = "9" Then
    Call Cri_Shift
    Rptname = "rzinshft.rpt"
    Me.vbxCrystal.WindowTitle = "Analysis of Incidents by Shift Code Report"
    xdbf = 10 + 4: xdbfw1 = 4: xdbfw2 = 10
End If
'MsgBox "CRI : " & glbstrSelCri
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & Rptname

' set to sorting/grouping criteria
x% = Cri_Sorts()   ' returns number of sections formated
  
'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    If Not glbSQL And Not glbOracle Then
        'Franks Feb 27,2003 #3703 duplicate records
        'glbstrSelCri = glbstrSelCri & " AND {HR_OCC_HEALTH_SAFETY.WRKEMP}='" & glbUserID & "'"
        glbstrSelCri = glbstrSelCri & " AND {HR_OCC_HEALTH_SAFETY.WRKEMP}='" & glbUserID & "'" & " AND {HREMP.ED_WRKEMP}='" & glbUserID & "'"
        'Franks Feb 27,2003 #3703 duplicate records
    End If
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If


If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For x% = 0 To xdbf
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        If x% = xdbfw1 Or x% = xdbfw2 Then
        Me.vbxCrystal.DataFiles(x%) = glbIHRDBW
        End If
    Next x%
End If

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

Private Sub SETWRK()
On Error GoTo EH

Dim rsHS As New ADODB.Recordset, rsTmp As New ADODB.Recordset
Dim rsCode As New ADODB.Recordset, rsBODY As New ADODB.Recordset
Dim SQLQ, xNum, xRecNum
Dim xFieldList
xFieldList = Get_Fields(gdbAdoIhr001W, "HREMP_HS", "KEY_EMPNBR,ED_WRKEMP,JB_GRPCD_TABL,JB_GRPCD,ED_ID,ED_HOMELINE_TABL,")

'On Error GoTo AttWrkError
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3
gdbAdoIhr001.CommandTimeout = 600
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).FloodPercent = 15

gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute "DELETE FROM HREMP_HS WHERE ED_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001W.CommitTrans

gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute "DELETE FROM HR_HS WHERE WRKEMP='" & glbUserID & "'"
gdbAdoIhr001W.CommitTrans

Call Pause(1)

MDIMain.panHelp(0).FloodPercent = 30
'for active employees
SQLQ = "INSERT INTO HREMP_HS (" & xFieldList & ",KEY_EMPNBR,ED_WRKEMP)"
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & "SELECT " & xFieldList
SQLQ = SQLQ & ",'1_' + TRIM(STR(HREMP.ED_EMPNBR) ) AS KEY_EMPNBR "
SQLQ = SQLQ & ",'" & glbUserID & "' AS ED_WRKEMP "
SQLQ = SQLQ & " FROM HREMP "

'Hemu - 09/29/2003 Begin - Taking too long to generate the report, esp. when Emp#
'                          entered
If Len(elpEEID.Text) > 0 Then
    SQLQ = SQLQ & "WHERE HREMP.ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
End If
'Hemu

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans
MDIMain.panHelp(0).FloodPercent = 50

'for terminated employees
SQLQ = "INSERT INTO HREMP_HS (" & xFieldList & ",KEY_EMPNBR,ED_WRKEMP)"
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & "SELECT " & xFieldList
SQLQ = SQLQ & ",'0_' +  TRIM(STR(Term_HREMP.TERM_SEQ) ) AS KEY_EMPNBR "
SQLQ = SQLQ & ",'" & glbUserID & "' AS ED_WRKEMP "
SQLQ = SQLQ & " FROM Term_HREMP "

'Hemu - 09/29/2003 Begin - Taking too long to generate the report, esp. when Emp#
'                          entered
If Len(elpEEID.Text) > 0 Then
    SQLQ = SQLQ & "WHERE Term_HREMP.ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
End If
'Hemu

gdbAdoIhr001X.BeginTrans
gdbAdoIhr001X.Execute SQLQ
gdbAdoIhr001X.CommitTrans

'Health & Safety
MDIMain.panHelp(0).FloodPercent = 70
'for active employees
SQLQ = "INSERT INTO HR_HS "
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & "SELECT " & Get_Fields(gdbAdoIhr001W, "HR_HS", "KEY_EMPNBR,WRKEMP,EC_AREADESC,EC_CODEDESC,EC_BODYDESC,EC_BODY2DESC,EC_EQUIDESC,EC_CLASSDESC,EC_CAUSEDESC")
SQLQ = SQLQ & ",'1_' + TRIM(STR(HR_OCC_HEALTH_SAFETY.EC_EMPNBR) ) AS KEY_EMPNBR "
SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
SQLQ = SQLQ & " FROM HR_OCC_HEALTH_SAFETY "

'Hemu - 09/29/2003 Begin - Taking too long to generate the report, esp. when Emp#
'                          entered
If Len(elpEEID.Text) > 0 Then
    SQLQ = SQLQ & "WHERE HR_OCC_HEALTH_SAFETY.EC_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
End If
'Hemu

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans
MDIMain.panHelp(0).FloodPercent = 80

'for terminated employees
SQLQ = "INSERT INTO HR_HS "
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & "SELECT " & Get_Fields(gdbAdoIhr001W, "HR_HS", "KEY_EMPNBR,WRKEMP,EC_AREADESC,EC_CODEDESC,EC_BODYDESC,EC_BODY2DESC,EC_EQUIDESC,EC_CLASSDESC,EC_CAUSEDESC")
SQLQ = SQLQ & ",'0_' +  TRIM(STR(Term_HR_OCC_HEALTH_SAFETY.TERM_SEQ) ) AS KEY_EMPNBR "
SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
SQLQ = SQLQ & " FROM Term_HR_OCC_HEALTH_SAFETY "

'Hemu - 09/29/2003 Begin - Taking too long to generate the report, esp. when Emp#
'                          entered
If Len(elpEEID.Text) > 0 Then
    SQLQ = SQLQ & "WHERE Term_HR_OCC_HEALTH_SAFETY.EC_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
End If
'Hemu

gdbAdoIhr001X.BeginTrans
gdbAdoIhr001X.Execute SQLQ
gdbAdoIhr001X.CommitTrans

MDIMain.panHelp(0).FloodPercent = 85
'If chkReport(0) Or optSum(2) Then
    MDIMain.panHelp(0).FloodPercent = 90
    Call Pause(2)
    MDIMain.panHelp(0).FloodPercent = 100
    Call Pause(2)
    
    'Hemu - 09/29/2003 Begin - Taking too long to generate the report, esp. when Emp#
    '                          entered
    If Len(elpEEID.Text) > 0 Then
        SQLQ = "SELECT * FROM HR_HS WHERE EC_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
    Else
        SQLQ = "SELECT * FROM HR_HS"
    End If
    rsHS.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    'rsHS.Open "HR_HS", gdbAdoIhr001W, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    'Hemu
    
    If Not rsHS.EOF Then
        MDIMain.panHelp(0).FloodPercent = 0
        rsHS.MoveLast
        rsHS.MoveFirst
        xRecNum = rsHS.RecordCount
        xNum = 0
    End If
    
    Do While Not rsHS.EOF
        MDIMain.panHelp(0).FloodPercent = xNum / (xRecNum + 1) * 100
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ECPA' "
        SQLQ = SQLQ & "AND TB_KEY = '" & rsHS("EC_AREA") & "' "
        rsTmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTmp.EOF Then
            rsHS("EC_AREADESC") = rsTmp("TB_DESC")
        End If
        rsTmp.Close
        'CODE
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ECCD' "
        SQLQ = SQLQ & "AND TB_KEY = '" & rsHS("EC_CODE") & "' "
        rsCode.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsCode.EOF Then
            rsHS("EC_CODEDESC") = rsCode("TB_DESC")
        End If
        rsCode.Close
        'BODY 1
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ECBS' "
        SQLQ = SQLQ & "AND TB_KEY = '" & rsHS("EC_PBODY") & "' "
        rsBODY.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsBODY.EOF Then
            rsHS("EC_BODYDESC") = rsBODY("TB_DESC")
        End If
        rsBODY.Close
        'BODY 2
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ECBS' "
        SQLQ = SQLQ & "AND TB_KEY = '" & rsHS("EC_SBODY") & "' "
        rsTmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTmp.EOF Then
            rsHS("EC_BODY2DESC") = rsTmp("TB_DESC")
        End If
        rsTmp.Close
        'EQUITMENT
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ECEQ' "
        SQLQ = SQLQ & "AND TB_KEY = '" & rsHS("EC_EQUIP") & "' "
        rsTmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTmp.EOF Then
            rsHS("EC_EQUIDESC") = rsTmp("TB_DESC")
        End If
        rsTmp.Close
        'CLASS
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ECCL' "
        SQLQ = SQLQ & "AND TB_KEY = '" & rsHS("EC_CLASS") & "' "
        rsTmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTmp.EOF Then
            rsHS("EC_CLASSDESC") = rsTmp("TB_DESC")
        End If
        rsTmp.Close
        'CAUSE
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ECCA' "
        SQLQ = SQLQ & "AND TB_KEY = '" & rsHS("EC_CAUSECD") & "' "
        rsTmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTmp.EOF Then
            rsHS("EC_CAUSEDESC") = rsTmp("TB_DESC")
        End If
        rsTmp.Close

        rsHS.Update
        rsHS.MoveNext
        xNum = xNum + 1
    Loop
    rsHS.Close
    Call Pause(1)
'End If

MDIMain.panHelp(0).FloodPercent = 100
Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
Exit Sub

EH:
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
End Sub

Private Sub Cri_Shift()
Dim ShftCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim OneSet%, x%

For x% = 0 To 3
    If Len(txtShift(x%)) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next x%
If OneSet% = 0 Then Exit Sub

For x% = 0 To 3
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

Private Sub Cri_WhatSort()
Dim SrtCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level

If optWhatSort(0) Then
   SrtCri = "WhatSort = 0"
Else
   SrtCri = "WhatSort = 1"
End If
Me.vbxCrystal.Formulas(0) = SrtCri

End Sub

Private Sub Cri_Comments()
    If chkComments.Visible = True Then
        If chkComments.Value = vbChecked Then
            Me.vbxCrystal.Formulas(2) = "chkComments=true"
        Else
            Me.vbxCrystal.Formulas(2) = "chkComments=false"
        End If
    End If
End Sub

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

For x% = 0 To 19
    If Not clpCode(x).ListChecker Then Exit Function
Next x%

'If lblEEName.Caption = "Enter Valid Employee #" Then
'    MsgBox "If Employee Entered - they must exist"
'    txtEEID.SetFocus
'    Exit Function
'End If
'Hemu - 05/13/2003 Begin - From Date and To Date
If (dlpDateRange(0).Text <> "" And IsDate(dlpDateRange(0).Text)) And (dlpDateRange(1).Text <> "" And IsDate(dlpDateRange(1).Text)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1).Text) < 0 Then
        MsgBox "To Date cannot be prior to From Date"
        dlpDateRange(1).SetFocus
        Exit Function
    End If
End If
'Hemu - 05/13/2003 End

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

Private Sub Form_Activate()
Call SET_UP_MODE
Report1.Visible = False
Report5.Visible = False
Report7.Visible = False
Report8.Visible = False
Report9.Visible = False
lblHSShift.Visible = True
txtShift(3).Visible = True
If ReportType.Caption = "1" Then Report1.Visible = True
If ReportType.Caption = "5" Then Report5.Visible = True
If ReportType.Caption = "7" Then Report7.Visible = True
If ReportType.Caption = "8" Then Report8.Visible = True
If ReportType.Caption = "9" Then
    Report9.Visible = True
    lblHSShift.Visible = False
    txtShift(3).Visible = False
End If

chkDivStart.Visible = True
chkDivStart.Caption = lStr(chkDivStart.Caption)

End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = "FRMRINJURY"

Screen.MousePointer = HOURGLASS
Call comGrpLoad
Call setRptCaption(Me)
If glbLinamar Then
    clpCode(17).MaxLength = 8
    'Ticket #14573
    lblClassification.Caption = "Type of Event"
    lblCause.Caption = "Root Cause"
End If
If glbCompSerial = "S/N - 2227W" Then clpCode(17).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
If glbCompSerial = "S/N - 2387W" Then  'Bird Packaging Limited 'Ticket #13636
    lblType.Caption = "Accident Type"
    lblCause.Caption = "Injury Class #1"
End If

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

Private Sub optWhatSort_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)     'laura jan 06, 1997
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub comGrpLoad()
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
    'comGroup(1).AddItem "Employee Name"
    'comGroup(1).ListIndex = 0

End Sub

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
'imbeded in report

Cri_Sorts = 0
'first set primary grouping
Y% = 0
grpField$ = getEGroup(comGroup(0).Text)

If comGroup(0) = "(none)" Then Exit Function
    
Y% = x% + 1
dscGroup$ = comGroup(x%).Text
dscGroup$ = "descGroup" & CStr(Y%) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(Y%) = dscGroup$

grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(x%) = grpCond$

strSFormat$ = "GH1;T;T;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1
strSFormat$ = "GF1;T;X;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1

'If ReportType.Caption = "3" Then
'    GrpIdx% = comGroup(1).ListIndex
'    Select Case GrpIdx%
'        Case 0: grpField$ = "{@EFullName}"
'    End Select
'    grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
'    Me.vbxCrystal.GroupCondition(1) = grpCond$
'End If
Cri_Sorts = z% ' next section number to format

End Function
