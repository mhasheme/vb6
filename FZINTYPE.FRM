VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRINTYPE 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Analysis of Incidents by Incident Types Report"
   ClientHeight    =   8175
   ClientLeft      =   75
   ClientTop       =   1410
   ClientWidth     =   13335
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
   ScaleHeight     =   8175
   ScaleWidth      =   13335
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtHRat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "EC_HAZARD"
      Height          =   315
      Left            =   3060
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   4980
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox cmbHRat 
      Height          =   315
      Left            =   2130
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Tag             =   "00-Hazard Rating - Choose A,B, or C"
      Top             =   4980
      Width           =   645
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
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Tag             =   "First Level of grouping records"
      Top             =   6810
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
      Left            =   5640
      TabIndex        =   25
      Top             =   5010
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   5580
      MaxLength       =   4
      TabIndex        =   16
      Tag             =   "00-Shift from H&S Incidents"
      Top             =   3330
      Width           =   450
   End
   Begin VB.CheckBox chkVDesc 
      Caption         =   "View Description"
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
      Left            =   7260
      TabIndex        =   10
      Top             =   2340
      Width           =   1845
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   7
      Left            =   1800
      TabIndex        =   11
      Tag             =   "00-Enter Status Code"
      Top             =   2670
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
      Tag             =   "00-Incident Type Code"
      Top             =   2340
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "ECTY"
   End
   Begin MSMask.MaskEdBox medDOW 
      Height          =   285
      Index           =   0
      Left            =   2130
      TabIndex        =   14
      Tag             =   "10-Day of week 1=Sunday, 2=Monday..."
      Top             =   3330
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   503
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
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   15
      Tag             =   "10-Day of week 1=Sunday, 2=Monday..."
      Top             =   3330
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   503
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
   Begin Threed.SSFrame FrmInType 
      Height          =   1725
      Left            =   9600
      TabIndex        =   41
      Top             =   300
      Visible         =   0   'False
      Width           =   3195
      _Version        =   65536
      _ExtentX        =   5636
      _ExtentY        =   3043
      _StockProps     =   14
      Caption         =   "Incident Type Description"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowColor     =   1
      ShadowStyle     =   1
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   495
         TabIndex        =   47
         Top             =   1350
         Width           =   45
      End
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   495
         TabIndex        =   46
         Top             =   1020
         Width           =   45
      End
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   495
         TabIndex        =   31
         Top             =   690
         Width           =   45
      End
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   495
         TabIndex        =   30
         Top             =   360
         Width           =   45
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "2."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   690
         Width           =   180
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "4."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   1350
         Width           =   180
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "3."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   1020
         Width           =   180
      End
   End
   Begin Threed.SSFrame fraSummaries 
      Height          =   1275
      Left            =   5160
      TabIndex        =   48
      Top             =   5880
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   2249
      _StockProps     =   14
      Caption         =   "Summaries By"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Begin Threed.SSOption optSum 
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   51
         TabStop         =   0   'False
         Tag             =   "Choose type of summary"
         Top             =   875
         Width           =   2000
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "          Shift"
         ForeColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optSum 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   50
         TabStop         =   0   'False
         Tag             =   "Choose type of summary"
         Top             =   575
         Width           =   2000
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "          Plant Area"
         ForeColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optSum 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   49
         Tag             =   "Choose type of summary"
         Top             =   240
         Width           =   1995
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "          Day of Week"
         ForeColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSOption chkReport 
      Height          =   255
      Index           =   0
      Left            =   1875
      TabIndex        =   27
      Tag             =   "Choose type of report"
      Top             =   5880
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "   Detail "
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
   Begin Threed.SSOption chkReport 
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "Choose type of report"
      Top             =   5880
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "   Summary "
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
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3360
      TabIndex        =   13
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3000
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   12
      Tag             =   "40-Date from and including this date forward"
      Top             =   3000
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
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
      Index           =   6
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
      Index           =   5
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
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Tag             =   "00-Incident Type Code"
      Top             =   2340
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "ECTY"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   4500
      TabIndex        =   8
      Tag             =   "00-Incident Type Code"
      Top             =   2340
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "ECTY"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   5880
      TabIndex        =   9
      Tag             =   "00-Incident Type Code"
      Top             =   2340
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "ECTY"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   9
      Left            =   1800
      TabIndex        =   18
      Tag             =   "00-Enter Administered By Code"
      Top             =   3990
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   10
      Left            =   6735
      TabIndex        =   21
      Tag             =   "00-Enter Section Code"
      Top             =   3960
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   8
      Left            =   1800
      TabIndex        =   17
      Tag             =   "00-Enter Region Code"
      Top             =   3660
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   8520
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   13
      Left            =   1800
      TabIndex        =   20
      Tag             =   "00-Classification of Incident"
      Top             =   4650
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECCL"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   14
      Left            =   6735
      TabIndex        =   23
      Tag             =   "00-Equipment being used when injured"
      Top             =   4650
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECEQ"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   12
      Left            =   6735
      TabIndex        =   22
      Tag             =   "01-Primary Cause of Injury"
      Top             =   4320
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECCA"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   11
      Left            =   1800
      TabIndex        =   19
      Tag             =   "01-Injury - Code"
      Top             =   4320
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECCD"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   15
      Left            =   1800
      TabIndex        =   26
      Tag             =   "01-Results of Claim - Code"
      Top             =   5360
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
      Left            =   120
      TabIndex        =   69
      Top             =   5405
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
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   67
      Tag             =   "00-Hazard Rationg"
      Top             =   5030
      Width           =   1620
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
      TabIndex        =   66
      Top             =   6480
      Width           =   1575
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
      Left            =   120
      TabIndex        =   65
      Top             =   6870
      Width           =   885
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
      TabIndex        =   64
      Top             =   4695
      Width           =   1155
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
      Left            =   5520
      TabIndex        =   63
      Top             =   4650
      Width           =   900
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
      Left            =   5520
      TabIndex        =   62
      Top             =   4320
      Width           =   1215
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
      TabIndex        =   61
      Top             =   4365
      Width           =   1410
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
      Left            =   4260
      TabIndex        =   60
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
      TabIndex        =   59
      Top             =   7560
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
      TabIndex        =   58
      Top             =   2010
      Width           =   630
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
      Left            =   5520
      TabIndex        =   57
      Top             =   4020
      Width           =   1260
   End
   Begin VB.Label lblAdmin 
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
      Index           =   1
      Left            =   120
      TabIndex        =   56
      Top             =   4035
      Width           =   1125
   End
   Begin VB.Label Label4 
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
      Index           =   1
      Left            =   120
      TabIndex        =   55
      Top             =   3690
      Width           =   510
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
      TabIndex        =   54
      Top             =   1020
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
      Left            =   0
      TabIndex        =   53
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Report"
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
      Left            =   120
      TabIndex        =   52
      Top             =   5910
      Width           =   1065
   End
   Begin VB.Label lblDOW 
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
      Left            =   2520
      TabIndex        =   40
      Top             =   3360
      Width           =   270
   End
   Begin VB.Label lblDOW 
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
      TabIndex        =   39
      Top             =   3330
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
      TabIndex        =   38
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   120
      TabIndex        =   37
      Top             =   2670
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
      TabIndex        =   36
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
      TabIndex        =   35
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
      TabIndex        =   34
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
      TabIndex        =   33
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
      TabIndex        =   32
      Top             =   360
      Width           =   555
   End
End
Attribute VB_Name = "frmRINTYPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkReport_Click(Index As Integer, Value As Integer)
Dim x%

If Index = 1 Then
    fraSummaries.Enabled = True
    optSum(1).Value = True
    For x% = 1 To 3
        optSum(x%).ForeColor = &H0&
    Next x%
Else
    fraSummaries.Enabled = False
    For x% = 1 To 3
        optSum(x%).Value = False
        optSum(x%).ForeColor = &H808080
    Next x%
End If

End Sub

Private Sub chkReport_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkVDesc_Click()
FrmInType.Visible = chkVDesc
End Sub

Private Sub clpCode_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
On Error Resume Next: lblCodeDesc(Index).Caption = clpCode(Index).Caption
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Analysis of Incidents by Incident Types Report Criteria", Me) Then Exit Sub
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Call set_PrintState(False)
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
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Call set_PrintState(False)
    
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

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    Select Case intIdx%
    Case 0: strCd$ = "HREMP.ED_LOC"
    Case 5: strCd$ = "HREMP.ED_ORG"
    Case 6: strCd$ = "HREMP.ED_EMP"
    Case 7: strCd$ = "HR_OCC_HEALTH_SAFETY.EC_RATEGRP"
    Case 8: strCd$ = "HREMP.ED_REGION"
    Case 9: strCd$ = "HREMP.ED_ADMINBY"
    Case 10: strCd$ = "HREMP.ED_SECTION"
    Case 11: strCd$ = "HR_OCC_HEALTH_SAFETY.EC_CODE"
    Case 12: strCd$ = "HR_OCC_HEALTH_SAFETY.EC_CAUSECD"
    Case 13: strCd$ = "HR_OCC_HEALTH_SAFETY.EC_CLASS"
    Case 14: strCd$ = "HR_OCC_HEALTH_SAFETY.EC_EQUIP"
    Case 15: strCd$ = "HR_OCC_HEALTH_SAFETY.EC_WCBRES"
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

Private Sub Cri_Incident()
Dim EECri As String, OneSet%, x%

OneSet% = False

For x% = 1 To 4
    If Len(clpCode(x%).Text) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next x%

If OneSet% = 0 Then Exit Sub

For x% = 1 To 4
    If Len(clpCode(x%).Text) > 0 Then
        EECri = EECri & "'" & clpCode(x%).Text & "' "
        OneSet% = OneSet% - 1
        If OneSet% > 0 Then
            EECri = EECri & ","
        End If
    End If
    'Franks 05/01/2003 Ticket# 3765
    'Error in Formula, move the following code up
    'If OneSet% > 0 Then
    '    EECri = EECri & ","
    'End If
Next x%
EECri$ = "{HR_OCC_HEALTH_SAFETY.EC_TYPE} in [" & EECri & "]"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

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

Private Sub Cri_Shift()
Dim ShftCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim OneSet%, x%

For x% = 0 To 0
    If Len(txtShift(x%)) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next x%
If OneSet% = 0 Then Exit Sub

For x% = 0 To 0
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

Private Sub SETWRK()
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
gdbAdoIhr001X.BeginTrans
gdbAdoIhr001X.Execute SQLQ
gdbAdoIhr001X.CommitTrans

MDIMain.panHelp(0).FloodPercent = 85
If chkReport(0) Or optSum(2) Then
    MDIMain.panHelp(0).FloodPercent = 90
    Call Pause(2)
    MDIMain.panHelp(0).FloodPercent = 100
    Call Pause(3)
    rsHS.Open "HR_HS", gdbAdoIhr001W, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
        rsHS.Update
        rsHS.MoveNext
        xNum = xNum + 1
    Loop
    rsHS.Close
    Call Pause(1)
End If

MDIMain.panHelp(0).FloodPercent = 100
Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
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
Dim x%
Dim xdbf, xdbfw1, xdbfw2

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

' Create query or view in database for display report
If Not glbSQL And Not glbOracle Then
    Call SETWRK
End If

glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_Incident
Call Cri_Code(0)  'Jaddy jun 21
Call Cri_Code(5)
Call Cri_Code(6)
Call Cri_Code(7)
Call Cri_Code(8)
Call Cri_Code(9)
Call Cri_Code(10)
Call Cri_PT
Call Cri_FTDates
Call Cri_DOW
Call Cri_Shift

Call Cri_Code(11)
Call Cri_Code(12)
Call Cri_Code(13)
Call Cri_Code(14)
Call Cri_Code(15)

Call Cri_Hazard

If chkDivStart Then
    Call Cri_DivStart
End If

'report name
If chkReport(0).Value = True Then
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzintype.rpt"
Else
    If optSum(1).Value = True Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzintped.rpt"
    Else
        If optSum(2).Value = True Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzintpea.rpt"
        Else
            If optSum(3).Value = True Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzintpes.rpt"
            End If
        End If
    End If
End If

'set to sorting/grouping criteria
x% = Cri_Sorts()   ' returns number of sections formated

'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    If Not glbSQL And Not glbOracle Then
        glbstrSelCri = glbstrSelCri & " AND {HR_OCC_HEALTH_SAFETY.WRKEMP}='" & glbUserID & "'"
    End If
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    ''DBEngine.RegisterDatabase "IHR001", "Microsoft Access Driver (*.mdb)", True, "DBQ=" & glbIHRDB & vbCr
    ''Me.vbxCrystal.Connect = "DSN=IHR001;PWD=petman;DSQ=" & glbIHRDB
    Me.vbxCrystal.Connect = "PWD=petman;"
    xdbf = 15: xdbfw1 = 2: xdbfw2 = 9
    For x% = 0 To xdbf
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        If x% = xdbfw1 Or x% = xdbfw2 Then
        Me.vbxCrystal.DataFiles(x%) = glbIHRDBW
        End If
    Next x%
End If

' window title if appropriate
Me.vbxCrystal.WindowTitle = "Analysis of Incidents by Incident Type Report"

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

For x% = 0 To 10       'laura jan 09, 1998
    If Not clpCode(x).ListChecker Then Exit Function
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
 'Hemu - 05/14/2003 Begin - From Date and To Date
If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then
        MsgBox "To Date cannot be prior to From Date"
        dlpDateRange(1).SetFocus
        Exit Function
    End If
End If

'Hemu - 05/14/2003 End
CriCheck = True

End Function

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
chkDivStart.Visible = True
chkDivStart.Caption = lStr(chkDivStart.Caption)
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = "FRMRINTYPE"

Screen.MousePointer = HOURGLASS
Call comGrpLoad
Call setRptCaption(Me)

If glbLinamar Then
    clpCode(8).MaxLength = 8
    'Ticket #14573
    lblClassification.Caption = "Type of Event"
    lblCause.Caption = "Root Cause"
End If
If glbCompSerial = "S/N - 2227W" Then clpCode(8).MaxLength = 6
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

Private Sub optSum_GotFocus(Index As Integer)
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
Me.vbxCrystal.Formulas(x%) = dscGroup$

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
