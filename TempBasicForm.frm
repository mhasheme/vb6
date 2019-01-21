VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12210
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   ScaleHeight     =   12210
   ScaleWidth      =   12900
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frOrganizational 
      BorderStyle     =   0  'None
      Caption         =   "Organizational"
      Height          =   3375
      Left            =   360
      TabIndex        =   91
      Top             =   5520
      Visible         =   0   'False
      Width           =   12200
      Begin VB.Frame frmWFCDIV 
         Height          =   330
         Left            =   6720
         TabIndex        =   93
         Top             =   2640
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox txtDouDiv 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   320
            MaxLength       =   4
            TabIndex        =   94
            Tag             =   "00-Bonus Reporting #"
            Top             =   0
            Width           =   870
         End
         Begin VB.Label lblDouDivDesc 
            Caption         =   "Unassigned"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1320
            TabIndex        =   95
            Top             =   0
            Width           =   2415
         End
         Begin VB.Image imgIDiv 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   0
            Picture         =   "TempBasicForm.frx":0000
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.TextBox txtDeptBonusCtr 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "ED_BONUSDEPT"
         Height          =   285
         Left            =   7530
         MaxLength       =   6
         TabIndex        =   92
         Tag             =   "00-Bonus Reporting #"
         Top             =   580
         Visible         =   0   'False
         Width           =   900
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         DataField       =   "ED_DIV"
         Height          =   285
         Left            =   1515
         TabIndex        =   96
         Tag             =   "00-Division"
         Top             =   920
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ED_ADMINBY"
         Height          =   285
         Index           =   3
         Left            =   1515
         TabIndex        =   97
         Tag             =   "00-Administered By"
         Top             =   1600
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         DataField       =   "ED_DEPTNO"
         Height          =   285
         Left            =   1515
         TabIndex        =   98
         Tag             =   "00-Department"
         Top             =   240
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   7
         LookupType      =   2
      End
      Begin INFOHR_Controls.CodeLookup clpGLNum 
         DataField       =   "ED_GLNO"
         Height          =   285
         Left            =   1515
         TabIndex        =   99
         Tag             =   "00-General Ledger - Code"
         Top             =   580
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   3
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   7215
         TabIndex        =   100
         Tag             =   "00-Region"
         Top             =   1260
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   7215
         TabIndex        =   101
         Tag             =   "00-Section"
         Top             =   1600
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.DateLookup dlpDivEDate 
         DataField       =   "ED_DIVEDATE"
         Height          =   285
         Left            =   7215
         TabIndex        =   102
         Tag             =   "40-Division Effective Date"
         Top             =   920
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDeptEDate 
         DataField       =   "ED_DEPTEDATE"
         Height          =   285
         Left            =   7215
         TabIndex        =   103
         Tag             =   "40-Department Effective Date"
         Top             =   240
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpHOME 
         Height          =   285
         Index           =   1
         Left            =   1515
         TabIndex        =   104
         Tag             =   "00-Home Operation Number"
         Top             =   1940
         Visible         =   0   'False
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "HMOP"
         MaxLength       =   12
      End
      Begin INFOHR_Controls.CodeLookup clpHOME 
         Height          =   285
         Index           =   2
         Left            =   1515
         TabIndex        =   105
         Tag             =   "00-Home Line"
         Top             =   2280
         Visible         =   0   'False
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "HMLN"
         MaxLength       =   12
      End
      Begin INFOHR_Controls.CodeLookup clpHOME 
         DataField       =   "ED_HOMEWRKCNT"
         Height          =   285
         Index           =   3
         Left            =   1515
         TabIndex        =   106
         Tag             =   "00-Home Work Center"
         Top             =   2620
         Visible         =   0   'False
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "HMWC"
         MaxLength       =   12
      End
      Begin INFOHR_Controls.CodeLookup clpHOME 
         DataField       =   "ED_HOMESHIFT"
         Height          =   285
         Index           =   4
         Left            =   1515
         TabIndex        =   107
         Tag             =   "00-Home Shift"
         Top             =   2960
         Visible         =   0   'False
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "HMSF"
         MaxLength       =   5
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ED_LOC"
         Height          =   285
         Index           =   1
         Left            =   1515
         TabIndex        =   108
         Tag             =   "00-Location - Code"
         Top             =   1260
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   5
         Left            =   10200
         TabIndex        =   109
         Tag             =   "00-Sub Department"
         Top             =   580
         Visible         =   0   'False
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "SUDE"
         MaxLength       =   20
      End
      Begin VB.Label lblRptNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Bonus Reporting No"
         Height          =   195
         Left            =   5580
         TabIndex        =   126
         Top             =   625
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "G/L #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   125
         Top             =   625
         Width           =   435
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
         TabIndex        =   124
         Top             =   285
         Width           =   990
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Administered By"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   25
         Left            =   120
         TabIndex        =   123
         Top             =   1645
         Width           =   1125
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   24
         Left            =   5265
         TabIndex        =   122
         Top             =   1305
         Width           =   1890
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   23
         Left            =   120
         TabIndex        =   121
         Top             =   1305
         Width           =   615
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   120
         Top             =   965
         Width           =   555
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Section"
         Height          =   195
         Index           =   26
         Left            =   6375
         TabIndex        =   119
         Top             =   1645
         Width           =   780
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Home Operation#"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   27
         Left            =   120
         TabIndex        =   118
         Top             =   1985
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Home Line"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   29
         Left            =   120
         TabIndex        =   117
         Top             =   2325
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Home Shift"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   28
         Left            =   120
         TabIndex        =   116
         Top             =   3005
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Home Work Center"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   30
         Left            =   120
         TabIndex        =   115
         Top             =   2665
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lblDivStart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Division Effective"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5325
         TabIndex        =   114
         Top             =   965
         Width           =   1830
      End
      Begin VB.Label lblDeptStart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Department Effective"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5655
         TabIndex        =   113
         Top             =   285
         Width           =   1500
      End
      Begin VB.Label lblPayroll 
         AutoSize        =   -1  'True
         Caption         =   "Payroll"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6780
         TabIndex        =   112
         Top             =   2280
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblPerson 
         AutoSize        =   -1  'True
         Caption         =   "Personnel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6780
         TabIndex        =   111
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblDeptBonusDesc 
         Caption         =   "Unassigned"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   8520
         TabIndex        =   110
         Top             =   595
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   7200
         Picture         =   "TempBasicForm.frx":014A
         Top             =   602
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Frame frMiscellaneous 
      BorderStyle     =   0  'None
      Caption         =   "Miscellaneous"
      Height          =   2415
      Left            =   360
      TabIndex        =   67
      Top             =   8760
      Visible         =   0   'False
      Width           =   12200
      Begin VB.ComboBox ComSmoker 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Tag             =   "00-Smoker Yes/No"
         Top             =   2010
         Width           =   855
      End
      Begin VB.TextBox txtSmoker 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_SMOKER"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2700
         MaxLength       =   2
         TabIndex        =   68
         Text            =   "Text14"
         Top             =   2010
         Visible         =   0   'False
         Width           =   450
      End
      Begin MSMask.MaskEdBox medCellPhone 
         DataField       =   "ED_CELLPHONE"
         Height          =   285
         Left            =   1830
         TabIndex        =   70
         Tag             =   "10-Cellular Telephone Number"
         Top             =   240
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(###) ###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPageNbr 
         DataField       =   "ED_PAGENBR"
         Height          =   285
         Left            =   7530
         TabIndex        =   71
         Tag             =   "10-Pager Number"
         Top             =   240
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(###) ###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medDRIVERLIC 
         DataField       =   "ED_DRIVERLIC"
         Height          =   285
         Left            =   1830
         TabIndex        =   72
         Tag             =   "00-Driver License Number"
         Top             =   594
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   30
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
      Begin MSMask.MaskEdBox medLICPLATE1 
         DataField       =   "ED_LICPLATE1"
         Height          =   285
         Left            =   1830
         TabIndex        =   73
         Tag             =   "00-License Plate #1"
         Top             =   1302
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   30
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
      Begin MSMask.MaskEdBox medLICPLATE2 
         DataField       =   "ED_LICPLATE2"
         Height          =   285
         Left            =   7530
         TabIndex        =   74
         Tag             =   "00-License Plate #2"
         Top             =   1302
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   30
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
      Begin MSMask.MaskEdBox medLOCKER 
         DataField       =   "ED_LOCKER"
         Height          =   285
         Left            =   1830
         TabIndex        =   75
         Tag             =   "00-Locker #"
         Top             =   1656
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   6
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
      Begin MSMask.MaskEdBox medCOMBINATION 
         DataField       =   "ED_COMBINATION"
         Height          =   285
         Left            =   7530
         TabIndex        =   76
         Tag             =   "00-Combination"
         Top             =   1656
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   8
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
      Begin MSMask.MaskEdBox medTYPEVEHICLE 
         DataField       =   "ED_TYPEVEHICLE"
         Height          =   285
         Left            =   7530
         TabIndex        =   77
         Tag             =   "00-Type of Vehicle"
         Top             =   594
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   30
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
      Begin MSMask.MaskEdBox medPARKPERMIT1 
         DataField       =   "ED_PARKPERMIT1"
         Height          =   285
         Left            =   1830
         TabIndex        =   78
         Tag             =   "00-Parking Permit #1"
         Top             =   948
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   30
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
      Begin MSMask.MaskEdBox medPARKPERMIT2 
         DataField       =   "ED_PARKPERMIT2"
         Height          =   285
         Left            =   7530
         TabIndex        =   79
         Tag             =   "00-Parking Permit #2"
         Top             =   948
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   30
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
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Parking Permit #2"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   40
         Left            =   5895
         TabIndex        =   90
         Top             =   993
         Width           =   1260
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Parking Permit #1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   39
         Left            =   120
         TabIndex        =   89
         Top             =   993
         Width           =   1260
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Vehicle"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   38
         Left            =   6045
         TabIndex        =   88
         Top             =   639
         Width           =   1110
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Combination"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   37
         Left            =   6270
         TabIndex        =   87
         Top             =   1701
         Width           =   870
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Locker #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   36
         Left            =   120
         TabIndex        =   86
         Top             =   1701
         Width           =   645
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "License Plate #2"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   35
         Left            =   5955
         TabIndex        =   85
         Top             =   1347
         Width           =   1200
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "License Plate #1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   34
         Left            =   120
         TabIndex        =   84
         Top             =   1347
         Width           =   1200
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Driver License #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   33
         Left            =   120
         TabIndex        =   83
         Top             =   639
         Width           =   1170
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pager Number"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   21
         Left            =   6015
         TabIndex        =   82
         Top             =   285
         Width           =   1140
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cellular Telephone"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   81
         Top             =   285
         Width           =   1320
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Smoker"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   80
         Top             =   2070
         Width           =   540
      End
   End
   Begin VB.Frame frPersonal 
      BorderStyle     =   0  'None
      Caption         =   "Personal"
      Height          =   5055
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   12200
      Begin VB.TextBox txtVadim1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   10080
         MaxLength       =   2
         TabIndex        =   32
         Top             =   1260
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.ComboBox comCountryOfEmp 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7530
         TabIndex        =   31
         Tag             =   "00-Country of Employment"
         Top             =   3605
         Width           =   1320
      End
      Begin VB.TextBox txtCountryOfEmp 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_WORKCOUNTRY"
         Height          =   285
         Left            =   8880
         MaxLength       =   10
         TabIndex        =   30
         Tag             =   "01-Country"
         Top             =   3620
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtEML 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_EML"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9240
         TabIndex        =   29
         Top             =   1260
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtCompany 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_COMPNO"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9240
         MaxLength       =   25
         TabIndex        =   28
         Top             =   1620
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtBadgeID 
         Appearance      =   0  'Flat
         DataField       =   "ED_BADGEID"
         Height          =   285
         Left            =   4680
         MaxLength       =   15
         TabIndex        =   27
         Tag             =   "00-Badge ID"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtMidName 
         Appearance      =   0  'Flat
         DataField       =   "ED_MIDNAME"
         Height          =   285
         Left            =   1830
         MaxLength       =   30
         TabIndex        =   26
         Tag             =   "00-Middle Name"
         Top             =   1592
         Width           =   3765
      End
      Begin VB.TextBox txtAlias 
         Appearance      =   0  'Flat
         DataField       =   "ED_ALIAS"
         Height          =   285
         Left            =   1830
         MaxLength       =   30
         TabIndex        =   25
         Tag             =   "00-Alias"
         Top             =   1930
         Width           =   3765
      End
      Begin VB.ComboBox comCountry 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7530
         TabIndex        =   24
         Tag             =   "00-Country"
         Top             =   3267
         Width           =   1320
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         DataField       =   "ED_TITLE"
         Height          =   285
         Left            =   1830
         MaxLength       =   4
         TabIndex        =   23
         Tag             =   "00-Courtesy Title - for example Mr. or Mrs."
         Top             =   578
         Width           =   875
      End
      Begin VB.TextBox txtCity 
         Appearance      =   0  'Flat
         DataField       =   "ED_CITY"
         Height          =   285
         Left            =   1830
         MaxLength       =   30
         TabIndex        =   22
         Tag             =   "01-City"
         Top             =   2944
         Width           =   2895
      End
      Begin VB.TextBox txtAdd2 
         Appearance      =   0  'Flat
         DataField       =   "ED_ADDR2"
         Height          =   285
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   21
         Tag             =   "00-Second Line in Address"
         Top             =   2606
         Width           =   4180
      End
      Begin VB.TextBox txtAdd1 
         Appearance      =   0  'Flat
         DataField       =   "ED_ADDR1"
         Height          =   285
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   20
         Tag             =   "01-First Line in Address"
         Top             =   2268
         Width           =   4180
      End
      Begin VB.TextBox txtFName 
         Appearance      =   0  'Flat
         DataField       =   "ED_FNAME"
         Height          =   285
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   19
         Tag             =   "01-First or Given Name"
         Top             =   1254
         Width           =   4180
      End
      Begin VB.TextBox txtSurname 
         Appearance      =   0  'Flat
         DataField       =   "ED_SURNAME"
         Height          =   285
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   18
         Tag             =   "01-Surname"
         Top             =   916
         Width           =   4180
      End
      Begin VB.TextBox txtPayrollID 
         Appearance      =   0  'Flat
         DataField       =   "ED_PAYROLL_ID"
         Height          =   285
         Left            =   1830
         MaxLength       =   15
         TabIndex        =   17
         Tag             =   "00-Payroll ID"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtMStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_MSTAT"
         Enabled         =   0   'False
         Height          =   300
         Left            =   3300
         MaxLength       =   1
         TabIndex        =   16
         Text            =   "T"
         Top             =   4303
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.ComboBox ComMStat 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "TempBasicForm.frx":0294
         Left            =   1830
         List            =   "TempBasicForm.frx":0296
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Tag             =   "Marital Status"
         Top             =   4296
         Width           =   1455
      End
      Begin VB.TextBox txtCountry 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_COUNTRY"
         Height          =   285
         Left            =   8880
         MaxLength       =   10
         TabIndex        =   14
         Tag             =   "01-Country"
         Top             =   3282
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Frame frmSex 
         Height          =   375
         Left            =   7530
         TabIndex        =   11
         Top             =   4230
         Width           =   2010
         Begin VB.OptionButton optGender 
            Alignment       =   1  'Right Justify
            Caption         =   "Male"
            Height          =   225
            Index           =   0
            Left            =   210
            TabIndex        =   13
            Tag             =   "41-Gender"
            Top             =   120
            Width           =   675
         End
         Begin VB.OptionButton optGender 
            Caption         =   "Female"
            Height          =   225
            Index           =   1
            Left            =   1050
            TabIndex        =   12
            Tag             =   "41-Gender"
            Top             =   120
            Width           =   930
         End
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_LDATE"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   9240
         MaxLength       =   25
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_LTIME"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   9600
         MaxLength       =   25
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_LUSER"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   9960
         MaxLength       =   25
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtUnion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   9600
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1260
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox txtENTOPT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_ENTOPTS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   9840
         MaxLength       =   2
         TabIndex        =   6
         Top             =   1620
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtENTOPT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_ENTOPT"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   9600
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1620
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtTDATE 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_ETDATES"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   9780
         MaxLength       =   11
         TabIndex        =   4
         Top             =   2370
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtTDATE 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_ETDATE"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   9780
         MaxLength       =   11
         TabIndex        =   3
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtFDATE 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_EFDATES"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   9240
         MaxLength       =   11
         TabIndex        =   2
         Top             =   2370
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtFDATE 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_EFDATE"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   9240
         MaxLength       =   11
         TabIndex        =   1
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSMask.MaskEdBox medSIN 
         DataField       =   "ED_SIN"
         Height          =   285
         Left            =   1830
         TabIndex        =   33
         Tag             =   "00-Social Insurance Number"
         Top             =   3958
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
      Begin INFOHR_Controls.DateLookup dlpDOB 
         DataField       =   "ED_DOB"
         Height          =   285
         Left            =   1515
         TabIndex        =   34
         Tag             =   "41-Birth Date"
         Top             =   3620
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1465
      End
      Begin INFOHR_Controls.CodeLookup clpProv 
         DataField       =   "ED_PROV"
         Height          =   285
         Left            =   7230
         TabIndex        =   35
         Tag             =   "31-Province - Code"
         Top             =   2944
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
      End
      Begin MSMask.MaskEdBox medPCode 
         DataField       =   "ED_PCODE"
         Height          =   285
         Left            =   1830
         TabIndex        =   36
         Tag             =   "40-Postal Code"
         Top             =   3282
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medTelephone 
         DataField       =   "ED_PHONE"
         Height          =   285
         Left            =   1830
         TabIndex        =   37
         Tag             =   "11-Telephone Number"
         Top             =   4665
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(###) ###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medTele2 
         DataField       =   "ED_BUSNBR"
         Height          =   285
         Left            =   7530
         TabIndex        =   38
         Tag             =   "10-Alternate Telephone Number"
         Top             =   4665
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(###) ###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medSSN 
         DataField       =   "ED_SSN"
         Height          =   285
         Left            =   5100
         TabIndex        =   39
         Tag             =   "00-Social Insurance Number"
         Top             =   3958
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataSource      =   " "
         Height          =   285
         Index           =   0
         Left            =   9840
         TabIndex        =   40
         Tag             =   "41-Original Hire Date "
         Top             =   3958
         Visible         =   0   'False
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1060
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Country of Employment"
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
         Index           =   44
         Left            =   5190
         TabIndex        =   66
         Top             =   3665
         Width           =   1950
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
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
         Index           =   19
         Left            =   6480
         TabIndex        =   65
         Top             =   3327
         Width           =   660
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   3840
         TabIndex        =   64
         Top             =   3665
         Width           =   330
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Badge ID"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   43
         Left            =   3600
         TabIndex        =   63
         Top             =   255
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   42
         Left            =   120
         TabIndex        =   62
         Top             =   1637
         Width           =   930
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Alias"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   41
         Left            =   120
         TabIndex        =   61
         Top             =   1975
         Width           =   330
      End
      Begin VB.Label lblMStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marital Status"
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   4356
         Width           =   960
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "S.S.N."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   22
         Left            =   4440
         TabIndex        =   59
         Top             =   4003
         Width           =   465
      End
      Begin VB.Label lblDOH 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8880
         TabIndex        =   58
         Top             =   4003
         Width           =   870
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Original Hire Date:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   7500
         TabIndex        =   57
         Top             =   4003
         Width           =   1290
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   17
         Left            =   4320
         TabIndex        =   56
         Top             =   3665
         Width           =   300
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Postal Code"
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
         Index           =   5
         Left            =   120
         TabIndex        =   55
         Top             =   3327
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Province"
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
         Index           =   6
         Left            =   6360
         TabIndex        =   54
         Top             =   2989
         Width           =   765
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Salutation"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   53
         Top             =   623
         Width           =   1305
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone #2 "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   6120
         TabIndex        =   52
         Top             =   4710
         Width           =   1050
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
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
         Index           =   7
         Left            =   120
         TabIndex        =   51
         Top             =   4710
         Width           =   915
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "S.I.N."
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
         Index           =   8
         Left            =   120
         TabIndex        =   50
         Top             =   4003
         Width           =   510
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
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
         Index           =   9
         Left            =   120
         TabIndex        =   49
         Top             =   3665
         Width           =   870
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "City"
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
         Index           =   4
         Left            =   120
         TabIndex        =   48
         Top             =   2989
         Width           =   330
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address 2"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   31
         Left            =   120
         TabIndex        =   47
         Top             =   2621
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Index           =   3
         Left            =   120
         TabIndex        =   46
         Top             =   2313
         Width           =   690
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
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
         Left            =   120
         TabIndex        =   45
         Top             =   1299
         Width           =   915
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Surname"
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
         Left            =   120
         TabIndex        =   44
         Top             =   961
         Width           =   750
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll ID"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   32
         Left            =   120
         TabIndex        =   43
         Top             =   255
         Width           =   1095
      End
      Begin VB.Image picPhoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   2535
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2625
      End
      Begin VB.Label PicNotF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Photo not Available"
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   6360
         TabIndex        =   42
         Top             =   1080
         Width           =   2115
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgHelp 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1440
         Picture         =   "TempBasicForm.frx":0298
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblWFCNote1 
         Caption         =   "Payroll ID must match ADP && Badge ID must match Tracker"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   9000
         TabIndex        =   41
         Top             =   240
         Visible         =   0   'False
         Width           =   4215
      End
   End
   Begin MSComctlLib.TabStrip tbDemographics 
      Height          =   11175
      Left            =   120
      TabIndex        =   127
      Top             =   360
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   19711
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Personal"
            Key             =   "tbPersonal"
            Object.ToolTipText     =   "Employee's Personal Information"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Organizational"
            Key             =   "tbOrganizational"
            Object.ToolTipText     =   "Employee's Organizational Information"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Miscellaneous"
            Key             =   "tbMiscellaneous"
            Object.ToolTipText     =   "Employee's Miscellaneous Information"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tbDemographics_Click()
    If tbDemographics.SelectedItem.Index = 1 Then
        frPersonal.Visible = True
        frPersonal.Top = 960
        frOrganizational.Visible = False
        frMiscellaneous.Visible = False
    ElseIf tbDemographics.SelectedItem.Index = 2 Then
        frPersonal.Visible = False
        frOrganizational.Visible = True
        frOrganizational.Top = 960
        frMiscellaneous.Visible = False
    ElseIf tbDemographics.SelectedItem.Index = 3 Then
        frPersonal.Visible = False
        frOrganizational.Visible = False
        frMiscellaneous.Visible = True
        frMiscellaneous.Top = 960
    End If
End Sub
