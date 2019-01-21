VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEEBASIC 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Demographics"
   ClientHeight    =   10905
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12900
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10905
   ScaleWidth      =   12900
   Tag             =   "Edit the information on this screen"
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame frMiscellaneous 
      BorderStyle     =   0  'None
      Caption         =   "Miscellaneous"
      Height          =   2655
      Left            =   4080
      TabIndex        =   145
      Top             =   7680
      Visible         =   0   'False
      Width           =   7515
      Begin VB.TextBox txtOtherEmail 
         Appearance      =   0  'Flat
         DataField       =   "ED_OTHREMAIL"
         DataSource      =   " "
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
         Left            =   7530
         MaxLength       =   60
         TabIndex        =   57
         Tag             =   "00-Other Email Address"
         Top             =   1680
         Width           =   4260
      End
      Begin VB.TextBox txtSmoker 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_SMOKER"
         Enabled         =   0   'False
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
         Left            =   2700
         MaxLength       =   2
         TabIndex        =   146
         Text            =   "Text14"
         Top             =   1650
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.ComboBox ComSmoker 
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
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Tag             =   "00-Smoker Yes/No"
         Top             =   1650
         Width           =   855
      End
      Begin MSMask.MaskEdBox medDRIVERLIC 
         DataField       =   "ED_DRIVERLIC"
         Height          =   285
         Left            =   1830
         TabIndex        =   48
         Tag             =   "00-Driver License Number"
         Top             =   240
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
         TabIndex        =   52
         Tag             =   "00-License Plate #1"
         Top             =   945
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
         TabIndex        =   53
         Tag             =   "00-License Plate #2"
         Top             =   945
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
         TabIndex        =   54
         Tag             =   "00-Locker #"
         Top             =   1290
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
         TabIndex        =   55
         Tag             =   "00-Combination"
         Top             =   1290
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
         TabIndex        =   49
         Tag             =   "00-Type of Vehicle"
         Top             =   240
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
         TabIndex        =   50
         Tag             =   "00-Parking Permit #1"
         Top             =   585
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
         TabIndex        =   51
         Tag             =   "00-Parking Permit #2"
         Top             =   585
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
      Begin MSMask.MaskEdBox medNetworkLogin 
         Height          =   285
         Left            =   1830
         TabIndex        =   58
         Tag             =   "00-Network Login"
         Top             =   2040
         Visible         =   0   'False
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   40
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
      Begin MSMask.MaskEdBox medVendorNo 
         Height          =   285
         Left            =   7530
         TabIndex        =   59
         Tag             =   "00-Vendor Number"
         Top             =   2040
         Visible         =   0   'False
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
      Begin INFOHR_Controls.CodeLookup clpVadim1 
         Height          =   285
         Left            =   1510
         TabIndex        =   60
         Top             =   2400
         Visible         =   0   'False
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDV1"
      End
      Begin VB.Label lblVadim1 
         AutoSize        =   -1  'True
         Caption         =   "Vadim Field 1"
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
         Left            =   120
         TabIndex        =   210
         Top             =   2460
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblOtherEmail 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Other Email Address"
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
         Left            =   5790
         TabIndex        =   209
         Top             =   1725
         Width           =   1425
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Number"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   56
         Left            =   5790
         TabIndex        =   206
         Top             =   2085
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Network Login"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   55
         Left            =   120
         TabIndex        =   205
         Top             =   2085
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Smoker"
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
         Index           =   16
         Left            =   120
         TabIndex        =   155
         Top             =   1710
         Width           =   540
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Driver License #"
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
         Index           =   33
         Left            =   120
         TabIndex        =   154
         Top             =   285
         Width           =   1170
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "License Plate #1"
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
         Index           =   34
         Left            =   120
         TabIndex        =   153
         Top             =   990
         Width           =   1200
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "License Plate #2"
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
         Index           =   35
         Left            =   5790
         TabIndex        =   152
         Top             =   990
         Width           =   1200
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Locker #"
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
         Index           =   36
         Left            =   120
         TabIndex        =   151
         Top             =   1335
         Width           =   645
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Combination"
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
         Index           =   37
         Left            =   5790
         TabIndex        =   150
         Top             =   1335
         Width           =   870
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Vehicle"
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
         Index           =   38
         Left            =   5790
         TabIndex        =   149
         Top             =   285
         Width           =   1110
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Parking Permit #1"
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
         Index           =   39
         Left            =   120
         TabIndex        =   148
         Top             =   630
         Width           =   1260
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Parking Permit #2"
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
         Index           =   40
         Left            =   5790
         TabIndex        =   147
         Top             =   630
         Width           =   1260
      End
   End
   Begin VB.Frame frOrganizational 
      BorderStyle     =   0  'None
      Caption         =   "Organizational"
      Height          =   2895
      Left            =   240
      TabIndex        =   124
      Top             =   7320
      Visible         =   0   'False
      Width           =   12075
      Begin VB.CommandButton cmdEditDiv 
         Appearance      =   0  'Flat
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   840
         TabIndex        =   207
         Tag             =   "Edit Transaction Date"
         Top             =   920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtDeptBonusCtr 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "ED_BONUSDEPT"
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
         Left            =   7880
         MaxLength       =   6
         TabIndex        =   29
         Tag             =   "00-Bonus Reporting #"
         Top             =   580
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Frame frmWFCDIV 
         Height          =   330
         Left            =   6720
         TabIndex        =   125
         Top             =   4560
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox txtDouDiv 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   320
            MaxLength       =   4
            TabIndex        =   32
            Tag             =   "00-Bonus Reporting #"
            Top             =   0
            Width           =   870
         End
         Begin VB.Image imgIDiv 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   0
            Picture         =   "febasic.frx":0000
            Top             =   0
            Width           =   240
         End
         Begin VB.Label lblDouDivDesc 
            Caption         =   "Unassigned"
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
            Height          =   255
            Left            =   1320
            TabIndex        =   126
            Top             =   0
            Width           =   2415
         End
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         DataField       =   "ED_DIV"
         Height          =   285
         Left            =   1515
         TabIndex        =   31
         Tag             =   "00-Division"
         Top             =   920
         Width           =   3960
         _ExtentX        =   6985
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
         TabIndex        =   36
         Tag             =   "00-Administered By"
         Top             =   1600
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
         MaxLength       =   10
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         DataField       =   "ED_DEPTNO"
         Height          =   285
         Left            =   1515
         TabIndex        =   26
         Tag             =   "00-Department"
         Top             =   240
         Width           =   3960
         _ExtentX        =   6985
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
         TabIndex        =   28
         Tag             =   "00-General Ledger - Code"
         Top             =   580
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   3
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   7560
         TabIndex        =   35
         Tag             =   "00-Region"
         Top             =   1260
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   7560
         TabIndex        =   37
         Tag             =   "00-Section"
         Top             =   1605
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.DateLookup dlpDivEDate 
         DataField       =   "ED_DIVEDATE"
         Height          =   285
         Left            =   7560
         TabIndex        =   33
         Tag             =   "40-Division Effective Date"
         Top             =   915
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDeptEDate 
         DataField       =   "ED_DEPTEDATE"
         Height          =   285
         Left            =   7560
         TabIndex        =   27
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
         TabIndex        =   38
         Tag             =   "00-Home Operation Number"
         Top             =   1935
         Visible         =   0   'False
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "HMOP"
         MaxLength       =   12
      End
      Begin INFOHR_Controls.CodeLookup clpHOME 
         Height          =   285
         Index           =   2
         Left            =   1515
         TabIndex        =   39
         Tag             =   "00-Home Line"
         Top             =   2280
         Visible         =   0   'False
         Width           =   4020
         _ExtentX        =   7091
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
         TabIndex        =   40
         Tag             =   "00-Home Work Center"
         Top             =   2620
         Visible         =   0   'False
         Width           =   4020
         _ExtentX        =   7091
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
         TabIndex        =   41
         Tag             =   "00-Home Shift"
         Top             =   2960
         Visible         =   0   'False
         Width           =   4020
         _ExtentX        =   7091
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
         TabIndex        =   34
         Tag             =   "00-Location - Code"
         Top             =   1260
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   5
         Left            =   10200
         TabIndex        =   30
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ED_ORGT1"
         Height          =   285
         Index           =   6
         Left            =   1515
         TabIndex        =   42
         Tag             =   "00-Orgranization - Code"
         Top             =   3300
         Visible         =   0   'False
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ORGN"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ED_ORGT2"
         Height          =   285
         Index           =   7
         Left            =   1515
         TabIndex        =   44
         Tag             =   "00-Orgranization - Code"
         Top             =   3640
         Visible         =   0   'False
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "OGN2"
      End
      Begin INFOHR_Controls.DateLookup dlpOrg1EDate 
         DataField       =   "ED_ORGT1EDATE"
         Height          =   285
         Left            =   7560
         TabIndex        =   43
         Tag             =   "40-Orgranization 1 Effective Date"
         Top             =   3300
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpOrg2EDate 
         DataField       =   "ED_ORGT2EDATE"
         Height          =   285
         Left            =   7560
         TabIndex        =   45
         Tag             =   "40-Organization 1 Effective Date"
         Top             =   3640
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpSalDist 
         DataField       =   "ED_SALDIST"
         Height          =   285
         Left            =   7560
         TabIndex        =   47
         Top             =   2280
         Visible         =   0   'False
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   6
         LookupType      =   8
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataSource      =   " "
         Height          =   285
         Index           =   0
         Left            =   7560
         TabIndex        =   46
         Tag             =   "00-Enter Union Code"
         Top             =   1935
         Visible         =   0   'False
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOR"
      End
      Begin VB.Label lblUnion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   5775
         TabIndex        =   164
         Top             =   1980
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblSalDist 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Distribution"
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
         Left            =   5775
         TabIndex        =   163
         Top             =   2325
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lblOrg2EffDate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Organization 2 Effective"
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
         Left            =   5775
         TabIndex        =   160
         Top             =   3685
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblOrg1EffDate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Organization 1 Effective"
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
         Left            =   5775
         TabIndex        =   159
         Top             =   3345
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Organization 2"
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
         Index           =   45
         Left            =   120
         TabIndex        =   158
         Top             =   3685
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Organization 1"
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
         Index           =   18
         Left            =   120
         TabIndex        =   157
         Top             =   3345
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   7560
         Picture         =   "febasic.frx":014A
         Top             =   600
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblDeptBonusDesc 
         Caption         =   "Unassigned"
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
         Height          =   255
         Left            =   8880
         TabIndex        =   143
         Top             =   600
         Visible         =   0   'False
         Width           =   3735
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
         TabIndex        =   142
         Top             =   3960
         Visible         =   0   'False
         Width           =   855
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
         TabIndex        =   141
         Top             =   4200
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblDeptStart 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Department Effective"
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
         Left            =   5775
         TabIndex        =   140
         Top             =   285
         Width           =   1500
      End
      Begin VB.Label lblDivStart 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Division Effective"
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
         Left            =   5775
         TabIndex        =   139
         Top             =   965
         Width           =   1230
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Home Work Center"
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
         Index           =   30
         Left            =   120
         TabIndex        =   138
         Top             =   2665
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Home Shift"
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
         Index           =   28
         Left            =   120
         TabIndex        =   137
         Top             =   3005
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lbltitle 
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
         Index           =   29
         Left            =   120
         TabIndex        =   136
         Top             =   2325
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbltitle 
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
         Index           =   27
         Left            =   120
         TabIndex        =   135
         Top             =   1985
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lbltitle 
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
         Index           =   26
         Left            =   5775
         TabIndex        =   134
         Top             =   1645
         Width           =   540
      End
      Begin VB.Label lbltitle 
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
         Index           =   13
         Left            =   120
         TabIndex        =   133
         Top             =   960
         Width           =   675
      End
      Begin VB.Label lbltitle 
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
         Index           =   23
         Left            =   120
         TabIndex        =   132
         Top             =   1305
         Width           =   615
      End
      Begin VB.Label lbltitle 
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
         Index           =   24
         Left            =   5775
         TabIndex        =   131
         Top             =   1305
         Width           =   510
      End
      Begin VB.Label lbltitle 
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
         Index           =   25
         Left            =   120
         TabIndex        =   130
         Top             =   1645
         Width           =   1125
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   129
         Top             =   285
         Width           =   990
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "G/L #"
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
         Left            =   120
         TabIndex        =   128
         Top             =   625
         Width           =   435
      End
      Begin VB.Label lblRptNo 
         Caption         =   "Employee not in Tracker"
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
         Left            =   5775
         TabIndex        =   127
         Top             =   630
         Visible         =   0   'False
         Width           =   1920
      End
   End
   Begin VB.Frame frAltPayIDs 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   11040
      TabIndex        =   168
      Top             =   8520
      Visible         =   0   'False
      Width           =   10935
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   10
         Left            =   960
         TabIndex        =   186
         Tag             =   "00-Enter pay period code"
         Top             =   1080
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "SDPP"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   8
         Left            =   960
         TabIndex        =   170
         Tag             =   "00-Enter pay period code"
         Top             =   360
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "SDPP"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   9
         Left            =   960
         TabIndex        =   178
         Tag             =   "00-Enter pay period code"
         Top             =   720
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "SDPP"
      End
      Begin MSMask.MaskEdBox medAltPayID 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   169
         Tag             =   "00-Alt. Payroll ID"
         Top             =   360
         Width           =   885
         _ExtentX        =   1561
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
      Begin MSMask.MaskEdBox medAltPayID 
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   177
         Tag             =   "00-Alt. Payroll ID"
         Top             =   720
         Width           =   885
         _ExtentX        =   1561
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
      Begin MSMask.MaskEdBox medAltPayID 
         Height          =   285
         Index           =   2
         Left            =   0
         TabIndex        =   185
         Tag             =   "00-Alt. Payroll ID"
         Top             =   1080
         Width           =   885
         _ExtentX        =   1561
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
      Begin MSMask.MaskEdBox medAltComp 
         Height          =   165
         Index           =   1
         Left            =   5760
         TabIndex        =   196
         Tag             =   "00-Alt. Company Code"
         Top             =   0
         Visible         =   0   'False
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   291
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BackColor       =   -2147483645
         PromptInclude   =   0   'False
         MaxLength       =   4
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
      Begin MSMask.MaskEdBox medAltComp 
         Height          =   165
         Index           =   2
         Left            =   5160
         TabIndex        =   197
         Tag             =   "00-Alt. Company Code"
         Top             =   0
         Visible         =   0   'False
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   291
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BackColor       =   -2147483645
         PromptInclude   =   0   'False
         MaxLength       =   4
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
      Begin MSMask.MaskEdBox medAltComp 
         Height          =   165
         Index           =   0
         Left            =   4560
         TabIndex        =   194
         Tag             =   "00-Alt. Company Code"
         Top             =   0
         Visible         =   0   'False
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   291
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BackColor       =   -2147483645
         PromptInclude   =   0   'False
         MaxLength       =   4
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   11
         Left            =   2280
         TabIndex        =   171
         Tag             =   "00-Location - Code"
         Top             =   360
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   12
         Left            =   2280
         TabIndex        =   179
         Tag             =   "00-Location - Code"
         Top             =   720
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   13
         Left            =   2280
         TabIndex        =   187
         Tag             =   "00-Location - Code"
         Top             =   1080
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpSalDis2 
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   172
         Top             =   360
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "n/a"
         MaxLength       =   6
         LookupType      =   8
      End
      Begin INFOHR_Controls.CodeLookup clpSalDis2 
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   180
         Top             =   720
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "n/a"
         MaxLength       =   6
         LookupType      =   8
      End
      Begin INFOHR_Controls.CodeLookup clpSalDis2 
         Height          =   285
         Index           =   2
         Left            =   3840
         TabIndex        =   188
         Top             =   1080
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "n/a"
         MaxLength       =   6
         LookupType      =   8
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataSource      =   " "
         Height          =   285
         Index           =   1
         Left            =   5280
         TabIndex        =   173
         Tag             =   "41-Original Hire Date "
         Top             =   360
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1060
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataSource      =   " "
         Height          =   285
         Index           =   2
         Left            =   5280
         TabIndex        =   181
         Tag             =   "41-Original Hire Date "
         Top             =   720
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1060
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataSource      =   " "
         Height          =   285
         Index           =   3
         Left            =   5280
         TabIndex        =   189
         Tag             =   "41-Original Hire Date "
         Top             =   1080
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1060
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   14
         Left            =   6720
         TabIndex        =   174
         Tag             =   "00-Region"
         Top             =   360
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   15
         Left            =   6720
         TabIndex        =   182
         Tag             =   "00-Region"
         Top             =   720
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   16
         Left            =   6720
         TabIndex        =   190
         Tag             =   "00-Region"
         Top             =   1080
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   17
         Left            =   7800
         TabIndex        =   175
         Tag             =   "00-Enter Status Code"
         Top             =   360
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDEM"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   18
         Left            =   7800
         TabIndex        =   183
         Tag             =   "00-Enter Status Code"
         Top             =   720
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDEM"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   19
         Left            =   7800
         TabIndex        =   191
         Tag             =   "00-Enter Status Code"
         Top             =   1080
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDEM"
      End
      Begin INFOHR_Controls.DateLookup dlpTermDate 
         Height          =   285
         Index           =   0
         Left            =   9000
         TabIndex        =   176
         Tag             =   "Date Terminated"
         Top             =   360
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1060
      End
      Begin INFOHR_Controls.DateLookup dlpTermDate 
         Height          =   285
         Index           =   1
         Left            =   9000
         TabIndex        =   184
         Tag             =   "Date Terminated"
         Top             =   720
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1060
      End
      Begin INFOHR_Controls.DateLookup dlpTermDate 
         Height          =   285
         Index           =   2
         Left            =   9000
         TabIndex        =   192
         Tag             =   "Date Terminated"
         Top             =   1080
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1060
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   54
         Left            =   6960
         TabIndex        =   203
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Termination Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   53
         Left            =   9000
         TabIndex        =   202
         Top             =   120
         Width           =   1470
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   52
         Left            =   8040
         TabIndex        =   201
         Top             =   120
         Width           =   555
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hire Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   51
         Left            =   5640
         TabIndex        =   200
         Top             =   120
         Width           =   825
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Distribution"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   50
         Left            =   3840
         TabIndex        =   199
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   49
         Left            =   2400
         TabIndex        =   198
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   48
         Left            =   960
         TabIndex        =   195
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll ID"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   47
         Left            =   0
         TabIndex        =   193
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.VScrollBar scrControl 
      Height          =   5565
      LargeChange     =   315
      Left            =   12600
      Max             =   100
      SmallChange     =   315
      TabIndex        =   75
      Top             =   900
      Width           =   300
   End
   Begin VB.Frame frPersonal 
      BorderStyle     =   0  'None
      Caption         =   "Personal"
      Height          =   5895
      Left            =   120
      TabIndex        =   80
      Top             =   960
      Width           =   12435
      Begin VB.CommandButton cmdEditPayID 
         Appearance      =   0  'Flat
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   960
         TabIndex        =   211
         Tag             =   "Edit Transaction Date"
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCandidate 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   4680
         MaxLength       =   15
         TabIndex        =   161
         Tag             =   "00-Badge ID"
         Top             =   578
         Visible         =   0   'False
         Width           =   1335
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataSource      =   " "
         Height          =   285
         Index           =   0
         Left            =   8520
         TabIndex        =   18
         Tag             =   "41-Original Hire Date "
         Top             =   3960
         Visible         =   0   'False
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1060
      End
      Begin VB.TextBox txtGender 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_SEX"
         Enabled         =   0   'False
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
         Left            =   10800
         MaxLength       =   1
         TabIndex        =   156
         Text            =   "Text14"
         Top             =   4320
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.ListBox lstEETables 
         Columns         =   2
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   144
         TabStop         =   0   'False
         Tag             =   "00-Employee Records found in these files"
         Top             =   5520
         Visible         =   0   'False
         Width           =   12375
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
         Left            =   7530
         TabIndex        =   15
         Tag             =   "00-Country of Employment"
         Top             =   3605
         Width           =   1320
      End
      Begin VB.ComboBox comCountry 
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
         Left            =   7530
         TabIndex        =   13
         Tag             =   "00-Country"
         Top             =   3267
         Width           =   1320
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
         TabIndex        =   97
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
         TabIndex        =   96
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
         TabIndex        =   95
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
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
         TabIndex        =   94
         Top             =   2370
         Visible         =   0   'False
         Width           =   1215
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
         TabIndex        =   93
         Top             =   1620
         Visible         =   0   'False
         Width           =   255
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
         TabIndex        =   92
         Top             =   1620
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtUnion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   9600
         MaxLength       =   2
         TabIndex        =   91
         Top             =   1260
         Visible         =   0   'False
         Width           =   450
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
         TabIndex        =   90
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
         TabIndex        =   89
         Top             =   840
         Visible         =   0   'False
         Width           =   330
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
         TabIndex        =   88
         Top             =   840
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Frame frmSex 
         Height          =   375
         Left            =   7530
         TabIndex        =   87
         Top             =   4230
         Width           =   3090
         Begin VB.OptionButton optGender 
            Caption         =   "Female"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   840
            TabIndex        =   21
            Tag             =   "41-Gender"
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton optGender 
            Caption         =   "Male"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Tag             =   "41-Gender"
            Top             =   120
            Width           =   675
         End
         Begin VB.OptionButton optGender 
            Caption         =   "Not Disclosed"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   1680
            TabIndex        =   204
            Tag             =   "41-Gender"
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.TextBox txtCountry 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_COUNTRY"
         Height          =   285
         Left            =   8880
         MaxLength       =   10
         TabIndex        =   86
         Tag             =   "01-Country"
         Top             =   3282
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.ComboBox ComMStat 
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
         ItemData        =   "febasic.frx":0294
         Left            =   1830
         List            =   "febasic.frx":0296
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Tag             =   "Marital Status"
         Top             =   4296
         Width           =   1455
      End
      Begin VB.TextBox txtMStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_MSTAT"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3300
         MaxLength       =   1
         TabIndex        =   85
         Text            =   "T"
         Top             =   4303
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox txtPayrollID 
         Appearance      =   0  'Flat
         DataField       =   "ED_PAYROLL_ID"
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
         Left            =   1830
         MaxLength       =   25
         TabIndex        =   0
         Tag             =   "00-Payroll ID"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtSurname 
         Appearance      =   0  'Flat
         DataField       =   "ED_SURNAME"
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
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   3
         Tag             =   "01-Surname"
         Top             =   916
         Width           =   4180
      End
      Begin VB.TextBox txtFName 
         Appearance      =   0  'Flat
         DataField       =   "ED_FNAME"
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
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   4
         Tag             =   "01-First or Given Name"
         Top             =   1254
         Width           =   4180
      End
      Begin VB.TextBox txtAdd1 
         Appearance      =   0  'Flat
         DataField       =   "ED_ADDR1"
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
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   7
         Tag             =   "01-First Line in Address"
         Top             =   2268
         Width           =   4180
      End
      Begin VB.TextBox txtAdd2 
         Appearance      =   0  'Flat
         DataField       =   "ED_ADDR2"
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
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   8
         Tag             =   "00-Second Line in Address"
         Top             =   2606
         Width           =   4180
      End
      Begin VB.TextBox txtCity 
         Appearance      =   0  'Flat
         DataField       =   "ED_CITY"
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
         Left            =   1830
         MaxLength       =   30
         TabIndex        =   9
         Tag             =   "01-City"
         Top             =   2944
         Width           =   2895
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         DataField       =   "ED_TITLE"
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
         Left            =   1830
         MaxLength       =   4
         TabIndex        =   2
         Tag             =   "00-Courtesy Title - for example Mr. or Mrs."
         Top             =   578
         Width           =   875
      End
      Begin VB.TextBox txtAlias 
         Appearance      =   0  'Flat
         DataField       =   "ED_ALIAS"
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
         Left            =   1830
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "00-Alias"
         Top             =   1930
         Width           =   3765
      End
      Begin VB.TextBox txtMidName 
         Appearance      =   0  'Flat
         DataField       =   "ED_MIDNAME"
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
         Left            =   1830
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "00-Middle Name"
         Top             =   1592
         Width           =   3765
      End
      Begin VB.TextBox txtBadgeID 
         Appearance      =   0  'Flat
         DataField       =   "ED_BADGEID"
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
         Left            =   4680
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "00-Badge ID"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtCompany 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_COMPNO"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9240
         MaxLength       =   25
         TabIndex        =   84
         Top             =   1620
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtEML 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_EML"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9240
         TabIndex        =   83
         Top             =   1260
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtCountryOfEmp 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_WORKCOUNTRY"
         Height          =   285
         Left            =   8880
         MaxLength       =   10
         TabIndex        =   82
         Tag             =   "01-Country"
         Top             =   3620
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtVadim1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   10080
         MaxLength       =   2
         TabIndex        =   81
         Top             =   1260
         Visible         =   0   'False
         Width           =   450
      End
      Begin MSMask.MaskEdBox medSIN 
         DataField       =   "ED_SIN"
         Height          =   285
         Left            =   1830
         TabIndex        =   16
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
         TabIndex        =   14
         Tag             =   "41-Birth Date"
         Top             =   3620
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1465
      End
      Begin MSMask.MaskEdBox medPCode 
         DataField       =   "ED_PCODE"
         Height          =   285
         Left            =   1830
         TabIndex        =   12
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
         TabIndex        =   22
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
         TabIndex        =   23
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
         Left            =   4020
         TabIndex        =   17
         Tag             =   "00-Social Insurance Number"
         Top             =   3960
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
      Begin INFOHR_Controls.CodeLookup clpProv 
         DataField       =   "ED_PROV"
         Height          =   285
         Left            =   7230
         TabIndex        =   10
         Tag             =   "31-Province - Code"
         Top             =   2944
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
      End
      Begin MSMask.MaskEdBox medCellPhone 
         DataField       =   "ED_CELLPHONE"
         Height          =   285
         Left            =   1830
         TabIndex        =   24
         Tag             =   "10-Cellular Telephone Number"
         Top             =   5010
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
         TabIndex        =   25
         Tag             =   "10-Pager Number"
         Top             =   5010
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
      Begin INFOHR_Controls.CodeLookup clpProvEmp 
         Height          =   285
         Left            =   11145
         TabIndex        =   11
         Tag             =   "30-Province Code"
         Top             =   2944
         Visible         =   0   'False
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Prov of Employment"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   57
         Left            =   9480
         TabIndex        =   208
         Top             =   2989
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pager Number"
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
         Index           =   21
         Left            =   6120
         TabIndex        =   166
         Top             =   5055
         Width           =   1020
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cellular Telephone"
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
         Index           =   20
         Left            =   120
         TabIndex        =   165
         Top             =   5055
         Width           =   1320
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Candidate"
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
         Index           =   46
         Left            =   3600
         TabIndex        =   162
         Top             =   630
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblWFCNote1 
         Caption         =   "Payroll ID must match ADP && Badge ID must match Tracker"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   9000
         TabIndex        =   123
         Top             =   240
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Image imgHelp 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1440
         Picture         =   "febasic.frx":0298
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   255
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
         TabIndex        =   122
         Top             =   1080
         Width           =   2115
         WordWrap        =   -1  'True
      End
      Begin VB.Image picPhoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   2535
         Left            =   6240
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2625
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll ID"
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
         Index           =   32
         Left            =   120
         TabIndex        =   121
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Surname"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   120
         Top             =   961
         Width           =   750
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   119
         Top             =   1299
         Width           =   915
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   118
         Top             =   2313
         Width           =   690
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address 2"
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
         Index           =   31
         Left            =   120
         TabIndex        =   117
         Top             =   2651
         Width           =   975
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   116
         Top             =   2989
         Width           =   330
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   115
         Top             =   3665
         Width           =   870
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "S.I.N."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   114
         Top             =   4003
         Width           =   510
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   113
         Top             =   4710
         Width           =   915
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone #2 "
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
         Left            =   6120
         TabIndex        =   112
         Top             =   4710
         Width           =   1050
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Salutation"
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
         TabIndex        =   111
         Top             =   630
         Width           =   705
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Province"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   6360
         TabIndex        =   110
         Top             =   2989
         Width           =   765
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Postal Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   109
         Top             =   3327
         Width           =   1035
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "00"
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
         Index           =   17
         Left            =   4020
         TabIndex        =   108
         Top             =   3660
         Width           =   180
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Original Hire Date"
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
         Index           =   15
         Left            =   5895
         TabIndex        =   107
         Top             =   4005
         Width           =   1245
      End
      Begin VB.Label lblDOH 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
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
         Left            =   7560
         TabIndex        =   106
         Top             =   4005
         Width           =   870
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "S.S.N."
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
         Index           =   22
         Left            =   3480
         TabIndex        =   105
         Top             =   4005
         Width           =   465
      End
      Begin VB.Label lblMStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marital Status"
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
         Left            =   120
         TabIndex        =   104
         Top             =   4356
         Width           =   960
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Alias"
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
         Index           =   41
         Left            =   120
         TabIndex        =   103
         Top             =   1975
         Width           =   330
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
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
         Index           =   42
         Left            =   120
         TabIndex        =   102
         Top             =   1637
         Width           =   930
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Badge ID"
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
         Index           =   43
         Left            =   3600
         TabIndex        =   101
         Top             =   285
         Width           =   975
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
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
         Index           =   14
         Left            =   3615
         TabIndex        =   100
         Top             =   3660
         Width           =   330
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   19
         Left            =   6480
         TabIndex        =   99
         Top             =   3327
         Width           =   660
      End
      Begin VB.Label lbltitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Country of Employment"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   44
         Left            =   5190
         TabIndex        =   98
         Top             =   3665
         Width           =   1950
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   68
      Top             =   10470
      Width           =   12900
      _Version        =   65536
      _ExtentX        =   22754
      _ExtentY        =   767
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
      Begin VB.CommandButton cmdCCLife 
         Caption         =   "CC Dep Life"
         Height          =   375
         Left            =   9600
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton cmdUnlockSmoker 
         Caption         =   "Unlock Smoker Status"
         Height          =   375
         Left            =   7320
         TabIndex        =   63
         Top             =   0
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.CommandButton cmdHide 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   300
         TabIndex        =   65
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdDeleteAll 
         Appearance      =   0  'Flat
         Caption         =   "Delete All"
         Height          =   375
         Left            =   1800
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdMiss 
         Caption         =   "&What is Missing"
         Height          =   375
         Left            =   5280
         TabIndex        =   62
         Top             =   0
         Width           =   1605
      End
      Begin VB.CommandButton cmdPhoto 
         Caption         =   "&Photo Off"
         Height          =   375
         Left            =   3990
         TabIndex        =   61
         Tag             =   "40-Photograph of Employee"
         Top             =   0
         Width           =   1185
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   11160
         Top             =   240
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
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   69
      Top             =   0
      Width           =   12900
      _Version        =   65536
      _ExtentX        =   22754
      _ExtentY        =   873
      _StockProps     =   15
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdCopyEmpByPayID 
         Caption         =   "Copy to New Payroll ID"
         Height          =   330
         Left            =   10200
         TabIndex        =   167
         Top             =   120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll ID must match ADP && Badge ID must match Tracker"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   7335
         TabIndex        =   78
         Top             =   180
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label lblEEProdLine 
         AutoSize        =   -1  'True
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
         Left            =   8040
         TabIndex        =   77
         Top             =   150
         Width           =   75
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   74
         Top             =   180
         Width           =   1005
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
         Left            =   1290
         TabIndex        =   72
         Top             =   150
         Width           =   1245
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblEEID"
         DataField       =   "ED_EMPNBR"
         DataSource      =   "Data1"
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   5790
         TabIndex        =   71
         Top             =   150
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         Left            =   2640
         TabIndex        =   70
         Top             =   150
         Width           =   1740
      End
   End
   Begin VB.Frame frmBlank 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   42375
      Left            =   12480
      TabIndex        =   67
      Top             =   8760
      Visible         =   0   'False
      Width           =   55665
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   11040
      Top             =   11040
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.TabStrip tbDemographics 
      Height          =   6495
      Left            =   0
      TabIndex        =   79
      Top             =   600
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   11456
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
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
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Alt. Payroll IDs"
            Key             =   "tbAltPayID"
            Object.ToolTipText     =   "Employee's Alt. Payroll IDs"
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
   Begin VB.Label lblTempDesc 
      AutoSize        =   -1  'True
      Caption         =   "use for programming"
      Height          =   195
      Left            =   2280
      TabIndex        =   76
      Top             =   10800
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label lblComp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Company #"
      Enabled         =   0   'False
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
      Left            =   1440
      TabIndex        =   73
      Top             =   10800
      Visible         =   0   'False
      Width           =   930
   End
End
Attribute VB_Name = "frmEEBASIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' change this to Public instead of Private when used in IHR.
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_TAB = &H9

'Dim snapDepts As New ADODB.Recordset
'Dim snapEmp As New ADODB.Recordset

Dim fglbNewEE As Integer
Dim fglbNew As Boolean
'Dim rsTA As New ADODB.Recordset

Dim oDept As String, ODeptD As String
Dim OGLNum As String, OGLNumD As String
Dim SavDiv, SavDept, oldEEId
Dim OFNAME, OSNAME, OADD1, OADD2, oGLNo, OCITY, oProv, oProvEmp, oDiv, OTEL, OSEX, OSMOKER
Dim OSection
Dim ODOB, oDOH, OSIN, OTITLE, OMSTAT, OBUSNBR, OPCODE, OPHONE
Dim oCountry As String 'ADDED BY RAUBREY 6/16/97
Dim oCountryEmployment As String
Dim oRegion, oAdminBy, oOrg1, oOrg2
Dim OCellPhone, OPageNbr, OSSN, oPayrollID
Dim OHOMELINE, OHOMESHIFT, OHOMEOPRTNBR, oHOMEWRKCNT
Dim ODivEdate, ODeptEDate
Dim SavLoc  'laura nov 4, 1997
Dim mbAddNewEmployee As Boolean
Dim UnloadForm As Boolean 'Jaddy 10/29/99
Dim glbPicDir, glbPicBMP  'Andy Sham July 20, 99
Dim RDept, RGLNum ''added by Jaddy Sep 20,99
Dim flagFrmLoad As Boolean   'carmen may 00
Dim oDRIVERLIC, oLICPLATE1, oLICPLATE2, oLOCKER, oCOMBINATION
Dim oTYPEVEHICLE, oPARKPERMIT1, oPARKPERMIT2
Dim oBadgeID, oMidName, oAlias
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim fDupSIN, fDupSSN
Dim fDupSIN_Term, fDupSSN_Term
Dim flgDupSINSSN_Term As Boolean
Dim SorocOPayrollID As String
Dim locFEDTAX, locPROVTAX, locPROVEMP, locUIC
Dim MailBody
Dim strNoAccessForms As String
Dim locUploadWithoutCheck As Boolean 'Ticket #19937 for Samuel -  Franks 05/06/2011
Dim AbortTerm As Boolean
'Ticket #22912 Franks 12/06/2012 - begin
Dim xFutureChgDeptNo As Boolean
Dim xFutureChgSection As Boolean
Dim xFutureChgRegion As Boolean
Dim xFutureDateDeptNo
Dim xFutureDateSection
Dim xFutureDateRegion
'Ticket #22912 Franks 12/06/2012 - end
'Ticket #23247 Franks 04/22/2013 - begin
Dim xBenGrpCode
Dim xWFCPayGroup
Dim xWFCNGSCode
'Ticket #23247 Franks 04/22/2013 - end
Dim xHRSoftPTCode, xHRSoftJob, xETHNICITY, xRACE
Dim SavOrg, oSalDist    'Ticket #24543 - Macaulay Child Development Centre
Dim oVadim1 'Ticket #29759 Franks 02/14/2017
'Ticket #24557 Franks 07/07/2015 - begin
Dim ADPBranchOld(3)
Dim ADPDeptOld(3)
'Ticket #24557 Franks 07/07/2015 end

'Ticket #28040 - To Track on New Hire if the user went into the Organizational tab at least once
Dim flgSwitchOrgTabNewHire As Boolean
Dim flgWFCDivChaFlag As Boolean

Private Function AUDITDEMO(Actn)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xProvNbr, xADD
Dim xBatchID
Dim HRChanges As New Collection
Dim UpdatePayrollID
'''On Error GoTo AUDIT_ERR

AUDITDEMO = False



rsTB.Open "SELECT NBR FROM HRPROV WHERE CODE= '" & clpProv.Text & "'", gdbAdoIhr001, adOpenKeyset  ', , adCmdTableDirect

If rsTB.EOF Then
    xProvNbr = "  "
Else
    xProvNbr = rsTB("NBR")
End If

Dim strFields As String

strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, AU_DOLENT_TABL, "
strFields = strFields & "AU_EARN_TABL"

'Number of fields makes using * worth it Ticket#9899
rsTA.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False

If Actn = "A" Then
    xADD = True
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_ADMINBY_TABL") = "EDAB": rsTA("AU_LANG1_TABL") = "EDL1":: rsTA("AU_LANG2_TABL") = "EDL1"
    
    rsTA("AU_DIV") = clpDiv.Text
    
    If glbSamuel And xFutureChgDeptNo Then 'Ticket #22912 create another audit for future date
    Else
        rsTA("AU_DEPTNO") = clpDept.Text
    End If
    
    rsTA("AU_TITLE") = txtTitle
    rsTA("AU_SURNAME") = txtSurname
    rsTA("AU_FNAME") = txtFName
    rsTA("AU_ADDR1") = txtAdd1
    
    If Trim$(txtAdd2) <> "" Then rsTA("AU_ADDR2") = txtAdd2
    
    rsTA("AU_CITY") = txtCity
    rsTA("AU_PROV") = clpProv.Text
    If glbLinamar Then 'Ticket #28846 Franks 07/13/2016
        rsTA("AU_PROVEMP") = clpProvEmp.Text
    End If
    rsTA("AU_COUNTRY") = txtCountry 'added by raubrey 6/16/97
    rsTA("AU_PCODE") = medPCode
    rsTA("AU_PHONE") = medTelephone
    rsTA("AU_BUSNBR") = IIf(medTele2 = "", Null, medTele2)
    rsTA("AU_DIVUPL") = clpDiv.Text
    rsTA("AU_SEX") = txtGender
    rsTA("AU_SMOKER") = ComSmoker
    rsTA("AU_DOB") = IIf(IsDate(dlpDOB), dlpDOB, Null)
    rsTA("AU_SIN") = medSIN
    rsTA("AU_DEPT_GL") = IIf(clpGLNum.Text = "", Null, clpGLNum.Text)
    rsTA("AU_PROVRES") = xProvNbr
    rsTA("AU_MSTAT") = txtMStatus
    rsTA("AU_NEWEMP") = "Y"
    rsTA("AU_PTUPL") = "FT" 'added by jrowland 9/19/97
    rsTA("AU_LOC") = clpCode(1).Text   'Laura nov 4, 1997
    
    If glbCountry <> "U.S.A." And glbEmpCountry <> "U.S.A." Then
        rsTA("AU_TD1") = "Y"
        rsTA("AU_TD1DOL") = locFEDTAX  'rsDATA("ED_TD1DOL")
        
        If clpProv.Text = "ON" Then
            rsTA("AU_PROVFORM") = "Y"
            rsTA("AU_PROVAMT") = locPROVTAX 'rsDATA("ED_PROVAMT")
        End If
        
        rsTA("AU_OLDTD1") = 0
    End If
    
    rsTA("AU_ADMINBY") = clpCode(3).Text
    
    If glbSamuel Then 'Ticket #22912 create another audit for future date
        If xFutureChgRegion Then
            'do nothing here, use SamuelFutureAudit
        Else
            rsTA("AU_REGION") = clpCode(2).Text
            If IsDate(xFutureDateRegion) Then rsTA("AU_LDAY") = xFutureDateRegion
        End If
    Else
        rsTA("AU_REGION") = clpCode(2).Text
    End If
    
    If glbSamuel Then 'Ticket #22912 create another audit for future date
        If xFutureChgSection Then
            'do nothing here, use SamuelFutureAudit
        Else
            rsTA("AU_SECTION") = clpCode(4).Text
            If IsDate(xFutureDateSection) Then rsTA("AU_FDAY") = xFutureDateSection
        End If
    Else
        rsTA("AU_SECTION") = clpCode(4).Text
    End If
    
    rsTA("AU_HOMEOPRTNBR") = clpHOME(1).Text
    rsTA("AU_HOMELINE") = clpHOME(2).Text
    rsTA("AU_HOMESHIFT") = clpHOME(4).Text
    rsTA("AU_HOMEWRKCNT") = clpHOME(3).Text
    rsTA("AU_CellPhone") = medCellPhone.Text
    rsTA("AU_PageNbr") = medPageNbr
    rsTA("AU_SSN") = medSSN
    rsTA("AU_Payroll_ID") = txtPayrollID
    
    If glbSamuel And xFutureChgDeptNo Then 'Ticket #22912
    Else
        If IsDate(dlpDeptEDate.Text) Then rsTA("AU_DEPTEDATE") = dlpDeptEDate.Text
    End If
    
    If IsDate(dlpDivEDate.Text) Then rsTA("AU_DIVEDATE") = dlpDivEDate.Text
    
    rsTA("AU_DRIVERLIC") = medDRIVERLIC
    rsTA("AU_LICPLATE1") = medLICPLATE1
    rsTA("AU_LICPLATE2") = medLICPLATE2
    
    If glbLinamar Then
        rsTA("AU_LOCKER") = medLOCKER
        rsTA("AU_COMBINATION") = medCOMBINATION
    End If
    
    rsTA("AU_TYPEVEHICLE") = medTYPEVEHICLE
    rsTA("AU_PARKPERMIT1") = medPARKPERMIT1
    rsTA("AU_PARKPERMIT2") = medPARKPERMIT2
    rsTA("AU_BADGEID") = txtBadgeID
    rsTA("AU_MIDNAME") = txtMidName
    rsTA("AU_ALIAS") = txtAlias
    
    If glbPayWeb Then
        rsTA("AU_WCB") = "Y"
    End If
    
    If glbInsync Then
        rsTA("AU_WCB") = "0"
        If glbCompSerial = "S/N - 2295W" Then
            rsTA("AU_CPP") = "O" '"0" Ticket 7118
        ElseIf glbCompSerial = "S/N - 2383W" Then 'Town of Orangeville Ticket #21492 Franks 01/25/2012
            rsTA("AU_CPP") = " "
        Else
            rsTA("AU_CPP") = "0"
        End If
    End If
    
    If glbCompSerial = "S/N - 2229W" Or glbCompSerial = "S/N - 2369W" Then 'Inscape Solutions 'TS TECH
        rsTA("AU_PROVEMP") = locPROVEMP
        rsTA("AU_UIC") = locUIC
    End If
    'Ticket #19067
    'If glbCompSerial = "S/N - 2382W" Then  ' Samuel - Ticket #18702
        rsTA("AU_DOH") = dlpDate(0).Text
    'End If
    If glbWFC Then ' Ticket #24695 Franks 11/26/2013
        rsTA("AU_NORMALR") = getWFCRetireDate(dlpDate(0).Text)
    End If
End If

If Actn = "M" Then
    
    Dim UpdateAudit As Boolean
    UpdateAudit = False
    
    Set HRChanges = New Collection
    If isChanged_Field(HRChanges, OTITLE, txtTitle) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OSNAME, txtSurname) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OFNAME, txtFName) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oMidName, txtMidName) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oAlias, txtAlias) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oProv, clpProv) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OCITY, txtCity) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OPCODE, medPCode) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oCountry, txtCountry) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OADD1, txtAdd1) Then UpdateAudit = True
    
    'Not for Town of Lasalle
    If glbCompSerial <> "S/N - 2379W" Then
        If isChanged_Field(HRChanges, OADD2, txtAdd2) Then UpdateAudit = True
    End If
    
    If isChanged_Field(HRChanges, OPHONE, medTelephone) Then UpdateAudit = True
    
    'Ticket #25469 - City of Campbell River wants Cell Phone # transferred instead as Telephone 2
    'Town of Lasalle wants Cell Phone # transferred instead as Phone 2
    If glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2458W" Then
        If isChanged_Field(HRChanges, OCellPhone, medCellPhone) Then UpdateAudit = True
    Else
        If isChanged_Field(HRChanges, OBUSNBR, medTele2) Then UpdateAudit = True
    End If
        
    'Ticket #19067
    'If glbCompSerial = "S/N - 2382W" Then  ' Samuel - Ticket #18702
        If isChanged_Field(HRChanges, oDOH, dlpDate(0)) Then UpdateAudit = True
    'End If
    
    If glbtermopen Then
        Call Passing_Changes(HRChanges, Demographices, "M", Date, glbTERM_ID, txtPayrollID)
    Else
        Call Passing_Changes(HRChanges, Demographices, "M", Date, glbLEE_ID)
    End If  'Hemu added this to open up the transfer of other changes of Term. employee - new
    
        Set HRChanges = New Collection
        If isChanged_Field(HRChanges, SavDiv, clpDiv) Then UpdateAudit = True
        If isChanged_Field(HRChanges, SavDept, clpDept) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OSEX, txtGender) Then UpdateAudit = True
        If isChanged_Field(HRChanges, ODOB, dlpDOB) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OSMOKER, ComSmoker) Then UpdateAudit = True
        
        'Town of Lasalle - Do not transfer GL # if Payment Type = S
        If glbCompSerial = "S/N - 2379W" Then
            If clpCode(2) <> "S" Then
                If isChanged_Field(HRChanges, oGLNo, clpGLNum) Then UpdateAudit = True
            End If
        Else
            If isChanged_Field(HRChanges, oGLNo, clpGLNum) Then UpdateAudit = True
        End If
        
        'If NOT City of Niagara Falls
        If glbCompSerial <> "S/N - 2276W" Then
            If isChanged_Field(HRChanges, OMSTAT, txtMStatus) Then UpdateAudit = True
        End If
        
        If isChanged_Field(HRChanges, OSIN, medSIN) Then UpdateAudit = True
        If isChanged_Field(HRChanges, SavLoc, clpCode(1)) Then UpdateAudit = True
        If isChanged_Field(HRChanges, oRegion, clpCode(2)) Then UpdateAudit = True
        If isChanged_Field(HRChanges, oAdminBy, clpCode(3)) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OSection, clpCode(4)) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OHOMELINE, clpHOME(2)) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OHOMESHIFT, clpHOME(4)) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OHOMEOPRTNBR, clpHOME(1)) Then UpdateAudit = True
        If isChanged_Field(HRChanges, oHOMEWRKCNT, clpHOME(3)) Then UpdateAudit = True
        If isChanged_Field(HRChanges, ODeptEDate, dlpDeptEDate) Then UpdateAudit = True
        If isChanged_Field(HRChanges, ODivEdate, dlpDivEDate) Then UpdateAudit = True
        
        'Ticket #25469 - City of Campbell River wants Cell Phone # transferred instead as Telephone 2.
        'It's already being transferred above, don't pass again
        If glbCompSerial <> "S/N - 2458W" Then
            If isChanged_Field(HRChanges, OCellPhone, medCellPhone) Then UpdateAudit = True
        End If
        
        If isChanged_Field(HRChanges, OPageNbr, medPageNbr) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OSSN, medSSN) Then UpdateAudit = True
        If isChanged_Field(HRChanges, oBadgeID, txtBadgeID) Then UpdateAudit = True
        If isChanged_Field(HRChanges, oPARKPERMIT1, medPARKPERMIT1) Then UpdateAudit = True
        If isChanged_Field(HRChanges, oPARKPERMIT2, medPARKPERMIT2) Then UpdateAudit = True
        
        'Ticket #24996 - City of Campbell River
        If glbCompSerial = "S/N - 2458W" Then
            If isChanged_Field(HRChanges, oOrg1, clpCode(6)) Then UpdateAudit = True
        End If
        
        If glbtermopen Then 'Hemu opened this up the transfer of other changes of Term. employee - new
            Call Passing_Changes(HRChanges, Demographices, "M", Date, glbTERM_ID, txtPayrollID) 'new
        Else 'new
            Call Passing_Changes(HRChanges, Demographices, "M", Date, glbLEE_ID)
        End If 'new
    
        Set HRChanges = New Collection
        If isChanged_Field(HRChanges, oPayrollID, txtPayrollID) Then UpdatePayrollID = True
        If UpdatePayrollID Then
            If glbtermopen Then 'Hemu opened this up the transfer of other changes of Term. employee - new
                Call Passing_Changes(HRChanges, Demographices, "R", Date, glbTERM_ID, oPayrollID)
            Else
                Call Passing_Changes(HRChanges, Demographices, "R", Date, glbLEE_ID)
            End If
            UpdateAudit = True
        End If
        
        'Ticket #24543 - Macaulay Child Development Centre
        If glbCompSerial = "S/N - 2420W" Then
            If SavOrg <> clpCode(0).Text Then UpdateAudit = True
            If oSalDist <> clpSalDist.Text Then UpdateAudit = True
        End If
        If glbLinamar Then 'Ticket #28846 Franks 07/13/2016
            If oProvEmp <> clpProvEmp.Text Then UpdateAudit = True
            If oVadim1 <> clpVadim1.Text Then UpdateAudit = True
        End If
    'End If 'Hemu commented this to open up the transfer of other changes of Term. employee

    If UpdateAudit Then GoTo MODUPD Else GoTo MODNOUPD
MODUPD:
    
    xADD = True
    
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN": rsTA("AU_UPLOAD") = "N"
    rsTA("AU_DIVUPL") = rsDATA("ED_DIV")
    rsTA("AU_TYPE") = "M"
    'rsta("AU_PTUPL") = rsDATA("ED_PT") ' added by jrowland 9/19/97 commented by Bryan 1/23/2007
    rsTA("AU_NEWEMP") = "N"
    
    If OSNAME <> txtSurname Or OFNAME <> txtFName Then
        rsTA("AU_SURNAME") = txtSurname
        rsTA("AU_FNAME") = txtFName
    End If
    
    If oMidName <> txtMidName Then rsTA("AU_MIDNAME") = txtMidName
    If oAlias <> txtAlias Then rsTA("AU_ALIAS") = txtAlias
    
 '------Changed by Jaddy 9/16/99
    DoEvents
    If OADD1 <> txtAdd1 Or OADD2 <> txtAdd2 Or OCITY <> txtCity Or OPCODE <> medPCode Or oProv <> clpProv.Text Then
        rsTA("AU_ADDR1") = txtAdd1
        If Len(txtAdd2) > 0 Then
            rsTA("AU_ADDR2") = txtAdd2
        Else
            If Not glbPayWeb Then
                rsTA("AU_ADDR2") = "-"
            End If
        End If
        rsTA("AU_CITY") = txtCity
        rsTA("AU_PCODE") = medPCode
        rsTA("AU_PROV") = clpProv.Text
        rsTA("AU_PROVRES") = xProvNbr
    End If
    If glbLinamar Then 'Ticket #28846 Franks 07/13/2016
        If oProvEmp <> clpProvEmp.Text Then rsTA("AU_PROVEMP") = clpProvEmp.Text
        If oVadim1 <> clpVadim1.Text Then rsTA("AU_VADIM1") = clpVadim1.Text
    End If
    If oCountry <> txtCountry Then rsTA("AU_COUNTRY") = txtCountry 'added by raubrey 6/16/97
    If OPHONE <> medTelephone Then rsTA("AU_PHONE") = medTelephone
    
    'Ticket #20351 Franks - AU_PAYROLL_ID should be updated for both active and term employees
    rsTA("AU_PAYROLL_ID") = txtPayrollID
    If oPayrollID <> txtPayrollID Then
        rsTA("AU_OLDPAYROLL_ID") = oPayrollID
    End If
    
    'Ticket #11737, Frank Sep 25, 2006
    If glbtermopen And glbCompSerial = "S/N - 2380W" Then   'VitalAire
        If SavDept <> clpDept.Text Then
            rsTA("AU_OLDDEPT") = SavDept
            rsTA("AU_DEPTNO") = clpDept.Text
        End If
    End If
    If Not glbtermopen Then
        If OTITLE <> txtTitle Then rsTA("AU_TITLE") = txtTitle
        If OBUSNBR <> medTele2 Then 'Added by Frank Nov 18,2003 Only pass GLNo when it was changed #5105
            rsTA("AU_BUSNBR") = IIf(medTele2 = "", Null, medTele2)
        End If
        If OSIN <> medSIN Then rsTA("AU_SIN") = medSIN   'added by jrowland 9/19/97
        If OSEX <> txtGender Then rsTA("AU_SEX") = txtGender
        If OSMOKER <> ComSmoker Then rsTA("AU_SMOKER") = ComSmoker
        If ODOB <> dlpDOB Then rsTA("AU_DOB") = dlpDOB
        If oGLNo <> clpGLNum Then 'Added by Frank Nov 18,2003 Only pass GLNo when it was changed #5105
            If clpGLNum.Text <> "" Then
                rsTA("AU_DEPT_GL") = clpGLNum.Text
            Else
                rsTA("AU_DEPT_GL") = Null
            End If
            If Len(oGLNo) > 0 Then 'Ticket #19946 Franks 03/03/2011
                rsTA("AU_OLD_GL") = oGLNo
            End If
            If glbCompSerial = "S/N - 2217W" Then 'City of Pickering
                'Ticket #20054 Franks 04/04/2011, keep old GLNO for ADP interface
                If Len(oGLNo) > 0 Then
                    rsTA("AU_VADIM1") = Left(oGLNo, 10)
                End If
            End If
        End If
        
        'Hemu - Commenting this line because it's contradicting with line above - confirmed with Frank
        'OGLNum contains the same value as oGLNo
        'If OGLNum <> clpGLNum.Text Then rsta("AU_DEPT_GL") = clpGLNum.Text
        
        If OMSTAT <> txtMStatus Then rsTA("AU_MSTAT") = txtMStatus
        If SavDiv <> clpDiv.Text Then
            rsTA("AU_OLDDIV") = SavDiv
            rsTA("AU_DIV") = clpDiv.Text
            rsTA("AU_DIVUPL") = clpDiv.Text
        End If
        If glbSamuel And xFutureChgDeptNo Then 'Ticket #22912 create another audit for future date
        Else
            If SavDept <> clpDept.Text Then
                rsTA("AU_OLDDEPT") = SavDept
                rsTA("AU_DEPTNO") = clpDept.Text
            End If
        End If
        If SavLoc <> clpCode(1).Text Then   'laura nov 4, 1997
            If SavLoc <> "" Then rsTA("AU_OLDLOC") = SavLoc
            If clpCode(1).Text <> "" Then rsTA("AU_LOC") = clpCode(1).Text
        End If
        If glbSamuel Then 'Ticket #22912 create another audit for future date
            If xFutureChgRegion Then
                'do nothing here, use SamuelFutureAudit
            Else
                If oRegion <> clpCode(2).Text Then rsTA("AU_REGION") = clpCode(2).Text
                If IsDate(xFutureDateRegion) Then rsTA("AU_LDAY") = xFutureDateRegion
            End If
        Else
            If oRegion <> clpCode(2).Text Then rsTA("AU_REGION") = clpCode(2).Text
        End If
        If oAdminBy <> clpCode(3).Text Then rsTA("AU_ADMINBY") = clpCode(3).Text
        If glbSamuel Then 'Ticket #22912 create another audit for future date
            If xFutureChgSection Then
                'do nothing here, use SamuelFutureAudit
            Else
                If OSection <> clpCode(4).Text Then rsTA("AU_SECTION") = clpCode(4).Text
                If IsDate(xFutureDateSection) Then rsTA("AU_FDAY") = xFutureDateSection
            End If
        Else
            If OSection <> clpCode(4).Text Then rsTA("AU_SECTION") = clpCode(4).Text
        End If
        If glbSamuel And xFutureChgDeptNo Then 'Ticket #22912
        Else
        If ODeptEDate <> dlpDeptEDate.Text Then If IsDate(dlpDeptEDate.Text) Then rsTA("AU_DEPTEDATE") = dlpDeptEDate.Text
        End If
        If ODivEdate <> dlpDivEDate.Text Then If IsDate(dlpDivEDate.Text) Then rsTA("AU_DIVEDATE") = dlpDivEDate.Text
        If OHOMELINE <> clpHOME(2) Then rsTA("AU_HOMELINE") = clpHOME(2)
        If OHOMESHIFT <> clpHOME(4) Then rsTA("AU_HOMESHIFT") = clpHOME(4)
        If OHOMEOPRTNBR <> clpHOME(1) Then rsTA("AU_HOMEOPRTNBR") = clpHOME(1)
        If oHOMEWRKCNT <> clpHOME(3) Then rsTA("AU_HOMEWRKCNT") = clpHOME(3)
        If OCellPhone <> medCellPhone Then rsTA("AU_CellPhone") = medCellPhone
        If OPageNbr <> medPageNbr Then rsTA("AU_PageNbr") = medPageNbr
        If OSSN <> medSSN Then rsTA("AU_SSN") = medSSN
        'Ticket #20351 Franks - AU_PAYROLL_ID should be updated for both active and term employees
        'If oPayrollID <> txtPayrollID Then
        '    rsTA("AU_PAYROLL_ID") = txtPayrollID
        '    rsTA("AU_OLDPAYROLL_ID") = oPayrollID
        'End If
        If oDRIVERLIC <> medDRIVERLIC Then rsTA("AU_DRIVERLIC") = medDRIVERLIC
        If oLICPLATE1 <> medLICPLATE1 Then rsTA("AU_LICPLATE1") = medLICPLATE1
        If oLICPLATE2 <> medLICPLATE2 Then rsTA("AU_LICPLATE2") = medLICPLATE2
        If oLOCKER <> medLOCKER Then rsTA("AU_LOCKER") = medLOCKER
        If oCOMBINATION <> medCOMBINATION Then rsTA("AU_COMBINATION") = medCOMBINATION
        If oTYPEVEHICLE <> medTYPEVEHICLE Then rsTA("AU_TYPEVEHICLE") = medTYPEVEHICLE
        If oPARKPERMIT1 <> medPARKPERMIT1 Then rsTA("AU_PARKPERMIT1") = medPARKPERMIT1
        If oPARKPERMIT2 <> medPARKPERMIT2 Then rsTA("AU_PARKPERMIT2") = medPARKPERMIT2
        If oBadgeID <> txtBadgeID Then rsTA("AU_BADGEID") = txtBadgeID
        'If glbSoroc Or glbSyndesis Then rsTA("AU_Payroll_ID") = txtPayrollID
        If Len(txtPayrollID) Then rsTA("AU_Payroll_ID") = txtPayrollID
        
        'Ticket #19067
        'If glbCompSerial = "S/N - 2382W" Then  ' Samuel - Ticket #18702
            If IsDate(oDOH) And IsDate(dlpDate(0).Text) Then
                 If CVDate(oDOH) <> CVDate(dlpDate(0).Text) Then rsTA("AU_DOH") = dlpDate(0).Text
            End If
        'End If
        
        'Ticket #24543 - Macaulay Child Development Centre
        If glbCompSerial = "S/N - 2420W" Then
            If SavOrg <> clpCode(0).Text Then
                If Len(clpCode(0).Text) > 0 Then
                    rsTA("AU_ORG") = clpCode(0).Text
                Else
                    rsTA("AU_ORG") = "-"
                End If
            End If
            If oSalDist <> clpSalDist.Text Then
                If Len(clpSalDist.Text) > 0 Then
                    rsTA("AU_SALDIST") = clpSalDist.Text
                End If
            End If
        End If
        
    End If
MODNOUPD:

End If

If Actn = "D" Then
    xADD = True
    Call DeletePayrollEmp(Date, glbLEE_ID)
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN": rsTA("AU_UPLOAD") = "N"
    rsTA("AU_SURNAME") = txtSurname
    rsTA("AU_FNAME") = txtFName
    rsTA("AU_DIVUPL") = clpDiv.Text
    rsTA("AU_NEWEMP") = "N"
    'rsta("AU_PTUPL") = rsDATA("ED_PT") ' commented by Bryan 1/23/07
    'If glbSoroc Or glbSyndesis Then rsTA("AU_Payroll_ID") = txtPayrollID
    If Len(txtPayrollID) Then rsTA("AU_Payroll_ID") = txtPayrollID
End If

If xADD Then
    If glbWFC Then 'Ticket #23564 Franks 04/15/2013
        If NewHireForms.count > 0 Then
            'If comCountryOfEmp.Text = "U.S.A." Then
                rsTA("AU_PTEDATE") = dlpDate(0).Text
            'End If
            If glbWFC_US_Ben_Trans Then 'Ticket #23247 Franks 04/22/2013
                If Len(glbTrsStatus) > 0 Then rsTA("AU_EMP") = glbTrsStatus
                If Len(glbTrsUnion) > 0 Then rsTA("AU_ORG") = glbTrsUnion
                'If Len(xBenGrpCode) > 0 Then rsTA("AU_ORG") = xBenGrpCode
                If Len(xWFCPayGroup) > 0 Then rsTA("AU_VADIM2") = xWFCPayGroup
                If Len(xWFCNGSCode) > 0 Then rsTA("AU_VADIM1") = xWFCNGSCode
            End If
        End If
    End If
    'rsTA("AU_PTUPL") = "FT" 'added by jrowland 9/19/97
    ' dkostka - 03/14/2002 - xAdd is always true, and we *don't* want to write FT to the Audit Master
    '   unless they don't have a valid FT/PT/etc entered on status/dates.
    'rsta("AU_PTUPL") = IIf(IsNull(rsDATA("ED_PT")), "FT", rsDATA("ED_PT"))' commented by Bryan 1/23/07
    rsTA("AU_DIVUPL") = clpDiv.Text
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = lblEEID
    'Ticket #19067
    'If glbCompSerial = "S/N - 2382W" Then  ' Samuel - Ticket #18702
        rsTA("AU_LDATE") = Date
        If IsDate(dlpDate(0).Text) Then
            If CVDate(dlpDate(0).Text) > Date Then
                rsTA("AU_LDATE") = dlpDate(0).Text
            End If
        End If
    'Else
    '    rsTA("AU_LDATE") = Date
    'End If
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = Actn
    rsTA.Update
    
    If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/21/2014 Franks
        Call FamilyDayAuditSync(glbLEE_ID, rsTA)
    End If
    
    ''Granite Club and Soroc city of st-thomas and County of Essex, DMuskoka,
    'glbWFC Or _ - Removed by Hemu - seems like Frank forgot to remove when doing Ticket #16616
    'Franks 07/13/2012, remove 2382 (samuel) since they don't use it
    'Add 2460 Oshawa Public Libraries ' Ticket #25323 Franks 12/16/2014
    If glbCompSerial = "S/N - 2241W" Or _
        glbSoroc Or _
        glbCompSerial = "S/N - 2191W" Or _
        glbCompSerial = "S/N - 2192W" Or _
        glbCompSerial = "S/N - 2460W" Or _
        glbCompSerial = "S/N - 2380W" Then 'Namasco 'VitalAire Canada
        If Actn = "M" Then
            If Len(glbChgTermReason) > 0 Then
                If glbChgTermReason = "***" And glbCompSerial = "S/N - 2380W" Then
                    Call DivBackInAudit(SavDiv, clpDiv.Text)
                Else
                    Call TermRehireAudit(rsTA)
                End If
            End If
        End If
    End If
End If

'Ticket #22912 Franks 12/06/2012 - begin
If glbSamuel Then
    Call SamuelFutureAudit
End If
'Ticket #22912 Franks 12/06/2012 - end
AUDITDEMO = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
'Resume Next
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '18June99 js
Resume Next
End Function

Private Sub DivBackInAudit(xOldDiv, xNewDiv)
Dim SQLQ
    SQLQ = "UPDATE HRAUDIT SET AU_DIVUPL = '" & xNewDiv & "' WHERE AU_EMPNBR = " & lblEEID & " AND AU_UPLOAD = 'N' AND AU_DIVUPL = '" & xOldDiv & "' "
    gdbAdoIhr001X.Execute SQLQ
    SQLQ = "UPDATE HRAUDIT SET AU_DIV = '" & xNewDiv & "' WHERE AU_EMPNBR = " & lblEEID & " AND AU_UPLOAD = 'N' AND AU_DIV = '" & xOldDiv & "' "
    gdbAdoIhr001X.Execute SQLQ
End Sub

Private Sub TermRehireAudit(rsTA As ADODB.Recordset)
Dim rsTC As New ADODB.Recordset
Dim SQLQ
Dim Langs 'George Apr 4,2006 #10574
    'Termination Data
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_NEWEMP") = "N"
    rsTA("AU_SURNAME") = txtSurname
    rsTA("AU_FNAME") = txtFName
    rsTA("AU_DOT") = glbChgTermDate
    rsTA("AU_TREAS") = glbChgTermReason
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = lblEEID '
    If glbCompSerial = "S/N - 2241W" Or glbCompSerial = "S/N - 2382W" Then 'Granite Club 'Namasco
        rsTA("AU_ADMINBY") = oAdminBy
    ElseIf glbCompSerial = "S/N - 2380W" Then 'VitalAire
        rsTA("AU_DIV") = SavDiv
    Else 'Soroc & WFC
        rsTA("AU_PAYROLL_ID") = oPayrollID
    End If
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire
        rsTA("AU_DIVUPL") = SavDiv
    Else
        rsTA("AU_DIVUPL") = clpDiv.Text
    End If
    If glbCompSerial = "S/N - 2460W" Then 'Oshawa Public Libraries = Ticket #25323 Franks 12/16/2014
        rsTA("AU_REGION") = oRegion
        rsTA("AU_VADIM2_TABL") = "TOLD" 'Term Old Company
    End If
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "T"
    rsTA.Update
    
    'New Hire Data
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_ADMINBY_TABL") = "EDAB": rsTA("AU_LANG1_TABL") = "EDL1":: rsTA("AU_LANG2_TABL") = "EDL1"
    rsTA("AU_DIV") = clpDiv.Text
    rsTA("AU_DEPTNO") = clpDept.Text
    rsTA("AU_TITLE") = txtTitle
    rsTA("AU_SURNAME") = txtSurname
    rsTA("AU_FNAME") = txtFName
    rsTA("AU_EMPNBR") = glbChgNewEmpnbr '
    rsTA("AU_PAYROLL_ID") = txtPayrollID
    rsTA("AU_ADDR1") = txtAdd1
    If Trim$(txtAdd2) <> "" Then rsTA("AU_ADDR2") = txtAdd2
    rsTA("AU_CITY") = txtCity
    rsTA("AU_PROV") = clpProv.Text
    rsTA("AU_COUNTRY") = txtCountry 'added by raubrey 6/16/97
    rsTA("AU_PCODE") = medPCode
    rsTA("AU_PHONE") = medTelephone
    rsTA("AU_BUSNBR") = IIf(medTele2 = "", Null, medTele2)
    rsTA("AU_DIVUPL") = clpDiv.Text
    rsTA("AU_SEX") = txtGender
    rsTA("AU_SMOKER") = ComSmoker
    rsTA("AU_DOB") = dlpDOB
    rsTA("AU_SIN") = medSIN
    rsTA("AU_DEPT_GL") = IIf(clpGLNum.Text = "", Null, clpGLNum.Text)
    'rsTA("AU_PROVRES") = xProvNbr
    rsTA("AU_MSTAT") = txtMStatus
    rsTA("AU_NEWEMP") = "Y"
    rsTA("AU_PTUPL") = "FT" 'added by jrowland 9/19/97
    rsTA("AU_LOC") = clpCode(1).Text   'Laura nov 4, 1997
    If glbCountry <> "U.S.A." And glbEmpCountry <> "U.S.A." Then
    rsTA("AU_TD1") = "Y"
    rsTA("AU_TD1DOL") = 7634 '7131
    If clpProv.Text = "ON" Then
        rsTA("AU_PROVFORM") = "Y"
        rsTA("AU_PROVAMT") = 7686
    End If
    rsTA("AU_OLDTD1") = 0
    End If
    rsTA("AU_ADMINBY") = clpCode(3).Text
    rsTA("AU_REGION") = clpCode(2).Text
    rsTA("AU_SECTION") = clpCode(4).Text
    rsTA("AU_HOMEOPRTNBR") = clpHOME(1).Text
    rsTA("AU_HOMELINE") = clpHOME(2).Text
    rsTA("AU_HOMESHIFT") = clpHOME(4).Text
    rsTA("AU_HOMEWRKCNT") = clpHOME(3).Text
    rsTA("AU_CellPhone") = medCellPhone.Text
    rsTA("AU_PageNbr") = medPageNbr
    rsTA("AU_SSN") = medSSN
    'rsTA("AU_Payroll_ID") = txtPayrollID

    If IsDate(dlpDeptEDate.Text) Then rsTA("AU_DEPTEDATE") = dlpDeptEDate.Text
    If IsDate(dlpDivEDate.Text) Then rsTA("AU_DIVEDATE") = dlpDivEDate.Text
    rsTA("AU_DRIVERLIC") = medDRIVERLIC
    rsTA("AU_LICPLATE1") = medLICPLATE1
    rsTA("AU_LICPLATE2") = medLICPLATE2
    rsTA("AU_TYPEVEHICLE") = medTYPEVEHICLE
    rsTA("AU_PARKPERMIT1") = medPARKPERMIT1
    rsTA("AU_PARKPERMIT2") = medPARKPERMIT2
    rsTA("AU_BADGEID") = txtBadgeID
    rsTA("AU_MIDNAME") = txtMidName
    rsTA("AU_ALIAS") = txtAlias
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Format(Time + 1, "HH:MM:SS")
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "A"
    rsTA.Update
    
    '------BANK Information Begin
    rsTC.Open "SELECT * FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic

    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_NEWEMP") = "N"
    'BANK 1
    rsTA("AU_DEPOSIT") = rsTC("ED_DEPOSIT")
    rsTA("AU_BRANCH") = rsTC("ED_BRANCH")
    rsTA("AU_BANK") = rsTC("ED_BANK")
    rsTA("AU_ACCOUNT") = rsTC("ED_ACCOUNT")
    rsTA("AU_TRANSITABA") = rsTC("ED_TRANSITABA")
    rsTA("AU_TRANSITABA2") = rsTC("ED_TRANSITABA2")
    rsTA("AU_TRANSITABA3") = rsTC("ED_TRANSITABA3")
    rsTA("AU_AMTDEPOSIT") = rsTC("ED_AMTDEPOSIT")
    rsTA("AU_PCDEPOSIT") = rsTC("ED_PCDEPOSIT")
    'BANK 2
    rsTA("AU_DEPOSIT2") = rsTC("ED_DEPOSIT2")
    rsTA("AU_BRANCH2") = rsTC("ED_BRANCH2")
    rsTA("AU_BANK2") = rsTC("ED_BANK2")
    rsTA("AU_ACCOUNT2") = rsTC("ED_ACCOUNT2")
    rsTA("AU_AMTDEPOSIT2") = rsTC("ED_AMTDEPOSIT2")
    'BANK3
    rsTA("AU_DEPOSIT3") = rsTC("ED_DEPOSIT3")
    rsTA("AU_BRANCH3") = rsTC("ED_BRANCH3")
    rsTA("AU_BANK3") = rsTC("ED_BANK3")
    rsTA("AU_ACCOUNT3") = rsTC("ED_ACCOUNT3")
    rsTA("AU_AMTDEPOSIT3") = rsTC("ED_AMTDEPOSIT3")
    rsTA("AU_PCDEPOSIT3") = rsTC("ED_PCDEPOSIT3")
    

    rsTA("AU_UIC") = rsTC("ED_UIC")
    rsTA("AU_WCBCODE") = rsTC("ED_WCBCODE")
    rsTA("AU_WCB") = rsTC("ED_WCB")
    rsTA("AU_CPP") = rsTC("ED_CPP")
    
    rsTA("AU_TD1CODE") = rsTC("ED_TD1CODE")
    rsTA("AU_TD1DOL") = rsTC("ED_TD1DOL")
    rsTA("AU_TD3") = rsTC("ED_TD3")
    rsTA("AU_TD1") = rsTC("ED_TD1")
    rsTA("AU_DDI") = rsTC("ED_DDI")
    rsTA("AU_PROVEMP") = rsTC("ED_PROVEMP")
    rsTA("AU_FedTax") = rsTC("ED_FedTax")
    rsTA("AU_ExtAmt") = rsTC("ED_ExtAmt")
    rsTA("AU_ProvForm") = rsTC("ED_ProvForm")
    rsTA("AU_ProvAmt") = rsTC("ED_ProvAmt")
    rsTA("AU_ExtraTax") = rsTC("ED_ExtraTax")
    rsTA("AU_ExtraTaxPC") = rsTC("ED_ExtraTaxPC")

    'Employee Status
    rsTA("AU_EMP") = rsTC("ED_EMP")
    rsTA("AU_DOH") = rsTC("ED_DOH")
    rsTA("AU_ORG") = rsTC("ED_ORG")
    rsTA("AU_SENDTE") = rsTC("ED_SENDTE")
    rsTA("AU_EMPTYPE") = rsTC("ED_EMPTYPE")
    'rsta("AU_PT") = rsTC("ED_PT")' commented by Bryan 1/23/07
    rsTA("AU_INTEL") = rsTC("ED_INTEL")
    'George Apr 4,2006 #10574
    'rsTA("AU_LANG1") = rsTC("ED_LANG1")
    'rsTA("AU_LANG2") = rsTC("ED_LANG2")
    If Len(glbChgNewEmpnbr) > 0 Then    'Ticket #22210 - causing 'Incorrect syntax near the keyword 'order' error
        Langs = Split(getLanguage(glbChgNewEmpnbr), "|")
        If Langs(0) <> "NoLang1" Then rsTA("AU_LANG1") = Langs(0) '0 is for ED_Lang1
        If Langs(1) <> "NoLang2" Then rsTA("AU_LANG2") = Langs(1) '1 is for ED_Lang2
    End If
    'George Apr 4,2006 #10574
    rsTA("AU_EMAIL") = rsTC("ED_EMAIL")
    rsTA("AU_FDAY") = rsTC("ED_FDAY")
    rsTA("AU_LDAY") = rsTC("ED_LDAY")
    rsTA("AU_OMDAY") = rsTC("ED_OMERS")
    rsTA("AU_USRDAT1") = rsTC("ED_USRDAT1")
    rsTA("AU_UNION") = rsTC("ED_UNION")
    rsTA("AU_LTHIRE") = rsTC("ED_LTHIRE")

    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbChgNewEmpnbr
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Format(Time + 1, "HH:MM:SS")
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "A"
    rsTA("AU_Payroll_ID") = txtPayrollID
    rsTA.Update
    rsTC.Close
    '------BANK Information End
    
    '------Job and Salary Information
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_NEWEMP") = "N"
    rsTC.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
    If Not rsTC.EOF Then
        rsTA("AU_JOB") = rsTC("JH_JOB")
        rsTA("AU_DHRS") = rsTC("JH_DHRS")
        rsTA("AU_PHRS") = rsTC("JH_PHRS")
        rsTA("AU_SJDATE") = rsTC("JH_SDATE")
        rsTA("AU_JREASON") = rsTC("JH_JREASON")
    End If
    rsTC.Close
    rsTC.Open "SELECT * FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
    If Not rsTC.EOF Then
        rsTA("AU_SALARY") = rsTC("SH_SALARY")
        rsTA("AU_WHRS") = rsTC("SH_WHRS")
        rsTA("AU_SALCD") = rsTC("SH_SALCD")
        rsTA("AU_SEDATE") = rsTC("SH_NEXTDAT")
        rsTA("AU_PAYP") = rsTC("SH_PAYP")
        rsTA("AU_SREASON") = rsTC("SH_SREAS1")
    End If
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbChgNewEmpnbr
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Format(Time + 1, "HH:MM:SS")
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "A"
    rsTA("AU_Payroll_ID") = txtPayrollID
    rsTA.Update
    rsTC.Close
    '------Job and Salary Information END
    
    '------Other Earnings Begin
    rsTC.Open "SELECT * FROM HREARN WHERE EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
    Do While Not rsTC.EOF
        rsTA.AddNew
        rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
        rsTA("AU_NEWEMP") = "N"
        rsTA("AU_EARN") = rsTC("EARN_TYPE")
        rsTA("AU_ADOLLAR") = rsTC("ACT_DOLLAR")
        rsTA("AU_COEFLAG") = IIf(rsTC("COST_OF_EMPLOYMENT"), "Y", "N")
        rsTA("AU_COMPNO") = "001"
        rsTA("AU_EMPNBR") = glbChgNewEmpnbr
        rsTA("AU_LDATE") = Date
        rsTA("AU_LUSER") = glbUserID
        rsTA("AU_LTIME") = Format(Time + 1, "HH:MM:SS")
        rsTA("AU_UPLOAD") = "N"
        rsTA("AU_TYPE") = "A"
        rsTA("AU_Payroll_ID") = txtPayrollID
        rsTA.Update
        rsTC.MoveNext
    Loop
    rsTC.Close
    '------Other Earnings End
    'Pay Period Code change
    If glbCompSerial = "S/N - 2241W" Or glbSoroc Then
    'If Pay Period code changed on Domographics screen, change Pay Period code on Salary screen too.
    If Left(oPayrollID, 3) <> Left(txtPayrollID, 3) Then
        SQLQ = "UPDATE HR_SALARY_HISTORY SET SH_PAYP = '" & Left(txtPayrollID, 3) & "' "
        SQLQ = SQLQ & "WHERE SH_EMPNBR= " & glbLEE_ID & " AND SH_CURRENT <>0"
        gdbAdoIhr001.Execute SQLQ
    End If
    End If
    'Pay Period Code change
End Sub

Private Function ifExistVadimPayrollID()
Dim X
Dim xBNo
Dim SQLQ
Dim rsVP As New ADODB.Recordset
'''On Error GoTo default_this
If Not glbVadim Then GoTo default_this
If txtPayrollID <> "" Then
    SQLQ = "SELECT EMP_NUM FROM EMPLOYEE WHERE EMP_NUM ='" & txtPayrollID & "'"
    rsVP.Open SQLQ, gdbPayroll, adOpenForwardOnly
    If Not rsVP.EOF Then
        ifExistVadimPayrollID = True
        rsVP.Close
        
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)

        MsgBox "Duplicate Payroll ID found in Vadim System."
        txtPayrollID.SetFocus
        Exit Function
    Else
        ifExistVadimPayrollID = False
        rsVP.Close
        Exit Function
    End If
End If
default_this:
    ifExistVadimPayrollID = False
End Function

Private Function chk_FEBASIC()
Dim VReturn%, X
Dim mSIN As String
Dim EditFlag  'ADDED BY RAUBREY 6/2/97
Dim Title$, DgDef, Response%, Msg As String
Dim xTemp
Dim xWorkVisaNo
Dim xWorkExpDate
Dim xTmpUnion
Dim xDV_ORGT1

EditFlag = True 'ADDED BY RAUBREY 6/2/97

If glbCompSerial = "S/N - 2207W" And clpProv.Text <> "ON" Then  'ADDED BY RAUBREY 6/2/97
  EditFlag = False
End If
'MsgBox glbcompserial
chk_FEBASIC = False
' dkostka - 04/09/2001 - Make Payroll ID mandatory for WFC and force format ####-####

'Bryan - Check Employee Number on Save
If fglbNewEE = True Then
    If Not chk_EMPNBR Then Exit Function
End If

If glbWFC Then
    If Len(Trim(lblEENum)) <> 8 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
    
        MsgBox "Employee Number is invalid format"
        txtPayrollID.SetFocus
        Exit Function
    End If
    
    'Ticket #29660 - Contractor Employee - Do not check for the Payroll ID validity as Contractors do not go to Payroll
    If NewHireForms.count = 0 Then
        If rsDATA("ED_EMP") = "CONP" Then
            'Do not check for Payroll ID
        Else
            If Len(txtPayrollID.Text) = 0 Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                
                MsgBox "Payroll ID is a required field"
                txtPayrollID.SetFocus
                Exit Function
            End If
            'Ticket #22481 Franks 08/27/2012
            If Len(txtPayrollID.Text) < 3 Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                
                MsgBox "Payroll ID cannot be less than 3 characters"
                txtPayrollID.SetFocus
                Exit Function
            End If
        End If
    Else
        If Len(txtPayrollID.Text) = 0 Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
            
            MsgBox "Payroll ID is a required field"
            txtPayrollID.SetFocus
            Exit Function
        End If
        'Ticket #22481 Franks 08/27/2012
        If Len(txtPayrollID.Text) < 3 Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
            
            MsgBox "Payroll ID cannot be less than 3 characters"
            txtPayrollID.SetFocus
            Exit Function
        End If
    End If
    
    'Ticket #29660 - Contractor Employee - Do not check for the Payroll ID validity as Contractors do not go to Payroll
    If NewHireForms.count = 0 Then
        If rsDATA("ED_EMP") = "CONP" Then
            'Do not check for Payroll ID
        Else
            If comCountryOfEmp.Text = "U.S.A." Then 'Ticket #16616
                'If clpCode(4).Text = "GREN" And Len(txtPayrollID.Text) < 6 Then
                If Not (Len(txtPayrollID.Text) = 6) Then
                    'Ticket #24164 - Re-ordering
                    tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                    
                    MsgBox "Invalid format for Payroll ID. Format must be ######"
                    txtPayrollID.SetFocus
                    Exit Function
                End If
            End If
        End If
    Else
        If comCountryOfEmp.Text = "U.S.A." Then 'Ticket #16616
            'If clpCode(4).Text = "GREN" And Len(txtPayrollID.Text) < 6 Then
            If Not (Len(txtPayrollID.Text) = 6) Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                
                MsgBox "Invalid format for Payroll ID. Format must be ######"
                txtPayrollID.SetFocus
                Exit Function
            End If
        End If
    End If
    
    'Ticket #24388 Franks 10/01/2013
    If Len(txtPayrollID.Text) > 0 Then
        xTemp = getDupEmpByPlantPayrollID(glbLEE_ID, txtPayrollID.Text, clpCode(4).Text)
        If xTemp > 0 Then
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
            MsgBox ("Duplicate Payroll ID found in the same Plant(Employee # " & xTemp & ")")
            txtPayrollID.SetFocus
            Exit Function
        End If
    End If
    
    
    'If Len(txtPayrollID.Text) <> 9 Then
    '    MsgBox "Invalid format for Payroll ID.  Format must be ####-####."
    '    txtPayrollID.SetFocus
    '    Exit Function
    'End If
    'If Not IsNumeric(Left(txtPayrollID.Text, 4)) Or Mid(txtPayrollID.Text, 5, 1) <> "-" Or Not IsNumeric(Right(txtPayrollID.Text, 4)) Then
    '    MsgBox "Invalid format for Payroll ID.  Format must be ####-####."
    '    txtPayrollID.SetFocus
    '    Exit Function
    'End If
    If Len(clpCode(4).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Section is a required field")
        clpCode(4).SetFocus
        Exit Function
    End If
    If lblDeptBonusDesc = "Unassigned" And Len(txtDeptBonusCtr.Text) > 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox "Invalid Bonus Reporting # entered."
        Exit Function
    End If
    
    If glbAdv And Not glbWFCFullRights Then 'Ticket #13867
        If Len(txtDeptBonusCtr.Text) = 0 Then
            If Len(txtBadgeID.Text) = 0 Then
                If NewHireForms.count > 0 Then
                    'Ticket #24164 - Re-ordering
                    tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                
                    MsgBox ("Badge ID is a required field")
                    txtBadgeID.SetFocus
                    Exit Function
                Else
                    'Ticket #19988 Franks 06/06/2011
                    'For wfc Adv change, the Badge ID can be blank if status = RET
                    If Not IsNull(rsDATA("ED_EMP")) Then
                        'If Not RSDATA("ED_EMP") = "RET" Then
                        If Not rsDATA("ED_EMP") = "RET" And Not glbtermopen Then 'Ticket #27476 Franks 08/31/2015
                            'Ticket #24164 - Re-ordering
                            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                            
                            MsgBox ("Badge ID is a required field")
                            txtBadgeID.SetFocus
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If Len(txtBadgeID.Text) > 0 Then 'Ticket #24317 Franks 09/18/2013
        xTemp = getDupEmpByPlantBadgeID(glbLEE_ID, clpCode(4).Text, txtBadgeID.Text)
        If xTemp > 0 Then
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
            MsgBox ("Duplicate Badge ID found in the same Plant(Employee # " & xTemp & ")")
            txtBadgeID.SetFocus
            Exit Function
        End If
    End If
    
    If (UCase(comCountry) = "CANADA" Or UCase(comCountry) = "U.S.A.") Then
        If Len(txtSurname.Text) > 0 Then 'Ticket #14154
            If Len(InvalidCharInStr(txtSurname.Text, glbWFCNameChars)) > 0 Then
                    'Ticket #24164 - Re-ordering
                    tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                    
                    MsgBox "Invalid character '" & InvalidCharInStr(txtSurname.Text, glbWFCNameChars) & "' in name field. "
                    txtSurname.SetFocus
                    Exit Function
            End If
        End If
        If Len(txtFName.Text) > 0 Then 'Ticket #14154
            If Len(InvalidCharInStr(txtFName.Text, glbWFCNameChars)) > 0 Then
                    'Ticket #24164 - Re-ordering
                    tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                    
                    MsgBox "Invalid character '" & InvalidCharInStr(txtFName.Text, glbWFCNameChars) & "' in name field. "
                    txtFName.SetFocus
                    Exit Function
            End If
        End If
    End If
    
    'Ticket #15396 - begin
    'If Len(txtTitle.Text) = 0 Then
    '    MsgBox "Salutation is a required field"
    '    txtTitle.SetFocus
    '    Exit Function
    'End If

    If Len(clpGLNum.Text) = 0 Then
        'MsgBox lStr("G/L Number is a required field") & lStr(" when Division is equal to ") & clpDiv.Text 'either CLIN or COMM")
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("G/L Number is a required field")
        clpGLNum.SetFocus
        Exit Function
    End If

    If Len(clpDiv.Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division is a required field")
        'clpDiv.SetFocus
        If frmWFCDIV.Visible Then txtDouDiv.SetFocus Else clpDiv.SetFocus 'Ticket #21543 Franks 02/08/2012
        Exit Function
    End If
    If Len(clpCode(1).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Location is a required field")
        clpCode(1).SetFocus
        Exit Function
    End If

    If Len(clpCode(3).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Administered By is a required field")
        clpCode(3).SetFocus
        Exit Function
    End If
        
    If Len(clpCode(2).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
        
    If Not (UCase(comCountry) = "CHINA") Then
        Msg = "Do not use ALL CAPS for name and address fields." & Chr(10) & "Please enter data using Proper/Title Case Only."
        If Len(txtSurname.Text) > 0 Then
            If AllCapitalString(txtSurname.Text) Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
            
                MsgBox Msg: txtSurname.SetFocus: Exit Function
            End If
        End If
        If Len(txtFName.Text) > 0 Then
            If AllCapitalString(txtFName.Text) Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                
                MsgBox Msg: txtFName.SetFocus: Exit Function
            End If
        End If
        'If Len(txtMidName.Text) > 0 Then 'Ticket #15832
        '    If AllCapitalString(txtMidName.Text) Then
        '        MsgBox Msg: txtMidName.SetFocus: Exit Function
        '    End If
        'End If
        If Len(txtAlias.Text) > 0 Then
            If AllCapitalString(txtAlias.Text) Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                
                MsgBox Msg: txtAlias.SetFocus: Exit Function
            End If
        End If
        If Len(txtAdd1.Text) > 0 Then
            If AllCapitalString(txtAdd1.Text, "RR") Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                
                MsgBox Msg: txtAdd1.SetFocus: Exit Function
            End If
        End If
        If Len(txtAdd2.Text) > 0 Then
            If AllCapitalString(txtAdd2.Text, "RR") Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                
                MsgBox Msg: txtAdd2.SetFocus: Exit Function
            End If
        End If
        If Len(txtCity.Text) > 0 Then
            If AllCapitalString(txtCity.Text) Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                
                MsgBox Msg: txtCity.SetFocus: Exit Function
            End If
        End If
        End If
    'Ticket #15396 - end
    
    ''Ticket #19266 Franks 01/13/2011
    ''Ticket #21119 Franks 11/14/2011 remove this logic - will turn off on 01/01/2012
    ''If Smoker = Yes, they shouldn't be able to enter the year. - Jerry
    'If ComSmoker.Text = "Yes" Then
    '    If Len(medTYPEVEHICLE.Text) > 0 Then
    '        'MsgBox "If Smoker = Yes the " & lStr("Type of Vehicle") & " must be blank."
    '        'medTYPEVEHICLE.SetFocus
    '        'Exit Function
    '        'Ticket #19955
    '        'Smoker NO to YES, affidavit field needs to be cleared automatically. Currently, they have to manually do this.
    '        medTYPEVEHICLE.Text = ""
    '    End If
    'End If
    
    ''Ticket #19266 Franks 12/23/2010
    ''Ticket #21119 Franks 11/14/2011 remove this logic - will turn off on 01/01/2012
    ''Call WFC_NGS_SmokerUpdate
    'Ticket #23301 Franks 02/20/2013
    Call WFC_SmokerChange
    
End If

If glbCompSerial = "S/N - 2443W" Then 'Walters Inc Ticket #22853 Franks 11/26/2012
    If Len(txtPayrollID.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
    
        MsgBox "Payroll ID is a required field"
        txtPayrollID.SetFocus
        Exit Function
    End If
    If Len(clpGLNum.Text) = 0 Then 'Ticket #25952 Franks 11/04/2014
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lStr("G/L Number is a required field")
        clpGLNum.SetFocus
        Exit Function
    End If
    If Len(clpCode(2).Text) = 0 Then 'Ticket #25952 Franks 11/04/2014
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lStr("Region") & " is a required field"
        clpCode(2).SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2191W" Or glbCompSerial = "S/N - 2192W" Or glbCompSerial = "S/N - 2370W" _
Or glbCompSerial = "S/N - 2396W" Or glbCompSerial = "S/N - 2410W" Or glbCompSerial = "S/N - 2436W" Then   'Or glbCompSerial = "S/N - 2373W" Then
'2396 - Oshawa CHC Ticket #17341
'2410 - Frontenac Ticket #18603
'2436 - Family Day Ticket #24729 01/22/2014 Franks
    If Len(txtPayrollID.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
        MsgBox lStr("Payroll ID is a required field")
        txtPayrollID.SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2436W" Then 'Family Day Ticket #24729 01/24/2014 Franks
    If Not Len(txtPayrollID.Text) = 9 Then
        MsgBox lStr("Payroll ID") & " must be 9 digits", vbCritical, "Error Occurred"
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        txtPayrollID.SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2350W" Then  'For Listowel
    If Len(clpDiv.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Division is a required field")
        'clpDiv.SetFocus
        If frmWFCDIV.Visible Then txtDouDiv.SetFocus Else clpDiv.SetFocus 'Ticket #21543 Franks 02/08/2012
        Exit Function
    End If
    
    If Len(clpGLNum.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("G/L Number is a required field")
        clpGLNum.SetFocus
        Exit Function
    End If
    
    If Len(clpCode(2).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
    If Len(clpCode(4).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Section is a required field")
        clpCode(4).SetFocus
        Exit Function
    End If
    
    If Len(dlpDivEDate.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division Effective Date is a required field")
        dlpDivEDate.SetFocus
        Exit Function
    End If
    
    If Len(dlpDeptEDate.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Department Effective Date is a required field")
        dlpDeptEDate.SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2453W" Then 'Gander Ticket #24518 Franks 12/05/2014
    If Len(clpGLNum.Text) = 0 Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lStr("G/L Number is a required field")
        clpGLNum.SetFocus
        Exit Function
    End If
    If Len(clpCode(1).Text) < 1 Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lStr("Location is a required field")
        clpCode(1).SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2487W" Then 'City of Kenora Ticket #30217 Franks 06/12/2017
    If Len(clpGLNum.Text) = 0 Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lStr("G/L Number is a required field")
        clpGLNum.SetFocus
        Exit Function
    End If
End If
If (glbCompSerial = "S/N - 2388W") Then   'For DNSSAB Ticket #14475
    If Len(clpDiv.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division is a required field")
        'clpDiv.SetFocus
        If frmWFCDIV.Visible Then txtDouDiv.SetFocus Else clpDiv.SetFocus 'Ticket #21543 Franks 02/08/2012
        Exit Function
    End If
    
    If Len(clpGLNum.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("G/L Number is a required field")
        clpGLNum.SetFocus
        Exit Function
    End If
    If Len(clpCode(1).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Location is a required field")
        clpCode(1).SetFocus
        Exit Function
    End If
    If Len(dlpDivEDate.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division Effective Date is a required field")
        dlpDivEDate.SetFocus
        Exit Function
    End If
    
    If Len(dlpDeptEDate.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Department Effective Date is a required field")
        dlpDeptEDate.SetFocus
        Exit Function
    End If
End If
If (glbCompSerial = "S/N - 2394W") Then   ' St. John's Rehab Hospital - Ticket #14572
    If Len(clpDiv.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division is a required field")
        'clpDiv.SetFocus
        If frmWFCDIV.Visible Then txtDouDiv.SetFocus Else clpDiv.SetFocus 'Ticket #21543 Franks 02/08/2012
        Exit Function
    End If
    
    If Len(clpGLNum.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("G/L Number is a required field")
        clpGLNum.SetFocus
        Exit Function
    End If
    If Len(clpCode(1).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Location is a required field")
        clpCode(1).SetFocus
        Exit Function
    End If
    If Len(clpCode(2).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
    If Len(clpCode(3).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Administered By is a required field")
        clpCode(3).SetFocus
        Exit Function
    End If
    If Len(dlpDivEDate.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division Effective Date is a required field")
        dlpDivEDate.SetFocus
        Exit Function
    End If
    
    If Len(dlpDeptEDate.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Department Effective Date is a required field")
        dlpDeptEDate.SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2259W" Then  'For County of Oxford
    If Len(clpDiv.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division is a required field")
        'clpDiv.SetFocus
        If frmWFCDIV.Visible Then txtDouDiv.SetFocus Else clpDiv.SetFocus 'Ticket #21543 Franks 02/08/2012
        Exit Function
    End If
    
    If Len(clpCode(1).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Location is a required field")
        clpCode(1).SetFocus
        Exit Function
    End If
    
'    If Len(clpCode(3).Text) = 0 Then
'        MsgBox lStr("Administered By is a required field")
'        clpCode(3).SetFocus
'        Exit Function
'    End If
    If Len(clpCode(2).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
    
    'Ticket #15590
    'If Len(clpGLNum.Text) = 0 Then
    '    MsgBox lStr("G/L Number is a required field")
    '    clpGLNum.SetFocus
    '    Exit Function
    'End If
    
End If
If glbCompSerial = "S/N - 2344W" Then 'Ticket #24988 Franks 01/28/2014 'cascade
    If Len(clpDiv.Text) = 0 Then
        MsgBox lStr("Division is a required field")
        'clpDiv.SetFocus
        If frmWFCDIV.Visible Then txtDouDiv.SetFocus Else clpDiv.SetFocus 'Ticket #21543 Franks 02/08/2012
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2418W" Then  'Ticket #17786
    If Len(clpDiv.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division is a required field")
        'clpDiv.SetFocus
        If frmWFCDIV.Visible Then txtDouDiv.SetFocus Else clpDiv.SetFocus 'Ticket #21543 Franks 02/08/2012
        Exit Function
    End If
    
    If Len(clpCode(1).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Location is a required field")
        clpCode(1).SetFocus
        Exit Function
    End If
    
    'If Len(clpCode(3).Text) = 0 Then
     '   MsgBox lStr("Administered By is a required field")
     '   clpCode(3).SetFocus
      '  Exit Function
   ' End If
    
    If Len(clpCode(4).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Section") & " is a required field"
        clpCode(4).SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2410W" Then 'Frontenac Ticket #18603
    'If Len(clpCode(3).Text) = 0 Then
    '    MsgBox lStr("Administered By is a required field")
    '    clpCode(3).SetFocus
    '    Exit Function
    'End If
    'Ticket #23857 Franks 05/30/2013
    If Len(clpCode(4).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Section") & " is a required field"
        clpCode(4).SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2383W" Then 'Town of Orangeville Ticket #18844 Franks 01/13/2011
    If Len(clpCode(3).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Administered By is a required field")
        clpCode(3).SetFocus
        Exit Function
    End If
    If Len(clpDiv.Text) = 0 Then 'Ticket #23189 Franks 02/07/2013
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division") & " is a required field"
        clpDiv.SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2485W" Then 'Mississaugas of Scugog Island First Nation -Ticket #28652  Franks 07/31/2017
    If Len(clpDiv.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lStr("Division") & " is a required field"
        clpDiv.SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2439W" Or glbCompSerial = "S/N - 2484W" Then
'OK Tire Ticket #22503 Franks 09/14/2012
'Ticket #28396 Franks 03/08/2017 PeterboroughFHT
    If Len(clpCode(3).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Administered By is a required field")
        clpCode(3).SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2483W" Then 'Scott Steel Ticket #28262 Franks 06/07/2016
    If Len(txtPayrollID) < 1 Then 'Ticket #29077 Franks 09/23/2016
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        MsgBox "Payroll ID is a required field"
        txtPayrollID.SetFocus
        Exit Function
    End If
    If Len(clpCode(3).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Administered By is a required field")
        clpCode(3).SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2409W" Then 'Delisle Youth Services - Ticket #27798
    If Len(txtPayrollID) < 1 Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        MsgBox "Payroll ID is a required field"
        txtPayrollID.SetFocus
        Exit Function
    End If
    If Len(clpCode(3).Text) = 0 Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lStr("Administered By is a required field")
        clpCode(3).SetFocus
        Exit Function
    End If
    If Len(clpCode(2).Text) = 0 Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
    If Len(clpCode(4).Text) = 0 Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lStr("Section is a required field")
        clpCode(4).SetFocus
        Exit Function
    End If
End If

If glbGP Then 'George Mar 8,2006 Great Plains 9965
    If (glbCompSerial = "S/N - 2259W") Then   'For County of Oxford
        If Len(clpGLNum.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
            MsgBox lStr("G/L Number is a required field")
            clpGLNum.SetFocus
            Exit Function
        End If
    End If
End If
If glbCompSerial = "S/N - 2373W" Then 'District Municipality of South Muskoka
    If Len(clpCode(2).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2454W" Then 'Showa Canada 'Ticket #24659
    If Len(txtBadgeID.Text) = 0 Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        MsgBox ("Badge ID is a required field")
        txtBadgeID.SetFocus
        Exit Function
    End If
    If Len(clpCode(2).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2366W" Or glbCompSerial = "S/N - 2443W" Then   ' FOR Family Youth Child Services of Muskoka or Walters Inc Ticket #23278
    If Len(clpDiv.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division is a required field")
        'clpDiv.SetFocus
        If frmWFCDIV.Visible Then txtDouDiv.SetFocus Else clpDiv.SetFocus 'Ticket #21543 Franks 02/08/2012
        Exit Function
    End If
End If

If glbVadim Then
    If glbCompSerial = "S/N - 2373W" Then   'Ticket #19113 - District Municipality of Muskoka
        'Do not allow to change Payroll ID
        If Len(txtPayrollID) < 1 Then
            'Payroll ID same as Employee #
            txtPayrollID.Text = glbLEE_ID
        End If
    Else
        If Len(txtPayrollID) < 1 Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
            MsgBox "Payroll ID is a required field"
            txtPayrollID.SetFocus
            Exit Function
        End If
    End If
    If Not VadimControl("Check") Then Exit Function
End If
If glbLambton Then
    If Len(clpDiv.Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division is a required field")
        'clpDiv.SetFocus
        If frmWFCDIV.Visible Then txtDouDiv.SetFocus Else clpDiv.SetFocus 'Ticket #21543 Franks 02/08/2012
        Exit Function
    End If
    If Len(clpCode(1).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Location is a required field")
        clpCode(1).SetFocus
        Exit Function
    End If
    If clpCode(2).Text = "S" Then
        If Len(clpGLNum) <> 0 Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
            MsgBox lStr("GL # must be empty for Salaried Employee (Region = ""S"")")
            clpGLNum.SetFocus
            Exit Function
        End If
    ElseIf InStr("H,C,P,F,", clpCode(2).Text & ",") <> 0 Then
        If Len(clpGLNum) = 0 Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
            MsgBox lStr("GL # is a required field for " & clpCode(2).Caption & " Employee")
            clpGLNum.SetFocus
            Exit Function
        End If
    End If
End If
If glbCompSerial = "S/N - 2363W" Then ' CITY OF K LAKES
    If Len(clpCode(2).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
    If InStr("H,C,P,F,", clpCode(1).Text & ",") <> 0 Then
        If Len(clpGLNum) = 0 Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
            MsgBox lStr("GL # is a required field for " & clpCode(1).Caption & " Employee")
            clpGLNum.SetFocus
            Exit Function
        End If
    End If
End If

'Ticket #25469 - City of Campbell River
If glbCompSerial = "S/N - 2458W" Then
    'They do not want this to be mandatory any more
    'If Len(clpGLNum.Text) = 0 Then
    '    'Ticket #24164 - Re-ordering
    '    tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    '
    '    MsgBox lStr("G/L Number is a required field")
    '    clpGLNum.SetFocus
    '    Exit Function
    'End If
    
    If Len(clpCode(2).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
    
    If Len(clpCode(4).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Section is a required field")
        clpCode(4).SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2351W" Then ' Burlington Technologies
    If Len(clpCode(1).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Location is a required field")
        clpCode(1).SetFocus
        Exit Function
    End If
End If

'Ticket #24396 - Macaulay Child Development Centre
If glbCompSerial = "S/N - 2420W" Then
    If Len(clpCode(1).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lStr("Location is a required field")
        clpCode(1).SetFocus
        Exit Function
    End If
    If Len(clpCode(2).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lStr("Region") & " is a required field"
        clpCode(2).SetFocus
        Exit Function
    End If
    If Len(clpSalDist.Text) < 1 Then 'Ticket #24557 Franks 12/11/2014
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lblSalDist.Caption & " is a required field"
        clpSalDist.SetFocus
        Exit Function
    End If
    If Len(txtPayrollID.Text) = 0 Then 'Ticket #24557 Franks 07/07/2015
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        MsgBox "Payroll ID is a required field"
        txtPayrollID.SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2326W" Or glbSyndesis Then 'Soroc
    If Len(txtPayrollID.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
    
        MsgBox "Payroll ID is a required field"
        txtPayrollID.SetFocus
        Exit Function
    End If
    'Don't check the duplicate for Syndesis, the Payroll ID may be duplicated for Canada and U.S.A
    If glbCompSerial = "S/N - 2326W" Then
        If Not rsDATA.EOF Then
            SorocOPayrollID = IIf(IsNull(rsDATA("ED_PAYROLL_ID")), "", rsDATA("ED_PAYROLL_ID"))
            If SorocOPayrollID <> txtPayrollID Then
                If Len(txtPayrollID) > 0 Then
                    Dim rsTmp As New ADODB.Recordset
                    Dim SQLQ
                    SQLQ = "SELECT ED_EMPNBR, ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR <> " & glbLEE_ID & " "
                    SQLQ = SQLQ & "AND ED_PAYROLL_ID = '" & txtPayrollID & "' "
                    rsTmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                    If Not rsTmp.EOF Then
                            'Ticket #24164 - Re-ordering
                            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                            
                            MsgBox "Employee #" & rsTmp("ED_EMPNBR") & " has the same Payroll ID already"
                            txtPayrollID.SetFocus
                            rsTmp.Close
                            Exit Function
                    Else
                        rsTmp.Close
                        SQLQ = "SELECT ED_EMPNBR, ED_PAYROLL_ID FROM Term_HREMP WHERE ED_EMPNBR <> " & glbLEE_ID & " "
                        SQLQ = SQLQ & "AND ED_PAYROLL_ID = '" & txtPayrollID & "' "
                        rsTmp.Open SQLQ, gdbAdoIhr001X, adOpenStatic
                        If Not rsTmp.EOF Then
                            'Ticket #24164 - Re-ordering
                            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                            
                            MsgBox "Employee #" & rsTmp("ED_EMPNBR") & " has the same Payroll ID already"
                            txtPayrollID.SetFocus
                            rsTmp.Close
                            Exit Function
                        End If
                    End If
                    rsTmp.Close
                End If
            End If
        End If
    End If
End If

'Ticket #24443 - North York Community House
If glbCompSerial = "S/N - 2391W" Then
    If Len(txtPayrollID.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
    
        MsgBox "Payroll ID is a required field"
        txtPayrollID.SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2174W" Or glbCompSerial = "S/N - 2469W" Then 'KH CAS Ticket #23382 Franks 07/11/2014
    If Len(txtPayrollID.Text) = 0 Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        MsgBox "Payroll ID is a required field"
        txtPayrollID.SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2466W" Then 'Chiefs of Ontario Ticket #25879 Franks 09/25/2014
    If Len(txtPayrollID.Text) = 0 Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        MsgBox "Payroll ID is a required field"
        txtPayrollID.SetFocus
        Exit Function
    End If
End If
'If glbCompSerial = "S/N - 2451W" Then 'Decor Ticket #23848
'    If Len(txtPayrollID.Text) = 0 Then
'        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
'        MsgBox "Payroll ID is a required field"
'        txtPayrollID.SetFocus
'        Exit Function
'    End If
'    If Len(clpCode(3).Text) = 0 Then
'        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
'        MsgBox lStr("Administered By is a required field")
'        clpCode(3).SetFocus
'        Exit Function
'    End If
'End If
If Len(txtSurname) < 1 Then
    'Ticket #24164 - Re-ordering
    tbDemographics.SelectedItem = tbDemographics.Tabs(1)
    
    MsgBox "Surname is a required field"
    txtSurname.SetFocus
    Exit Function
End If

If Len(txtFName) < 1 Then
    'Ticket #24164 - Re-ordering
    tbDemographics.SelectedItem = tbDemographics.Tabs(1)

    MsgBox lStr("First Name is a required field")
    txtFName.SetFocus
    Exit Function
End If


If (Not gSec_Show_ADDRESS) And (Len(txtAdd1) < 1) Then
    'Ticket #24164 - Re-ordering
    tbDemographics.SelectedItem = tbDemographics.Tabs(1)
    
    MsgBox "First Address Line is a required field." & vbCrLf & "To allow the data entry for 'Address', please check the Security Setup for 'Address'." & vbCrLf & vbCrLf & "Cancelling the changes."
    Exit Function
ElseIf gSec_Show_ADDRESS Then
    If Len(txtAdd1) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
        MsgBox "First Address Line is a required field"
        txtAdd1.SetFocus
        Exit Function
    End If
End If

''Ticket #24565- District Municipality of South Muskoka
'If glbCompSerial = "S/N - 2373W" And Len(txtAdd2) > 0 Then
'    'Address Line 2 can only be numeric
'    If Not IsNumeric(txtAdd2) Then
'        'Ticket #24164 - Re-ordering
'        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
'
'        MsgBox "Address 2 can only by numeric"
'        txtAdd2.SetFocus
'        Exit Function
'    End If
'End If

If (Not gSec_Show_ADDRESS) And (Len(txtCity) < 1) Then
    'Ticket #24164 - Re-ordering
    tbDemographics.SelectedItem = tbDemographics.Tabs(1)
    
    MsgBox "City is a required field." & vbCrLf & "To allow the data entry for 'Address', please check the Security Setup for 'Address'."
    Exit Function
ElseIf gSec_Show_ADDRESS Then
    If Len(txtCity) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
        MsgBox "City is a required field"
        txtCity.SetFocus
        Exit Function
    End If
End If

If (Not gSec_Show_ADDRESS) And ((Len(clpProv.Text) < 1) Or (clpProv.Caption = "Unassigned")) Then
    'Ticket #24164 - Re-ordering
    tbDemographics.SelectedItem = tbDemographics.Tabs(1)
    
    MsgBox "Province is a required field." & vbCrLf & "To allow the data entry for 'Address', please check the Security Setup for 'Address'."
    Exit Function
ElseIf gSec_Show_ADDRESS Then
    If Len(clpProv.Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
    
        MsgBox "Province is a required field"
        clpProv.SetFocus
        Exit Function
    Else
        If clpProv.Caption = "Unassigned" Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
            MsgBox "Invalid Province"
            clpProv.SetFocus
            Exit Function
        End If
    End If
End If

If glbLinamar Then 'Ticket #28846 Franks 07/13/2016
    'Ticket #29759 Franks 02/21/2017           ********************************************
    If NewHireForms.count > 0 Then 'New Hire
        If Len(txtPayrollID.Text) > 0 Then
            'check if Payroll ID is duplicate
            If IsLinDupPayrollID(lblEEID, txtPayrollID.Text, "Y", "Y", "") Then
                MsgBox "Duplicate Payroll ID."
                Exit Function
            End If
        End If
    End If
    'Ticket #29759 Franks 02/21/2017 for Edit - begin
    If txtPayrollID.Enabled Then 'Edit
        If NewHireForms.count = 0 Then
            If Len(txtPayrollID.Text) > 0 Then
                'check if Payroll ID is duplicate
                If IsLinDupPayrollID(lblEEID, txtPayrollID.Text, "N", "Y", "") Then
                    MsgBox "Duplicate Payroll ID."
                    Exit Function
                End If
            End If
        End If
    End If
    'Ticket #29759 Franks 02/21/2017 for Edit - end *****************************************
    
    If Len(clpProvEmp.Text) = 0 Then clpProvEmp.Text = clpProv.Text
    If clpProvEmp.Caption = "Unassigned" Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        MsgBox "Invalid Province of Employment."
        If clpProvEmp.Enabled Then clpProvEmp.SetFocus
        Exit Function
    End If
    
    'Ticket #29759 Franks 02/14/2017
    If Len(clpVadim1.Text) > 0 And clpVadim1.Caption = "Unassigned" Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(3)
        MsgBox lStr("Vadim Field 1") & " must be valid."
        If clpVadim1.Enabled Then clpVadim1.SetFocus
        Exit Function
    End If

End If

If EditFlag Then
    If (Not gSec_Show_ADDRESS) And (Len(medPCode) < 1) Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
        MsgBox "Postal/Zip Code is a required field." & vbCrLf & "To allow the data entry for 'Address', please check the Security Setup for 'Address'."
        Exit Function
    ElseIf gSec_Show_ADDRESS Then
        If Len(medPCode) < 1 Then
            If comCountry = "U.S.A." Or comCountry = "MEXICO" Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                
                MsgBox "Zip Code is a required field"
                medPCode.SetFocus
                Exit Function
            End If
            If comCountry = "CANADA" Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                
                MsgBox "Postal Code is a required field"
                medPCode.SetFocus
                Exit Function
            End If
        End If
    End If
End If

If (Not gSec_Show_DOB) And Len(dlpDOB) < 1 Then
    'Ticket #24164 - Re-ordering
    tbDemographics.SelectedItem = tbDemographics.Tabs(1)
    
    MsgBox "Birth Date is a required field." & vbCrLf & "To allow the data entry for 'Birth Date', please check the Security Setup for 'Birth Date'." & vbCrLf & vbCrLf & "Cancelling the changes."
    'Call cmdCancel_Click
    Exit Function
ElseIf gSec_Show_DOB Then
    If Len(dlpDOB) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
        MsgBox "Birth Date is a required field"
        dlpDOB.SetFocus
        Exit Function
    Else
        If Not IsDate(dlpDOB) Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
            
            MsgBox "Invalid Birth Date"
            dlpDOB.SetFocus
            Exit Function
        ElseIf Year(dlpDOB) > 2070 Or Year(dlpDOB) < 1900 Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
            MsgBox "Invalid Year of Birth Date"
            dlpDOB.SetFocus
            Exit Function
        'ElseIf CVDate(Format(dlpDOB, "mm/dd/yyyy")) > CVDate(Format(Now, "mm/dd/yyyy")) Then
        'Ticket #26814 Franks 03/16/2015 - the function above not work for date format dd/MM/yyyy
        ElseIf CVDate(dlpDOB.Text) > CVDate(Now) Then
            'Ticket #24338 - Date Validation
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
            MsgBox "Birth Date cannot be greater than today"
            dlpDOB.SetFocus
            Exit Function
        ElseIf fglbNewEE Then
            If (DateDiff("yyyy", CVDate(Format(dlpDOB, "mm/dd/yyyy")), CVDate(Format(Now, "mm/dd/yyyy"))) < 18) Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
            
                MsgBox "Employee's Age is less than 18 years", vbExclamation, "Warning"
            ElseIf (DateDiff("yyyy", CVDate(Format(dlpDOB, "mm/dd/yyyy")), CVDate(Format(Now, "mm/dd/yyyy"))) > 71) Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                
                MsgBox "Employee's Age is greater than 71 years", vbExclamation, "Warning"
            End If
        ElseIf IsDate(dlpDate(0)) Then
            'Ticket #24338 - Date Validation
            If CVDate(Format(dlpDOB, "mm/dd/yyyy")) >= CVDate(Format(dlpDate(0), "mm/dd/yyyy")) Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
            
                MsgBox "Birth Date cannot be greater than " & lStr("Original Hire")
                dlpDOB.SetFocus
                Exit Function
            End If
        Else
            'WFC Pension Outstanding Tasks By Dec1009.doc in W:\2008 Projects\Pension\Pension Phase II
            'Frank 12/16/2009
            'Ticket #18804 - remove this logic
            'If glbWFC Then
            '    If Not (ODOB = dlpDOB.Text) Then
            '        glbAccessPswd = False
            '        frmAccessPswd.Show 1
            '        If glbAccessPswd = False Then   'Access Denied
            '            MsgBox "Can not change Original Hire Date."
            '            dlpDOB.SetFocus
            '            Exit Function
            '        End If
            '    End If
            'End If
        End If
    End If
End If

'Ticket #26340 - Do not allow blank
If Len(comCountry.Text) = 0 Then
    'Ticket #24164 - Re-ordering
    tbDemographics.SelectedItem = tbDemographics.Tabs(1)
    
    MsgBox ("Country is a required field")
    comCountry.SetFocus
    Exit Function
End If

'Ticket #13240
'Make Make County of Employment mandatory
If Len(comCountryOfEmp.Text) = 0 Then
    If Len(comCountry.Text) > 0 Then
        comCountryOfEmp.Text = comCountry.Text
        txtCountryOfEmp.Text = comCountry.Text
    Else
        'Ticket #26340 - Do not allow blank
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
        MsgBox ("Country of Employment is a required field")
        comCountryOfEmp.SetFocus
        Exit Function
    End If
    'MsgBox ("Country of Employment is a required field")
    'comCountryOfEmp.SetFocus
    'Exit Function
End If

If EditFlag Then 'line ADDED BY RAUBREY 6/2/97
    If Len(medSIN) < 1 Then
        If Not glbCompSerial = "S/N - 2415W" Then 'Ticket #16982 SPC- Volunteer System
            If gSec_Show_SIN_SSN Then
                If comCountry = "CANADA" Then
                    'Ticket #24164 - Re-ordering
                    tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                    
                    MsgBox "Social Insurance Number is a required field"
                    medSIN.SetFocus
                    Exit Function
                End If
                If comCountry = "U.S.A." Or comCountry = "MEXICO" Then
                    'Ticket #24164 - Re-ordering
                    tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                    
                    MsgBox "Social Security Number is a required field"
                    medSIN.SetFocus
                    Exit Function
                End If
            End If
        End If
    Else
        Dim ValidSIN
        ValidSIN = False
        If comCountry = "BAHAMAS" Then
            If Len(medSIN) <> 8 Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                
                MsgBox "Invalid National Ins" & IIf(glbLinamar, "", " - if Unassigned set to 99999999")
                medSIN.SetFocus
                Exit Function
            End If
        Else
            If Len(medSIN) <> 9 Then
                If comCountry = "CANADA" Then
                    'Ticket #24164 - Re-ordering
                    tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                    
                    MsgBox "Invalid SIN" & IIf(glbLinamar, "", " - if Unassigned set to 999-999-999")
                    medSIN.SetFocus
                    Exit Function
                ElseIf comCountry = "U.S.A." Then 'Or comCountry = "MEXICO" Then
                    'Ticket #24164 - Re-ordering
                    tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                    
                    MsgBox "Invalid SSN" & IIf(glbLinamar, "", " - if Unassigned set to 999-99-9999")
                    medSIN.SetFocus
                    Exit Function
                ElseIf comCountry = "MEXICO" Then
                    If Len(medSIN) <> 11 Then
                        'Ticket #24164 - Re-ordering
                        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                        
                        MsgBox "Invalid SSN" & IIf(glbLinamar, "", " - if Unassigned set to 999-99-9999")
                        medSIN.SetFocus
                        Exit Function
                    End If
                Else
                    'MsgBox "Invalid National Ins - if Unassigned set to 999999999"
                End If
                'MedSIN.SetFocus
                'Exit Function
            Else
                If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/21/2014 Franks
                    If medSIN.Text = "999999999" Then
                        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                        MsgBox "Cannot use '999999999' as SIN"
                        medSIN.SetFocus
                        Exit Function
                    End If
                End If
                If comCountry = "CANADA" And (medSIN <> "999999999" Or glbLinamar) Then
                    If Not SIN_chk(medSIN.Text) Then
                        'Ticket #24164 - Re-ordering
                        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                        
                        MsgBox "Invalid SIN" & IIf(glbLinamar, "", "- if Unassigned set to 999-999-999")
                        medSIN.SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    If Len(medSSN) > 0 Then
        If Len(medSSN) <> 9 And glbCompSerial <> "S/N - 2376W" Then
            If comCountry = "CANADA" Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                
                MsgBox "Invalid SSN" & IIf(glbLinamar, "", " - if Unassigned set to 999-99-9999")
                medSSN.SetFocus
                Exit Function
            ElseIf comCountry = "U.S.A." Or comCountry = "MEXICO" Then
                'Ticket #24164 - Re-ordering
                tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                
                MsgBox "Invalid SIN" & IIf(glbLinamar, "", " - if Unassigned set to 999-999-999")
                medSSN.SetFocus
                Exit Function
            ElseIf comCountry = "BAHAMAS" Then
                'MsgBox "Invalid SSN - if Unassigned set to 99999999"
            Else
                'MsgBox "Invalid SSN - if Unassigned set to 999999999"
            End If
        Else
            If glbCompSerial = "S/N - 2376W" And Len(medSSN) <> 10 Then
                'Ticket #17394
                'MsgBox "Invalid Status Number - if Unassigned set to 9999999999"
                'medSSN.SetFocus
                'Exit Function
            ElseIf comCountry = "U.S.A." And medSSN <> "999999999" Then  'Or comCountry = "MEXICO"
                'Ticket #17394
                'If Not SIN_chk(medSSN.Text) Then
                '    MsgBox "Invalid SIN" & IIf(glbLinamar, "", " - if Unassigned set to 999-999-999")
                '    medSSN.SetFocus
                '    Exit Function
                'End If
            End If
        End If
    End If
End If 'line ADDED BY RAUBREY 6/2/97

'Add by Franks Jan 29,2002 for checking duplicate SIN/SSN
If (Len(medSIN) > 0 And (medSIN <> "999999999" Or glbLinamar)) Then
    If CheckSINSSN(lblEEID, medSIN, "SIN") Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
        Load frmMsgBox
        frmMsgBox.cmdCancel.Caption = "No"
        frmMsgBox.cmdOk.Caption = "Yes"
        frmMsgBox.Caption = "Duplicate " & Mid(medSIN.Tag, 4) & " number found "
        Msg$ = ""
        If glbLinamar Then
            frmMsgBox.cmdPrint.Visible = False
            frmMsgBox.txtLongMsg.Visible = False
        Else
            If (glbCompSerial = "S/N - 2350W" Or glbCompSerial = "S/N - 2382W") And flgDupSINSSN_Term Then
                'Listowel - Ticket #14040
                'Ticket #19937 2382W for Samuel -  Franks 05/06/2011
                If Len(Trim(fDupSIN)) > 0 Then
                    frmMsgBox.txtLongMsg = fDupSIN & vbNewLine & vbNewLine & fDupSIN_Term
                Else
                    frmMsgBox.txtLongMsg = fDupSIN_Term
                End If
            Else
                frmMsgBox.txtLongMsg = fDupSIN
            End If
        End If
        
        locUploadWithoutCheck = False 'Ticket #19937
        'If False And glbCompSerial = "S/N - 2382W" And flgDupSINSSN_Term And NewHireForms.count > 0 And gSec_Inq_Rehire Then
        If glbCompSerial = "S/N - 2382W" And flgDupSINSSN_Term And NewHireForms.count > 0 And gSec_Inq_Rehire Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
            
            'Ticket #19937 2382W for Samuel -  Franks 05/06/2011 - begin
            Msg$ = "Found same S.I.N in Terminated employee list."
            Msg$ = Msg$ & Chr(10) & "You need to do Rehire for this employee."
            Msg$ = Msg$ & Chr(10) & "Press Yes to Rehire or No to edit"
            frmMsgBox.lblQuestion = Msg$
            frmMsgBox.Show 1
            If glbMsgBoxResult = vbCancel Then
                medSIN.SetFocus
                Exit Function
            Else
                Call funGoToRehire(medSIN.Text)
                'Call funGoToRehire("466066859")
                locUploadWithoutCheck = True
                chk_FEBASIC = True
                Exit Function
                
                'Unload Me
            End If
            'Ticket #19937 2382W for Samuel -  Franks 05/06/2011 - end
        ElseIf glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/21/2014 Franks
            'Multiple Payroll IDs for 1 Employee Options
        Else
            
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
            
            Msg$ = "Are you sure you wish to accept it?"
            Msg$ = Msg$ & Chr(10) & "Press Yes to accept or No to edit"
            frmMsgBox.lblQuestion = Msg$
            frmMsgBox.Show 1
            If glbMsgBoxResult = vbCancel Then
                medSIN.SetFocus
                Exit Function
            End If
        End If
    End If
End If

If (Len(medSSN) > 0 And medSSN <> "999999999") Then
    If CheckSINSSN(lblEEID, medSIN, "SSN") Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
        Load frmMsgBox
        frmMsgBox.cmdCancel.Caption = "No"
        frmMsgBox.Caption = "Duplicate " & Mid(medSSN.Tag, 4) & " number found "
        Msg$ = ""
        If glbLinamar Then
            frmMsgBox.cmdPrint.Visible = False
            frmMsgBox.txtLongMsg.Visible = False
        Else
            If (glbCompSerial = "S/N - 2350W" Or glbCompSerial = "S/N - 2382W") And flgDupSINSSN_Term Then
                'Listowel - Ticket #14040
                'Ticket #19937 2382W for Samuel -  Franks 05/06/2011
                If Len(Trim(fDupSSN)) > 0 Then
                    frmMsgBox.txtLongMsg = fDupSSN & vbNewLine & vbNewLine & fDupSSN_Term
                Else
                    frmMsgBox.txtLongMsg = fDupSSN_Term
                End If
            Else
                frmMsgBox.txtLongMsg = fDupSSN
            End If
        End If
        
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
        Msg$ = "Are you sure you wish to accept it?"
        Msg$ = Msg$ & Chr(10) & "Press Yes to accept or No to edit"
        frmMsgBox.lblQuestion = Msg$
        frmMsgBox.Show 1
        If glbMsgBoxResult = vbCancel Then
            medSSN.SetFocus
            Exit Function
        End If
    End If
End If

'Ticket #19067
'Surrey Place Centre - Ticket #19067
'If (glbCompSerial = "S/N - 2347W") Then
    If Len(dlpDate(0).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
        MsgBox lStr("Original Hire Date") & " is a required field."
        dlpDate(0).SetFocus
        Exit Function
    Else
        If Not IsDate(dlpDate(0).Text) Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
            
            MsgBox "Invalid " & lStr("Original Hire Date")
            dlpDate(0).SetFocus
            Exit Function
        ElseIf Len(Year(dlpDate(0).Text)) = 3 Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
            
            MsgBox "Invalid Year of " & lStr("Original Hire Date")
            dlpDate(0).SetFocus
            Exit Function
        End If
    End If
'End If


'Ticket #25469 - City of Campbell River - Not Mandatory
'Add by Franks Jan 29,2002
If glbCompSerial <> "S/N - 2332W" And glbCompSerial <> "S/N - 2458W" Then   'Town of Fort Frances
    If Len(medTelephone) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
        MsgBox "Telephone Number is a required field"
        medTelephone.SetFocus
        Exit Function
    End If
End If

If Len(clpDept.Text) < 1 Then
    'Ticket #24164 - Re-ordering
    tbDemographics.SelectedItem = tbDemographics.Tabs(2)

    MsgBox lStr("Department is a required field")
    clpDept.SetFocus
    Exit Function
Else
    If clpDept.Caption = "Unassigned" Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox "Department Code must be valid"
        Call Get_Dept(False)
        If Len(glbDept) < 1 Then
             clpDept.Text = oDept
             clpGLNum.Text = OGLNum
             clpDept.Caption = ODeptD
             clpGLNum.Caption = OGLNumD
             clpDept.Visible = True
           '  clpGLNum.Visible = True
        Else
             clpDept.Text = glbDept
             clpGLNum.Text = glbGLNum
             clpDept.Caption = glbDeptDesc
        End If
        clpDept.SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2242W" Then   'C.C.A.C. London & Middlesex - Ticket #6718
    If Len(Trim(clpGLNum.Text)) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("G/L Number is a required field")
        clpGLNum.SetFocus
        Exit Function
    End If
End If

If Len(clpGLNum.Text) > 0 And clpGLNum.Caption = "Unassigned" Then
    'Ticket #24164 - Re-ordering
    tbDemographics.SelectedItem = tbDemographics.Tabs(2)

    MsgBox lStr("If G/L Number is entered it must be valid")
    clpGLNum.SetFocus
    Exit Function
End If

If glbCompSerial = "S/N - 2347W" Then  'For Surrey Place
    If Len(Trim(clpGLNum.Text)) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("G/L Number is a required filed")
        clpGLNum.SetFocus
        Exit Function
    End If
End If

'Hamilton CAS
'If glbCompSerial = "S/N - 2257W" Then
'Granite Club
If glbCompSerial = "S/N - 2241W" Then
    If Len(dlpDeptEDate.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Department Effective Date is a required field")
        dlpDeptEDate.SetFocus
        Exit Function
    End If
    If Not IsDate(dlpDeptEDate.Text) Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox "Invalid " & lStr("Department Effective Date")
        dlpDeptEDate.SetFocus
        Exit Function
    ElseIf Len(Year(dlpDeptEDate.Text)) = 3 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox "Invalid Year of " & lStr("Department Effective Date")
        dlpDeptEDate.SetFocus
        Exit Function
    End If
End If

If (glbCompSerial = "S/N - 2241W") And Len(clpDiv) = 0 Then ' for Granite Club
    'Ticket #24164 - Re-ordering
    tbDemographics.SelectedItem = tbDemographics.Tabs(2)

    MsgBox lStr("Division is a required field")
    'clpDiv.SetFocus
    If frmWFCDIV.Visible Then txtDouDiv.SetFocus Else clpDiv.SetFocus 'Ticket #21543 Franks 02/08/2012
    Exit Function
End If

If Len(clpDiv.Text) > 1 And clpDiv.Caption = "Unassigned" Then
    If Not glbLinamar Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("If Division is entered it must be valid")
        'clpDiv.SetFocus
        If frmWFCDIV.Visible Then txtDouDiv.SetFocus Else clpDiv.SetFocus 'Ticket #21543 Franks 02/08/2012
        Exit Function
    End If
End If

'C.C.A.C. London & Middlesex - Ticket #6718; Hamilton CAS; Lanark
If (glbCompSerial = "S/N - 2242W") Or (glbCompSerial = "S/N - 2257W") Or (glbCompSerial = "S/N - 2172W") Then
    If Len(clpDiv) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division is a required field")
        'clpDiv.SetFocus
        If frmWFCDIV.Visible Then txtDouDiv.SetFocus Else clpDiv.SetFocus 'Ticket #21543 Franks 02/08/2012
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2393W" Then  ' KTH Shelburne Mfg. Inc. - Ticket #14613
    If Len(clpDiv.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division is a required field")
        'clpDiv.SetFocus
        If frmWFCDIV.Visible Then txtDouDiv.SetFocus Else clpDiv.SetFocus 'Ticket #21543 Franks 02/08/2012
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2408W" Then  ' Township of Wilmot - Ticket #15785
    If Len(clpDiv.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division is a required field")
        clpDiv.SetFocus
        Exit Function
    End If
End If

'CCAC London and Hamilton CAS and Granite Club
If glbCompSerial = "S/N - 2242W" Or glbCompSerial = "S/N - 2241W" Then ' Or glbCompSerial = "S/N - 2257W" Then   'C.C.A.C. London & Middlesex - Ticket #6718
    If Len(dlpDivEDate) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division Effective Date is a required field")
        dlpDivEDate.SetFocus
        Exit Function
    End If
    If Not IsDate(dlpDivEDate) Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division Effective Date must be valid date")
        dlpDivEDate.SetFocus
        Exit Function
    ElseIf Len(Year(dlpDivEDate.Text)) = 3 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox "Invalid Year of " & lStr("Division Effective Date")
        dlpDivEDate.SetFocus
        Exit Function
    End If
End If

For X = 1 To 4
    'Hemu - Begin - Since for Surrey Place the location code has been moved to
    '               Status/Dates screen by Frank, On Save - Location code validity
    '               should not be checked here. Ticket # 4972
    If (X = 1) And (glbCompSerial = "S/N - 2347W") Then GoTo nextcode
    If (X = 2) And (glbCompSerial = "S/N - 2192W") Then GoTo nextcode
    'Hemu - End
    If Len(clpCode(X).Text) > 0 And clpCode(X).Caption = "Unassigned" Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox "If code entered it must be known"
        clpCode(X).SetFocus
        Exit Function
    End If
nextcode:
Next X

If (glbCompSerial = "S/N - 2347W") Then
    If Len(clpCode(4).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox "Vacation is a required field"
        clpCode(4).SetFocus
        Exit Function
    End If
End If
'Ticket #16235 Samuel
If (glbCompSerial = "S/N - 2382W") Then
    If Len(clpCode(4).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Section") & " is a required field"
        clpCode(4).SetFocus
        Exit Function
    End If
    If Len(clpCode(1).Text) = 0 Then 'Ticket #18090
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Location is a required field")
        clpCode(1).SetFocus
        Exit Function
    End If
    'Ticket #18702 - begin
    If Len(dlpDate(0).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
    
        MsgBox lStr("Original Hire Date") & " is a required field."
        dlpDate(0).SetFocus
        Exit Function
    Else
        If Not IsDate(dlpDate(0).Text) Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
            MsgBox "Invalid " & lStr("Original Hire Date")
            dlpDate(0).SetFocus
            Exit Function
        ElseIf Len(Year(dlpDate(0).Text)) = 3 Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
            MsgBox "Invalid Year of " & lStr("Original Hire Date")
            dlpDate(0).SetFocus
            Exit Function
        End If
    End If
    'Ticket #18702 - end
    'Ticket #20319 Franks 05/19/2011 - begin
    If Len(dlpDivEDate.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division Effective") & " is a required field."
        dlpDivEDate.SetFocus
        Exit Function
    Else
        If Not IsDate(dlpDivEDate.Text) Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
            MsgBox "Invalid " & lStr("Division Effective Date")
            dlpDivEDate.SetFocus
            Exit Function
        ElseIf Len(Year(dlpDivEDate.Text)) = 3 Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
            MsgBox "Invalid Year of " & lStr("Division Effective Date")
            dlpDivEDate.SetFocus
            Exit Function
        End If
    End If
    'Ticket #20319 Franks 05/19/2011 - end
    
    'Ticket #20885 Franks 11/10/2011 - begin
    If NewHireForms.count = 0 Then 'for change only
        If Not glbtermopen Then 'active only
            Call CheckReptAuth
        End If
    End If
    'Ticket #20885 Franks 11/10/2011 - end
End If
'Release 8.0 - Ticket #22682: Jerry said to open up these fields for everyone
'Ticket #24164 - Re-ordering and new fields - Organization fields
'If (glbCompSerial = "S/N - 2382W") Then
    For X = 6 To 7
        If Len(clpCode(X).Text) > 0 And clpCode(X).Caption = "Unassigned" Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
            MsgBox "If code entered it must be known"
            clpCode(X).SetFocus
            Exit Function
        End If
    Next X
    If Len(dlpOrg1EDate.Text) > 0 Then
        If Not IsDate(dlpOrg1EDate.Text) Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
            MsgBox "Invalid " & lStr("Organization 1 Effective Date")
            dlpOrg1EDate.SetFocus
            Exit Function
        ElseIf Len(Year(dlpOrg1EDate.Text)) = 3 Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
            MsgBox "Invalid Year of " & lStr("Organization 1 Effective Date")
            dlpOrg1EDate.SetFocus
            Exit Function
        End If
    End If
    If Len(dlpOrg2EDate.Text) > 0 Then
        If Not IsDate(dlpOrg2EDate.Text) Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
            MsgBox "Invalid " & lStr("Organization 2 Effective Date")
            dlpOrg2EDate.SetFocus
            Exit Function
        ElseIf Len(Year(dlpOrg2EDate.Text)) = 3 Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
            MsgBox "Invalid Year of " & lStr("Organization 2 Effective Date")
            dlpOrg2EDate.SetFocus
            Exit Function
        End If
    End If
'End If

If glbCompSerial = "S/N - 2173W" Then 'for town of Ajex Ticket# 6685
    If Len(clpCode(2).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2257W" Then 'HCCAS Ticket #25786 Franks 07/25/2014
    If Len(clpCode(2).Text) < 1 Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
End If
'Franks Nov 6, 2002 for Casey House #3148
If glbCompSerial = "S/N - 2214W" Then
    If clpDept.Text = "1425" Then clpCode(2) = 3
    
    If Len(clpCode(2).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
    If Len(clpDiv.Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Division is a required field")
        clpDiv.SetFocus
        Exit Function
    End If
    'Franks 02/10/04 ticket #5522
    'Franks 03/07/08 ticket 14550
    'If Not (clpDiv.Text = "AIDS" Or clpDiv.Text = "FDTN") Then
        If Len(clpGLNum.Text) = 0 Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
            
            'MsgBox lStr("G/L Number is a required field") & lStr(" when Division is equal to ") & clpDiv.Text 'either CLIN or COMM")
            MsgBox lStr("G/L Number is a required field")
            clpGLNum.SetFocus
            Exit Function
        End If
        If Len(clpCode(3).Text) = 0 Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
            
            'MsgBox lStr("Administered By is a required field") & lStr(" when Division is equal to ") & clpDiv.Text 'either CLIN or COMM")
            MsgBox lStr("Administered By is a required field")
            clpCode(3).SetFocus
            Exit Function
        End If
    'End If
    If Len(clpCode(1).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Location is a required field")
        clpCode(1).SetFocus
        Exit Function
    End If
    If Len(clpCode(4).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Section is a required field")
        clpCode(4).SetFocus
        Exit Function
    End If
    If Not IsDate(dlpDeptEDate) Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Department Effective Date is a required field")
        dlpDeptEDate.SetFocus
        Exit Function
    ElseIf Len(Year(dlpDeptEDate.Text)) = 3 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox "Invalid Year of " & lStr("Department Effective Date")
        dlpDeptEDate.SetFocus
        Exit Function
    End If
    If Not IsDate(dlpDivEDate) Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Division Effective Date is a required field")
        dlpDivEDate.SetFocus
        Exit Function
    ElseIf Len(Year(dlpDivEDate.Text)) = 3 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox "Invalid Year of " & lStr("Division Effective Date")
        dlpDivEDate.SetFocus
        Exit Function
    End If
    
End If
If glbLinamar Then
    If Len(clpCode(2).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
    For X = 1 To 4
        If Len(clpHOME(X)) > 0 And clpHOME(X).Caption = "Unassigned" Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
            MsgBox "If code entered it must be known"
            clpHOME(X).SetFocus
            Exit Function
        End If
    Next X
End If
If Len(dlpDeptEDate.Text) > 0 Then
    If Not IsDate(dlpDeptEDate.Text) Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox "Invalid Department Effective Date"
        dlpDeptEDate.SetFocus
        Exit Function
    ElseIf Len(Year(dlpDeptEDate.Text)) = 3 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox "Invalid Year of " & lStr("Department Effective Date")
        dlpDeptEDate.SetFocus
        Exit Function
    End If
End If
If Len(dlpDivEDate.Text) > 0 Then
    If Not IsDate(dlpDivEDate.Text) Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox "Invalid Division Effective Date"
        dlpDivEDate.SetFocus
        Exit Function
    ElseIf Len(Year(dlpDivEDate.Text)) = 3 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox "Invalid Year of " & lStr("Division Effective Date")
        dlpDivEDate.SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2357W" And comCountry = "CANADA" Then   'I.T. Xchange
    If Len(clpCode(2).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
End If

'Town of Fort Frances and Hamilton CAS, CCAC London, FCY Muskoka, Granite Club
If glbCompSerial = "S/N - 2332W" Or glbCompSerial = "S/N - 2257W" Or glbCompSerial = "S/N - 2242W" Or glbCompSerial = "S/N - 2366W" Or (glbCompSerial = "S/N - 2241W") Then
    If Len(clpCode(3)) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Administered By is a required field")
        clpCode(3).SetFocus
        Exit Function
    End If
End If

'Hemu - Begin - City of Timmins - Ticket #9557
If glbCompSerial = "S/N - 2375W" Then
    If Len(clpDiv.Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Division is a required field")
        clpDiv.SetFocus
        Exit Function
    End If
    If Len(clpCode(1).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Location is a required field")
        clpCode(1).SetFocus
        Exit Function
    End If
    'If Len(clpCode(4).Text) < 1 Then   - Ticket 9877
    '    MsgBox lStr("Section is a required field")
    '    clpCode(4).SetFocus
    '    Exit Function
    'End If
End If
'Hemu - End

If glbCompSerial = "S/N - 2182W" Or glbCompSerial = "S/N - 2453W" Or glbCompSerial = "S/N - 2493W" Then    'Ticket #25694 Franks 07/07/2014
'2453W - Town of Gander Ticket #25716 Franks 07/28/2014
    If Len(clpCode(4).Text) < 1 Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lStr("Section is a required field")
        clpCode(4).SetFocus
        Exit Function
    End If
End If

'added by Bryan 26/Oct/05 Ticket#9627
If glbCompSerial = "S/N - 2369W" Then
    If Len(clpDiv.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Division is a required field")
        clpDiv.SetFocus
        Exit Function
    End If
    
    If Len(clpGLNum.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("G/L Number is a required field")
        clpGLNum.SetFocus
        Exit Function
    End If
End If

'Ticket #11148
If glbCompSerial = "S/N - 2382W" Then  'Namasco
    'Ticket #16235 - Begin
    'If Len(txtBadgeID.Text) = 0 Then
    '    MsgBox ("Badge ID is a required field")
    '    txtBadgeID.SetFocus
    '    Exit Function
    'End If
    If Len(clpDiv.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Division is a required field")
        clpDiv.SetFocus
        Exit Function
    End If
    If Len(clpCode(2).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
    If Len(dlpDeptEDate.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Department Effective Date is a required field")
        dlpDeptEDate.SetFocus
        Exit Function
    End If
    'Ticket #16235 - End
    If Len(clpCode(3).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Administered By is a required field")
        clpCode(3).SetFocus
        Exit Function
    End If
    'Ticket #20045 Franks 04/01/2011
    If clpCode(3).Text = "5322" Or clpCode(3).Text = "2158" Then
        If Len(clpGLNum.Text) = 0 Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
            
            MsgBox lStr("G/L #") & " is a required field if " & lStr("Administered By") & " is '5322' or '2158'."
            clpGLNum.SetFocus
            Exit Function
        End If
    End If
End If

'Ticket #15793
If glbCompSerial = "S/N - 2390W" Then  'Collectcorp Inc
    If Len(txtBadgeID.Text) = 0 Then
        txtBadgeID.Text = glbLEE_ID
    End If
End If

If glbCompSerial = "S/N - 2385W" Then  'Conservation Halton 'Ticket #13063
    If Len(clpDiv.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Division is a required field")
        clpDiv.SetFocus
        Exit Function
    End If
    If Len(clpCode(4).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Section is a required field")
        clpCode(4).SetFocus
        Exit Function
    End If
End If
If glbCompSerial = "S/N - 2335W" Then   'Mitchell Plastics Ltd. - Ticket #20982
    If Len(clpCode(4).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Section is a required field")
        clpCode(4).SetFocus
        Exit Function
    End If
End If


'Ticket #24443 - North York Community House
If glbCompSerial = "S/N - 2391W" Then
    If Len(clpCode(2).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2380W" Then   'VitalAire Ticket #13505
    If Len(clpCode(4).Text) > 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox (lStr("Section") & " only accepts one digit Code")
        clpCode(4).SetFocus
        Exit Function
    End If
End If


If glbCompSerial = "S/N - 2386W" Then  'The Walter Fedy Partnership 'Ticket #13828
    If Len(clpDiv.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Division is a required field")
        clpDiv.SetFocus
        Exit Function
    End If
    If Len(clpCode(1).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Location is a required field")
        clpCode(1).SetFocus
        Exit Function
    End If
    If Len(clpCode(2).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
    If Len(clpCode(3).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Administered By is a required field")
        clpCode(3).SetFocus
        Exit Function
    End If
    If Len(clpCode(4).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Section is a required field")
        clpCode(4).SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2415W" Then 'Ticket #16982 SPC- Volunteer System
    If Len(clpDiv.Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Division is a required field")
        clpDiv.SetFocus
        Exit Function
    End If
    If Len(clpCode(1).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Location is a required field")
        clpCode(1).SetFocus
        Exit Function
    End If
    If Len(clpCode(3).Text) < 1 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Administered By is a required field")
        clpCode(3).SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2233W" Then   'Leeds-Grenville F&CS - Ticket #16737
    If Len(clpGLNum.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("G/L Number is a required field")
        clpGLNum.SetFocus
        Exit Function
    End If
    
    If Len(clpCode(2).Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2393W" Then   'KTH Shelburne Ticket #17289
    If Len(medPARKPERMIT2.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(3)
        
        MsgBox lStr("Parking Permit #2 is a required field")
        medPARKPERMIT2.SetFocus
        Exit Function
    End If
    If Len(medLICPLATE2.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(3)
    
        MsgBox lStr("License Plate #2 is a required field")
        medLICPLATE2.SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2335W" Then 'Mitchell Plastics Ticket #21866 Franks 04/05/2012
    If Len(txtPayrollID.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
        MsgBox "Payroll ID is a required field"
        txtPayrollID.SetFocus
        Exit Function
    End If
    If Len(txtBadgeID.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
    
        MsgBox ("Badge ID is a required field")
        txtBadgeID.SetFocus
        Exit Function
    End If
    If Len(clpDiv.Text) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    
        MsgBox lStr("Division is a required field")
        clpDiv.SetFocus
        Exit Function
    End If
End If

'Ticket #22409 Frank 08/08/2012
If glbWFC Then
    'Ticket #25088 Franks 02/18/2014 - begin
    If NewHireForms.count = 0 Then 'modify only
        If Len(SavDiv) > 0 Then
            If Not (SavDiv = clpDiv.Text) Then
                If flgWFCDivChaFlag Then
                    'Ticket #28808 Franks 06/20/2016
                    '"   Add an   button beside Division. When pressed, a screen pops up to ask for a password.
                Else
                    MsgBox "Division and Plant cannot be changed on this screen. " & Chr(10) & "Use the Transfer Out/In facility if the Division and Plant needs to change. "
                    tbDemographics.SelectedItem = tbDemographics.Tabs(2)
                    'clpDiv.SetFocus
                    If frmWFCDIV.Visible Then txtDouDiv.SetFocus Else clpDiv.SetFocus
                    Exit Function
                End If
            End If
        End If
        If Len(OSection) > 0 Then
            If Not (OSection = clpCode(4).Text) Then
                MsgBox "Division and Plant cannot be changed on this screen. " & Chr(10) & "Use the Transfer Out/In facility if the Division and Plant needs to change. "
                tbDemographics.SelectedItem = tbDemographics.Tabs(2)
                clpCode(4).SetFocus
                Exit Function
            End If
        End If
    End If
    'Ticket #25088 Franks 02/18/2014 - end
    
    'Ticket #24421 Franks 10/08/2013
    '"   If HRsoft has no gender, the user must select either Male or Female before saving the record. It's defaulting to female.
    If Not optGender(0).Value And Not optGender(1).Value And Not optGender(2).Value Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        MsgBox ("Please select Male, Female or Not Disclosed")
        Exit Function
    End If
    
    'Ticket #22409 Frank 08/08/2012
    Call WFC_NGS_SmokerUptWithEmail
End If

'Ticket #24543 - Macaulay Child Development Centre
If glbCompSerial = "S/N - 2420W" Then
    If Len(clpCode(0).Text) > 0 Then
        If clpCode(0).Caption = "Unassigned" Then
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
            MsgBox lStr("Union code must be valid")
            clpCode(0).SetFocus
            Exit Function
        End If
    End If
    
    If clpSalDist.Caption = "Unassigned" Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lblSalDist.Caption & " must be valid"
        clpSalDist.SetFocus
        Exit Function
    End If
End If

'Ticket #24543 - Macaulay Child Development Centre
If glbCompSerial = "S/N - 2420W" Then
    'SIN begins with 9 then Work Visa # and Expiration Date should be entered before the Save
    If Left(medSIN, 1) = "9" Then
        'Dim xWorkVisaNo
        'Dim xWorkExpDate
        
        'Initialize
        glbWorkVisaNo = ""
        glbWorkExpDate = ""
        
        'Check if Work Visa and Expiration Date is entered
        xWorkVisaNo = get_EmpOtherByField(glbLEE_ID, "ER_VISAPERMITNO")
        xWorkExpDate = get_EmpOtherByField(glbLEE_ID, "ER_VISAPERMITDATE")
        If IsNull(xWorkVisaNo) Or IsNull(xWorkExpDate) Then
            'Info not found, enter it
            
            'Initialize
            glbWorkVisaNo = ""
            glbWorkExpDate = ""
            
            'Display the info if found.
            If Len(xWorkVisaNo) > 0 Then
                glbWorkVisaNo = xWorkVisaNo
            End If
            If Len(xWorkExpDate) > 0 Then
                glbWorkExpDate = xWorkExpDate
            End If
            
            frmMsgWorkPermit.Show 1
            
            'Cannot save employee information without Work Visa # and Expiration Date
            If glbWorkVisaNo = "" Or glbWorkExpDate = "" Then
                MsgBox "This employee's record cannot be saved without the Work Visa # and Expiration Date.", vbOKOnly, "Work Visa # and Expiration Date Required"
                Exit Function
            End If
        End If
    End If
End If

If glbLinamar Then 'Ticket #28875 Franks 07/13/2016
    'Required field if SIN starts with a 9.  Pop up a window asking for an Expiry Date and store under Other Information's Visa/Work Permit Expiration Date.
    If Left(medSIN, 1) = "9" Then
        'Initialize
        glbWorkExpDate = ""
        'Check if Work Visa and Expiration Date is entered
        xWorkExpDate = get_EmpOtherByField(glbLEE_ID, "ER_VISAPERMITDATE")
        'If IsNull(xWorkVisaNo) Or IsNull(xWorkExpDate) Then
        If IsNull(xWorkExpDate) Then
            'Info not found, enter it
            
            'Initialize
            glbWorkExpDate = ""
            
            'Display the info if found.
            If Len(xWorkExpDate) > 0 Then
                glbWorkExpDate = xWorkExpDate
            End If
            
            frmMsgWorkPermit.Show 1
            
            'Cannot save employee information without Work Visa # and Expiration Date
            If glbWorkExpDate = "" Then
                MsgBox "This employee's record cannot be saved without the Visa/Work Permit Expiration Date.", vbOKOnly, "Visa/Work Permit Expiration Date Required"
                Exit Function
            End If
        End If
    End If
End If

If glbCompSerial = "S/N - 2460W" Then 'OPL 'Ticket #25323 Franks 12/16/2014
    If Len(txtPayrollID.Text) = 0 Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        MsgBox "Payroll ID is a required field"
        txtPayrollID.SetFocus
        Exit Function
    End If
    If Len(clpCode(2).Text) < 1 Then
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        MsgBox lStr("Region is a required field")
        clpCode(2).SetFocus
        Exit Function
    End If
End If

'Ticket #26672 - County of Perth
If glbCompSerial = "S/N - 2417W" Then
    'Type of Vehicle
    If Len(Trim(medTYPEVEHICLE.Text)) = 0 Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(3)
    
        MsgBox lStr("Type of Vehicle") & " is a required field"
        medTYPEVEHICLE.SetFocus
        Exit Function
    End If
End If

If glbWFC Then 'Ticket #28637 Franks 05/18/2016
    'If clpCode(4).Text = "MISS" Or clpCode(4).Text = "TROY" Then
    If NewHireForms.count > 0 Then 'new hire only
        xTmpUnion = glbTrsUnion
    Else
        xTmpUnion = glbUNION
    End If
    'If xTmpUnion = "NONE" Or xTmpUnion = "EXEC" Then
    If (xTmpUnion = "NONE" Or xTmpUnion = "EXEC") And Not glbtermopen Then  'Ticket #29836 Franks 02/24/2017 - not for term
        Call WFCNetworkLoginSetup 'Ticket #28772 Franks 06/21/2016
        If Len(medNetworkLogin.Text) = 0 Then
            tbDemographics.SelectedItem = tbDemographics.Tabs(3)
            MsgBox "Network Login is a required field"
            medNetworkLogin.SetFocus
            Exit Function
        End If
        If Len(medVendorNo.Text) = 0 Then
            'Ticket #30012 Franks 04/07/2017 - begin
            'Jerry: If the Division Master does not have a Locator Code, Vendor Number does not need to be entered. They should not get the message below
            If Len(clpDiv.Text) > 0 Then
                xDV_ORGT1 = getDivField(clpDiv.Text, "DV_ORGT1")
            Else
                xDV_ORGT1 = ""
            End If
            If Len(xDV_ORGT1) > 0 Then
                tbDemographics.SelectedItem = tbDemographics.Tabs(3)
                MsgBox "Vendor Number is a required field"
                medVendorNo.SetFocus
                Exit Function
            End If
            'Ticket #30012 Franks 04/07/2017 - end
        End If
    End If
End If

chk_FEBASIC = True

End Function

Private Function chkForEEData(TabName$, EEIDAlias$, EEID&)
Dim snapTEE As New ADODB.Recordset
Dim gdbESS As New ADODB.Connection
Dim SQLQ As String

chkForEEData = 0

'''On Error GoTo chkPopErr

If EEIDAlias$ <> "" Then
    SQLQ = "Select " & TabName$ & "." & EEIDAlias$ & " FROM " & TabName$
    
    If InStr(EEIDAlias$, "USER") > 0 Then
        SQLQ = SQLQ & " WHERE " & EEIDAlias$ & " = '" & EEID& & "'"
    Else
        SQLQ = SQLQ & " WHERE " & EEIDAlias$ & " = " & EEID&
    End If
End If

chkForEEData = 0

If SQLQ <> "" Then
    'Users were getting error when the ESS mdb was not there for MS Access users.
    If Not glbSQL And Not glbOracle And (TabName = "HR_TIMESHEET" Or TabName = "HR_TIMESHEET_MODS" Or TabName = "HR_VACTIMEOFF_REQ" Or TabName = "HR_VACTIMEOFF_REQ_ARCHIVE") Then
        If gdbESS = "" Then
            gdbESS.Open Replace(glbAdoIHRDB, "IHR001", "IHRESS")
        End If
    End If

    If Not glbSQL And Not glbOracle And (TabName = "HR_TIMESHEET" Or TabName = "HR_TIMESHEET_MODS" Or TabName = "HR_VACTIMEOFF_REQ" Or TabName = "HR_VACTIMEOFF_REQ_ARCHIVE") Then
        If gdbESS <> "" Then
            snapTEE.Open SQLQ, gdbESS, adOpenStatic
        Else
            GoTo Skip
        End If
    Else
        snapTEE.Open SQLQ, gdbAdoIhr001, adOpenStatic
    End If
    
    If IsNull(snapTEE.RecordCount) Then
        Exit Function
    End If
    
    If snapTEE.RecordCount > 0 Then
        chkForEEData = snapTEE.RecordCount
    End If
    
    snapTEE.Close
    
Skip:

End If
Exit Function

chkPopErr:
If Err.Number = -2147467259 Then
    gdbESS = ""
    Resume Next
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HR_TABLES Error", "TabName$", "Select")
Call RollBack '21June99 js

End Function

Sub cmdCancel_Click()
Dim X
'''On Error GoTo Can_Err

Call SetCountries

RDept = SavDept

If Not (rsDATA.EOF And rsDATA.BOF) Then
    rsDATA.CancelUpdate
End If

'Ticket #24543 - Macaulay Child Development Centre
If glbCompSerial = "S/N - 2420W" Or glbLinamar Then
    'Initialize
    glbWorkVisaNo = ""
    glbWorkExpDate = ""
End If

Call Display_Value

If Not rsDATA.EOF Then Call getCodes

'Hemu - 11/21/2003 Begin - For the first time it prompts to associate G/L with Dept
'                  even when the Dept. Code has not changed, this was because the
'                   RDept value was empty for the first time
RDept = clpDept.Text
'Hemu - 11/21/2003 End

If fglbNewEE = True Then
    fglbNewEE = False
    glbLEE_ID = oldEEId
    For X = 1 To NewHireForms.count
        NewHireForms.Remove 1
    Next
    
    If glbCompSerial = "S/N - 2241W" Then ' Not Granite Club
        Call Check_EMPLOYEE_Number(glbNextEmpl)
    End If
    
    Unload Me
    Exit Sub
End If
Call ST_UPD_MODE(True)

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
Call RollBack '21June99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEEBASIC" Then glbOnTop = ""
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdDelete_Click()
Dim a As Integer, Msg As String, DtTm As Variant
Dim Title$, DgDef, Response%
DtTm = Now

'Ticket #24164 - Re-ordering
tbDemographics.SelectedItem = tbDemographics.Tabs(1)

'''On Error GoTo Del_Err
If glbtermopen Then
    'Ticket #24164 - Re-ordering
    tbDemographics.SelectedItem = tbDemographics.Tabs(1)

    Msg$ = Msg$ & Chr(10) & "Are you sure you wish to delete "
    Msg$ = Msg$ & Chr(10) & "termination records for this employee?"
    Title$ = "Delete Termination Records for Employee"
    DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then    ' Evaluate response
        Exit Sub
    End If
    Screen.MousePointer = HOURGLASS
    If Not modNukeEETerm(glbTERM_Seq + 0) Then MsgBox "Employee remains in Termination file"
    Screen.MousePointer = DEFAULT
    glbTERM_ID = 0
    Unload Me
    Exit Sub
End If

'Ticket #24164 - Re-ordering
'lstEETables.Top = clpDept.Top
tbDemographics.Height = tbDemographics.Height + 2500
frPersonal.Height = frPersonal.Height + 2500

'Ticket #24164 - Re-ordering
'lstEETables.Height = fraDetail.Height - lstEETables.Top
lstEETables.Height = frPersonal.Height - lstEETables.Top - 200

lstEETables.Visible = True  'js-01Apr99
cmdHide.Visible = True      '
cmdDeleteAll.Visible = True '
 
'cmdClose.Visible = False    '
'cmdModify.Visible = False   '
'cmdOK.Visible = False       '
'cmdCancel.Visible = False   '
'cmdNew.Visible = False      '
'cmdDelete.Visible = False   '
cmdPhoto.Visible = False
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Searching for Employee Information"
If Review_EE_Tables(glbLEE_ID) Then       ' records found
    MDIMain.panHelp(0).Caption = ""
    Screen.MousePointer = DEFAULT
    Exit Sub
End If
DoEvents


MDIMain.panHelp(0).Caption = ""
Screen.MousePointer = DEFAULT
glbBasicChg% = True

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HREMP", "Delete")
Call RollBack '21June99 js

End Sub

Private Sub clpCode_GotFocus(Index As Integer)
Dim xTransDiv As String
If glbSamuel Then 'Ticket #21106 Franks 11/07/2011
    If Index = 4 Then
        xTransDiv = GetTransDiv(1)  'GetTransDiv(Index)
        'If Len(xTransDiv) > 0 Then
            clpCode(Index).TransDiv = xTransDiv
        'End If
    End If
    If Index = 2 Then 'Ticket #22423 Franks 08/30/2012
        xTransDiv = GetTransDiv(2)
        'If Len(xTransDiv) > 0 Then
            clpCode(Index).TransDiv = xTransDiv
        'End If
    End If
End If
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
If glbCompSerial = "S/N - 2382W" Then 'Ticket #20695 for Samuel Franks 09/23/2011
    If Index = 3 Then
        Call Samuel_GL
    End If
End If
End Sub

Private Sub clpDept_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub clpDept_LostFocus()
     Call Dept_GL
End Sub

Private Sub clpDIV_Change()
If glbWFC Then 'Ticket #21544 Franks 02/06/2012
    txtDouDiv.Text = clpDiv
End If
End Sub

Private Sub clpDiv_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub clpDiv_LostFocus()
'Ticket #23902 Franks 07/05/2013
'If glbCompSerial = "S/N - 2376W" Then 'Assembly of First Nations Ticket #15648
'    clpCode(2).Text = clpDiv.Text
'End If
End Sub

Private Sub cmdCCLife_Click()
Dim Msg$, DgDef As Variant, Response%
Dim SQLQ As String

'Ticket #24164 - Re-ordering
tbDemographics.SelectedItem = tbDemographics.Tabs(1)

Msg$ = "This function will update CC and GTLD benefits for this employee: "
Msg$ = Msg$ & Chr(10) & "   Benefit Effective Date = Original Hire Date "
'Msg$ = Msg$ & Chr(10) & "   2. Relationship equal to 'Son' or 'Daughter' on the Dependents screen "
Msg$ = Msg & Chr(10) & Chr(10) & "you want to proceed?"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2
Response% = MsgBox(Msg$, DgDef, "Confirm")

If Response% = IDYES Then
    Call WFC_UptUSBenByEmp(glbLEE_ID, CVDate(dlpDate(0).Text), 0, "Y", "Y", "Y")
End If
End Sub

Private Sub cmdCopyEmpByPayID_Click()
    frmSecuCopy.IsCopyByPayID = True
    frmSecuCopy.elpEmpLookup = glbLEE_ID
    frmSecuCopy.txtFromUserID = txtPayrollID.Text
    frmSecuCopy.Show 1
End Sub

Private Sub cmdDeleteAll_Click()
Dim Msg$, DgDef As Variant, Response%
Dim rsT_PARCO As New ADODB.Recordset
Dim SQLQ As String
Dim xCurPosition
'Ticket #24164 - Re-ordering
tbDemographics.SelectedItem = tbDemographics.Tabs(1)

Msg$ = "Warning, you are about to delete "
Msg$ = Msg$ & Chr(10) & "ALL information about this employee."
Msg$ = Msg$ & Chr(10) & "No information will remain to forward to "
Msg$ = Msg$ & Chr(10) & "ANY other system (i.e. Payroll interface)."
Msg$ = Msg$ & Chr(10) & Chr(10) & "Are you ABSOLUTELY sure "
Msg$ = Msg & Chr(10) & "you want to proceed?"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2
Response% = MsgBox(Msg$, DgDef, "Warning - No Recovery Delete!")

If Response% = IDYES Then
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).Caption = "Removing all traces of this Employee"
    Do While NewHireForms.count > 0
        NewHireForms.Remove 1
    Loop
    
    If glbWFC Then 'Ticket #25956 Franks 09/02/2014
        Call WFCCandidateDele(txtCandidate.Text)
        xCurPosition = getEmpPostion(glbLEE_ID) 'Ticket #27820 Franks 11/25/2015
    End If
    
    Call Employee_Master_Integration(glbLEE_ID, , True)
    
    If Not AUDITDEMO("D") Then MsgBox "ERROR : AUDIT FILE"
    If glbtermopen Then
    Else
        Call NukeEE(glbLEE_ID)
        'Ticket #23116 Franks 01/23/2013 for WFC
        Call NukeEE_SerialNo(glbLEE_ID)
        
        'Ticket #19478 Frank 11/26/2010
        'delete HRAUDIT records too
        SQLQ = "DELETE FROM HRAUDIT WHERE AU_EMPNBR = " & glbLEE_ID
        gdbAdoIhr001X.Execute SQLQ
    End If
    Screen.MousePointer = DEFAULT
    MDIMain.panHelp(0).Caption = " "
    
    glbLEE_ID = 0
    glbLEE_FName = ""
    glbLEE_SName = ""
    If glbLinamar Then
        glbLEE_ProdLine = "" 'Ticket #14775
    End If
    
    Call UpdMaxEmpNbr     'laura 03/03/98
    rsT_PARCO.Open "HRPARCO", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #27829
        rsT_PARCO("PC_NUMBER_EMPLOYEES") = modECount_FamilyDay
    Else
        rsT_PARCO("PC_NUMBER_EMPLOYEES") = rsT_PARCO("PC_NUMBER_EMPLOYEES") - 1 'UPDATE FIELD WITH ACTUAL COUNT
    End If
    rsT_PARCO.Update
    rsT_PARCO.Close
    
    If glbWFC Then 'Ticket #27820 Franks 11/25/2015
        Call mod_Upd_Pos_Budget_WFC(xCurPosition, "")
    End If

    Call UnloadFrms
    Unload Me
    
    Exit Sub
End If

lstEETables.Visible = False  'js-01Apr99
cmdHide.Visible = False      '
cmdDeleteAll.Visible = False '

'Ticket #24164 - Re-ordering
frPersonal.Height = 5895    '5535
tbDemographics.Height = 6495    '6000

'cmdClose.Visible = True      '
'cmdModify.Visible = True     '
'cmdOK.Visible = True         '
'cmdCancel.Visible = True     '
'cmdNew.Visible = True        '
'cmdDelete.Visible = True     '
cmdPhoto.Visible = True
End Sub

Private Sub cmdEditDiv_Click() 'Ticket #28808 Franks 06/20/2016
    flgWFCDivChaFlag = False
    glbAccessPswd = False
    frmAccessPswd.Show 1
    If glbAccessPswd = False Then   'Access Denied
        Exit Sub
    End If
    flgWFCDivChaFlag = True
    'clpDiv.Enabled = True
    'clpDiv.SetFocus
End Sub

Private Sub cmdEditPayID_Click() 'Ticket #29759 Franks 02/21/2017
    glbAccessPswd = False
    frmAccessPswd.Show 1
    If glbAccessPswd = False Then   'Access Denied
        Exit Sub
    End If
    txtPayrollID.Enabled = True
    txtPayrollID.SetFocus
End Sub

Private Sub cmdHide_Click()

lstEETables.Visible = False  'js-01Apr99
cmdHide.Visible = False      '

'Ticket #24164 - Re-ordering
frPersonal.Height = 5895    '5535
tbDemographics.Height = 6495    '6000


cmdDeleteAll.Visible = False '

'cmdClose.Visible = True      '
'cmdModify.Visible = True     '
'cmdOK.Visible = True         '
'cmdCancel.Visible = True     '
'cmdNew.Visible = True        '
'cmdDelete.Visible = True     '
cmdPhoto.Visible = True
End Sub

Private Sub cmdMiss_Click()
Dim rsTN As New ADODB.Recordset
Dim rsTA As New ADODB.Recordset
Dim xMiss, SQLQ, xTITLE, xMissMSG

Screen.MousePointer = vbHourglass
SQLQ = "SELECT HRNEWHIRE.*,INFO_HR_TABLES.Empnbr_Alias, INFO_HR_TABLES.TERMINATION_TABLE FROM HRNEWHIRE"
If glbOracle Then
    SQLQ = SQLQ & " ,INFO_HR_TABLES "
    SQLQ = SQLQ & " WHERE HRNEWHIRE.TableName=INFO_HR_TABLES.Table_Name "
    SQLQ = SQLQ & " AND HRNEWHIRE.NEWHIRE<>0"
Else
    SQLQ = SQLQ & " INNER JOIN INFO_HR_TABLES "
    SQLQ = SQLQ & " ON HRNEWHIRE.TableName=INFO_HR_TABLES.Table_Name "
    SQLQ = SQLQ & " WHERE HRNEWHIRE.NEWHIRE<>0"
End If
' danielk - 12/31/2002 - Changed <>0 to =0, it was selecting ONLY term tables instead of NOT term tables.
SQLQ = SQLQ & " AND TERMINATION_TABLE=0"
SQLQ = SQLQ & " ORDER BY ID"
rsTN.Open SQLQ, glbAdoIHRDB
xMiss = ""
xTITLE = ""
Do Until rsTN.EOF
    If Not ((UCase(rsTN("TableName")) = "HRCOBRA" Or UCase(rsTN("TableName")) = "HREEO") And glbCountry = "CANADA") Then
        SQLQ = "SELECT * FROM " & rsTN("TableName") & " Where " & rsTN("Empnbr_Alias") & "=" & glbLEE_ID
        If Not IsNull(rsTN("KeyField")) Then SQLQ = SQLQ & " AND " & rsTN("KeyField") & " is not NULL "
        rsTA.Open SQLQ, glbAdoIHRDB
        If rsTA.EOF Then
            If IsNull(rsTN("MenuTitle")) Then
                xTITLE = ""
            Else
                If xTITLE <> rsTN("MenuTitle") Then
                    'If glbWFC And UCase(rsTN("TableName")) = "HREEO" Then 'Ticket #24422 Franks 10/11/2013
                    If UCase(rsTN("TableName")) = "HREEO" Then  'Ticket #24422 Franks 10/11/2013
                        xMiss = xMiss & "EEO" & Chr(10)
                    Else
                        xMiss = xMiss & rsTN("MenuTitle") & Chr(10)
                    End If
                End If
                xTITLE = rsTN("MenuTitle")
            End If
            If UCase(rsTN("TableName")) = "HREEO" Then 'Ticket #24422 Franks 10/11/2013
                'xMiss = xMiss & IIf(Len(xTitle) > 0, vbTab, "") & "EEO" & Chr(10)
            Else
                xMiss = xMiss & IIf(Len(xTITLE) > 0, vbTab, "") & rsTN("MenuItem") & Chr(10)
            End If
        End If
        rsTA.Close
    End If
    rsTN.MoveNext
Loop
rsTN.Close
Screen.MousePointer = vbDefault
If Len(xMiss) = 0 Then xMiss = "All screens have been completed based on the New Hire Procedure."
MsgBox xMiss, , "What is Missing"
End Sub

Sub cmdModify_Click()

'''On Error GoTo Mod_Err

OSNAME = txtSurname
OFNAME = txtFName
OADD1 = txtAdd1
OADD2 = txtAdd2
OTITLE = txtTitle
OCITY = txtCity
oProv = clpProv.Text
oProvEmp = clpProvEmp.Text
oCountry = txtCountry
oCountryEmployment = txtCountryOfEmp.Text
OPCODE = medPCode
OPHONE = medTelephone
OBUSNBR = medTele2
OSEX = txtGender
OSIN = medSIN
OSMOKER = ComSmoker
ODOB = dlpDOB
oDOH = dlpDate(0).Text
oGLNo = clpGLNum.Text
OMSTAT = txtMStatus
oRegion = clpCode(2).Text
oAdminBy = clpCode(3).Text
OSection = clpCode(4).Text
OHOMELINE = clpHOME(2)
OHOMESHIFT = clpHOME(4)
OHOMEOPRTNBR = clpHOME(1)
oHOMEWRKCNT = clpHOME(3)
ODeptEDate = dlpDeptEDate.Text
ODivEdate = dlpDivEDate.Text
OCellPhone = medCellPhone
OPageNbr = medPageNbr
OSSN = medSSN
oPayrollID = txtPayrollID
SavDept = clpDept.Text
SavDiv = clpDiv.Text
SavLoc = clpCode(1).Text
oDRIVERLIC = medDRIVERLIC
oLICPLATE1 = medLICPLATE1
oLICPLATE2 = medLICPLATE2
oLOCKER = medLOCKER
oCOMBINATION = medCOMBINATION
oTYPEVEHICLE = medTYPEVEHICLE
oPARKPERMIT1 = medPARKPERMIT1
oPARKPERMIT2 = medPARKPERMIT2
oBadgeID = txtBadgeID
oMidName = txtMidName
oAlias = txtAlias
OGLNum = clpGLNum

oOrg1 = clpCode(6).Text

'Ticket #24543 - Macaulay Child Development Centre
If glbCompSerial = "S/N - 2420W" Then
    SavOrg = clpCode(0).Text
    oSalDist = clpSalDist
End If

If glbLinamar Then 'Ticket #29759 Franks 02/14/2017
    oVadim1 = clpVadim1.Text
End If

'Ticket #22912 Franks 12/06/2012 - begin
If glbSamuel Then
    xFutureChgDeptNo = False
    xFutureChgSection = False
    xFutureChgRegion = False
    xFutureDateDeptNo = ""
    xFutureDateSection = ""
    xFutureDateRegion = ""
End If
'Ticket #22912 Franks 12/06/2012 - end

'Ticket #24543 - Macaulay Child Development Centre
If glbCompSerial = "S/N - 2420W" Then
    'Initialize
    glbWorkVisaNo = ""
    glbWorkExpDate = ""
End If

'Ticket #24557 Franks 07/07/2015 - begin
If glbCompSerial = "S/N - 2420W" Then  'Macaulay Child Development
    ADPBranchOld(0) = clpCode(1).Text 'main Branch
    ADPBranchOld(1) = clpCode(11).Text 'Alt Payroll ID - Branch 1
    ADPBranchOld(2) = clpCode(12).Text 'Alt Payroll ID - Branch 2
    ADPBranchOld(3) = clpCode(13).Text 'Alt Payroll ID - Branch 3
    ADPDeptOld(0) = clpSalDist.Text  'main dept
    ADPDeptOld(1) = clpSalDis2(0).Text 'Alt Payroll ID - dept 1
    ADPDeptOld(2) = clpSalDis2(1).Text 'Alt Payroll ID - dept 2
    ADPDeptOld(3) = clpSalDis2(2).Text 'Alt Payroll ID - dept 3
End If
'Ticket #24557 Franks 07/07/2015 end

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '21June99 js

End Sub


Public Sub cmdNew_Click()
Dim X%
Dim intRet

If Not modECountChk() Then
    MsgBox "You have reached the maximum number of employees for your license"
    Exit Sub
End If

MDIMain.MainToolBar.ButtonS(8).Enabled = True
MDIMain.MainToolBar.ButtonS(9).Enabled = True

'Ticket #24164 - Re-ordering
tbDemographics.SelectedItem = tbDemographics.Tabs(1)

lstEETables.Visible = False
cmdHide.Visible = False
cmdDeleteAll.Visible = False '

'Ticket #24164 - Re-ordering
frPersonal.Height = 5895    '5535
tbDemographics.Height = 6495    '6000

fglbNewEE = True
frmBlank.Visible = True: frmBlank.Top = 0: frmBlank.Left = 0
If glbNo Then
    glbNo = False
    Exit Sub
End If

'Ticket #28040 - To Track on New Hire if the user went into the Organizational tab at least once.
'Initialize
flgSwitchOrgTabNewHire = True

'Get the list of forms user does not have access to
Call No_Security_Rights_on_Forms

'Ask user if they want to proceed with new when they don't have access to certain new hire screens
If strNoAccessForms <> "" Then
    strNoAccessForms = "You do not have access to the following new hire screen(s):" & vbCrLf & strNoAccessForms
    strNoAccessForms = strNoAccessForms & vbCrLf & vbCrLf & "Proceed with New Hire? "
    intRet = MsgBox(strNoAccessForms, vbYesNo, "New Hire")
    If intRet = vbNo Then
        fglbNewEE = False
        Unload frmEEBASIC
        Exit Sub
    End If
End If

X% = CR_NEW_EE()
mbAddNewEmployee = False

If UnloadForm Then Unload Me: Exit Sub

If Not fglbNewEE Then Exit Sub

Call get_NewHireForms

lblDOH = ""

If IsNull(glbCountry) Then
    comCountry = "CANADA"
Else
    comCountry = glbCountry
End If

If glbCompSerial = "S/N - 2214W" Then 'for casey house Ticket #14550
    clpProv.Text = "ON"
End If

'Ticket #20873 - Defaults for Kerry's Place and Conservation Halton
If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2385W" Then
    clpProv.Text = "ON"
    comCountry = "CANADA"
    comCountryOfEmp = "CANADA"
End If

'Ticket #24543 - Macaulay Child Development Centre
If glbCompSerial = "S/N - 2420W" Or glbLinamar Then
    'Initialize
    glbWorkVisaNo = ""
    glbWorkExpDate = ""
End If

'Ticket #20834 Franks 08/22/11
If glbSQL Or glbOracle Then
    If cmdPhoto.Caption = "&Photo Off" Then
        picPhoto.Visible = False
        PicNotF.Visible = True
    End If
End If

If glbLinamar Then
    Select Case Right(glbLEE_ID, 3)
    Case "501", "505"
        comCountry = "MEXICO"
        clpProv = "CO"
    Case "502"
        comCountry = "MEXICO"
        clpProv = "DU"
    Case "117"
        comCountry = "U.S.A."
        clpProv = "IA"
    Case "118"
        comCountry = "U.S.A."
        clpProv = "IL"
    Case "120"
        comCountry = "U.S.A."
        clpProv = "IA"
    Case "345", "310", "370"
        comCountry = "U.S.A."
        clpProv = ""
    Case "360"
        comCountry = "U.S.A."
        clpProv = "KY"
    Case "430"
        comCountry = "GERMANY"
        clpProv = "HE"
    Case Else
        comCountry = "CANADA"
        
        'Ticket #20628 - Province is coming from HR_Division now so no defaulting to ON
        'clpProv = "ON"
    End Select

    'Ticket #29759 Franks 02/15/2017
    'get Auto Payroll ID - get auto Payroll ID after click Save
    txtPayrollID.Text = getNextLinPayrollID("N")
    clpVadim1.Text = "N" 'Ticket #29902 Franks 03/03/2017
End If
If glbCompSerial = "S/N - 2192W" Then 'County of Essex
    txtPayrollID = glbLEE_ID
End If

'Ticket #15793
If glbCompSerial = "S/N - 2390W" Then  'Collectcorp Inc
    txtBadgeID.Text = glbLEE_ID
End If

If glbCompSerial = "S/N - 2259W" Then  'For County of Oxford Ticket #16542
    clpCode(4).Text = "N"
End If

If glbCompSerial = "S/N - 2288W" Then
    txtEML = 10
End If

'default values on new hire
If glbCompSerial = "S/N - 2393W" Then 'KTH Shelburne Ticket #17289
    medLICPLATE2.Text = "KSM" 'Company
    medPARKPERMIT2.Text = "H" 'Status
End If

OSMOKER = "" 'Ticket #23491 Franks 04/02/2013
If glbWFC Then 'Ticket #19266 Franks 12/13/2010
    ComSmoker.ListIndex = 1 'default Somker to Yes
End If

SavDept = ""
SavDiv = ""
SavLoc = ""

'Ticket #24557 Franks 07/07/2015 - begin
If glbCompSerial = "S/N - 2420W" Then  'Macaulay Child Development
    ADPBranchOld(0) = "" 'main Branch
    ADPBranchOld(1) = "" 'Alt Payroll ID - Branch 1
    ADPBranchOld(2) = "" 'Alt Payroll ID - Branch 2
    ADPBranchOld(3) = "" 'Alt Payroll ID - Branch 3
    ADPDeptOld(0) = ""  'main dept
    ADPDeptOld(1) = "" 'Alt Payroll ID - dept 1
    ADPDeptOld(2) = "" 'Alt Payroll ID - dept 2
    ADPDeptOld(3) = "" 'Alt Payroll ID - dept 3
End If
'Ticket #24557 Franks 07/07/2015 end

'Ticket #24259 - Adding Dept on New Hire Window
If Len(glbTrsDept) > 0 Then
    clpDept = glbTrsDept
End If

If Len(glbTrsDIV) > 0 Then
    Call SetOtherFieldsFromDiv(glbTrsDIV)
End If

If glbWFC Then 'Ticket #24184 Franks 09/11/2013
    Call WFCHRSoftDispValues
End If

'WDGPHU - Ticket #27899
If glbCompSerial = "S/N - 2411W" Then
    'Multi-Position Employee
    clpCode(6).Text = "NO"
End If

Call SetCountries

Call getCodes

ODeptEDate = dlpDeptEDate.Text
ODivEdate = dlpDivEDate.Text
frmBlank.Visible = False: frmBlank.Top = 100000: frmBlank.Left = 100000

'City of Kawartha Lakes - Remove this pop up (Ticket #11562)
'If glbCompSerial = "S/N - 2363W" Then
'    glbTrsDIV = ""
'    glbTrsVadim1 = ""
'    glbUnionDemog = True
'    frmNewEmployee.Show 1
'    If glbTrsVadim1 <> "Cancel" Then
'        txtUnion.DataField = "ED_ORG"
'        txtVadim1.DataField = "ED_VADIM1"
'        txtUnion.Text = glbTrsDIV
'        txtVadim1.Text = glbTrsVadim1
'    End If
'    glbUnionDemog = False
'End If

'Ticket #9780 - Jerry allowed to make the change to Payroll ID for Aurora
'If txtPayrollID.Enabled And txtPayrollID.Visible And Me.Visible And glbCompSerial <> "S/N - 2378W" Then txtPayrollID.SetFocus

If txtPayrollID.Enabled And txtPayrollID.Visible And Me.Visible Then txtPayrollID.SetFocus

If glbWFC Then 'Ticket #28664 Franks 05/30/2016
    'If clpCode(4).Text = "MISS" Or clpCode(4).Text = "TROY" Then
    If (clpCode(4).Text = "MISS" Or clpCode(4).Text = "TROY") And Not glbtermopen Then 'Ticket #29836 Franks 02/24/2017 - not for term
        lbltitle(55).FontBold = True
        lbltitle(56).FontBold = True
    Else
        lbltitle(55).FontBold = False
        lbltitle(56).FontBold = False
    End If
    'Ticket #30491 Franks 09/07/2017
    medNetworkLogin.Text = ""
    medVendorNo.Text = ""
End If

End Sub

Private Sub SetOtherFieldsFromDiv(xDiv)
Dim rsODiv As New ADODB.Recordset
Dim SQLQ
    'Exit Sub
    clpDiv = xDiv
    SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & xDiv & "' "
    rsODiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsODiv.EOF Then
        If Not IsNull(rsODiv("DV_LOC")) Then
            clpCode(1) = rsODiv("DV_LOC")
        End If
        If Not IsNull(rsODiv("DV_REGION")) Then
            clpCode(2) = rsODiv("DV_REGION")
        End If
        If Not IsNull(rsODiv("DV_ADMINBY")) Then
            clpCode(3) = rsODiv("DV_ADMINBY")
        End If
        If Not IsNull(rsODiv("DV_SECTION")) Then
            clpCode(4) = rsODiv("DV_SECTION")
        End If
        If Not IsNull(rsODiv("DV_COUNTRY")) Then
            txtCountryOfEmp = rsODiv("DV_COUNTRY")
            comCountryOfEmp = txtCountryOfEmp
            txtCountry = rsODiv("DV_COUNTRY")
            comCountry = txtCountry
        End If
        If Not IsNull(rsODiv("DV_BONUSDEPT")) Then
            If glbWFC Then
                'Ticket #27609 Franks 10/07/2015 - comment it out
                'If Not glbWFCHrsSal Then
                '    txtDeptBonusCtr = rsODiv("DV_BONUSDEPT")
                'End If
            Else
                txtDeptBonusCtr = rsODiv("DV_BONUSDEPT")
            End If
        End If
        If glbWFC Then 'Ticket #28637 Franks 05/18/2016
            If Not IsNull(rsODiv("DV_ORGT1")) Then
                clpCode(6).Text = rsODiv("DV_ORGT1")
            End If
        End If
    End If
    rsODiv.Close
    
    If glbWFC Then 'Ticket #27983 Franks 02/10/2015
        If Not glbWFCHrsSal Then 'Salaried employee
            txtDeptBonusCtr = "000000" 'new hire default
        End If
    End If
End Sub
'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdOK_Click()
Dim rc%, DtTm As Variant, X%
Dim xDept, xDiv, ctylist
Dim xDeptEDate, xDivEDate
Dim Msg$
Dim strSQL As String
Dim SQLQ As String
Dim rsJobHis As New ADODB.Recordset
Dim xOldPayID As String
Dim xSmoker 'Ticket #21118
Dim xORG
Dim a%

'''On Error GoTo Add_Err
DtTm = Now

If glbCompSerial = "S/N - 2382W" Then  'Namasco
    'Ticket #16235
    'If Len(txtBadgeID) = 0 Then
    '    txtBadgeID = lblEEID
    'End If
End If

'Ticket #28040 - To Track on New Hire if the user went into the Organizational tab at least once.
'This is to ensure users have filled in values on the Organizational tab - the non mandatory fields.
If fglbNewEE And flgSwitchOrgTabNewHire Then
    tbDemographics.SelectedItem = tbDemographics.Tabs(2)
    Exit Sub
End If

If Not chk_FEBASIC() Then Exit Sub

'Franks 07/13/2012 - reset them to blank since they may carry values from Termination screen
glbChgTermDate = ""
glbChgTermReason = ""

'Ticket #19937 2382W for Samuel Franks 05/08/2011
If glbCompSerial = "S/N - 2382W" Then
    If locUploadWithoutCheck Then
        Call ToRehireSub
        Exit Sub
    End If
End If

If glbVadim Then
    If isTransfer(Demographices) Then
        If fglbNewEE Then
            If ifExistVadimPayrollID Then
                Exit Sub
            End If
        Else
            Screen.MousePointer = DEFAULT
            If oPayrollID <> txtPayrollID Then
                If ifExistVadimPayrollID Then
                    Exit Sub
                Else
                    'Ticket #9780 - Jerry allowed them (Town of Aurora) to make the change to Payroll ID
                    'Jerry said to make City of Kawartha Lakes same as Town of Aurora setup
                    'Ticket #14285 - City of Niagara Falls wants to use the same logic as City of Kawartha Lakes
                    'Ticket #19113 - District Municipality of Muskoka cannot change Payroll ID
                    If glbCompSerial <> "S/N - 2378W" And glbCompSerial <> "S/N - 2363W" And glbCompSerial <> "S/N - 2276W" And glbCompSerial <> "S/N - 2373W" Then
                        'Ticket #24164 - Re-ordering
                        tbDemographics.SelectedItem = tbDemographics.Tabs(1)
                    
                        Msg$ = "A change in the Payroll ID will affect Vadim." & vbNewLine
                        Msg$ = Msg$ & "Is this change due to a data entry mistake?" & vbNewLine & vbNewLine
                        Msg$ = Msg$ & "If no, then Payroll ID can not be changed on this screen"
                        X% = MsgBox(Msg$, vbYesNo)
                        If X% = vbNo Then
                           txtPayrollID = oPayrollID
                           Exit Sub
                        End If
                    Else
                        'City of Kawartha Lakes - Send Email on Payroll ID & Employee Name change
                        'City of Niagara Falls
                        If glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2276W" Then
                            MailBody = ""
                            If (glbCompSerial = "S/N - 2363W") And (OSNAME <> txtSurname Or OFNAME <> txtFName) Then
                                If gSec_Show_ADDRESS Then
                                    If (OADD1 <> txtAdd1) Or (OADD2 <> txtAdd2) Or (OCITY <> txtCity) Or (oProv <> clpProv) Or (OPCODE <> medPCode) Or (oCountry <> comCountry) Then
                                        MailBody = "The Payroll ID, Employee Name and Address has been changed." & vbCrLf & vbCrLf
                                    Else
                                        MailBody = "The Payroll ID and Employee Name has been changed." & vbCrLf & vbCrLf
                                    End If
                                Else
                                    MailBody = "The Payroll ID and Employee Name has been changed." & vbCrLf & vbCrLf
                                End If
                            Else
                                MailBody = "The Payroll ID has been changed." & vbCrLf & vbCrLf
                            End If
                            MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
                            'MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
                            MailBody = MailBody & vbCrLf
                            MailBody = MailBody & "Old Payroll ID: " & oPayrollID & vbCrLf
                            MailBody = MailBody & "New Payroll ID: " & txtPayrollID & vbCrLf
                            
                            If (glbCompSerial = "S/N - 2363W") And (OSNAME <> txtSurname Or OFNAME <> txtFName) Then
                                MailBody = MailBody & vbCrLf
                                MailBody = MailBody & "Old Name: " & OSNAME & ", " & OFNAME & vbCrLf
                                MailBody = MailBody & "New Name: " & txtSurname.Text & ", " & txtFName.Text & vbCrLf
                                
                                If gSec_Show_ADDRESS Then
                                    If (glbCompSerial = "S/N - 2363W") And ((OADD1 <> txtAdd1) Or (OADD2 <> txtAdd2) Or (OCITY <> txtCity) Or (oProv <> clpProv) Or (OPCODE <> medPCode) Or (oCountry <> comCountry)) Then
                                        MailBody = MailBody & vbCrLf
                                        MailBody = MailBody & "Old Address: " & vbCrLf
                                        MailBody = MailBody & "Address: " & OADD1 & vbCrLf
                                        MailBody = MailBody & "Address 2: " & OADD2 & vbCrLf
                                        MailBody = MailBody & "City: " & OCITY & vbCrLf
                                        MailBody = MailBody & "Province: " & oProv & vbCrLf
                                        MailBody = MailBody & "Postal Code: " & OPCODE & vbCrLf
                                        MailBody = MailBody & "Country: " & oCountry & vbCrLf
                                        MailBody = MailBody & vbCrLf
                                        MailBody = MailBody & "New Address: " & vbCrLf
                                        MailBody = MailBody & "Address: " & txtAdd1 & vbCrLf
                                        MailBody = MailBody & "Address 2: " & txtAdd2 & vbCrLf
                                        MailBody = MailBody & "City: " & txtCity & vbCrLf
                                        MailBody = MailBody & "Province: " & clpProv & vbCrLf
                                        MailBody = MailBody & "Postal Code: " & medPCode & vbCrLf
                                        MailBody = MailBody & "Country: " & comCountry & vbCrLf
                                    End If
                                End If
                            Else
                                If gSec_Show_ADDRESS Then
                                    If (glbCompSerial = "S/N - 2363W") And ((OADD1 <> txtAdd1) Or (OADD2 <> txtAdd2) Or (OCITY <> txtCity) Or (oProv <> clpProv) Or (OPCODE <> medPCode) Or (oCountry <> comCountry)) Then
                                        MailBody = MailBody & vbCrLf
                                        MailBody = MailBody & "Old Address: " & vbCrLf
                                        MailBody = MailBody & "Address: " & OADD1 & vbCrLf
                                        MailBody = MailBody & "Address 2: " & OADD2 & vbCrLf
                                        MailBody = MailBody & "City: " & OCITY & vbCrLf
                                        MailBody = MailBody & "Province: " & oProv & vbCrLf
                                        MailBody = MailBody & "Postal Code: " & OPCODE & vbCrLf
                                        MailBody = MailBody & "Country: " & oCountry & vbCrLf
                                        MailBody = MailBody & vbCrLf
                                        MailBody = MailBody & "New Address: " & vbCrLf
                                        MailBody = MailBody & "Address: " & txtAdd1 & vbCrLf
                                        MailBody = MailBody & "Address 2: " & txtAdd2 & vbCrLf
                                        MailBody = MailBody & "City: " & txtCity & vbCrLf
                                        MailBody = MailBody & "Province: " & clpProv & vbCrLf
                                        MailBody = MailBody & "Postal Code: " & medPCode & vbCrLf
                                        MailBody = MailBody & "Country: " & comCountry & vbCrLf
                                    End If
                                End If
                            End If
                            'Screen.MousePointer = DEFAULT
                            'Call imgEmail_Click
                        End If
                    End If
                End If
            Else
                'City of Kawartha Lakes - Employee Name change - Send email
                If glbCompSerial = "S/N - 2363W" Then
                    If OSNAME <> txtSurname Or OFNAME <> txtFName Then
                        If gSec_Show_ADDRESS Then
                            If (OADD1 <> txtAdd1) Or (OADD2 <> txtAdd2) Or (OCITY <> txtCity) Or (oProv <> clpProv) Or (OPCODE <> medPCode) Or (oCountry <> comCountry) Then
                                MailBody = "The Employee Name and Address has been changed." & vbCrLf & vbCrLf
                            Else
                                MailBody = "The Employee Name has been changed." & vbCrLf & vbCrLf
                            End If
                        End If
                        
                        MailBody = MailBody & "Payroll ID: " & txtPayrollID & vbCrLf
                        MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
                        MailBody = MailBody & vbCrLf
                        MailBody = MailBody & "Old Name: " & OSNAME & ", " & OFNAME & vbCrLf
                        MailBody = MailBody & "New Name: " & txtSurname.Text & ", " & txtFName.Text & vbCrLf
                        
                        If gSec_Show_ADDRESS Then
                            If (OADD1 <> txtAdd1) Or (OADD2 <> txtAdd2) Or (OCITY <> txtCity) Or (oProv <> clpProv) Or (OPCODE <> medPCode) Or (oCountry <> comCountry) Then
                                MailBody = MailBody & vbCrLf
                                MailBody = MailBody & "Old Address: " & vbCrLf
                                MailBody = MailBody & "Address: " & OADD1 & vbCrLf
                                MailBody = MailBody & "Address 2: " & OADD2 & vbCrLf
                                MailBody = MailBody & "City: " & OCITY & vbCrLf
                                MailBody = MailBody & "Province: " & oProv & vbCrLf
                                MailBody = MailBody & "Postal Code: " & OPCODE & vbCrLf
                                MailBody = MailBody & "Country: " & oCountry & vbCrLf
                                MailBody = MailBody & vbCrLf
                                MailBody = MailBody & "New Address: " & vbCrLf
                                MailBody = MailBody & "Address: " & txtAdd1 & vbCrLf
                                MailBody = MailBody & "Address 2: " & txtAdd2 & vbCrLf
                                MailBody = MailBody & "City: " & txtCity & vbCrLf
                                MailBody = MailBody & "Province: " & clpProv & vbCrLf
                                MailBody = MailBody & "Postal Code: " & medPCode & vbCrLf
                                MailBody = MailBody & "Country: " & comCountry & vbCrLf
                            End If
                        End If
                    Else
                        If gSec_Show_ADDRESS Then
                            If (OADD1 <> txtAdd1) Or (OADD2 <> txtAdd2) Or (OCITY <> txtCity) Or (oProv <> clpProv) Or (OPCODE <> medPCode) Or (oCountry <> comCountry) Then
                                MailBody = "The Employee Address has been changed." & vbCrLf & vbCrLf
                                MailBody = MailBody & "Payroll ID: " & txtPayrollID & vbCrLf
                                MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
                                MailBody = MailBody & vbCrLf
                                MailBody = MailBody & "Old Address: " & vbCrLf
                                MailBody = MailBody & "Address: " & OADD1 & vbCrLf
                                MailBody = MailBody & "Address 2: " & OADD2 & vbCrLf
                                MailBody = MailBody & "City: " & OCITY & vbCrLf
                                MailBody = MailBody & "Province: " & oProv & vbCrLf
                                MailBody = MailBody & "Postal Code: " & OPCODE & vbCrLf
                                MailBody = MailBody & "Country: " & oCountry & vbCrLf
                                MailBody = MailBody & vbCrLf
                                MailBody = MailBody & "New Address: " & vbCrLf
                                MailBody = MailBody & "Address: " & txtAdd1 & vbCrLf
                                MailBody = MailBody & "Address 2: " & txtAdd2 & vbCrLf
                                MailBody = MailBody & "City: " & txtCity & vbCrLf
                                MailBody = MailBody & "Province: " & clpProv & vbCrLf
                                MailBody = MailBody & "Postal Code: " & medPCode & vbCrLf
                                MailBody = MailBody & "Country: " & comCountry & vbCrLf
                            Else
                                MailBody = ""
                            End If
                        Else
                            MailBody = ""
                        End If
                    End If
                End If
            End If
            Screen.MousePointer = HOURGLASS
            
        End If
    End If
End If
'Frank Sep 19,2003

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Canada Inc Ticket #13512
    glbChgTermDate = ""
    glbChgTermReason = ""
    glbChgNewEmpnbr = lblEEID
    Screen.MousePointer = DEFAULT
    If SavDiv <> clpDiv.Text Then
        If Len(SavDiv) > 0 And Len(clpDiv.Text) > 0 Then 'Ticket #13512
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
            'Dim a%
            Msg$ = lStr("Division") & " change will cause an employee termination for old company" & Chr(10) & "and a new hire for new company in ADP system"
            Msg$ = Msg$ & Chr(10) & "Are you sure you want to do it? "
            a% = MsgBox(Msg, 36, "Confirm Update")
            If a% <> 6 Then
                glbChgTermReason = "***"
            Else
                frmMsgTerm.Show 1
            End If
        End If
    End If
    Screen.MousePointer = HOURGLASS
End If

If glbCompSerial = "S/N - 2460W" Then 'Oshawa Public Libraries = Ticket #25323 Franks 12/16/2014
    glbChgTermDate = ""
    glbChgTermReason = ""
    'glbChgNewEmpnbr = lblEEID
    Screen.MousePointer = DEFAULT
    If oRegion <> clpCode(2).Text Then
        If Len(oRegion) > 0 And Len(clpCode(2).Text) > 0 Then
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
            'Dim a%
            Msg$ = lStr("Region") & " change will cause an employee termination for old company and a new hire for new company in Easypay system"
            Msg$ = Msg$ & Chr(10) & Chr(10) & "Are you sure you want to do it? "
            a% = MsgBox(Msg, 36, "Confirm Update")
            If a% <> 6 Then
                glbChgTermReason = "***"
            Else
                frmMsgTerm.Show 1
            End If
        End If
    End If
End If

If glbSoroc Or glbCompSerial = "S/N - 2191W" Or glbCompSerial = "S/N - 2192W" Then 'Or glbCompSerial = "S/N - 2373W" Then    ' Soroc 'City of St-Thomas 'County of Essex 'DMuskoka
    glbChgTermDate = ""
    glbChgTermReason = ""
    glbChgNewEmpnbr = lblEEID
    If rsDATA.EOF Then
        SorocOPayrollID = ""
    Else
        SorocOPayrollID = IIf(IsNull(rsDATA("ED_PAYROLL_ID")), "", rsDATA("ED_PAYROLL_ID"))
    End If
    Screen.MousePointer = DEFAULT
    If SorocOPayrollID <> txtPayrollID Then
        If Len(SorocOPayrollID) > 0 And Len(txtPayrollID) Then
            frmMsgTerm.Show 1
        End If
    End If
    Screen.MousePointer = HOURGLASS
End If

'Frank Jul 16,2004
'Ticket #16616 Frank Apr 16, 2009 -  Remove Termination/New Hire process from the Payroll ID change
'If glbWFC Then 'WFC
'    If clpCode(4).Text <> "GREN" Then
'        If Not fglbNewEE Then
'            glbChgTermDate = ""
'            glbChgTermReason = ""
'            glbChgNewEmpnbr = lblEEID
'            SorocOPayrollID = IIf(IsNull(rsDATA("ED_PAYROLL_ID")), "", rsDATA("ED_PAYROLL_ID"))
'            Screen.MousePointer = DEFAULT
'            If SorocOPayrollID <> txtPayrollID Then
'                If Len(SorocOPayrollID) > 0 And Len(txtPayrollID) Then
'                    xOldPayID = "OldPayID" & Trim(SorocOPayrollID)
'                    frmMsgTerm.Show 1
'                End If
'            End If
'            Screen.MousePointer = HOURGLASS
'        End If
'    End If
'End If
'Jaddy Mar 24, 2004
'If glbCompSerial = "S/N - 2241W" Or glbCompSerial = "S/N - 2382W" Then
'Ticket #21791 Franks 03/29/2012 - Samuel not want this function
If glbCompSerial = "S/N - 2241W" Then
'2241W - Granite Club : 2382W - Namasco Ltd.
    glbChgTermDate = ""
    glbChgTermReason = ""
    glbChgNewEmpnbr = lblEEID
    If rsDATA.EOF Then
        oAdminBy = ""
    Else
        oAdminBy = IIf(IsNull(rsDATA("ED_ADMINBY")), "", rsDATA("ED_ADMINBY"))
    End If
    Screen.MousePointer = DEFAULT
    If oAdminBy <> clpCode(3) And oAdminBy <> "" Then
        frmMsgTerm.Show 1
    End If
    Screen.MousePointer = HOURGLASS
End If

rsDATA.Requery

If fglbNewEE Then rsDATA.AddNew

Screen.MousePointer = HOURGLASS

'If Len(txtDeptEDate) = 0 Then txtDeptEDate = Date
'If Len(txtDivEDate) = 0 And Len(txtDiv) > 0 Then txtDivEDate = Date

If SavDept <> clpDept.Text Then xDept = clpDept.Text Else xDept = ""
If SavDiv <> clpDiv.Text Then xDiv = clpDiv.Text Else xDiv = "*"
If ODeptEDate <> dlpDeptEDate.Text And IsDate(dlpDeptEDate.Text) Then
    xDeptEDate = CVDate(dlpDeptEDate.Text)
    
    'Ticket #24129 - If Dept has not changed then Old and New Dept will be same
    If xDept = "" Then
        xDept = clpDept.Text
    End If
Else
    xDeptEDate = Date
End If

If glbSamuel Then 'Ticket #22912 Franks 12/06/2012
    If SavDept <> clpDept.Text Then
        If IsDate(dlpDeptEDate.Text) Then
            If CVDate(dlpDeptEDate.Text) > Date Then
                xFutureChgDeptNo = True: xFutureDateDeptNo = CVDate(dlpDeptEDate.Text)
            End If
        End If
    End If
End If
If ODivEdate <> dlpDivEDate.Text And IsDate(dlpDivEDate.Text) Then
    xDivEDate = CVDate(dlpDivEDate.Text)
    
    'Ticket #24129 - If Divison has not changed then Old and New Division will be same
    If xDiv = "*" Then
        xDiv = clpDiv.Text
    End If
Else
    xDivEDate = Date
End If

'Ticket #24543 - Macaulay Child Development Centre
If glbCompSerial = "S/N - 2420W" And SavOrg <> clpCode(0).Text Then
    xORG = ""
    If Len(clpCode(2).Text) > 0 Then xORG = clpCode(2).Text Else xORG = "*"
End If


'' dkostka - 11/21/2001 - Keypress events don't fire for gender option buttons, don't know why, but this
''   fixes it.
'If optGender(0).Value = True Then
'    txtGender = "M"
'Else
'    txtGender = "F"
'End If


'Hemu - 06/10/2003 Begin - txtMStatus was being re-assigned to the first char. of the status
'                          this was giving logical error cause it was assigning same
'                          char. for 'Partner' and 'Parent(Single)'. Added the same logic
'                          as its on Lost_focus of ComMStatus
    If ComMStat = "Partner" Or ComMStat = "Same-Sex" Then
        txtMStatus = UCase(Right(ComMStat.Text, 1))
    ElseIf ComMStat = "Separated" Then
        txtMStatus = UCase(Mid(ComMStat.Text, 4, 1))
    Else
        txtMStatus = Left(ComMStat.Text, 1)
    End If
'Hemu - 06/10/2003 End

txtCountry = comCountry
glbEmpCountry = comCountry
txtCountryOfEmp = comCountryOfEmp

If ComSmoker = "Yes" Then txtSmoker = "-1" Else txtSmoker = "0"

ctylist = CountryList

Call UpdUStats(Me)

If fglbNewEE Then
    If SavDept <> clpDept.Text Or ODeptEDate <> dlpDeptEDate.Text Then
        If Not EmpHisCalc(2, glbLEE_ID, xDept, "", "", "", "", "", "", xDeptEDate) Then MsgBox "EMPHIS Error "
    End If
    If SavDiv <> clpDiv.Text Or ODivEdate <> dlpDivEDate.Text Then
        If Len(Trim(clpDiv.Text)) > 0 Then 'Ticket #23837 Franks 05/28/2013 - no Div then no history
            If Not EmpHisCalc(2, glbLEE_ID, "", xDiv, "", "", "", "", "", xDivEDate) Then MsgBox "EMPHIS Error "
        End If
    End If
    'Ticket #23837 Franks 05/28/2013 - begin
    'If SavLoc <> clpCode(1) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", Date, "LOC", clpCode(1)) Then MsgBox "EMPHIS Error "
    'If glbLinamar Then
    '    If oRegion <> clpCode(2) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", Date, "REGION", getProductLineCodeforLinamar(clpCode(2).TransDiv & clpCode(2).Text)) Then MsgBox "EMPHIS Error "
    'Else
    '    If oRegion <> clpCode(2) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", Date, "REGION", clpCode(2)) Then MsgBox "EMPHIS Error "
    'End If
    'If oAdminBy <> clpCode(3) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", Date, "ADMINBY", clpCode(3)) Then MsgBox "EMPHIS Error "
    'If OSection <> clpCode(4) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", Date, "SECTION", clpCode(4)) Then MsgBox "EMPHIS Error "
    If SavLoc <> clpCode(1) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", dlpDate(0).Text, "LOC", clpCode(1)) Then MsgBox "EMPHIS Error "
    If glbLinamar Then
        If oRegion <> clpCode(2) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", Date, "REGION", getProductLineCodeforLinamar(clpCode(2).TransDiv & clpCode(2).Text)) Then MsgBox "EMPHIS Error "
    Else
        If oRegion <> clpCode(2) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", dlpDate(0).Text, "REGION", clpCode(2)) Then MsgBox "EMPHIS Error "
    End If
    If oAdminBy <> clpCode(3) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", dlpDate(0).Text, "ADMINBY", clpCode(3)) Then MsgBox "EMPHIS Error "
    If OSection <> clpCode(4) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", dlpDate(0).Text, "SECTION", clpCode(4)) Then MsgBox "EMPHIS Error "
    'Ticket #23837 Franks 05/28/2013 - end
    
    'Ticket #24543 - Macaulay Child Development Centre
    If glbCompSerial = "S/N - 2420W" Then
        If xORG <> "*" And Len(xORG) > 0 Then
            If Not EmpHisCalc(1, glbLEE_ID, "", "", "", "", xORG, "", "", dlpDate(0).Text) Then MsgBox "EMPHIS Error"
        End If
    End If
    
    'Ticket #28794 - Add Marital Status to Employee History for everyone
    If Not OMSTAT = txtMStatus Then
        If Not EmpHisCalc(6, glbLEE_ID, "", "", "", "", "", "", "", Date, , txtMStatus.Text) Then MsgBox "EMPHIS Error "
    End If
    
    If glbCountry <> "U.S.A." And glbEmpCountry <> "U.S.A." And Not (glbWFC And glbEmpCountry = "Canada") Then
        ' bank defaults removed for the states Bryan 17/Mar/06
        Dim rsPA As New ADODB.Recordset
        If Not (glbCompSerial = "S/N - 2370W") Then 'David Chapman's
            rsPA.Open "select PC_NEXT_AVAILABLE_NBR,PC_FEDTAX,PC_PROVTAX from HRPARCO", gdbAdoIhr001, adOpenStatic, adLockPessimistic
            
            rsDATA("ED_TD1") = "Y"
            'If rsPA!PC_FEDTAX = 0 Then
            '    RSDATA("ED_TD1DOL") = 7634 'Ticket #25879 Franks 12/10/2014 don't use this
            'Else
                rsDATA("ED_TD1DOL") = rsPA!PC_FEDTAX '7634 '7131
            'End If
            locFEDTAX = rsDATA("ED_TD1DOL")
            locPROVTAX = ""
            If clpProv.Text = "ON" Then
                rsDATA("ED_PROVFORM") = "Y"
                'If rsPA!PC_PROVTAX = 0 Then
                '   RSDATA("ED_PROVAMT") = 7686 'Ticket #25879 Franks 12/10/2014 don't use this
                'Else
                    rsDATA("ED_PROVAMT") = rsPA!PC_PROVTAX '7686
                'End If
                locPROVTAX = rsDATA("ED_PROVAMT")
            Else
                If IsNumeric(rsPA!PC_PROVTAX) Then
                    rsDATA("ED_PROVAMT") = rsPA!PC_PROVTAX '7686
                End If
            End If
            rsPA.Close
        End If
    End If
    
    If glbCompSerial = "S/N - 2229W" Or glbCompSerial = "S/N - 2369W" Then  'Inscape Solutions
        rsDATA("ED_PROVEMP") = "ON"
        rsDATA("ED_UIC") = "1"
        locPROVEMP = "ON"
        locUIC = "1"
    End If
        
    If glbSamuel Then 'Ticket #21195 Franks 11/14/2011
        If clpCode(2).Text = "GO" Then
            rsDATA("ED_UIC") = "P"
        End If
        
        If NewHireForms.count = 0 Then  'non New Hire only 'Ticket #23490 Franks 03/28/2013
            'Ticket #22912 Franks 12/06/2012
            Call SamuelSectionChg
            Call SamuelRegionChg
        End If
    End If
    
        'Ticket #21134 - Since on Status/Dates screen the From Date is autopopulated with Hire Date, I am
        'doing the same here otherwise if no date on Status/Dates screen is updated or lost_focus'd on
        'then the From Date is not auto-populated.
        If NewHireForms.count > 0 And IsDate(dlpDate(0).Text) Then
            rsDATA("ED_SFDATE") = dlpDate(0).Text
        End If
        
        If glbWFC Then 'Ticket #23564 Franks 04/15/2013
            xBenGrpCode = "" 'Ticket #23247 Franks 04/22/2013
            xWFCPayGroup = "" 'Ticket #23247 Franks 04/22/2013
            xWFCNGSCode = "" 'Ticket #23247 Franks 04/22/2013
            If NewHireForms.count > 0 Then
                'If comCountryOfEmp.Text = "U.S.A." Then
                    rsDATA("ED_PTEDATE") = dlpDate(0).Text
                    rsDATA("ED_NORMALR") = getWFCRetireDate(dlpDOB.Text) ' Ticket #24695 Franks 11/26/2013
                'End If
                    If glbWFC_US_Ben_Trans Then 'Ticket #23247 Franks 04/22/2013
                        Call WFC_EmpFieldsForUSBen(rsDATA, glbLEE_ID)
                        If glbTrsStatus = "COOP" Or glbTrsStatus = "STUD" Then 'Ticket #25352 Franks 04/16/2014
                            '"   If Employment Status = COOP or STUD, don't update the Benefit Master Code or NGS Sub Group
                            xWFCNGSCode = ""
                        End If
                    End If
                    'Ticket #24184 Franks 09/12/2013 - from HRSoft New Hire
                    If glbCandidate > 0 Then
                        rsDATA("ED_CANDIDATE") = glbCandidate
                        If Len(xHRSoftPTCode) > 0 Then rsDATA("ED_PT") = xHRSoftPTCode
                    End If
            End If
        End If
        
        If Not AUDITDEMO("A") Then MsgBox "ERROR : AUDIT FILE"
        
        If glbPayWeb Then
            rsDATA("ED_WCB") = "Y"
        End If
        If glbInsync Then
            rsDATA("ED_WCB") = "0"
            If glbCompSerial = "S/N - 2295W" Then
                rsDATA("ED_CPP") = "O" '"0"
            ElseIf glbCompSerial = "S/N - 2383W" Then 'Town of Orangeville Ticket #21492 Franks 01/25/2012
                rsDATA("ED_CPP") = " "
            Else
                rsDATA("ED_CPP") = "0"
            End If
        End If
        If glbVadim Then
            If glbCompSerial = "S/N - 2373W" Then 'Ticket #24565- District Municipality of South Muskoka
                rsDATA("ED_WCB") = "1"
            Else
                rsDATA("ED_WCB") = "N"
            End If
            rsDATA("ED_CPP") = "Y"
            'If glbCompSerial <> "S/N - 2379W" Then   'Ticket #23795 - Not Town of LaSalle (values are not Y/N) - Getting an error, need to pass default
                rsDATA("ED_UIC") = "Y"
            'End If
            rsDATA("ED_GROSSCD") = "Y"
        End If
        
        'Wellington-Dufferin-Guelph Public Health - Ticket #17129
        If glbCompSerial = "S/N - 2411W" Then
            rsDATA("ED_UIC") = "0"
            rsDATA("ED_GROSSCD") = "0"
        End If
        
        'Automatic Email - County of Lambton
        If glbLambton Then
            If Not rsDATA.EOF Then
                rsDATA("ED_EMAIL") = LCase(txtFName.Text) & "." & LCase(txtSurname.Text) & "@county-lambton.on.ca"
            End If
        End If
        
        'Ticket #24565- District Municipality of South Muskoka
        If glbCompSerial = "S/N - 2373W" Then
            rsDATA("ED_PROVEMP") = "ON"
            rsDATA("ED_WCBCODE") = "1"
        End If
        
        'City of Kawartha Lakes
        If glbCompSerial = "S/N - 2363W" Then
            rsDATA("ED_PROVEMP") = "ON"
        End If
        
        'Ticket #28786 - Goodmans
        If glbCompSerial = "S/N - 2290W" Then
            rsDATA("ED_PROVEMP") = "ON"
        End If
        
        If glbCompSerial = "S/N - 2381W" Then 'Ticket #13603
            rsDATA("ED_UIC") = "Y"
            rsDATA("ED_CPP") = "Y"
        End If
        
        'Ticket #21504 - Kerry's Place
        If glbCompSerial = "S/N - 2433W" Then
            rsDATA("ED_UIC") = "1"
            rsDATA("ED_CPP") = "Y"
        End If
        
        If glbCompSerial = "S/N - 2353W" Then   'Let's Talk Science Ticket #27072 10/14/2015
            rsDATA("ED_UIC") = "1"
            rsDATA("ED_CPP") = "1"
        End If
        
        If rsDATA("ED_PROVEMP") = "" Or IsNull(rsDATA("ED_PROVEMP")) Then
            rsDATA("ED_PROVEMP") = rsDATA("ED_PROV")
        End If
    
        'City of Timmins - For RPP # (Vadim)
        'If glbCompSerial = "S/N - 2375W" Then
        '    rsDATA("ED_PENSION") = "1"
        '    rsDATA("ED_NORMALR") = DateAdd("yyyy", 65, CVDate(dlpDOB))
        'End If
    
    Call UpdCodes
    Call Set_Control("U", Me, rsDATA)
    If glbCompSerial = "S/N - 2380W" Then   'VitalAire Ticket #12142
        rsDATA("ED_EMPTYPE") = rsDATA("ED_SECTION")
    End If
    If glbLinamar Then 'Ticket #18770
        If NewHireForms.count > 0 Then
            rsDATA("ED_VADIM2") = "Yes"
            If Len(clpVadim1.Text) = 0 Then 'Ticket #29759 Franks 02/14/2017
                rsDATA("ED_VADIM1") = "N"
            End If
        End If
    End If
    rsDATA.Update
    
    'If glbWFC Then 'Ticket# 7185
        Call InsBlankEmpOther(glbLEE_ID)
    'End If
    
    Call AddNewPayrollEmp(Demographices, Date, glbLEE_ID, txtPayrollID)
    glbNextEmpl = glbNextEmpl + 1
    If glbCompSerial = "S/N - 2241W" Then ' Not Granite Club
        Call Check_EMPLOYEE_Number(glbNextEmpl)
    End If
    
    If glbWFC Then 'Ticket #23491 Franks 04/02/2013
        If Not OSMOKER = ComSmoker Then
            xSmoker = ComSmoker.Text
            If NewHireForms.count > 0 Then ''Ticket #23837 Franks 05/28/2013
                If Not EmpHisCalc(5, glbLEE_ID, "", "", "", "", "", "", "", dlpDate(0).Text, , xSmoker, , , OSMOKER) Then MsgBox "EMPHIS Error "
            Else
                If Not EmpHisCalc(5, glbLEE_ID, "", "", "", "", "", "", "", Date, , xSmoker, , , OSMOKER) Then MsgBox "EMPHIS Error "
            End If
        End If
        
        If NewHireForms.count > 0 Then
            Call WFC_UptUSBenByEmp(glbLEE_ID, CVDate(dlpDate(0).Text), glbTrsHourWeek)  'Ticket #23247 Franks 04/22/2013
            
            If glbCandidate > 0 Then 'Ticket #24695 Franks 11/28/2013
                Call uptEEO_Fields(glbLEE_ID, "New", , , xHRSoftJob, xETHNICITY, xRACE)
            End If
        End If
        
        Call UptWFCEMPOTHER   '#28637 Franks 05/18/2016
    End If
    
    fglbNewEE = False
    Call ST_UPD_MODE(True)
    SET_UP_MODE
    X% = modECount(True)
    
    'If SavLoc <> clpCode(1).Text Then
    '    Call Update_Overtime_Bank(clpCode(1).Text)
    'End If
    'If oRegion <> clpCode(2).Text Then
    '    Call Update_Overtime_Bank(clpCode(2).Text)
    'End If
    'If oAdminBy <> clpCode(3).Text Then
    '    Call Update_Overtime_Bank(clpCode(3).Text)
    'End If
    'If OSection <> clpCode(4).Text Then
    '    Call Update_Overtime_Bank(clpCode(4).Text)
    'End If

    If NewHireForms.count > 0 Then  'New Hire only
        If Not glbtermopen Then
            Call UPDOvertime_Overview
        End If
    End If

    If gsEMAIL_ONNEWHIRE Then
        If NewHireForms.count > 0 Then 'new hire
            If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
                'MailBody = "Employee # " & lblEENum.Caption & " "
                'MailBody = MailBody & "- " & lblEEName.Caption & " "
                'If Len(clpCode(3).Text) > 0 Then
                '    MailBody = MailBody & "in " & lStr("Administered By") & " " & clpCode(3).Text & " "
                'End If
                'If Len(clpCode(4).Text) > 0 Then
                '    MailBody = MailBody & "in " & lStr("Section") & " " & clpCode(4).Text & " "
                'End If
                'MailBody = MailBody & "was hired on " & "DOH" & " "
                'Screen.MousePointer = DEFAULT
                'Call imgEmail_Click
                'Screen.MousePointer = HOURGLASS
                '--------- Move the email sending function to Status/Dates screen, they need DOH in email
            ElseIf glbCompSerial = "S/N - 2433W" Then  'Kerry's Place - Ticket #24692
                '**************Do not send email now, send once the Position is added.***************
            ElseIf glbWFC Then 'Ticket #28763 Franks 06/21/2016
                '**************Do not send email now, send once the Position is added.***************
            Else
                MailBody = "The new employee has been hired." & vbCrLf & vbCrLf
                MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
                MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
                Screen.MousePointer = DEFAULT
                Call imgEmail_Click
                Screen.MousePointer = HOURGLASS
            End If
        End If
    End If
    
    Call NextForm
    
    '7.9 - Nice to Haves - If new hires in Canada and SIN # starting with 9 then Other Information screen
    'needs to be filled in
    If Left(medSIN, 1) = "9" And UCase(txtCountry.Text) = "CANADA" Then
        'Add Other Information form in the NewHireForms collection
        Call get_NewHireForms_Add_OtherInformation_Form
    End If
    
Else
    If Not glbtermopen Then
        If SavDept <> clpDept.Text Or ODeptEDate <> dlpDeptEDate.Text Then
            If Not EmpHisCalc(1, glbLEE_ID, xDept, "", "", "", "", "", "", xDeptEDate) Then MsgBox "EMPHIS Error "
        End If
        If SavDiv <> clpDiv.Text Or ODivEdate <> dlpDivEDate.Text Then
            If Not EmpHisCalc(1, glbLEE_ID, "", xDiv, "", "", "", "", "", xDivEDate) Then MsgBox "EMPHIS Error "
        End If
        If SavLoc <> clpCode(1) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", Date, "LOC", clpCode(1)) Then MsgBox "EMPHIS Error "
        If glbLinamar Then
            If oRegion <> clpCode(2) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", Date, "REGION", getProductLineCodeforLinamar(clpCode(2).TransDiv & clpCode(2).Text)) Then MsgBox "EMPHIS Error "
        Else
            If oRegion <> clpCode(2) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", Date, "REGION", clpCode(2)) Then MsgBox "EMPHIS Error "
        End If
        If oAdminBy <> clpCode(3) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", Date, "ADMINBY", clpCode(3)) Then MsgBox "EMPHIS Error "
        If OSection <> clpCode(4) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", Date, "SECTION", clpCode(4)) Then MsgBox "EMPHIS Error "
        If glbWFC Then 'Ticket #21118 Franks 10/27/2011
            If Not OSMOKER = ComSmoker Then
                'If ComSmoker.Text = "Yes" Then xSMOKER = 1 Else xSMOKER = 0
                xSmoker = ComSmoker.Text
                'If Not EmpHisCalc(5, glbLEE_ID, "", "", "", "", "", "", "", Date, , xSMOKER) Then MsgBox "EMPHIS Error "
                'Ticket #23491 Franks 04/02/2013
                If Not EmpHisCalc(5, glbLEE_ID, "", "", "", "", "", "", "", Date, , xSmoker, , , OSMOKER) Then MsgBox "EMPHIS Error "
            End If
        End If
                
        'Ticket #28794 - Add Marital Status to Employee History for everyone
        If Not OMSTAT = txtMStatus Then
            If Not EmpHisCalc(6, glbLEE_ID, "", "", "", "", "", "", "", Date, , txtMStatus.Text) Then MsgBox "EMPHIS Error "
        End If
        
        'Ticket #24543 - Macaulay Child Development Centre
        If glbCompSerial = "S/N - 2420W" Then
            If xORG <> "*" And Len(xORG) > 0 Then
                If Not EmpHisCalc(1, glbLEE_ID, "", "", "", "", xORG, "", "", Date) Then MsgBox "EMPHIS Error"
            End If
        End If
        
        If glbSamuel Then 'Ticket #22912 Franks 12/06/2012
            If NewHireForms.count = 0 Then  'non New Hire only 'Ticket #23490 Franks 03/28/2013
                Call SamuelSectionChg
                Call SamuelRegionChg
            End If
        End If
        
        If Not AUDITDEMO("M") Then MsgBox "ERROR : AUDIT FILE"
        
        Call SecurityMod
        
        'Kerry's Place - Ticket #24692 - Update matching Current Positions fields
        If glbCompSerial = "S/N - 2433W" Then
            If SavDept <> clpDept.Text Or SavDiv <> clpDiv.Text Then
                Call UpdateCurrentPosition_KerrysPlace
            End If
        Else
            Call UpdateCurrentPosition
        End If
        
        'Ticket #19785 - Update employee's Current Position's multiposition fields with default values for employees
        'with non multi position setting (ED_SECTION <> "Y").
        If glbCompSerial = "S/N - 2259W" Then  'For County of Oxford
            If clpCode(4).Text <> "Y" Then
                Call UpdateOxfordCurrentPosition
            End If
        End If
        'Ticket #27899 - Also for WDGPHU as above
        If glbCompSerial = "S/N - 2411W" Then   'For WDGPHU
            If clpCode(6).Text <> "YES" Then
                Call UpdateOxfordCurrentPosition
            End If
        End If
    Else
        If Not AUDITDEMO("M") Then MsgBox "ERROR : AUDIT FILE"
    End If
    
    If comCountry = "CANADA" Then
        rsDATA("ED_SIN") = Format(Val(medSIN), "#########")
    Else
        rsDATA("ED_SIN") = medSIN
    End If
    If rsDATA("ED_PROVEMP") = "" Or IsNull(rsDATA("ED_PROVEMP")) Then
        rsDATA("ED_PROVEMP") = rsDATA("ED_PROV")
    End If
                
    Call UpdCodes
    Call Set_Control("U", Me, rsDATA)
    
    If glbCompSerial = "S/N - 2380W" Then   'VitalAire Ticket #12142
        rsDATA("ED_EMPTYPE") = rsDATA("ED_SECTION")
    End If
    rsDATA.Update
    
    Call ST_UPD_MODE(True)
    
    If glbCompSerial = "S/N - 2241W" Then 'Granite Club
        If glbChgTermReason <> "" And oAdminBy <> "" Then
            MsgBox "Please check the employee's current Position/Salary and make the changes if necessary."
            If glbLEE_ID <> glbChgNewEmpnbr Then
                Call ChangeEmpnbr
                Call UnloadFrms
            End If
        End If
    End If
    
    If glbCompSerial = "S/N - 2191W" Then 'city of st-thomas
        If glbChgTermReason <> "" And oPayrollID <> "" Then
            MsgBox "Please check the employee's current Position/Salary and make the changes if necessary."
            If glbLEE_ID <> glbChgNewEmpnbr Then
                Call ChangeEmpnbr
                Call UnloadFrms
            End If
        End If
    End If
    'COMMENTED BY SAM AS THIS FUNCTION SHOULD BE CALLED ON STATUS AND DATES SCREEN ONLY 08/16/2006
    'Call UPDEML
    
    'City of Kawartha Lakes - Send email on Payroll ID change
    'City of Niagara Falls - Ticket #14285 - Send email on Payroll ID change
    If gsEMAIL_ONNEWHIRE Then
        If glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2276W" Then
            If Len(MailBody) > 0 Then
                Screen.MousePointer = DEFAULT
                Call imgEmail_OtherChanges_Click
            End If
        End If
    End If
 '   Call EERetrieve
'    txtPayrollID.SetFocus

    'Ticket #18235 - Hours/Day/Week/Period - Samuel, Son & Co., Limited
    If glbCompSerial = "S/N - 2382W" And (clpCode(3).Text <> oAdminBy Or clpCode(4).Text <> OSection) Then
        '''Get Payroll ID and Branch #
        ''Select Case clpCode(3).Text     'Payroll #
        ''    Case "2158"
        ''        Call Update_Position_Hours_DWP("8", "40", "40")
        ''    Case "5230"
        ''        If clpCode(4).Text <> "15" Then     'Branch Code
        ''            Call Update_Position_Hours_DWP("7.5", "37.5", "81.25")
        ''        Else
        ''            Call Update_Position_Hours_DWP("8", "40", "86.67")
        ''        End If
        ''    Case "5231"
        ''        Call Update_Position_Hours_DWP("7.5", "37.5", "81.25")
        ''    Case "5232"
        ''        Call Update_Position_Hours_DWP("7.4", "37", "80.17")
        ''    Case "5322"
        ''        If clpCode(4).Text <> "02" And clpCode(4).Text <> "Z" Then  'Branch Code
        ''            Call Update_Position_Hours_DWP("8", "40", "40")
        ''        Else
        ''            Call Update_Position_Hours_DWP("12", "60", "60")
        ''        End If
        ''End Select
        'Ticket #21652 Franks 03/20/2012 they don't use the default hours, use SAM_POS_ITEMS_MATRIX on Position screen
    End If

    'Ticket #24543 - Macaulay Child Development Centre - Union moved to this screen
    'Ticket #22847 - Add the call to procedure UPDOvertime_Overview to add Overtime Master record for employees whose
    'Location, Region, Admin By or Section changes. And doing the same from Status/Dates screen when Employment Status,
    'Category or Union changes
    If Not glbtermopen Then
        If SavLoc <> clpCode(1) Or oRegion <> clpCode(2) Or oAdminBy <> clpCode(3) Or OSection <> clpCode(4) Or (glbCompSerial = "S/N - 2420W" And SavOrg <> clpCode(0).Text) Then
            Call UPDOvertime_Overview   'Ticket #22847
            Call Update_Overtime_Bank(clpCode(2).Text)
        End If
    End If
    
    'Ticket #25609 - Training Plan by Department
    'Update employee's Training Plan, if any, with courses matching employee's Department, if the Courses has Department Code assigned
    If Not glbtermopen Then
        'Not Friesens Corporation as their logic is very different and this enhancement not designed keeping in mind
        'their logic - some tweaking will be required but no testing of this logic has been done with their database.
        If glbCompSerial <> "S/N - 2279W" Then
            Call Track_Courses_Renewal_Update("Delete", "C", GetJHData(glbLEE_ID, "JH_JOB", ""))
            'Call Update_Employee_Job_Training_List(clpJob.Text, "Current")
        End If
    End If
End If

If glbWFC Then 'Ticket #16395 update employee names on Work Flow table
    If Len(OSNAME) > 0 And Len(OFNAME) > 0 Then
        If OSNAME <> txtSurname Or OFNAME <> txtFName Then
            Call WorkFlowNameChange(glbLEE_ID, txtSurname, txtFName)
        End If
        'Ticket #19286 Frank 11/23/2010 - begin
        If OSNAME <> txtSurname Then
            Call Upt_WFCPENSIOND_NAME(glbLEE_ID, medSIN, rsDATA("ED_EMPTYPE"), OSNAME, txtSurname, "Surname")
        End If
        If OFNAME <> txtFName Then
            Call Upt_WFCPENSIOND_NAME(glbLEE_ID, medSIN, rsDATA("ED_EMPTYPE"), OFNAME, txtFName, "Fname")
        End If
        'Ticket #19286 Frank 11/23/2010 - end
    End If
    'Ticket #19286 Frank 11/10/2010
    If Len(OSIN) > 0 Then 'change only, not include new hire
        If OSIN <> medSIN Then
            Call Upt_WFCPENSIOND_SIN(glbLEE_ID, rsDATA("ED_EMPTYPE"), OSIN, medSIN)
        End If
    End If
    
    'Ticket #19266
    Call AUDIT_NGS_TRANS

End If
                    
If glbCompSerial = "S/N - 2439W" Then   'OK Tire - Ticket #21518 Franks 05/03/2012
    Call AUDIT_GWL_TRANS
End If
If glbSamuel Then 'Ticket #20885 Franks 12/01/2011
    Call AUDIT_SAMUEL_TRANS
    
    '''Ticket #22261 Franks 07/27/2012 - begin
    ''Call SamuelSectionChg
    ''Call SamuelRegionChg
End If
If glbCompSerial = "S/N - 2420W" Then 'Ticket #24557 Franks 07/07/2015 Macaulay
    Call AUDIT_ADP_PAYATWORK_TRAN
End If

''Ticket #19473
'If glbCompSerial = "S/N - 2259W" Then  'For County of Oxford
'    Call UpdateGPPayCodeOxford
'End If

'Ticket #18790 - Update EEO record
Call EEO_Process

'Ticket #20406 Franks 07/08/2011
If glbCompSerial = "S/N - 2259W" Then 'County of Oxford
    Call UpdateGPMainBenDed
End If

If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/21/2014 Franks
    Call FamilyDayEmpSync(glbLEE_ID)
    Call FamilyDayEphotoSync(glbLEE_ID)
End If

'Ticket #24543 - Macaulay Child Development Centre
If glbCompSerial = "S/N - 2420W" And Not glbtermopen Then
    'SIN begins with 9 then Work Visa # and Expiration Date should be entered before the Save
    If Left(medSIN, 1) = "9" Then
        'Save the Work Visa info.
        If Len(glbWorkVisaNo) > 0 And Len(glbWorkExpDate) > 0 Then
            Call Update_WorkVisa_Info(glbLEE_ID)
        End If
    End If

    'Initialize
    glbWorkVisaNo = ""
    glbWorkExpDate = ""
End If
If glbLinamar And Not glbtermopen Then   'Ticket #28875 Franks 07/13/2016
    'SIN begins with 9 then Work Visa # and Expiration Date should be entered before the Save
    If Left(medSIN, 1) = "9" Then
        'Save the Work Visa info.
        If Len(glbWorkExpDate) > 0 Then
            Call Update_WorkVisa_Info(glbLEE_ID)
        End If
    End If

    'Initialize
    glbWorkVisaNo = ""
    glbWorkExpDate = ""
End If

If glbWFC Then
    Call UptWFCEMPOTHER   '#28637 Franks 05/18/2016
End If

If glbLinamar Then 'Ticket #29759 Franks 02/21/2017 -
    'After user change it then reset the next Payroll ID
    If Not (oPayrollID = txtPayrollID.Text) Then
        Call getNextLinPayrollID("Y")
    End If
End If

Screen.MousePointer = DEFAULT

glbLEE_SName = txtSurname
glbLEE_FName = txtFName

If glbLinamar Then
    glbLEE_ProdLine = clpCode(2).Text & " - " & GetTABLDesc("EDRG", clpCode(2).TransDiv & clpCode(2).Text) 'Ticket #14775
End If

Call EERetrieve

DoEvents

'note for programmers: the following function for Gander must be called before cmdModify_Click 'Ticket #24518 Franks 12/15/2014
If glbCompSerial = "S/N - 2453W" Then 'Town of Gander Ticket #24518 Franks 12/15/2014 - if Location was changed then call Salary_Integration
    If NewHireForms.count = 0 Then 'non new hire
        If Not (SavLoc = clpCode(1).Text) Then
            Call Salary_Integration(glbLEE_ID, , False, False, 0)
        End If
    End If
End If

Call cmdModify_Click

If glbtermopen Then
    Call Employee_Master_Integration(lblEEID, , , glbTERM_Seq)
Else
    If glbWFC And Len(xOldPayID) > 0 Then 'Ticket #12542
        'Call Employee_Master_Integration(glbLEE_ID, xOldPayID)
        If glbAdv Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
            Msg$ = "The Payroll ID was changed from " & SorocOPayrollID & " to " & txtPayrollID & " for this employee"
            Msg$ = Msg$ & Chr(10) & "Please change Employee Code in Advanced Tracker"
            MsgBox Msg$
        End If
    Else
        'If (glbCompSerial = "S/N - 2394W" Or glbCompSerial = "S/N - 2355W" Or glbCompSerial = "S/N - 2410W") And NewHireForms.count > 0 Then   ' St. John's Rehab Hospital - Ticket #14572
        If NewHireForms.count > 0 Then    'Ticket #28995 Franks 01/25/2017 - 'Do not transfer to GP yet since this new hire does not have enough (Emp Type)
            'Do not transfer to MediPay yet since this new hire does not have enough
            'information to create records in the staging tables.
            
            'Ticket #15718 for County of Lambton
            'The employee table in Employee Tracker is being populated with null entries
            'Do not populate the employee information in new hire, do it on Status/Dates screen
            
            'Ticket #18602 County of Frontenac
            'Do not transfer to GP yet since this new hire does not have enough (Emp Type)
        Else
            Call Employee_Master_Integration(glbLEE_ID)
            If glbMediPay Then Call Position_Integration(glbLEE_ID) 'Ticket #14752
        End If
    End If
End If

'Hemu - Not sure why was this added and who added this. It is causing an issue. The Status/Dates screen is getting skipped and
'there is already NextForm call above. So the NextForm call calls the Status/Dates screen and then right then
'the below NextForm call calls another screen after that. I am commenting this out for now.
'If glbWFC Then
'    Call NextForm
'End If
'Ticket #24767 Franks 12/11/2013 - WFC needs mini new hire after Rehire from HRsoft, the forms are Demo, Status/date, Position, Salary and Banking
If glbWFC Then
    If Len(glbCandidate) > 0 And glbHRSoftType = "ReHire" And glbHRSoftAction = "ReHireEmp" Then
        If Not glbtermopen Then
            Call NextForm
        End If
    End If
    flgWFCDivChaFlag = False
End If

'update Alt. Payroll IDs
If glbCompSerial = "S/N - 2420W" Then 'Macaulay Ticket #25016 Franks 04/01/2014
    Call UptAltPayrollIDs
End If

Exit Sub

Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREMP", "Update")
Call RollBack '21June99 js
Resume Next
End Sub
Private Sub EEO_Process()
    
    If Not (oCountryEmployment = txtCountryOfEmp.Text) Then
        If oCountryEmployment = "U.S.A." Then
            'from USA to other country should delete EEO - Ticket #18906
            Call uptEEO_Fields(glbLEE_ID, "Delete")
        End If
    End If
    If comCountryOfEmp.Text = "U.S.A." Then
        If Not glbtermopen Then
            If Not (oCountryEmployment = txtCountryOfEmp.Text) Then
                'Countr of Employment was changed from another country to "U.S.A.",
                'then create an  EEO record - Ticket #18906
                Call uptEEO_Fields(glbLEE_ID, "New")
            Else
                If IsEEOChg Then
                    Call uptEEO_Fields(glbLEE_ID, "Update")
                End If
            End If
        End If
    End If
End Sub
Private Function IsEEOChg()
Dim retval As Boolean
    retval = False
    If Not OSNAME = txtSurname Then retval = True
    If Not OFNAME = txtFName Then retval = True
    If Not OSEX = txtGender Then retval = True
    If Not OSIN = medSIN Then retval = True
    If Not oRegion = clpCode(2).Text Then retval = True
    If Not SavLoc = clpCode(1).Text Then retval = True
    If IsDate(ODOB) And IsDate(dlpDOB) Then
        If Not CVDate(ODOB) = CVDate(dlpDOB) Then
            retval = True
        End If
    End If
    IsEEOChg = retval
End Function

Private Sub WorkFlowNameChange(glbLEE_ID, xSurname, xFName)
Dim rsWF As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM HRWORKFLOW_EMPLOYEE WHERE PE_EMPNBR = " & glbLEE_ID
    rsWF.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsWF.EOF
        If Len(xSurname) > 0 Then
            rsWF("PE_SURNAME") = Trim(xSurname)
            rsWF.Update
        End If
        If Len(xFName) > 0 Then
            rsWF("PE_FNAME") = Trim(xFName)
            rsWF.Update
        End If
        rsWF.MoveNext
    Loop
    rsWF.Close
End Sub

Private Sub InsBlankEmpOther(xEmpNo)
Dim SQLQ As String
Dim rsHREmpOther As New ADODB.Recordset

    'Check if the record already exists for any reason; we do keep on getting Primary Key violati'''On Error from clients
    SQLQ = "SELECT ER_EMPNBR FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo
    rsHREmpOther.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsHREmpOther.EOF Then
        SQLQ = "INSERT INTO HREMP_OTHER (ER_COMPNO,ER_EMPNBR) "
        SQLQ = SQLQ & "VALUES ( '001'," & xEmpNo & ") "
        gdbAdoIhr001.Execute SQLQ
    End If
    rsHREmpOther.Close
    Set rsHREmpOther = Nothing
End Sub

Private Sub SecurityMod()
    If OSNAME <> txtSurname Or OFNAME <> txtFName Then
        Dim RsSecurity As New ADODB.Recordset
        Dim SQLQ
        SQLQ = "SELECT * FROM HR_SECURE_BASIC WHERE EMPNBR = " & lblEEID
        RsSecurity.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not RsSecurity.EOF Then
            RsSecurity("USERNAME") = Trim(txtSurname) & "," & Trim(txtFName)
            RsSecurity.Update
        End If
        RsSecurity.Close
    Else
        Exit Sub
    End If
End Sub
'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPhoto_Click() '21June99-js-from VB3
    'cmdPicture_Click
    Call SubPicture
End Sub


Public Sub cmdPrint_Click()
Dim RHeading As String, xReport, X%

'cmdPrint.Enabled = False
Me.vbxCrystal.Reset
Me.vbxCrystal.Destination = crptToPrinter
RHeading = lblEEName & "'s Basic Information"

Me.vbxCrystal.WindowTitle = lblEEName & "'s Basic Information Report"
Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
Call setRptLabel(Me, 1)
If Not glbtermopen Then
    Me.vbxCrystal.Connect = RptODBC_SQL
    xReport = glbIHRREPORTS & "rgbasic.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{HREMP.ED_EMPNBR}=" & glbLEE_ID & " "
Else
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 9
            If X% = 3 Or X% = 9 Then
                Me.vbxCrystal.DataFiles(X%) = glbIHRAUDIT
            Else
                Me.vbxCrystal.DataFiles(X%) = glbIHRDB
            End If
        Next
    End If
    xReport = glbIHRREPORTS & "rgbasic2.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{Term_HREMP.TERM_SEQ}=" & glbTERM_Seq & " "
    
End If


Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True


End Sub
Public Sub cmdView_Click()
Dim RHeading As String, xReport, X%

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

'cmdPrint.Enabled = False

RHeading = lblEEName & "'s Basic Information"
Me.vbxCrystal.Reset
Me.vbxCrystal.Destination = crptToWindow
Me.vbxCrystal.WindowTitle = lblEEName & "'s Basic Information Report"
Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
Call setRptLabel(Me, 1)
If Not glbtermopen Then
    Me.vbxCrystal.Connect = RptODBC_SQL

    xReport = glbIHRREPORTS & "rgbasic.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{HREMP.ED_EMPNBR}=" & glbLEE_ID & " "
Else
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 9
            If X% = 3 Or X% = 9 Then
                Me.vbxCrystal.DataFiles(X%) = glbIHRAUDIT
            Else
                Me.vbxCrystal.DataFiles(X%) = glbIHRDB
            End If
        Next
    End If
    xReport = glbIHRREPORTS & "rgbasic2.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{Term_HREMP.TERM_SEQ}=" & glbTERM_Seq & " "
    
End If


Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub

Private Sub clpDept_Change()
'    If Not cmdOK.Enabled Then RDept = clpDept    'added by Jaddy Sep 20,99

End Sub

Private Sub cmdUnlockSmoker_Click()
    ComSmoker.Enabled = True
End Sub

Private Sub comCountry_Click()
Call SetCountries
End Sub

Private Sub comCountry_GotFocus() 'RAUBREY 6/16/97
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comCountry_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub comCountry_LostFocus() 'RAUBREY 6/16/97
    Call SetCountries
    If Len(comCountry) > 0 Then
        If Len(comCountryOfEmp) = 0 Then
            comCountryOfEmp = comCountry
        End If
    End If
End Sub

Private Sub comCountryOfEmp_Change()
Call SetCountries
End Sub

Private Sub comCountryOfEmp_GotFocus()
 Call SetPanHelp(Me.ActiveControl)
End Sub

Sub ComMstat_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Sub ComMStat_Click()
    If ComMStat = "Partner" Or ComMStat = "Same-Sex" Then
        txtMStatus = UCase(Right(ComMStat.Text, 1))
    ElseIf ComMStat = "Separated" Then
        txtMStatus = UCase(Mid(ComMStat.Text, 4, 1))
    Else
        txtMStatus = Left(ComMStat.Text, 1)
    End If
End Sub

Private Sub comSmoker_Click()
If ComSmoker = "Yes" Then txtSmoker = "-1" Else txtSmoker = "0"
End Sub

Private Sub comSmoker_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub UpdMaxEmpNbr()

Dim SQLQ As String, countr As Integer, SQLQ2 As String
Dim Desc As String
Dim DMaxNum, DMaxNumX
Dim snapEmpX As New ADODB.Recordset
Dim snapEmp As New ADODB.Recordset
Dim rsPA As New ADODB.Recordset
'''On Error GoTo Emp_Err
    
    If glbCompSerial = "S/N - 2241W" And glbSysGen = True Then Exit Sub ' Granite Club
    
    DMaxNum = 0
    SQLQ = "Select MAX(ED_EMPNBR) AS MAXEMPNBR from HREMP"
    If glbCompSerial = "S/N - 2351W" Then ' Burlington Technologies
        SQLQ = SQLQ & " WHERE ED_EMPNBR< 90000000"
    End If
    snapEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not IsNull(snapEmp("MAXEMPNBR")) Then DMaxNum = snapEmp("MAXEMPNBR")
    
    If glbCompSerial <> "S/N - 2151W" Then
        DMaxNumX = 0
        SQLQ = "SELECT MAX(ED_EMPNBR) AS MAXEMPNBR FROM Term_HREMP"
        If glbCompSerial = "S/N - 2351W" Then ' Burlington Technologies
            SQLQ = SQLQ & " WHERE ED_EMPNBR< 90000000"
        End If
        snapEmpX.Open SQLQ, gdbAdoIhr001X, adOpenStatic
        If Not IsNull(snapEmpX("MAXEMPNBR")) Then DMaxNumX = snapEmpX("MAXEMPNBR")
        
        If DMaxNum < DMaxNumX Then DMaxNum = DMaxNumX Else DMaxNum = DMaxNum
    End If
    
    rsPA.Open "select PC_NEXT_AVAILABLE_NBR from HRPARCO", gdbAdoIhr001, adOpenStatic, adLockPessimistic
    
    If rsPA("PC_NEXT_AVAILABLE_NBR") <= DMaxNum Then
        rsPA("PC_NEXT_AVAILABLE_NBR") = DMaxNum + 1
        rsPA.Update
    ElseIf rsPA("PC_NEXT_AVAILABLE_NBR") > DMaxNum + 1 Then
        rsPA("PC_NEXT_AVAILABLE_NBR") = DMaxNum + 1
        rsPA.Update
    End If
    rsPA.Close

Exit Sub

Emp_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Employees", "HREMP", "SELECT")
Call RollBack '21June99 js

End Sub


Function CR_NEW_EE()
Dim SQLQ As String
Dim Msg As String, Msg1 As String, Title As String, Answer As String
Dim FEENum  As Long
Dim MsgC As String, DgDef As Integer, Response%

CR_NEW_EE = False
'''On Error GoTo CRNEWEE_ERR

Call UnloadFrms  'laura nov 21, 1997
If glbSysGen = True Then  'laura nov 28, 199
    'Comment by Frank on Jul 18, 2005 as Jerry's request. Ticket# 8805
    'MsgBox "The next available number is " & glbNextEmpl & " "
    Call ST_UPD_MODE(True)
    glbLEE_ID = CLng(glbNextEmpl)
    
    Dim rsPA As New ADODB.Recordset
    rsPA.Open "select PC_NEXT_AVAILABLE_NBR,PC_FEDTAX,PC_PROVTAX from HRPARCO", gdbAdoIhr001, adOpenStatic, adLockPessimistic
    glbNextEmpl = glbNextEmpl + 1
    If glbCompSerial = "S/N - 2241W" Or glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2377W" Then 'Granite Club or Timmins or Town of Grimsby
        Call Check_EMPLOYEE_Number(glbNextEmpl)
    Else
        rsPA("PC_NEXT_AVAILABLE_NBR") = glbNextEmpl
        rsPA.Update
    End If
    rsPA.Close
    
    
    'Added by Franks May 12,2003 begin - Ticket# 4148
    SQLQ = "Select " & FldList & " from HREMP"
    SQLQ = SQLQ & " where ED_EMPNBR = " & glbLEE_ID & ""
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
    'Added by Franks May 12,2003 end
    
    'added by Bryan Aug 24, 2003 Ticket# 9191
        Data1.RecordSource = SQLQ
        Data1.Refresh
    'end bryan
    
    CR_NEW_EE = True
Else
    Msg1 = "Enter a new, unique "
    Msg1 = Msg1 & Chr(10) & "Employee ID." & Chr(10)
    If Not glbWFC Then
    Msg1 = Msg1 & "  Value between "
    Msg1 = Msg1 & Chr(10) & "  1 and 999999999" & Chr(10)
    End If
    Msg = Msg1 & Chr(10) & Chr(10)
    Title = "Create New Employee" ' Set title.
    
    Do
        'Comment by Frank on Mar 15, 2005 for ticket# 8046
        'If glbLinamar Then
            glbUserUploadMode = UploadFormWithoutCheck
            DoEvents
            If glbWFC And glbCandidate > 0 And (glbHRSoftType = "ReHire" Or glbHRSoftType = "NewHire") Then 'Ticket #24184 Franks 09/24/2013
            'If glbWFC And glbCandidate > 0 And (glbHRSoftType = "ReHire") Then  'Ticket #24184 Franks 09/24/2013
                Answer = glbTrsEE_ID
            Else
                Answer = GetNewEmpnbr()
            End If
        'Else
        '    Answer = InputBox$(Msg, Title)   ' Get user input.
        'End If
        If Len(Answer) > 0 Then
            If Not IsNumeric(Answer) Then
                Msg = Msg1 & "Sorry, must be numeric."
                GoTo NEW_NG
            End If
            If Len(Answer) > 9 Then
                Msg = Msg1 & "Number must be between 1 and 999999999"
                GoTo NEW_NG
            End If
            FEENum = CLng(Answer)
            If FEENum < 1 Or FEENum > 999999999 Then
                Msg = Msg1 & "Number must be between 1 and 999999999"
                GoTo NEW_NG
            End If
            If glbWFC Then
                If Len(Answer) <> 8 Then
                    Msg = Msg1 & "Employee Number is not valid format" & Chr(10)
                    Msg = Msg & "It should be Division Number + 4 digit Employee Number"
                    GoTo NEW_NG
                End If
            End If
            
            Dim rsEmp As New ADODB.Recordset
            SQLQ = "Select " & FldList & " from HREMP"
            SQLQ = SQLQ & " where ED_EMPNBR = " & FEENum & ""
            
            Data1.RecordSource = SQLQ
            Data1.Refresh
            
            rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            
            If rsEmp.BOF And rsEmp.EOF Then
                Call SET_UP_MODE
                If glbLinamar Or (glbCompSerial = "S/N - 2390W") Then '2390W - Collectcorp Ticket #16312
                    If EmpNoInTerm(FEENum) Then
                        Msg = "Sorry, Employee # " & ShowEmpnbr(Answer)
                        Msg = Msg & Chr(10) & "Already exists in terminated employee list."
                        rsEmp.Close
                        GoTo NEW_NG
                    End If
                End If
                If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
                rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
                
                Call ST_UPD_MODE(True)
                
                glbLEE_ID = CLng(Answer)
                
                CR_NEW_EE = True
                rsEmp.Close
                
                Exit Do
            Else
                Msg = "Sorry, Employee # " & ShowEmpnbr(Answer)
                Msg = Msg & Chr(10) & rsEmp("ED_SURNAME")
                Msg = Msg & Chr(10) & "Already exists."
                rsEmp.Close
                GoTo NEW_NG
            End If
        Else
            MsgBox "Add of new Employee Aborted"
            fglbNewEE = False
            glbLEE_ID = 0
            
            Dim X
            For X = 1 To NewHireForms.count
                NewHireForms.Remove 1
            Next
                
            glbNo = False
            
            MDIMain.panHelp(0).Caption = "Select function from the menu."   'Jaddy 10/22/99
            
            UnloadForm = True
            
            Exit Do
        End If
NEW_NG:
        MsgBox Msg, , Title
    Loop
    
    MDIMain.MainToolBar.ButtonS(8).Visible = True
    MDIMain.MainToolBar.ButtonS(9).Visible = True
    MDIMain.MainToolBar.ButtonS(8).Enabled = True
    MDIMain.MainToolBar.ButtonS(9).Enabled = True
    MDIMain.MainToolBar.ButtonS(1).Enabled = True
End If

If CR_NEW_EE Then
    fglbNewEE = True
    
    Call SET_UP_MODE
    
    Call Set_Control("B", Me)
    
    txtENTOPT(0) = glbEntOutStanding$
    txtENTOPT(1) = glbEntOutStandingS$
    
    'Vacation & Sick date range from Entitlement Master since v7.6
    'If txtENTOPT(0) = "1" Then
    '    txtFDATE(0) = glbCompEdFrom
    '    txtTDATE(0) = glbCompEdTo
    'End If
    'If txtENTOPT(1) = "1" Then
    '    txtFDATE(1) = glbCompEdFromS
    '    txtTDATE(1) = glbCompEdToS
    'End If
    
    lblEEID = glbLEE_ID
    lblEENum = ShowEmpnbr(lblEEID)
    
    'Ticket #23795 - Town of Lasalle
    If glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2379W" Then
        txtPayrollID.Text = Right("00000" & glbLEE_ID, 5)
    ElseIf glbCompSerial = "S/N - 2373W" Then   'Ticket #19113 - District Municipality of Muskoka
        txtPayrollID.Text = glbLEE_ID
    End If
    
    lblEEName.Caption = "New Employee"
    
    If glbCompSerial = "S/N - 2391W" Then   'North York Community House - Ticket #15933
        txtGender = "F"
        txtMStatus = "F"
    Else
        txtGender = "M"
        txtMStatus = "M"
    End If
    
    If IsNull(glbCountry) Then
        comCountry = "CANADA"
        '8.0 - Ticket #22682
        If Len(comCountryOfEmp) = 0 Then
            comCountryOfEmp = comCountry
        End If
    Else
        comCountry = glbCountry
        '8.0 - Ticket #22682
        If Len(comCountryOfEmp) = 0 Then
            comCountryOfEmp = comCountry
        End If
    End If
    
    If glbLinamar Then
        'Ticket #20628 - Province is coming from HR_Division so no defaulting to ON
        'clpProv.Text = "ON"
        clpDiv.Text = Right(glbLEE_ID, 3)
        clpProv.Text = Get_Division_Name(clpDiv.Text, "DV_PROV")
    End If
    
    txtCompany = "001"
    clpGLNum.TextBoxWidth = 1500
End If

MDIMain.MainToolBar.ButtonS(8).Visible = True
MDIMain.MainToolBar.ButtonS(9).Visible = True
MDIMain.MainToolBar.ButtonS(8).Enabled = True
MDIMain.MainToolBar.ButtonS(9).Enabled = True
MDIMain.MainToolBar.ButtonS(1).Enabled = True

clpGLNum.TextBoxWidth = 1500

If glbWFC Then ComSmoker.Enabled = True 'Ticket #21119 Franks 11/14/2011

Exit Function

CRNEWEE_ERR:
If Err.Number = 364 Then
    UnloadForm = True
    Exit Function
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_NEW_EE", "HREMP", "ADD RECORD")
Call RollBack '21June99 js
Resume Next
End Function

Public Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

'''On Error GoTo EERError
If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
If glbtermopen Then
    SQLQ = "Select " & FldList & " from Term_HREMP"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "Select " & FldList & " from HREMP"
    SQLQ = SQLQ & " where ED_EMPNBR = " & glbLEE_ID
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

Data1.RecordSource = SQLQ
Data1.Refresh

If rsDATA.EOF Or rsDATA.BOF Then Exit Function

If Not rsDATA.EOF Then Call getCodes

EERetrieve = True

'8.0 - Ticket #22682 - Photo from a folder now
'If Len(glbPicDir) < 1 Then
If Not gsEMPLOYEEPHOTO Then
    If glbSQL Or glbOracle Then
        If cmdPhoto.Caption = "&Photo Off" Then
            picPhoto.Visible = False
            PicNotF.Visible = True
    
            If glbtermopen Then
                Call FillPhoto(Val(glbTERM_ID))
            Else
                Call FillPhoto(Val(glbLEE_ID))
            End If
        Else
            picPhoto.Visible = False
            PicNotF.Visible = False
        End If
        If glbWFC Then 'Ticket #21119 Franks 11/14/2011
            If IsNull(Data1.Recordset("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = Data1.Recordset("ED_VADIM1")
        End If
    End If
Else
    If Len(glbPicDir) < 1 Then
        picPhoto.Visible = False
    Else
        If cmdPhoto.Caption = "&Photo Off" Then
            picPhoto.Visible = False
            PicNotF.Visible = True
            If glbtermopen Then
                Call LoadPhoto(Val(glbTERM_ID))
            Else
                Call LoadPhoto(Val(glbLEE_ID))
            End If
        Else
            picPhoto.Visible = False
            PicNotF.Visible = False
        End If
    End If
End If

If glbCompSerial = "S/N - 2357W" And comCountry = "CANADA" Then   'I.T. Xchange
    lbltitle(24).FontBold = True
ElseIf glbCompSerial = "S/N - 2357W" And comCountry <> "CANADA" Then
    lbltitle(24).FontBold = False
End If

If glbWFC Then 'Ticket #27336 Franks 07/22/2015
    lbltitle(8).FontBold = True
    If Not IsNull(Data1.Recordset("ED_WORKCOUNTRY")) Then
        If UCase(Data1.Recordset("ED_WORKCOUNTRY")) = "INDIA" Then
            lbltitle(8).FontBold = False
        End If
    End If
End If

EERetrieve = True
Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HREMP", "SELECT")
Call RollBack '21June99 js

Resume Next
End Function

Private Sub dlpDate_LostFocus(Index As Integer)
'Ticket #23837 Franks 05/28/2013
If NewHireForms.count > 0 And Index = 0 Then  'New Hire only & DOH
    If IsDate(dlpDate(0).Text) Then
        dlpDeptEDate.Text = dlpDate(0).Text
        dlpDivEDate.Text = dlpDate(0).Text
    End If
End If
End Sub

Private Sub Form_Activate()

glbOnTop = "FRMEEBASIC"

If glbCompSerial = "S/N - 2227W" Then
    clpCode(2).MaxLength = 6
End If

If mbAddNewEmployee Then
    cmdNew_Click
    If UnloadForm Then Exit Sub
    mbAddNewEmployee = False
Else

'    Me.cmdModify_Click
    Call SET_UP_MODE
End If

End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEEBASIC"
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim Answer, DefVal, Msg$, Title  ' Declare variables.
Dim RFound As Integer, VReturn%, X%, xPIC

glbOnTop = "FRMEEBASIC"
oldEEId = glbLEE_ID
flagFrmLoad = True
locUploadWithoutCheck = False 'Ticket #19937 for Samuel -  Franks 05/06/2011

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If
Data1.RecordSource = "SELECT " & FldList & " FROM HREMP"

'cmdPrint.Visible = glbSQL

'8.0 - Jerry is allowing clients to retain photos in the database.
'8.0 - Ticket #22682 - Move Photo out of database into a folder ---------------------------------------------------
xPIC = glbIHRREPORTS & "IHRPICS.MTR"
'Get Photo Path
If gsEMPLOYEEPHOTO Then
    xPIC = GetComPreferEmail("EMPLOYEEPHOTOPATH")
    If Len(xPIC) > 0 And Right(xPIC, 1) <> "\" Then xPIC = xPIC & "\"
End If

'Ticket #20367 - Jerry said to show photos for Terminated employees. Have made code change to not delete
'photo on termination and on rehire the photo is brought back.
If (Dir(xPIC) = "" And Not glbOracle And Not glbSQL) Then   'Or glbtermopen Then
'If xPIC = "" Then   'Ticket #22682
    PicNotF.Visible = False
    cmdPhoto.Enabled = False 'Jaddy 10/28/99
    cmdPhoto.Caption = "&Photo"
    picPhoto.Visible = False
    glbPicDir = ""
Else
    PicNotF.Visible = True
    cmdPhoto.Enabled = True 'Jaddy 10/28/99
    picPhoto.Visible = False
    'glbPicDir = glbIHRREPORTS  'Ticket #22682
    glbPicDir = xPIC
End If
'8.0  Ticket #22682 - Move Photo out of database into a folder ---------------------------------------------------

'Ticket #19375 - Mostafa
If glbCompSerial = "S/N - 2241W" Then 'Granite Club
    clpCode(4).MaxLength = 10
    'clpCode(4).TextBoxWidth = 1500
    clpCode(4).Width = 2700
    'clpCode(4).ShowDescription = False
    'clpCode(4).ShowDescription = True
End If
If Not glbCompSerial = "S/N - 2420W" Then 'Macaulay Ticket #25016 Franks 04/01/2014
    Call tbDemographics.Tabs.Remove(4)
Else
    Call MacaulayAltPayIDScreen 'Ticket #24557 Franks 09/03/2014
End If

If glbCompSerial = "S/N - 2482W" Then 'Windsor Family Credit Union Ticket #28515 Franks 04/26/2016
    lbltitle(24).Visible = False
    clpCode(2).Visible = False
End If

Screen.MousePointer = HOURGLASS
Call setCaption(lbltitle(2))
Call setCaption(lbltitle(11))
Call setCaption(lbltitle(12))
Call setCaption(lbltitle(13))

Call setCaption(lbltitle(15))

Call setCaption(lbltitle(23))
Call setCaption(lbltitle(24))
Call setCaption(lbltitle(25))
Call setCaption(lbltitle(26))
Call setCaption(lblDivStart)
Call setCaption(lblDeptStart)

Call setCaption(lbltitle(33))
Call setCaption(lbltitle(34))
Call setCaption(lbltitle(35))
Call setCaption(lbltitle(36))
Call setCaption(lbltitle(37))
Call setCaption(lbltitle(38))
Call setCaption(lbltitle(39))
Call setCaption(lbltitle(40))

'Ticket #23537 and Release 8.0 - Demographics labels
lbltitle(10).Caption = lStr("Telephone #2")
lbltitle(20).Caption = lStr("Cellular Telephone")
lbltitle(21).Caption = lStr("Pager Number")

'Release 8.0 - Ticket #2268: Add Payroll ID to Label Master
lbltitle(32).Caption = lStr("Payroll ID")

'Ticket #29375 - Add to Label Master
lblOtherEmail.Caption = lStr("Other Email Address")

'Ticket #24164 - Re-ordering and new Organizaton fields
'Samuel only
If (glbCompSerial = "S/N - 2382W") Then
    lbltitle(18).Caption = lStr("Organization 1")
    lbltitle(45).Caption = lStr("Organization 2")
    lblOrg1EffDate.Caption = lStr("Organization 1 Effective")
    lblOrg2EffDate.Caption = lStr("Organization 2 Effective")
    
    'Reorder the Organization fields
    lbltitle(18).Top = lbltitle(27).Top
    clpCode(6).Top = clpHOME(1).Top
    lbltitle(45).Top = lbltitle(29).Top
    clpCode(7).Top = clpHOME(2).Top
    lblOrg1EffDate.Top = lbltitle(18).Top
    dlpOrg1EDate.Top = clpCode(6).Top
    lblOrg2EffDate.Top = lbltitle(45).Top
    dlpOrg2EDate.Top = clpCode(7).Top
    
    lbltitle(18).Visible = True
    clpCode(6).Visible = True
    lbltitle(45).Visible = True
    clpCode(7).Visible = True
    lblOrg1EffDate.Visible = True
    dlpOrg1EDate.Visible = True
    lblOrg2EffDate.Visible = True
    dlpOrg2EDate.Visible = True
Else
    'Release 8.0 - Ticket #22682: Jerry said to open up these fields for everyone.
    lbltitle(18).Caption = lStr("Organization 1")
    lbltitle(45).Caption = lStr("Organization 2")
    lblOrg1EffDate.Caption = lStr("Organization 1 Effective")
    lblOrg2EffDate.Caption = lStr("Organization 2 Effective")
        
    'Reorder the Organization fields
    lbltitle(18).Top = lbltitle(27).Top
    clpCode(6).Top = clpHOME(1).Top
    lbltitle(45).Top = lbltitle(29).Top
    clpCode(7).Top = clpHOME(2).Top
    lblOrg1EffDate.Top = lbltitle(18).Top
    dlpOrg1EDate.Top = clpCode(6).Top
    lblOrg2EffDate.Top = lbltitle(45).Top
    dlpOrg2EDate.Top = clpCode(7).Top
        
    lbltitle(18).Visible = True
    clpCode(6).Visible = True
    lbltitle(45).Visible = True
    clpCode(7).Visible = True
    lblOrg1EffDate.Visible = True
    dlpOrg1EDate.Visible = True
    lblOrg2EffDate.Visible = True
    dlpOrg2EDate.Visible = True
End If

'Ticket #24543 - Macaulay Child Development Centre
If glbCompSerial = "S/N - 2420W" Then
    lblUnion.Visible = True
    lblSalDist.Visible = True
    lblSalDist.FontBold = True 'Ticket #24557 Franks 12/11/2014
    clpCode(0).DataField = "ED_ORG"
    clpCode(0).Visible = True
    clpSalDist.Visible = True
    Call setCaption(lblUnion)
    Call setCaption(lblSalDist)

    'Reorder the Organization fields
    lbltitle(18).Top = lbltitle(30).Top
    clpCode(6).Top = clpHOME(3).Top
    lbltitle(45).Top = lbltitle(28).Top
    clpCode(7).Top = clpHOME(4).Top
    lblOrg1EffDate.Top = lbltitle(18).Top
    dlpOrg1EDate.Top = clpCode(6).Top
    lblOrg2EffDate.Top = lbltitle(45).Top
    dlpOrg2EDate.Top = clpCode(7).Top

End If

If glbCompSerial = "S/N - 2332W" Then   'Town of Fort Frances
    lbltitle(25).FontBold = True
    lbltitle(7).FontBold = False
End If
If glbCompSerial = "S/N - 2458W" Then   'Ticket #25469 - City of Campbell River
    lbltitle(7).FontBold = False
End If
If glbCompSerial = "S/N - 2383W" Then 'Town of Orangeville Ticket #18844 Franks 01/13/2011
    lbltitle(25).FontBold = True
    lbltitle(13).FontBold = True 'Ticket #23189 Franks 02/07/2013
End If
If glbCompSerial = "S/N - 2485W" Then 'Mississaugas of Scugog Island First Nation -Ticket #28652  Franks 07/31/2017
    lbltitle(13).FontBold = True
End If
If glbCompSerial = "S/N - 2439W" Or glbCompSerial = "S/N - 2484W" Then 'OK Tire Ticket #22503 Franks 09/14/2012 'Ticket #28396 Franks 03/08/2017 PeterboroughFHT
    lbltitle(25).FontBold = True
End If
If glbCompSerial = "S/N - 2483W" Then 'Scott Steel Ticket #28262 Franks 06/07/2016
    lbltitle(32).FontBold = True 'Ticket #29077 Franks 09/23/2016
    lbltitle(25).FontBold = True
End If
If (glbCompSerial = "S/N - 2347W") Or glbWFC Or (glbCompSerial = "S/N - 2382W") Or (glbCompSerial = "S/N - 2182W") Or (glbCompSerial = "S/N - 2453W") Then    '2382W - Samuel '2453W Gander
    '2182 = Town of Caledon
    lbltitle(26).FontBold = True 'Section
End If
If (glbCompSerial = "S/N - 2493W") Then lbltitle(26).FontBold = True 'Section
If glbCompSerial = "S/N - 2453W" Then 'Gander Ticket #24518 Franks 12/05/2014
    lbltitle(12).FontBold = True    'G/L #
    lbltitle(23).FontBold = True    'Location
End If
If glbCompSerial = "S/N - 2487W" Then 'City of Kenora Ticket #30217 Franks 06/12/2017
    lbltitle(12).FontBold = True    'G/L #
End If
If (glbCompSerial = "S/N - 2241W") Then ' for Granite Club
    lbltitle(25).FontBold = True    'Payroll # (Admin By)
    lbltitle(13).FontBold = True
    Call modECountChk   'For New Hires to pick the Next Employee #. and the value for glbSysGen
    lblDeptStart.FontBold = True    'Dept Effective Date
    lblDivStart.FontBold = True     'Division Effective Date
End If
If (glbCompSerial = "S/N - 2388W") Then ' DNSSAB Ticket #14475
    lbltitle(12).FontBold = True    'G/L #
    lbltitle(13).FontBold = True    'Division
    lbltitle(23).FontBold = True    'Location
    lblDeptStart.FontBold = True    'Dept Effective Date
    lblDivStart.FontBold = True     'Division Effective Date
End If
If glbCompSerial = "S/N - 2214W" Then 'for casey house Ticket #14550
    lbltitle(24).FontBold = True    'Region
    lbltitle(12).FontBold = True    'G/L #
    lbltitle(13).FontBold = True    'Division
    lbltitle(23).FontBold = True    'Location
    lbltitle(25).FontBold = True    'Admin By
    lbltitle(26).FontBold = True    'Section
    lblDeptStart.FontBold = True    'Dept Effective Date
    lblDivStart.FontBold = True     'Division Effective Date
End If
If glbCompSerial = "S/N - 2173W" Then 'for town of Ajex
    lbltitle(24).FontBold = True
End If
If glbCompSerial = "S/N - 2409W" Then 'Delisle Youth Services - Ticket #27798 04/25/2016
    lbltitle(25).FontBold = True    'Admin By
    lbltitle(32).FontBold = True    'Payroll ID Ticket #27798 09/09/2016
    lbltitle(26).FontBold = True    'Section    Ticket #27798 09/09/2016
    lbltitle(24).FontBold = True    'Region     Ticket #27798 09/09/2016
End If
If glbCompSerial = "S/N - 2454W" Then 'Showa Canada 'Ticket #24659
    lbltitle(43).FontBold = True
    lbltitle(24).FontBold = True
End If
If glbCompSerial = "S/N - 2178W" Or glbCompSerial = "S/N - 2443W" Then 'Community Living Dufferin or Walters Inc Ticket #23278
    lbltitle(13).FontBold = True
End If
If glbCompSerial = "S/N - 2257W" Then 'Hamilton CAS
    lbltitle(13).FontBold = True
    lbltitle(25).FontBold = True
    lbltitle(24).FontBold = True 'Ticket #25786 Franks 07/25/2014
    'lblDeptStart.FontBold = True
    'lblDivStart.FontBold = True
End If
If glbCompSerial = "S/N - 2259W" Then  'For County of Oxford
    lbltitle(13).FontBold = True
    lbltitle(23).FontBold = True
    'lblTitle(25).FontBold = True
    lbltitle(24).FontBold = True
    'lblTitle(12).FontBold = True   'Ticket #15590
End If

If glbGP Then 'George Mar 8,2006 Great Plains 9965
    If glbCompSerial = "S/N - 2259W" Then
        lbltitle(12).FontBold = True
    End If
End If

'Hemu - Begin - County of Essex - Modifications  - Ticket # 6549
If glbCompSerial = "S/N - 2192W" Then
    lbltitle(24).Visible = False
    clpCode(2).Visible = False
End If
'Hemu - End

If glbCompSerial = "S/N - 2366W" Then  ' FOR Family Youth Child Services of Muskoka
    lbltitle(13).FontBold = True
    lbltitle(23).FontBold = True
    lbltitle(25).FontBold = True
End If

If glbCompSerial = "S/N - 2415W" Then 'Ticket #16982 SPC- Volunteer System
    lbltitle(13).FontBold = True
    lbltitle(23).FontBold = True
    lbltitle(25).FontBold = True
End If

If glbWFC Then
    Call WFCScreenSetup
End If
If glbCompSerial = "S/N - 2382W" Then 'Ticket #20600 Franks 09/02/2011
    'lblRptNo.Caption = "Actual Branch" '"Sub Dept"
    lblRptNo.Caption = "Physical Branch" 'Ticket #24285 Franks 09/1/2013
    lblRptNo.Visible = True
    'imgIcon.Visible = True
    'txtDeptBonusCtr.Visible = True
    'lblDeptBonusDesc.Visible = True
    clpCode(5).DataField = "ED_SUBDEPT"
    clpCode(5).Left = clpCode(2).Left
    clpCode(5).Top = txtDeptBonusCtr.Top
    clpCode(5).Visible = True
End If

'Hemu - Begin - C.C.A.C. London & Middlesex - Ticket #6718
If glbCompSerial = "S/N - 2242W" Then
    lbltitle(12).FontBold = True
    lbltitle(13).FontBold = True
    lbltitle(25).FontBold = True 'Ticket 8189
    lblDivStart.FontBold = True
End If
'Hemu - End

'Hemu - Begin - City of Timmins - Ticket #9557
If glbCompSerial = "S/N - 2375W" Then
    lbltitle(13).FontBold = True
    lbltitle(23).FontBold = True
    'lblTitle(26).FontBold = True   - Ticket 9877
End If
'Hemu - End

If glbCompSerial = "S/N - 2191W" Or glbCompSerial = "S/N - 2396W" Then
'City of St-Thomas, Oshawa CHC
    lbltitle(32).FontBold = True
End If

'Ticket #24443 - North York Community House
If glbCompSerial = "S/N - 2391W" Then
    lbltitle(32).FontBold = True
    lbltitle(24).FontBold = True
End If

If glbCompSerial = "S/N - 2410W" Then 'Frontenac Ticket #18603
    lbltitle(32).FontBold = True
    'lblTitle(25).FontBold = True
    'Ticket #23857 Franks 05/30/2013
    lbltitle(26).FontBold = True
End If
If glbCompSerial = "S/N - 2451W" Then 'Decor Ticket #23848
    lbltitle(32).FontBold = True 'payroll id
    'lblTitle(25).FontBold = True 'Administered By
End If

If glbCompSerial = "S/N - 2378W" Then 'Aurora
    lbltitle(12).FontBold = True
End If

If glbCompSerial = "S/N - 2382W" Then  'Namasco -> Samuel
    'Ticket #16235 - Begin
    'lblTitle(43).FontBold = True
    lbltitle(13).FontBold = True
    lbltitle(24).FontBold = True
    lblDeptStart.FontBold = True
    'Ticket #16235 - End
    lbltitle(25).FontBold = True
    lbltitle(23).FontBold = True
    'Ticket #18702 - begin 06/22/2010 Frank
    lbltitle(15).FontBold = True
    dlpDate(0).Left = lblDOH.Left + 400
    lblDOH.Visible = False
    dlpDate(0).Visible = True
    dlpDate(0).DataField = "ED_DOH"
    'Ticket #18702 - end
    'Ticket #20319 Franks 05/19/2011
    lblDivStart.FontBold = True
End If

If glbCompSerial = "S/N - 2335W" Then 'Mitchell Plastics Ticket #21866 Franks 04/05/2012
    lbltitle(32).FontBold = True
    lbltitle(43).FontBold = True
    lbltitle(13).FontBold = True
End If

If glbCompSerial = "S/N - 2443W" Then 'Walters Inc Ticket #22853 Franks 11/26/2012
    lbltitle(32).FontBold = True
    lbltitle(12).FontBold = True 'Ticket #25952 Franks 11/04/2014
    lbltitle(24).FontBold = True 'Ticket #25952 Franks 11/04/2014
End If

If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/20/2014
    cmdCopyEmpByPayID.Left = 7500
    cmdCopyEmpByPayID.Visible = True
    lbltitle(32).FontBold = True
End If
If glbCompSerial = "S/N - 2174W" Then 'KH CAS Ticket #23382 Franks 07/11/2014
    lbltitle(32).FontBold = True
End If
If glbCompSerial = "S/N - 2466W" Then 'Chiefs of Ontario Ticket #25879 Franks 09/25/2014
    lbltitle(32).FontBold = True
End If
If glbCompSerial = "S/N - 2469W" Then 'Rideau FHT - Ticket #26423 Franks 03/06/2015
    lbltitle(32).FontBold = True
End If

Call TabOrderSetup 'Ticket #18090

If (glbCompSerial = "S/N - 2380W") Then  'VitalAire
    clpCode(1).MaxLength = 5
    'Simona - VitalAire - ticket #14995 - not show ED_SECTION
    lbltitle(26).Visible = False
    clpCode(4).Visible = False
End If

If (glbCompSerial = "S/N - 2411W") Then  'Wellington-Dufferin-Guelph Public Health - Ticket #17129
    clpCode(1).MaxLength = 6
End If

If (glbCompSerial = "S/N - 2394W") Then ' St. John's Rehab Hospital - Ticket #14572
    lbltitle(12).FontBold = True    'G/L #
    lbltitle(13).FontBold = True    'Division
    lbltitle(23).FontBold = True    'Location
    lbltitle(24).FontBold = True    'Region
    lbltitle(25).FontBold = True    'Administered By
    lblDeptStart.FontBold = True    'Dept Effective Date
    lblDivStart.FontBold = True     'Division Effective Date
    txtAdd2.Enabled = False         'Ticket #15408
End If

If glbCompSerial = "S/N - 2393W" Then  ' KTH Shelburne Mfg. Inc. - Ticket #14613
    lbltitle(13).FontBold = True
End If

If glbCompSerial = "S/N - 2408W" Or glbCompSerial = "S/N - 2172W" Then 'Township of Wilmot - Ticket #15785; Lanark
    lbltitle(13).FontBold = True
End If

If glbCompSerial = "S/N - 2415W" Then 'Ticket #16982 SPC- Volunteer System
    lbltitle(8).FontBold = False
End If

If glbCompSerial = "S/N - 2233W" Then   'Leeds-Grenville F&CS - Ticket #16737
    lbltitle(24).FontBold = True    'Region
    lbltitle(12).FontBold = True    'G/L #
End If
If glbCompSerial = "S/N - 2393W" Then   'KTH Shelburne Ticket #17289
    lbltitle(35).FontBold = True    'License Plate #2
    lbltitle(40).FontBold = True    'Parking Permit #2
End If

'Ticket #25469 - City of Campbell River
If glbCompSerial = "S/N - 2458W" Then
    'lblTitle(12).FontBold = True    'G/L #     'They do not want this to be mandatory any more
    lbltitle(24).FontBold = True    'Region
    lbltitle(26).FontBold = True    'Section
End If

'Ticket #26672 - County of Perth
If glbCompSerial = "S/N - 2417W" Then
    lbltitle(38).FontBold = True    'Type of Vehicle
End If

'Ticket #19067
'Surrey Place Centre - Ticket #19067
'If glbCompSerial = "S/N - 2347W" Then
    lbltitle(15).FontBold = True
    dlpDate(0).Left = 7220 'lblDOH.Left + 400
    lblDOH.Visible = False
    dlpDate(0).Visible = True
    dlpDate(0).DataField = "ED_DOH"
'End If

'If glbAdv Then
'    lblTitle(43).FontBold = True
'End If
'removed at jerry request
Call addItems
Call UpdMaxEmpNbr
Call INI_Controls(Me)

'Frank 09/03/03 Ticket #4684, move Location to Stats/Date screen
'2347W - For Surrey Place
'2410W - Ticket #18603 - Frontenac
If glbCompSerial = "S/N - 2347W" Or glbCompSerial = "S/N - 2410W" Then
    lbltitle(23).Visible = False
    lbltitle(25).Top = 7470 '7140
    clpCode(3).Top = 7440 '7110
    If glbCompSerial = "S/N - 2410W" Then 'Ticket #24498 Franks 11/06/2013
    Else
        lbltitle(12).FontBold = True
    End If
End If
If glbCompSerial = "S/N - 2345W" Then  'For North Bay & District Health Unit
    lbltitle(21).Caption = "Home Cell"
End If
If glbCompSerial = "S/N - 2377W" Then  'Town of Grimsby
    lbltitle(10).Caption = "Town Cell Phone"
End If
If glbCompSerial = "S/N - 2350W" Then  'For Listowel
    lbltitle(13).FontBold = True
    lbltitle(12).FontBold = True
    lbltitle(24).FontBold = True
    lbltitle(26).FontBold = True
    lblDeptStart.FontBold = True
    lblDivStart.FontBold = True
End If
If glbCompSerial = "S/N - 2363W" Then ' CITY OF K LAKES
    lbltitle(23).FontBold = True
    lbltitle(24).FontBold = True
End If
If glbCompSerial = "S/N - 2401W" Then  'Assessment Strategies Inc.
    lbltitle(21).Caption = "Cottage Number"
    medPageNbr.Tag = "Cottage Number"
End If

If glbVadim Then
    Call VadimControl("Show")
    lbltitle(32).FontBold = True
End If
If glbLambton Then 'Ticket# 6355
    lbltitle(24).FontBold = True
    lbltitle(13).FontBold = True
    lbltitle(23).FontBold = True
End If
If glbCompSerial = "S/N - 2351W" Then ' Burlington Technologies
    lbltitle(23).FontBold = True 'Ticket# 7768
End If
Screen.MousePointer = DEFAULT

If glbCompSerial = "S/N - 2357W" And comCountry = "CANADA" Then   'I.T. Xchange
    lbltitle(24).FontBold = True
End If

If glbCompSerial = "S/N - 2369W" Then 'TS Tech Ticket#9627
    lbltitle(12).FontBold = True
    lbltitle(13).FontBold = True
End If

If (glbCompSerial = "S/N - 2385W") Then ' Conservation Halton 'Ticket #13063
    lbltitle(13).FontBold = True    'Division
    lbltitle(26).FontBold = True    'Section
End If
If glbCompSerial = "S/N - 2335W" Then   'Mitchell Plastics Ltd. - Ticket #20982
    lbltitle(26).FontBold = True    'Section
End If
If glbCompSerial = "S/N - 2386W" Then  'The Walter Fedy Partnership 'Ticket #13828
    lbltitle(13).FontBold = True
    lbltitle(23).FontBold = True
    lbltitle(24).FontBold = True
    lbltitle(25).FontBold = True
    lbltitle(26).FontBold = True
End If

'Hemu - CollectCorp Inc. - Ticket #14247
If glbCompSerial = "S/N - 2390W" Then
    lbltitle(23).FontBold = True
    lbltitle(24).FontBold = True
    lbltitle(26).FontBold = True
End If

If (glbCompSerial = "S/N - 2418W") Then ' charton hobbs - Ticket #17786
    lbltitle(13).FontBold = True    'Division
    lbltitle(23).FontBold = True    'Location
    'lblTitle(25).FontBold = True    'Admin by
    lbltitle(26).FontBold = True    'Section
End If

If glbCompSerial = "S/N - 2420W" Then   'Ticket #24396 - Macaulay Child Development
    lbltitle(23).FontBold = True
    lbltitle(24).FontBold = True
    lbltitle(32).FontBold = True 'Ticket #24557 Franks 07/07/2015
End If

Call ScreenSetupNew 'Ticket #25323 Franks 12/16/2014

glbEEFIND_New = True ' allow adding of new employees
UnloadForm = False
If glbLinamar Then
    Call ctrlSetup
Else
    clpCode(2).DataField = "ED_REGION"
    clpCode(4).DataField = "ED_SECTION"
End If
If fglbNewEE Then
    glbNo = False
'    mbAddNewEmployee = True
    frmBlank.Visible = True: frmBlank.Top = 0: frmBlank.Left = 0
    Exit Sub
End If


If Not glbtermopen Then
    If glbLEE_ID = 0 Then
        'Ticket #22682 - Release 8.0: Add New Hire security
        If gSec_Upd_Basic And gSec_Add_NewHire Then  '~~ May99 js
            Msg$ = "Do you wish to add an Employee?"
            VReturn% = MsgBox(Msg$, MB_YESNO, "Add New Employee?")
            If VReturn% = IDYES Then  'YES add a new employee
                If Not modECountChk() Then
                    MsgBox "Maximum number of employees for the license has been reached. Contact HR Systems Strategies Inc."
                    ' DK - 12/07/1999 - Made change to unload form if license has been exceeded
                    Unload Me
                Else
                    glbNo = False
                    mbAddNewEmployee = True
                    frmBlank.Visible = True: frmBlank.Top = 0: frmBlank.Left = 0 'Jaddy 10/27/99
                End If
                Exit Sub
            Else    'NO do not add new employee
                frmEEFIND.Show 1  'find EE if non present, else retrieve it
            End If
        Else
            frmEEFIND.Show 1
        End If                  '~~
    End If
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

frmEEBASIC.Enabled = True
If Not EERetrieve() Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
End If

lblDeptBonusDesc = ""

Call Display_Value
Call getCodes
If glbCompSerial = "S/N - 2344W" Then  'cascade
    lbltitle(39) = "Green Shield"
    medPARKPERMIT1.Tag = "00-Green Shield"
    lbltitle(34) = "Unum"
    medLICPLATE1.Tag = "00-Unum"
    lbltitle(13).FontBold = True 'Ticket #24988 Franks 01/28/2014
End If

If glbCompSerial = "S/N - 2351W" Then ' Burlington Technologies
    lbltitle(40) = "Type of Vehicle #2"
    medPARKPERMIT2.Tag = "00-Type of Vehicle #2"
End If

'Hemu - 11/21/2003 Begin - For the first time it prompts to associate G/L with Dept
'                  even when the Dept. Code has not changed, this was because the
'                   RDept value was empty for the first time
RDept = clpDept.Text
'Hemu - 11/21/2003 End

If glbWFC Then
    'Ticket #29660 - Contractor Employee - Do not check for the Payroll ID validity as Contractors do not go to Payroll
    If NewHireForms.count = 0 Then
        If rsDATA("ED_EMP") = "CONP" Then
            lbltitle(32).FontBold = False    'Payroll ID
        Else
            lbltitle(32).FontBold = True    'Payroll ID
        End If
    Else
        lbltitle(32).FontBold = True    'Payroll ID
    End If
End If

'Added call by Bryan 11/Jan/06 Ticket#10069
    Call SET_UP_MODE
    'Call ST_UPD_MODE(False)
clpGLNum.TextBoxWidth = 1500
End Sub

Private Sub TabOrderSetup() 'Ticket #18090 , Ticket #20045
    If glbCompSerial = "S/N - 2382W" Then  'Samuel
        txtSurname.TabIndex = 0
        txtFName.TabIndex = 1
        txtAdd1.TabIndex = 2
        txtCity.TabIndex = 3
        clpProv.TabIndex = 4
        medPCode.TabIndex = 5
        comCountry.TabIndex = 6
        dlpDOB.TabIndex = 7
        comCountryOfEmp.TabIndex = 8
        medSIN.TabIndex = 9
        dlpDate(0).TabIndex = 10 'doh
        optGender(0).TabIndex = 11
        medTelephone.TabIndex = 12
        clpDept.TabIndex = 13
        dlpDeptEDate.TabIndex = 14
        clpGLNum.TabIndex = 15
        clpCode(5).TabIndex = 16 'Ticket #22066 Franks 05/23/2012
        clpDiv.TabIndex = 17
        dlpDivEDate.TabIndex = 18
        clpCode(1).TabIndex = 19
        clpCode(2).TabIndex = 20
        clpCode(3).TabIndex = 21
        clpCode(4).TabIndex = 22
    End If
End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Resize()
'fraDetail.Height = 8400
'If glbLinamar Then fraDetail.Height = 9800
'
'If Me.Height >= fraDetail.Height + panControls.Height + SSPanel1.Height + 230 Then
'    scrControl.Value = 0
'    fraDetail.Top = 230
'    scrControl.Visible = False
'    Exit Sub
'End If
'If Me.Height < scrControl.Top + panControls.Height + 400 Then Exit Sub
'scrControl.Visible = True
'scrControl.Max = fraDetail.Height + panControls.Height + SSPanel1.Height + 250 - Me.Height
'scrControl.Left = Me.Width - scrControl.Width - 120
'If Me.Height - scrControl.Top - panControls.Height - 500 > 0 Then
'    scrControl.Height = Me.Height - scrControl.Top - panControls.Height - 500
'Else
'    scrControl.Height = 0
'End If
'------------------------------------------------------------------------------------------------
'Ticket #24164 - Re-ordering
'fraDetail.Height = 8400
tbDemographics.Height = 6495    '6000
'If glbLinamar Then fraDetail.Height = 9800

'If Me.Height >= fraDetail.Height + panControls.Height + SSPanel1.Height + 230 Then
If Me.Height >= tbDemographics.Height + panControls.Height + SSPanel1.Height + 3000 Then
    scrControl.Value = 0
    
    'fraDetail.Top = 230
    frPersonal.Top = 960
    frOrganizational.Top = 960
    frMiscellaneous.Top = 960
    
    frPersonal.Left = 0
    frOrganizational.Left = 0
    frMiscellaneous.Left = 0
    
    frMiscellaneous.Height = 2775

    scrControl.Visible = False
    Exit Sub
End If

If Me.Height < scrControl.Top + panControls.Height + 400 Then Exit Sub
scrControl.Visible = True
'scrControl.Max = fraDetail.Height + panControls.Height + SSPanel1.Height + 250 - Me.Height
'If frPersonal.Visible Then
    scrControl.Max = frPersonal.Height + panControls.Height + SSPanel1.Height - 3050 '+ 650 '- Me.Height
'ElseIf frOrganizational.Visible Then
'    scrControl.Max = frOrganizational.Height + panControls.Height + SSPanel1.Height + 650 - Me.Height
'ElseIf frMiscellaneous.Visible Then
'    scrControl.Max = frMiscellaneous.Height + panControls.Height + SSPanel1.Height + 650 - Me.Height
'End If

scrControl.Left = Me.Width - scrControl.Width - 220
If Me.Height - scrControl.Top - panControls.Height - 500 > 0 Then
    scrControl.Height = Me.Height - scrControl.Top - panControls.Height - 700
Else
    scrControl.Height = 0
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

MDIMain.panHelp(0).Caption = "Select function from the menu."
Set frmEEBASIC = Nothing
Call NextForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean

If locUploadWithoutCheck Then 'Ticket #19937 for Samuel -  Franks 05/06/2011
    Exit Sub
End If
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub imgHelp_Click()
Dim MsgStr As String
    'Ticket #24164 - Re-ordering
    tbDemographics.SelectedItem = tbDemographics.Tabs(1)

    'Ticket #22556 WFC
    MsgStr = "Payroll ID must match ADP & Badge ID must match Tracker"
    MsgBox MsgStr, vbInformation
End Sub

Private Sub imgIcon_Click()
    Call txtDeptBonusCtr_DblClick
End Sub

Private Sub imgIDiv_Click()
Call txtDouDiv_DblClick
End Sub

Private Sub lblEEID_Change()
If flagFrmLoad = False Then Exit Sub  'carmen may 00
If fglbNewEE Then Exit Sub
ComMStat.ListIndex = 0
ComSmoker.ListIndex = 0
If rsDATA.RecordCount = 0 Then GoTo EndChanges
If lblEEID = "" Then Exit Sub

lblEENum.Caption = ShowEmpnbr(lblEEID)

If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If

If IsNull(rsDATA("ED_SMOKER")) Then Exit Sub
If rsDATA("ED_SMOKER") Then ComSmoker.ListIndex = 1

If fglbNewEE Then GoTo EndChanges

If IsNull(rsDATA("ED_MSTAT")) Then
  txtMStatus = "S"
Else
  txtMStatus = rsDATA("ED_MSTAT")
End If

If glbWFC Then 'Ticket #23653 Franks 05/07/2013
    ComMStat.ListIndex = 5 'by default
    If rsDATA("ED_MSTAT") = "S" Then ComMStat.ListIndex = 0
    If rsDATA("ED_MSTAT") = "M" Then ComMStat.ListIndex = 1
    If rsDATA("ED_MSTAT") = "D" Then ComMStat.ListIndex = 2
    If rsDATA("ED_MSTAT") = "W" Then ComMStat.ListIndex = 3
    If rsDATA("ED_MSTAT") = "C" Then ComMStat.ListIndex = 4
    If rsDATA("ED_MSTAT") = "O" Then ComMStat.ListIndex = 5
    If rsDATA("ED_MSTAT") = "A" Then ComMStat.ListIndex = 6
ElseIf glbCompSerial = "S/N - 2373W" Then 'Ticket #24565- District Municipality of South Muskoka
    'If RSDATA("ED_MSTAT") = "M" Then ComMStat.ListIndex = 1
    If rsDATA("ED_MSTAT") = "S" Then ComMStat.ListIndex = 0
    If rsDATA("ED_MSTAT") = "F" Then ComMStat.ListIndex = 1
    'If RSDATA("ED_MSTAT") = "P" Then ComMStat.ListIndex = 3
    'If RSDATA("ED_MSTAT") = "D" Then ComMStat.ListIndex = 4
    'If RSDATA("ED_MSTAT") = "W" Then ComMStat.ListIndex = 5
    'If RSDATA("ED_MSTAT") = "C" Then ComMStat.ListIndex = 6
    'If RSDATA("ED_MSTAT") = "R" Then ComMStat.ListIndex = 7
    'If RSDATA("ED_MSTAT") = "X" Then ComMStat.ListIndex = 8
    'If RSDATA("ED_MSTAT") = "O" Then ComMStat.ListIndex = 9
    'If RSDATA("ED_MSTAT") = "A" Then ComMStat.ListIndex = 10
ElseIf glbCompSerial = "S/N - 2482W" Then 'Ticket #28794 - Windsor Family Credit Union
    If rsDATA("ED_MSTAT") = "M" Then ComMStat.ListIndex = 1
    If rsDATA("ED_MSTAT") = "C" Then ComMStat.ListIndex = 2
    If rsDATA("ED_MSTAT") = "W" Then ComMStat.ListIndex = 3
    If rsDATA("ED_MSTAT") = "A" Then ComMStat.ListIndex = 4
    If rsDATA("ED_MSTAT") = "D" Then ComMStat.ListIndex = 5
Else
    If rsDATA("ED_MSTAT") = "M" Then ComMStat.ListIndex = 1
    If rsDATA("ED_MSTAT") = "F" Then ComMStat.ListIndex = 2
    If rsDATA("ED_MSTAT") = "P" Then ComMStat.ListIndex = 3
    If rsDATA("ED_MSTAT") = "D" Then ComMStat.ListIndex = 4
    If rsDATA("ED_MSTAT") = "W" Then ComMStat.ListIndex = 5
    If rsDATA("ED_MSTAT") = "C" Then ComMStat.ListIndex = 6
    If rsDATA("ED_MSTAT") = "R" Then ComMStat.ListIndex = 7
    If rsDATA("ED_MSTAT") = "X" Then ComMStat.ListIndex = 8
    If rsDATA("ED_MSTAT") = "O" Then ComMStat.ListIndex = 9
    If rsDATA("ED_MSTAT") = "A" Then ComMStat.ListIndex = 10
End If

If Not IsNull(rsDATA("ED_COUNTRY")) Then comCountry = rsDATA("ED_COUNTRY")
If Not IsNull(rsDATA("ED_WORKCOUNTRY")) Then
    comCountryOfEmp = rsDATA("ED_WORKCOUNTRY")
Else
    comCountryOfEmp = ""
End If

EndChanges:
    Call SetCountries

End Sub

Private Sub medAltPayID_Change(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

'Private Sub lblLgrDesc_Click()
'Dim Leager$, SQLQ$, Msg$
'Dim snapLeagers As New ADODB.Recordset
'SQLQ$ = "SELECT GL_DESCR from HRGL "
'SQLQ$ = SQLQ$ & "WHERE GL_NO = '" &  clpGLNum.Text & "'"
'snapLeagers.Open SQLQ$, gdbAdoIhr001, adOpenStatic
'If snapLeagers.BOF And snapLeagers.EOF Then
'    snapLeagers.Close
'Else
'    Msg$ = "This Leager already exists"
'    MsgBox Msg$
'    snapLeagers.Close
'    Exit Sub
'End If
'Exit Sub
'End Sub


Private Sub medCellPhone_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medCOMBINATION_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medDRIVERLIC_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medLICPLATE1_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medLICPLATE2_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medLOCKER_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medNetworkLogin_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medPageNbr_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medPARKPERMIT1_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medPARKPERMIT2_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medPCode_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub


Private Sub medPCode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub medSIN_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medSSN_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medTele2_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medTelephone_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medTelephone_KeyPress(KeyAscii As Integer)
'    If glbCompSerial = "S/N - 2382W" And KeyAscii = 9 Then  'Samuel
'        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
'        clpDept.SetFocus
'    End If
End Sub

Private Sub medTelephone_LostFocus()
    If glbCompSerial = "S/N - 2382W" Then
        If GetTabState Then
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
            clpDept.SetFocus
        End If
    End If
End Sub

Private Sub medTYPEVEHICLE_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medVendorNo_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optGender_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Function Review_EE_Tables(EEID&)
Dim snapEETables As New ADODB.Recordset
Dim SQLQ As String
Dim TabName$, EEIDAlias$, TabDescription$

Review_EE_Tables = False

'''On Error GoTo Review_EE__Err

SQLQ = "SELECT * FROM INFO_HR_TABLES "
SQLQ = SQLQ & " WHERE Employee_Keyed <>0 "
SQLQ = SQLQ & " AND TERMINATION_TABLE=0"
'Ticket #20415 - Add Serial # to the select statement so custom tables also gets employee # changed.
'Serial 9999 is by default for all standard info:HR table.
SQLQ = SQLQ & " AND (SERIAL = 'S/N - 9999W' OR SERIAL = '" & glbCompSerial & "')"

snapEETables.Open SQLQ, gdbAdoIhr001, adOpenStatic

If snapEETables.RecordCount < 1 Then Exit Function
snapEETables.MoveFirst
lstEETables.Clear

While Not snapEETables.EOF
    TabName$ = snapEETables("Table_Name")
    If UCase(Right(TabName$, 3)) <> "WRK" Then
        If glbtermopen Then
            EEIDAlias$ = "TERM_SEQ"
        Else
            If IsNull(snapEETables("EMPNBR_Alias")) Then
                EEIDAlias$ = ""
            Else
                EEIDAlias$ = snapEETables("EMPNBR_Alias")
            End If
        End If
        If IsNull(snapEETables("Table_Description")) Then
            TabDescription$ = ""
        Else
            TabDescription$ = snapEETables("Table_Description")
        End If
        If chkForEEData(TabName$, EEIDAlias$, EEID&) > 0 Then
            lstEETables.AddItem TabDescription$ & "  (" & TabName$ & ")"
            Review_EE_Tables = True
        End If
    End If
    snapEETables.MoveNext
Wend

snapEETables.Close

If glbCompSerial = "S/N - 2279W" Then  'Friesens Corporation
    SQLQ = "SELECT PH_EMPNBR FROM HR_PERFORM_FRIESEN WHERE PH_EMPNBR =" & EEID& & " "
    snapEETables.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not snapEETables.EOF Then
        lstEETables.AddItem "Staff Profile (HR_PERFORM_FRIESEN)"
    End If
End If

Exit Function

Review_EE__Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HR_TABLES Error", "TabName$", "Search")
Call RollBack '21June99 js

End Function

Private Sub SetCountries()
'Jaddy fixed the problem for country changed. Ticket #3196 Nov 18, 2002
      
If UCase(comCountry) = "CANADA" Then
    lbltitle(6) = "Province"                    '
    clpProv.Tag = "31-Province - Code"          '
    
    lbltitle(5) = "Postal Code"                 '
    medPCode.MaxLength = 7
    medPCode.Mask = "?#? #?#"
    medPCode.Tag = "01-Postal Code"
    
    lbltitle(8) = "S.I.N"                       '
    medSIN.MaxLength = 11
    medSIN.Mask = "###-###-###"
    medSIN.Tag = "11-Social Insurance Number"   '
    
    If glbCompSerial = "S/N - 2376W" Then
        lbltitle(22) = "Status Number"
        medSSN.Mask = "&&&&&&&&&&" 'Ticket #16291
        medSSN.Tag = "10-Status Number"
    Else
        lbltitle(22) = "S.S.N"
        medSSN.Mask = "###-##-####"
        medSSN.Tag = "10-Social Security Number"   '
    End If
          '
    medTelephone.MaxLength = 14
    medTelephone.Mask = "(###) ###-####"
    
    medTele2.MaxLength = 27
    medTele2.Mask = "(###) ###-####   Ext(#####)"
    medCellPhone.MaxLength = 14
    medCellPhone.Mask = "(###) ###-####"
    medPageNbr.MaxLength = 14
    medPageNbr.Mask = "(###) ###-####"
    
ElseIf comCountry = "U.S.A." Then
    
    lbltitle(6) = "State"
    clpProv.Tag = "31-State - Code"         '
    
    lbltitle(5) = "Zip Code"
    medPCode.MaxLength = 10
    medPCode.Mask = "AAAAA-AAAA"
    medPCode.Tag = "01-Zip Code"            '
    
    lbltitle(8) = "S.S.N."
    medSIN.MaxLength = 11
    medSIN.Mask = "###-##-####"
    medSIN.Tag = "11-Social Security Number" '
    
    If glbCompSerial = "S/N - 2376W" Then
        lbltitle(22) = "Status Number"
        'medSSN.Mask = "###-##-####"
        medSSN.Tag = "10-Status Number"
    Else
        lbltitle(22) = "S.I.N."
        medSSN.Mask = "###-###-###"
        medSSN.Tag = "10-Social Insurance Number"   '
    End If
    
    medTelephone.MaxLength = 14
    medTelephone.Mask = "(###) ###-####"
    medTele2.MaxLength = 27
    medTele2.Mask = "(###) ###-####   Ext(#####)"
    
    medCellPhone.MaxLength = 14
    medCellPhone.Mask = "(###) ###-####"
    medPageNbr.MaxLength = 14
    medPageNbr.Mask = "(###) ###-####"

ElseIf comCountry = "MEXICO" Then
    
    lbltitle(6) = "State"
    clpProv.Tag = "31-State - Code"         '
    
    lbltitle(5) = "Zip Code"
    medPCode.MaxLength = 10
    medPCode.Mask = "AAAAA-AAAA"
    medPCode.Tag = "01-Zip Code"            '
    
    lbltitle(8) = "S.S.N."
    medSIN.MaxLength = 15
    If glbLinamar Then
        medSIN.Mask = "###############"
    Else
        medSIN.Mask = "##########-#"
    End If
    medSIN.Tag = "15-Social Security Number" '
    
    lbltitle(22) = "S.I.N."
    medSSN.Mask = "###-###-###"
    medSSN.Tag = "10-Social Insurance Number"   '
    
    If glbLinamar Then
        medTelephone.MaxLength = 25
        medTelephone.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
        medTele2.MaxLength = 25
        medTele2.Format = ""
        medTele2.Mask = ""
        medTele2.Format = "&&&&&&&&&&&&&&&&&&&&&&&&&"
        medTele2.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
        
        medCellPhone.MaxLength = 25
        medCellPhone.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
        medPageNbr.MaxLength = 25
        medPageNbr.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
    Else
        medTelephone.MaxLength = 25
        medTelephone.Mask = "(###) ###-####"
        medTele2.MaxLength = 27
        medTele2.Mask = "(###) ###-####   Ext(#####)"
        
        medCellPhone.MaxLength = 25
        medCellPhone.Mask = "(###) ###-####"
        medPageNbr.MaxLength = 25
        medPageNbr.Mask = "(###) ###-####"
    End If
    
ElseIf UCase(comCountry) = "BAHAMAS" Then
    lbltitle(6) = "Island"                      '
    clpProv.Tag = "30-Island - Code"            '
    
    lbltitle(5) = "Postal Code"                 '
    medPCode.MaxLength = 8
    medPCode.Mask = "AAAAAAAA"
    medPCode.Tag = "01-Postal Code"             '
    
    lbltitle(8) = "National Ins."               '
    medSIN.MaxLength = 15
    medSIN.Mask = "###############"
    medSIN.Tag = "15-National Insurance Number" '
    
    lbltitle(22) = "S.S.N."
    medSSN.Mask = "###-##-####"
    medSSN.Tag = "10-Social Security Number"   '

    medTelephone.MaxLength = 25
    medTelephone.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
    medTele2.MaxLength = 25
    medTele2.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
    medTele2.Format = "&&&&&&&&&&&&&&&&&&&&&&&&&"

    medCellPhone.MaxLength = 25
    medCellPhone.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
    medPageNbr.MaxLength = 25
    medPageNbr.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
    
Else
    lbltitle(6) = "Province"                '
    clpProv.Tag = "31-Province - Code"      '
    
    lbltitle(5) = "Postal Code"             '
    medPCode.Mask = "&&&&&&&&&&&&&&&"
    medPCode.MaxLength = 15 ' 10
    medPCode.Tag = "01-Postal Code"         '
    
    lbltitle(8) = "National Ins." '"S.I.N"                   '
    medSIN.MaxLength = 25   '15 ' 11    ''Ticket #18668
    medSIN.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
    medSIN.Tag = "25-National Insurance Number" '
    
    lbltitle(22) = "S.S.N"                   '
    medSSN.MaxLength = 25   'Ticket #18668
    medSSN.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
    medSSN.Tag = "25-Social Security Number"   '

    medTelephone.MaxLength = 25
    medTelephone.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
    medTele2.MaxLength = 25
    medTele2.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
    medTele2.Format = "&&&&&&&&&&&&&&&&&&&&&&&&&"
    
    medCellPhone.MaxLength = 25
    medCellPhone.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
    medPageNbr.MaxLength = 25
    medPageNbr.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
End If

End Sub

Private Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF
'cmdPrint.Enabled = FT
'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdMiss.Enabled = FT
comCountry.Enabled = TF
comCountryOfEmp.Enabled = TF
'Ticket #21119 Franks 11/14/2011 for NGS employees
'- will turn off on 01/01/2012
'Ticket #21330 12/19/2011 Franks
If glbWFC And Len(glbWFCNGSSubGroup) > 0 Then
    ComSmoker.Enabled = False
    'cmdUnlockSmoker.Visible = True 'Ticket #22409 Frank 08/08/2012
    'Ticket #22533 Franks 09/10/2012
    cmdUnlockSmoker.Visible = gSec_WFC_UnlockSmokerStatus
Else
    ComSmoker.Enabled = TF
    cmdUnlockSmoker.Visible = False 'Ticket #22409 Frank 08/08/2012
End If

'Ticket #23247 Franks 07/22/2013
If glbWFC And glbUNION = "U959" Then
    cmdCCLife.Visible = True
Else
    cmdCCLife.Visible = False
End If

ComMStat.Enabled = TF
frmSex.Enabled = TF
medPCode.Enabled = TF
medSIN.Enabled = TF
medSSN.Enabled = TF
medTele2.Enabled = TF
medTelephone.Enabled = TF
medCellPhone.Enabled = TF
medPageNbr.Enabled = TF
If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/21/2014 Franks
    txtBadgeID.Enabled = False
Else
    txtBadgeID.Enabled = TF
End If
txtMidName.Enabled = TF
txtAlias.Enabled = TF
txtAdd1.Enabled = TF
txtAdd2.Enabled = TF
txtCity.Enabled = TF
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
clpCode(4).Enabled = TF
txtCountry.Enabled = TF
clpDept.Enabled = TF
clpDiv.Enabled = TF And Not glbLinamar
dlpDOB.Enabled = TF
txtFName.Enabled = TF
clpGLNum.Enabled = TF
dlpDate(0).Enabled = TF

'Ticket #9780 - Jerry asked to allow them to make the change
'If glbCompSerial = "S/N - 2378W" Then
'    txtPayrollID.Enabled = False
'Else
If glbCompSerial = "S/N - 2396W" And Not fglbNewEE And Len(txtPayrollID.Text) > 0 Then
    'Ticket #17341 Oshawa CHC, not new hire, there is Payroll ID
    txtPayrollID.Enabled = False
ElseIf glbLinamar Then  'Ticket #29759 Franks 02/21/2017
    txtPayrollID.Enabled = False
ElseIf glbCompSerial = "S/N - 2373W" Then   'Ticket #19113 - District Municipality of Muskoka
    txtPayrollID.Enabled = False
Else
    txtPayrollID.Enabled = TF
End If
'End If

clpProv.Enabled = TF
txtSurname.Enabled = TF
txtTitle.Enabled = TF
clpHOME(2).Enabled = TF
clpHOME(4).Enabled = TF
clpHOME(1).Enabled = TF
clpHOME(3).Enabled = TF
dlpDeptEDate.Enabled = TF
dlpDivEDate.Enabled = TF
medDRIVERLIC.Enabled = TF
medLICPLATE1.Enabled = TF
medLICPLATE2.Enabled = TF
medLOCKER.Enabled = TF
medCOMBINATION.Enabled = TF
medTYPEVEHICLE.Enabled = TF
medPARKPERMIT1.Enabled = TF
medPARKPERMIT2.Enabled = TF
If glbtermopen Then
'    cmdNew.Enabled = False
    cmdMiss.Visible = False
    cmdPhoto.Visible = False
End If
If glbLinamar Then
    lbltitle(27).Visible = True
    lbltitle(28).Visible = True
    lbltitle(29).Visible = True
    lbltitle(30).Visible = True
     clpHOME(2).Visible = True
     clpHOME(4).Visible = True
     clpHOME(1).Visible = True
     clpHOME(3).Visible = True
    dlpDeptEDate.Visible = False
    dlpDivEDate.Visible = False
    lblDeptStart.Visible = False
    lblDivStart.Visible = False
End If
If glbLinamar Then
    'Ticket #29327 - Allow them to enter the Country now
    'comCountry.Enabled = False
    ' clpProv.Enabled = False
    clpCode(3).Enabled = False 'Hemu - Disable the Division Group (Admin By) - Ticket #8457
End If
If Not gSec_Show_SIN_SSN Then
    medSIN.Visible = False
    medSSN.Visible = False
End If
If Not gSec_Show_DOB Then
    dlpDOB.Visible = False
    lbltitle(17).Visible = False
End If
If Not gSec_Show_ADDRESS Then
    txtAdd1.Visible = False
    txtAdd2.Visible = False
    txtCity.Visible = False
    clpProv.Visible = False
    medPCode.Visible = False
    comCountry.Visible = False
End If
If Not gSec_Show_Marital Then
    lblMStatus.Visible = False
    ComMStat.Visible = False
End If
If Not gSec_Upd_Basic Then         'May99 js
'    cmdModify.Enabled = False   '
'    cmdNew.Enabled = False       '
'   cmdDelete.Enabled = False
End If                          '

'Oxford Ticket #15590
'If glbCompSerial = "S/N - 2259W" Then
'    lblTitle(12).Enabled = False
'    clpGLNum.Enabled = False
'End If

'Ticket #24543 - Macaulay Child Development Centre
If glbCompSerial = "S/N - 2420W" Then
    'clpCode(0).Enabled = TF
    'clpSalDist.Enabled = TF
    Call MacaulayST_UPD(TF) 'Ticket #24557 Franks 09/08/2014
End If
If glbWFC Then 'Ticket #28637 Franks 05/18/2016
    medNetworkLogin.Enabled = TF
    medVendorNo.Enabled = TF
End If
End Sub

Private Sub MacaulayST_UPD(TF As Integer)
Dim I As Integer
    clpCode(0).Enabled = TF
    clpSalDist.Enabled = TF
    medAltPayID(0).Enabled = TF
    medAltPayID(1).Enabled = TF
    medAltPayID(2).Enabled = TF
    
    clpSalDis2(0).Enabled = TF
    clpSalDis2(1).Enabled = TF
    clpSalDis2(2).Enabled = TF
    
    dlpDate(1).Enabled = TF
    dlpDate(2).Enabled = TF
    dlpDate(3).Enabled = TF
    
    dlpTermDate(0).Enabled = TF
    dlpTermDate(1).Enabled = TF
    dlpTermDate(2).Enabled = TF
    
    For I = 8 To 19
        clpCode(I).Enabled = TF '8 - 19
    Next
    
    If TF Then
        'Once the termination date is entered, do not allow any changes to the Alt. Payroll ID
        If IsDate(dlpTermDate(0).Text) Then '
            medAltPayID(0).Enabled = False
            clpCode(8).Enabled = False
            clpCode(11).Enabled = False
            clpSalDis2(0).Enabled = False
            dlpDate(1).Enabled = False
            clpCode(14).Enabled = False
            clpCode(17).Enabled = False
            'dlpTermDate(0).Enabled = False
        End If
        If IsDate(dlpTermDate(1).Text) Then '
            medAltPayID(1).Enabled = False
            clpCode(9).Enabled = False
            clpCode(12).Enabled = False
            clpSalDis2(1).Enabled = False
            dlpDate(2).Enabled = False
            clpCode(15).Enabled = False
            clpCode(18).Enabled = False
            'dlpTermDate(1).Enabled = False
        End If
        
        If IsDate(dlpTermDate(2).Text) Then '
            medAltPayID(2).Enabled = False
            clpCode(10).Enabled = False
            clpCode(13).Enabled = False
            clpSalDis2(2).Enabled = False
            dlpDate(3).Enabled = False
            clpCode(16).Enabled = False
            clpCode(19).Enabled = False
            'dlpTermDate(2).Enabled = False
        End If
    End If
End Sub

Sub SubPicture()
Dim xPIC
Dim Msg As String
Dim xHeight, xWidth

'''On Error GoTo cmdPic_ERR

If glbtermopen Then Exit Sub

'8.0 - Ticket #22682 - Photo from a folder now
If Not gsEMPLOYEEPHOTO Then
    'Ticket #24164 - Re-ordering
    tbDemographics.SelectedItem = tbDemographics.Tabs(1)
    
    If glbSQL Or glbOracle Then
        If cmdPhoto.Caption = "&Photo Off" Then
            picPhoto.Visible = False
            PicNotF.Visible = False
            cmdPhoto.Caption = "&Photo"
        Else
            picPhoto.Visible = False
            PicNotF.Visible = True
            cmdPhoto.Caption = "&Photo Off"
            Call FillPhoto(Val(lblEEID))
        End If
    End If
Else
    'Ticket #24164 - Re-ordering
    tbDemographics.SelectedItem = tbDemographics.Tabs(1)

    If Len(glbPicDir) < 1 Then
        picPhoto.Visible = False
        Exit Sub
    End If
    If cmdPhoto.Caption = "&Photo Off" Then
        picPhoto.Visible = False
        PicNotF.Visible = False
        picPhoto = LoadPicture()
        cmdPhoto.Caption = "&Photo"
    Else
        picPhoto.Visible = False
        PicNotF.Visible = True
        cmdPhoto.Caption = "&Photo Off"
        Call LoadPhoto(Val(lblEEID))
    End If
End If

Exit Sub

cmdPic_ERR:
If Err Then
  PicNotF.Visible = True
  Exit Sub
End If

End Sub

Private Sub optGender_KeyPress(Index As Integer, KeyAscii As Integer)
If optGender(0).Value = True Then
    txtGender = "M"
ElseIf optGender(1).Value = True Then
    txtGender = "F"
Else
    txtGender = "N" 'Ticket #26152 Franks 11/26/2014
End If
End Sub

Private Sub optGender_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If optGender(0).Value = True Then
    txtGender = "M"
ElseIf optGender(1).Value = True Then
    txtGender = "F"
Else
    txtGender = "N" 'Ticket #26152 Franks 11/26/2014
End If
End Sub

Private Sub scrControl_Change()
'Ticket #24164 - Re-ordering
'fraDetail.Top = 240 - scrControl.Value
If frPersonal.Visible Then
    frPersonal.Top = 1000 - scrControl.Value
ElseIf frOrganizational.Visible Then
    frOrganizational.Top = 1000 - scrControl.Value
ElseIf frMiscellaneous.Visible Then
    frMiscellaneous.Top = 1000 - scrControl.Value
End If
End Sub

Private Sub tbDemographics_Click()
    If tbDemographics.SelectedItem.Index = 1 Then
        frOrganizational.Visible = False
        frMiscellaneous.Visible = False
        frAltPayIDs.Visible = False
        frPersonal.Visible = True
        frPersonal.Top = 960
        frPersonal.Left = 120
        frPersonal.Height = 5895
        frPersonal.Width = 12435
    ElseIf tbDemographics.SelectedItem.Index = 2 Then
        frPersonal.Visible = False
        frMiscellaneous.Visible = False
        frAltPayIDs.Visible = False
        frOrganizational.Visible = True
        frOrganizational.Top = 960
        frOrganizational.Left = 120
        frOrganizational.Height = 5055
        frOrganizational.Width = 12435
        
        'Ticket #28040 - To Track on New Hire if the user went into the Organizational tab at least once.
        If fglbNewEE Then
            flgSwitchOrgTabNewHire = False
        End If
    ElseIf tbDemographics.SelectedItem.Index = 3 Then
        frPersonal.Visible = False
        frOrganizational.Visible = False
        frAltPayIDs.Visible = False
        frMiscellaneous.Visible = True
        frMiscellaneous.Top = 960
        frMiscellaneous.Left = 120
        frMiscellaneous.Height = 2415
        frMiscellaneous.Width = 12435
    ElseIf tbDemographics.SelectedItem.Index = 4 Then 'Ticket #25016 Franks 04/01/2014 for Macaulay
        frPersonal.Visible = False
        frOrganizational.Visible = False
        frMiscellaneous.Visible = False
        frAltPayIDs.Visible = True
        frAltPayIDs.Top = 960
        frAltPayIDs.Left = 120
        frAltPayIDs.Height = 2295
        frAltPayIDs.Width = 10935 ' 7095
    End If
End Sub

Private Sub txtAdd1_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtAdd2_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtAlias_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtBadgeID_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtCity_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Samuel_GL()
If NewHireForms.count > 0 Then  'New Hire only
    If clpCode(3).Text = "5322" Or clpCode(3).Text = "2158" Then
        If Len(clpGLNum.Text) = 0 Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
            
            clpGLNum.SetFocus
            MsgBox lbltitle(12).Caption & " is required if " & lbltitle(25).Caption & " is '5322' or '2158'"
            clpGLNum.SetFocus
        End If
    End If
End If
End Sub

Private Sub Dept_GL()
Dim Response%, Msg$, Title$, DgDef As Double
Dim SQLQ As String
Dim rsDEPT As New ADODB.Recordset
'''On Error GoTo Dept_GL_Err

If Len(clpDept.Text) > 0 Then
    rsDEPT.Open "SELECT DF_GLNO FROM HRDEPT WHERE DF_NBR='" & clpDept.Text & "'", gdbAdoIhr001
    If Not rsDEPT.EOF Then
        RGLNum = rsDEPT("DF_GLNO")
        If glbCompSerial = "S/N - 2443W" Then 'Walters Inc Ticket #25952 Franks 11/04/2014
            If Len(clpGLNum.Text) = 0 Then
                If Not IsNull(rsDEPT("DF_GLNO")) Then
                    clpGLNum.Text = rsDEPT("DF_GLNO")
                End If
            End If
        Else
            If RDept <> clpDept Then
                If IsNull(RGLNum) Then
                    RGLNum = ""
                Else
                    If Not glbCompSerial = "S/N - 2394W" Then 'Ticket #14572 St. John's Rehab Hospital
                        'Ticket #24164 - Re-ordering
                        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
                    
                        Msg$ = lStr("Do you want the associated G/L #?")
                        Title$ = "info:HR"
                        DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
                        Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                        If Response% = IDYES Then clpGLNum.Text = RGLNum
                    End If
                End If
                RDept = clpDept.Text
            End If
        End If
    End If
End If

Exit Sub

Dept_GL_Err:
If Err = 94 Then
     clpGLNum.Text = ""
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Dept Snap", "DEPT", "SELECT")
Call RollBack '21June99 js
End Sub




Private Sub dlpDOB_Change()
If flagFrmLoad = False Then Exit Sub  'carmen may 00
Dim birthdate
Dim Age As Double
 
If Not IsDate(dlpDOB) Then
    lbltitle(17) = ""
Else
    birthdate = CVDate(dlpDOB)
    Age = DateDiff("m", birthdate, Now)
    If month(birthdate) = month(Now) Then
        If Day(Now) < Day(birthdate) Then
            Age = Age - 1
        End If
    End If
    Age = CDbl(Age / 12)

    lbltitle(17) = Format(Age, "#0.0")
    
    'Ticket #25500 - Goodmans - LTD Ends Date -> 65th Birthday - 90days -> get the last day of the month
    If glbCompSerial = "S/N - 2290W" Then
        If Year(CVDate(Format(dlpDOB, "mm/dd/yyyy"))) > 1900 Then
            Call Update_Age65_LTD_Benefit_EndDate(glbLEE_ID, dlpDOB)
        End If
    End If
        
End If

End Sub

Private Sub txtDeptBonusCtr_Change()
If glbWFC Then
    lblDeptBonusDesc = GetBonusRptDesc(txtDeptBonusCtr)
    If Len(txtDeptBonusCtr) > 0 And Len(txtDeptBonusCtr) > 0 Then
        If lblDeptBonusDesc = "" Then lblDeptBonusDesc = "Unassigned"
    End If
End If
End Sub

Private Sub txtDeptBonusCtr_DblClick()
    frmDEPTSBonus.cmdSelect.Enabled = True
    glbDeptInhSel% = False
    frmDEPTSBonus.Show 1
    If Len(frmDEPTSBonus.DeptNbr) > 0 Then
        txtDeptBonusCtr = frmDEPTSBonus.DeptNbr
        lblDeptBonusDesc = frmDEPTSBonus.DeptDesc
    End If
End Sub

Private Sub txtDeptBonusCtr_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtDouDiv_Change()
    clpDiv.Text = txtDouDiv.Text
    lblDouDivDesc.Caption = getDivDesc(clpDiv.Text)
End Sub
Private Function getDivDesc(xCode)
Dim rsDiv As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = "Unassigned"
    If Not IsNull(xCode) Then
        SQLQ = "SELECT DIV,Division_Name FROM HR_DIVISION WHERE DIV = '" & xCode & "' "
        rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsDiv.EOF Then
            xRetVal = rsDiv("Division_Name")
        End If
        rsDiv.Close
    End If
    getDivDesc = xRetVal
End Function
Private Sub txtDouDiv_DblClick()
    'Ticket #21544 Franks 02/08/2012
    Call Get_Div(False)
    If Len(glbDiv) > 0 Then
        txtDouDiv.Text = glbDiv
    End If
    'frmDIVISIONS.cmdSelect.Enabled = True
    'glbDeptInhSel% = False
    'frmDIVISIONS.Show 1
    'If Len(frmDEPTSBonus.DeptNbr) > 0 Then
    '    txtDeptBonusCtr = frmDEPTSBonus.DeptNbr
    '    lblDeptBonusDesc = frmDEPTSBonus.DeptDesc
    'End If
End Sub

Private Sub txtFName_Change()
If flagFrmLoad = False Then Exit Sub  'carmen may 00
If Len(txtFName.Text) > 0 Then  ' dont do on add new until in
    Me.lblEEName = RTrim$(txtSurname) & ", " & RTrim$(txtFName)
End If

End Sub

Private Sub txtFName_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtFName_LostFocus()
    'If glbWFC Then 'Ticket #28772 Franks 06/21/2016
    '    Call WFCNetworkLoginSetup
    'End If
End Sub

Private Sub WFCNetworkLoginSetup()
Dim xSurname, xFName
Dim xStr, xLogID
Dim I As Integer
    'If clpCode(4).Text = "MISS" Or clpCode(4).Text = "TROY" Then
    'Ticket #29552 Franks 12/13/2016 for Salaried employees
    If NewHireForms.count > 0 Then 'new hire only
        If glbTrsUnion = "NONE" Or glbTrsUnion = "EXEC" Then
            If Len(medNetworkLogin.Text) = 0 Then
                If Len(txtFName.Text) > 0 And Len(txtSurname.Text) > 0 Then
                    '''medNetworkLogin.Text = Trim(Left(txtFName.Text, 1) & Trim(Left(txtSurname.Text, 7)))
                    ''medNetworkLogin.Text = getWFCNetworkLogin(txtFName.Text, txtSurname.Text) 'Ticket #28772 Franks 06/22/2016
                    'xStr = getWFCNetworkLogin(txtFName.Text, txtSurname.Text)
                    
                    'If there are two surnames then use the first one(such as "Rambo Rodney" or "Rambo-Rodney")
                    xSurname = Trim(txtSurname.Text)
                    I = InStr(1, xSurname, "-") 'check "-"
                    If I > 0 Then
                        xSurname = Trim(Left(xSurname, I - 1))
                    End If
                    I = InStr(1, xSurname, " ") 'check space
                    If I > 0 Then
                        xSurname = Trim(Left(xSurname, I))
                    End If
                    
                    'If Alias is populate use it as Frist Name
                    xFName = ""
                    If Len(Trim(txtAlias.Text)) > 0 Then
                        xFName = Trim(txtAlias.Text)
                    End If
                    If Len(xFName) = 0 Then
                        xFName = txtFName.Text
                    End If
                    
                    'If there are two first names then use the first one
                    xFName = Trim(xFName)
                    I = InStr(1, xFName, "-") 'check "-"
                    If I > 0 Then
                        xFName = Trim(Left(xFName, I - 1))
                    End If
                    I = InStr(1, xFName, " ") 'check space
                    If I > 0 Then
                        xFName = Trim(Left(xFName, I))
                    End If

                
                    xStr = getWFCNetworkLogin(xFName, xSurname)
                    
                    xLogID = xStr
                    
                    'New function with duplicate check
                    xLogID = getWFCNetworkLoginNoDupicate(xFName, xSurname, txtMidName, glbLEE_ID, xStr)
                    
                    medNetworkLogin.Text = xLogID
                End If
            End If
            If Len(medVendorNo.Text) = 0 Then
                medVendorNo.Text = "N/A"
            End If
        End If
        If Len(medNetworkLogin.Text) = 0 Then ''Ticket #30491 Franks 09/07/2017
            medNetworkLogin.Text = "N/A"
        End If
    End If
End Sub

Private Sub txtGender_Change()
If Len(txtGender.Text) > 0 Then
    If txtGender = "M" Then
        optGender(0) = True
        optGender(1) = False
        optGender(2) = False
    ElseIf txtGender = "F" Then
        optGender(0) = False
        optGender(1) = True
        optGender(2) = False
    Else 'Ticket #26152 Franks 11/26/2014
        optGender(0) = False
        optGender(1) = False
        optGender(2) = True
    End If
End If

End Sub

Private Sub txtMidName_GotFocus()
    Call SetPanHelp(Me.ActiveControl)

End Sub

Private Sub txtOtherEmail_DblClick()
On Error GoTo Email_Err
    If gsEMAIL_SENDING Then
        If Len(txtOtherEmail.Text) > 0 Then
            frmSendEmail.txtTo.Text = txtOtherEmail.Text
            frmSendEmail.Tag = ""
            frmSendEmail.Show 1
        Else
            MsgBox "Other Email Address is blank."
        End If
    End If
    Exit Sub
    
Email_Err:
    If Err.Number = 364 Then
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send Other EMail", "SMTP", "SENDEMAIL")
    Resume Next
End Sub

Private Sub txtOtherEmail_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtPayrollID_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtPayrollID_KeyPress(KeyAscii As Integer)
If glbVadim Then KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtPayrollID_LostFocus()
    'Ticket #19157
    If glbVadim And Len(txtPayrollID) > 0 Then
        txtPayrollID.Text = Trim(txtPayrollID.Text)
    End If
End Sub

Private Sub txtSurname_Change()
If flagFrmLoad = False Then Exit Sub  'carmen may 00
If Len(txtSurname.Text) > 0 Then  ' dont do on add new until in
    'frmEEBASIC.Caption = "Demographics - " & IIf(mbAddNewEmployee, "New Employee", Left$(txtSurname, 5))    'Jaddy 10/27/99
    'frmEEBASIC.lblEEName = RTrim$(txtSurname) & ", " & RTrim$(txtFName)
    Me.Caption = "Demographics - " & IIf(mbAddNewEmployee, "New Employee", Left$(txtSurname, 5))    'Jaddy 10/27/99
    Me.lblEEName = RTrim$(txtSurname) & ", " & RTrim$(txtFName)
End If
End Sub

Private Sub txtSurname_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtSurname_LostFocus()
    'If glbWFC Then 'Ticket #28772 Franks 06/21/2016
    '    Call WFCNetworkLoginSetup
    'End If
End Sub

Private Sub txtTitle_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Function RollBack()
'''On Error GoTo rr
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
rr:
End Function

Private Sub get_NewHireForms()
Dim rsTN As New ADODB.Recordset, X
rsTN.Open "SELECT * FROM HRNEWHIRE WHERE NewHire<>0 ORDER BY ID", glbAdoIHRDB
For X = 1 To NewHireForms.count: NewHireForms.Remove 1: Next
Do Until rsTN.EOF
    If Trim(strNoAccessForms) <> "" Then
        'Skip the screens the user does not have access to
        If InStr(strNoAccessForms, Trim(rsTN!MenuItem)) = 0 Then
            NewHireForms.Add Trim(rsTN!FormName)
        End If
    Else
        NewHireForms.Add Trim(rsTN!FormName)
    End If
    rsTN.MoveNext
Loop
rsTN.Close
End Sub

Private Sub get_NewHireForms_Add_OtherInformation_Form()
Dim X
Dim rsTN As New ADODB.Recordset

'Check if Other Information form already in New Hire Procedure
'If so then don't need to add that form in the NewHireForms collection
rsTN.Open "SELECT * FROM HRNEWHIRE WHERE NewHire<>0 AND FormName = 'frmEmpOther' ORDER BY ID", glbAdoIHRDB
If rsTN.EOF Then
    'Add the form in the new hire forms collection if the user has access to that screen.
    If Trim(strNoAccessForms) <> "" Then
        'User has access to limited screens
        'Skip the screen if user does not have access to
        If InStr(strNoAccessForms, "frmEmpOther") = 0 Then
            'Has access to Other Information screen
            NewHireForms.Add Trim("frmEmpOther"), , , 2
        End If
    Else
        'User has access to all the screens
        NewHireForms.Add Trim("frmEmpOther"), , , 2
    End If
End If
rsTN.Close
Set rsTN = Nothing
End Sub

Private Sub addItems()
Dim ctylist, X
ComSmoker.AddItem "No"
ComSmoker.AddItem "Yes"
ComSmoker.ListIndex = 0

ctylist = CountryList
X = 1
Do While X > 0
    X = InStr(ctylist, "&")
    If X > 0 Then
        comCountry.AddItem Left(ctylist, X - 1)
        comCountryOfEmp.AddItem Left(ctylist, X - 1)
        ctylist = Mid(ctylist, X + 1)
    Else
        comCountry.AddItem ctylist
        comCountryOfEmp.AddItem ctylist
    End If
Loop

comCountry.ListIndex = 0        '

If glbWFC Then 'Ticket #23653 Franks 05/07/2013
    ComMStat.AddItem "Single"
    ComMStat.AddItem "Married"
    ComMStat.AddItem "Divorced"
    ComMStat.AddItem "Widowed"
    ComMStat.AddItem "Common-Law"
    ComMStat.AddItem "Other"
    ComMStat.AddItem "Separated"
    ComMStat.ListIndex = 0
ElseIf glbCompSerial = "S/N - 2373W" Then 'District Municipality of South Muskoka
    ComMStat.AddItem "Single"
    ComMStat.AddItem "Family"
    ComMStat.ListIndex = 0
ElseIf glbCompSerial = "S/N - 2482W" Then 'Ticket #28794 - Windsor Family Credit Union -
    ComMStat.AddItem "Single"
    ComMStat.AddItem "Married"
    ComMStat.AddItem "Common-Law"
    ComMStat.AddItem "Widow/Widower"
    ComMStat.AddItem "Separated"
    ComMStat.AddItem "Divorced"
    ComMStat.ListIndex = 0
Else
    ComMStat.AddItem "Single"
    ComMStat.AddItem "Married"
    ComMStat.AddItem "Family"
    ComMStat.AddItem "Parent(Single)"
    ComMStat.AddItem "Divorced"
    ComMStat.AddItem "Widowed"
    ComMStat.AddItem "Common-Law"
    ComMStat.AddItem "Partner"          'Jdy for screen changes 4/20/00
    ComMStat.AddItem "Same-Sex"         'Jdy for screen changes 4/20/00
    ComMStat.AddItem "Other"
    ComMStat.AddItem "Separated"
    ComMStat.ListIndex = 0
End If
End Sub

Private Function FldList()
Dim SQLQ
SQLQ = ""
SQLQ = SQLQ & "ED_COMPNO, ED_EMPNBR, ED_PAYROLL_ID, ED_TITLE,"
SQLQ = SQLQ & "ED_ENTOPT, ED_EFDATE, ED_ETDATE, ED_SURNAME,"
SQLQ = SQLQ & "ED_ENTOPTS, ED_EFDATES, ED_ETDATES, ED_FNAME,"
SQLQ = SQLQ & "ED_ADDR1, ED_ADDR2, ED_CITY, ED_PROV, ED_PCODE,"
SQLQ = SQLQ & "ED_COUNTRY, ED_DOB, ED_DOH, ED_SIN, ED_SSN,"
SQLQ = SQLQ & "ED_MSTAT, ED_SMOKER, ED_SEX, ED_PHONE, ED_BUSNBR,"
SQLQ = SQLQ & "ED_CELLPHONE, ED_PAGENBR, ED_DEPTNO, ED_GLNO,"
SQLQ = SQLQ & "ED_DIV, ED_LOC, ED_REGION, ED_ADMINBY, ED_SECTION,"
SQLQ = SQLQ & "ED_TD1, ED_TD1DOL, ED_PROVFORM, ED_PROVAMT, ED_PT,"
SQLQ = SQLQ & "ED_DEPTEDATE,ED_DIVEDATE,"
SQLQ = SQLQ & "ED_HOMELINE,ED_HOMESHIFT,ED_HOMEOPRTNBR,ED_HOMEWRKCNT,"
SQLQ = SQLQ & "ED_DRIVERLIC,ED_LICPLATE1,ED_LICPLATE2,"
SQLQ = SQLQ & "ED_TYPEVEHICLE,ED_PARKPERMIT1,ED_PARKPERMIT2,"
SQLQ = SQLQ & "ED_BADGEID,ED_MIDNAME,ED_ALIAS,ED_EML,"
SQLQ = SQLQ & "ED_WCB,ED_CPP,ED_UIC,ED_GROSSCD,ED_VADIM2,ED_SFDATE,"
SQLQ = SQLQ & "ED_LOCKER,ED_COMBINATION,ED_WORKCOUNTRY,ED_BONUSDEPT,"
SQLQ = SQLQ & "ED_LDATE, ED_LTIME,ED_LUSER, ED_PROVEMP"
SQLQ = SQLQ & ",ED_EMP" 'Ticket #19988 Franks 06/06/2011
SQLQ = SQLQ & ",ED_FDAY,ED_LDAY" 'Ticket #22261 Franks 07/27/2012 '
SQLQ = SQLQ & ",ED_NORMALR" 'Ticket #24695 Franks 11/26/2013
SQLQ = SQLQ & ",ED_PTEDATE,ED_BENEFIT_GROUP" 'Ticket #23564 Franks 04/15/2013

'Ticket #24164 - Re-ordering and new fields - Organization fields
'If glbCompSerial = "S/N - 2382W" Then
SQLQ = SQLQ & ",ED_ORGT1, ED_ORGT1EDATE, ED_ORGT2, ED_ORGT2EDATE"


If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"
If glbCompSerial = "S/N - 2259W" Then SQLQ = SQLQ & ",ED_LTHIRE"  'For County of Oxford
If glbCompSerial = "S/N - 2363W" Or glbWFC Then SQLQ = SQLQ & ",ED_ORG,ED_VADIM1,ED_CANDIDATE"     'City of Kawartha Lakes
If glbCompSerial = "S/N - 2375W" Then SQLQ = SQLQ & ",ED_PENSION" ',ED_NORMALR"  'For City of Timmins
If glbCompSerial = "S/N - 2380W" Or glbWFC Then SQLQ = SQLQ & ",ED_EMPTYPE"  'For VitalAire
If glbLambton Then SQLQ = SQLQ & ",ED_EMAIL"
If glbCompSerial = "S/N - 2382W" Then SQLQ = SQLQ & ",ED_SUBDEPT" 'Samuel Ticket #20600 Franks 09/21/2011
If glbLinamar Then SQLQ = SQLQ & ",ED_VADIM1 " 'Ticket #29759 Franks 02/14/2017

'Ticket #24543 - Macaulay Child Development Centre
'If glbCompSerial = "S/N - 2420W" Then SQLQ = SQLQ & ",ED_ORG, ED_SALDIST"
SQLQ = SQLQ & ",ED_ORG, ED_SALDIST, ED_WCBCODE "

'Ticket #29375
SQLQ = SQLQ & ",ED_OTHREMAIL "

FldList = SQLQ
End Function

Private Function CountryList() As String
Dim xCountryList As String, ctyFile
xCountryList = ""
ctyFile = glbIHRREPORTS & "CountryList.MTF"

'''On Error GoTo ErrorHandler

If File(ctyFile) Then
    Open ctyFile For Input As #1
    Input #1, xCountryList
    Close #1
End If

ResumeHere:
'If InStr(xCountryList, BasicCountry) = 0 Then
'    xCountryList = BasicCountry
'End If
If InStr(xCountryList, comCountry) = 0 And comCountry <> "" Then
    xCountryList = xCountryList & "&" & comCountry
    comCountry.AddItem comCountry
    comCountryOfEmp.AddItem comCountry
End If
Open ctyFile For Output As #1
Print #1, xCountryList
Close #1
CountryList = xCountryList
Exit Function

ErrorHandler:
If Err.Number = 62 Then
    'Ticket #24164 - Re-ordering
    tbDemographics.SelectedItem = tbDemographics.Tabs(1)

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


Private Sub ctrlSetup()
txtAdd1.MaxLength = 45
'lblTitle(31).Visible = False: txtAdd2.Visible = False
lblPayroll.Visible = True: lblPerson.Visible = True
lbltitle(24).FontBold = True

picPhoto.Top = picPhoto.Top - 50
Call ctrlTop(lbltitle(4), txtCity, 1)
Call ctrlTop(lbltitle(6), clpProv, 1)
Call ctrlTop(lbltitle(5), medPCode, 2)
Call ctrlTop(lbltitle(19), comCountry, 2)
Call ctrlTop(lbltitle(9), dlpDOB, 3)
Call ctrlTop(lbltitle(44), comCountryOfEmp, 3)
Call ctrlTop(lbltitle(14), lbltitle(17), 3)
Call ctrlTop(lbltitle(8), medSIN, 4)
Call ctrlTop(lbltitle(22), medSSN, 4)
Call ctrlTop(lbltitle(15), lblDOH, 4)
Call ctrlTop(lblMStatus, ComMStat, 5)
Call ctrlTop(lbltitle(16), ComSmoker, 5)
Call ctrlTop(txtGender, frmSex, 5)
Call ctrlTop(lbltitle(7), medTelephone, 6)
Call ctrlTop(lbltitle(10), medTele2, 6)
Call ctrlTop(lbltitle(20), medCellPhone, 7)
Call ctrlTop(lbltitle(21), medPageNbr, 7)
Call ctrlTop(lbltitle(33), medDRIVERLIC, 8)
Call ctrlTop(lbltitle(38), medTYPEVEHICLE, 8)
Call ctrlTop(lbltitle(39), medPARKPERMIT1, 9)
Call ctrlTop(lbltitle(40), medPARKPERMIT2, 9)
Call ctrlTop(lbltitle(34), medLICPLATE1, 10)
Call ctrlTop(lbltitle(35), medLICPLATE2, 10)
'Call ctrlTop(lblTitle(36), medLOCKER, 11)
'Call ctrlTop(lblTitle(37), medCOMBINATION, 11)

'Ticket #24164 - Re-ordering
'lblPayroll.Top = lblTitle(37).Top + 330
'lblPerson.Top = lblTitle(37).Top + 330
'lblPayroll.Left = 0
'lblPerson.Left = lblPayroll.Left + 6200

lblPayroll.Top = lbltitle(11).Top - 50
lblPerson.Top = lbltitle(11).Top - 50
lblPayroll.Left = 0
lblPerson.Left = lblPayroll.Left + 5600


Call CodesTop(lbltitle(24), clpCode(2), 1)
Call CodesTop(lbltitle(27), clpHOME(1), 2)
Call CodesTop(lbltitle(29), clpHOME(2), 3)
Call CodesTop(lbltitle(12), clpGLNum, 4)
Call CodesTop(lbltitle(13), clpDiv, 5)
Call CodesTop(lbltitle(23), clpCode(1), 6)
Call CodesTop(lbltitle(25), clpCode(3), 7)
Call CodesTop(lbltitle(30), clpHOME(3), 8)
Call CodesTop(lbltitle(28), clpHOME(4), 9)

'Organization fields
Call CodesTop(lbltitle(18), clpCode(6), 10)
Call CodesTop(lbltitle(45), clpCode(7), 11)
lblOrg1EffDate.Top = lbltitle(18).Top
dlpOrg1EDate.Top = clpCode(6).Top
lblOrg2EffDate.Top = lbltitle(45).Top
dlpOrg2EDate.Top = clpCode(7).Top

Call CodesTop(lbltitle(11), clpDept, 1)
'Ticket #24164 - Re-ordering
'Call CodesTop(lblTitle(26), clpCode(4), 2)
Call CodesTop(lbltitle(26), clpCode(4), 1)

frOrganizational.Height = 4500

'Ticket #28846 Franks 07/13/2016 - begin
clpProv.Width = 2535 - 400
lbltitle(57).FontBold = True
lbltitle(57).Visible = True
clpProvEmp.Visible = True
clpProvEmp.DataField = "ED_PROVEMP"
'Ticket #28846 Franks 07/13/2016 - end

'Ticket #29759 Franks 02/14/2017 - begin
lblVadim1.Top = 2085
lblVadim1.Caption = lStr("Vadim Field 1")
lblVadim1.Visible = True
clpVadim1.Top = 2040
clpVadim1.Visible = True
clpVadim1.DataField = "ED_VADIM1"

txtPayrollID.Enabled = False
cmdEditPayID.Visible = True
'Ticket #29759 Franks 02/14/2017 - end

End Sub

Private Sub ctrlTop(lbl As Control, Ctrl As Control, xLine)
Dim chgTop
'chgTop = 320 + xLine * 0
'If lbl.name = "lblTitle" Then
'    If lbl.Index = 36 Or lbl.Index = 37 Then
'        chgTop = 5 * 330  '+320
'    End If
'End If

lbl.Top = lbl.Top - chgTop
Ctrl.Top = Ctrl.Top - chgTop
End Sub

Private Sub CodesTop(lbl As Control, Ctrl As Control, xLine)
Dim chgTop, chgLeft
Static X As Integer

'Ticket #24164 - Re-ordering
'chgTop = lbl.Top - (lblTitle(37).Top + 330) - (330 * xLine)
chgTop = lbl.Top - (lbltitle(11).Top + 50) - (350 * xLine)

lbl.Top = lbl.Top - chgTop
Ctrl.Top = Ctrl.Top - chgTop
If lbl.Index = 24 Then
    lbl.Alignment = 0
    lbl.Left = lbltitle(13).Left
    
    'Ticket #24164 - Re-ordering
    'Ctrl.Left = dlpDOB.Left
    Ctrl.Left = dlpDOB.Left + 50
    
    Ctrl.MaxLength = 8
    
    'Ticket #24164 - Re-ordering
    'Ctrl.Width = 2800
ElseIf lbl.Index = 11 Then
    'Department
    'Ticket #24164 - Re-ordering
    lbl.Top = lbltitle(24).Top
    Ctrl.Top = clpCode(2).Top
    
    'lbl.Left = lbl.Left + 5500
    lbl.Left = lbltitle(26).Left
    'lbl.Alignment = 1
    
    Ctrl.Left = clpCode(4).Left
ElseIf lbl.Index = 26 Then
    'Section
    'Ticket #24164 - Re-ordering
    lbl.Top = lbltitle(27).Top
    Ctrl.Top = clpHOME(1).Top
    'lbl.Alignment = 1
    'Ctrl.Width = 3800
Else
    'Ticket #24164 - Re-ordering
    Ctrl.Left = clpCode(2).Left
End If

X = X + 1

'Ticket #24164 - Re-ordering
'Ctrl.TabIndex = medCOMBINATION.TabIndex + X
Ctrl.TabIndex = medTele2.TabIndex + X


End Sub

Private Sub UpdCodes()
    If glbLinamar Then
        If Trim(clpHOME(1)) <> "" Then
            rsDATA("ED_HOMEOPRTNBR") = clpHOME(1).TransDiv & clpHOME(1)
        Else
            rsDATA("ED_HOMEOPRTNBR") = Null
        End If
        If Trim(clpHOME(2)) <> "" Then
            rsDATA("ED_HOMELINE") = clpHOME(2).TransDiv & clpHOME(2)
        Else
            rsDATA("ED_HOMELINE") = Null
        End If
        If Trim(clpCode(2).Text) <> "" Then
            rsDATA("ED_REGION") = getProductLineCodeforLinamar(clpCode(2).TransDiv & clpCode(2).Text)
        Else
            rsDATA("ED_REGION") = ""
        End If
        If Trim(clpCode(4).Text) <> "" Then
            rsDATA("ED_SECTION") = clpCode(4).TransDiv & clpCode(4).Text
        Else
            rsDATA("ED_SECTION") = ""
        End If
    End If

End Sub

Sub getCodes()
If rsDATA.EOF Then Exit Sub
If glbLinamar Then
    
    clpCode(2).TransDiv = clpDiv
    clpCode(4).TransDiv = clpDiv
    clpHOME(1).TransDiv = clpDiv
    clpHOME(2).TransDiv = clpDiv

    If Not IsNull(rsDATA("ED_HOMEOPRTNBR")) Then
         clpHOME(1) = Mid(rsDATA("ED_HOMEOPRTNBR"), 4)
    Else
         clpHOME(1) = ""
    End If
    If Not IsNull(rsDATA("ED_HOMELINE")) Then
         clpHOME(2) = Mid(rsDATA("ED_HOMELINE"), 4)
    Else
         clpHOME(2) = ""
    End If
    If Not IsNull(rsDATA("ED_REGION")) Then
        clpCode(2).Text = Mid(rsDATA("ED_REGION"), 4)
    Else
        clpCode(2).Text = ""
    End If
    If Not IsNull(rsDATA("ED_SECTION")) Then
        clpCode(4).Text = Mid(rsDATA("ED_SECTION"), 4)
    Else
        clpCode(4).Text = ""
    End If
Else
    If Not IsNull(rsDATA("ED_REGION")) Then
        clpCode(2).Text = rsDATA("ED_REGION")
    Else
        clpCode(2).Text = ""
    End If
    If Not IsNull(rsDATA("ED_SECTION")) Then
        clpCode(4).Text = rsDATA("ED_SECTION")
    Else
        clpCode(4).Text = ""
    End If
End If
End Sub

Private Function LoadPhoto(zEMPNBR As Long)
Dim xHeight, xWidth
glbPicBMP = glbPicDir & zEMPNBR & ".JPG"

'Hemu
If Not IsNull(glbPicBMP) Then
    If Not (Dir(glbPicBMP) = "") Then
        picPhoto = LoadPicture(glbPicBMP)
    Else
        Exit Function
    End If
Else
    Exit Function
End If
'If Not IsNull(glbPicBMP) Then picPhoto = LoadPicture(glbPicBMP)
'Hemu

picPhoto.Stretch = False
xHeight = picPhoto.Height
xWidth = picPhoto.Width
picPhoto.Stretch = True
picPhoto.Height = 2325
picPhoto.Width = (xWidth * picPhoto.Height) / xHeight
picPhoto.Stretch = True
picPhoto.Visible = True
PicNotF.Visible = False
End Function

Private Function FillPhoto(zEMPNBR As Long)
    '''On Error GoTo ErrHandler:
    Dim rsPHOTO As New ADODB.Recordset
    Dim byteChunk() As Byte

    Dim Offset As Long
    Dim Totalsize As Long
    Dim Remainder As Long

    Dim FieldSize As Long
    Dim FileNumber As Integer
    Const HeaderSize As Long = 78
    Const ChunkSize As Long = 100
    Dim TempFile As String
    Dim TempDir As String * 255
    
    GetTempPath 255, TempDir
    TempFile = Replace(Replace(TempDir, Chr(0), "") & "\tempfile.tmp", "\\", "\")
    
    picPhoto.Picture = Nothing
    If zEMPNBR = 0 Then Exit Function
    rsPHOTO.Open "select * from HR_PHOTO WHERE PT_EMPNBR=" & zEMPNBR, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If rsPHOTO.EOF Then Exit Function
    
    
    FileNumber = FreeFile
    Open TempFile For Binary Access Write As FileNumber
    
    ReDim byteChunk(rsPHOTO("PT_PHOTO").ActualSize)
    byteChunk() = rsPHOTO("PT_PHOTO").GetChunk(rsPHOTO("PT_PHOTO").ActualSize)
    Put FileNumber, , byteChunk()

    Close FileNumber
    picPhoto.Picture = LoadPicture(TempFile)
    Kill (TempFile)
    rsPHOTO.Close
    Dim xHeight, xWidth
    picPhoto.Stretch = False
    xHeight = picPhoto.Height
    xWidth = picPhoto.Width
    picPhoto.Stretch = True
    picPhoto.Height = 2325
    picPhoto.Width = (xWidth * picPhoto.Height) / xHeight
    picPhoto.Stretch = True
    picPhoto.Visible = True
    PicNotF.Visible = False ' Ticket #18190
    Exit Function
    
ErrHandler:
    MsgBox Err.Description, , "Error "
    
End Function

Private Function ChangeEmpnbr()
Dim dyn_Table As New ADODB.Recordset
Dim xCount, xx
Dim SQLQ, X%, xFldTitle, xFld As String, xTable As String
ChangeEmpnbr = False
'''On Error GoTo ChangeEmpnbr_cmdUpdErr
Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 0
SQLQ = "SELECT * FROM INFO_HR_TABLES WHERE TERMINATION_TABLE=0"
'Ticket #20415 - Add Serial # to the select statement so custom tables also gets employee # changed.
'Serial 9999 is by default for all standard info:HR table.
SQLQ = SQLQ & " AND (SERIAL = 'S/N - 9999W' OR SERIAL = '" & glbCompSerial & "')"

dyn_Table.Open SQLQ, gdbAdoIhr001, adOpenStatic
MDIMain.panHelp(0).FloodPercent = 10
xCount = dyn_Table.RecordCount
xx = 0
Do Until dyn_Table.EOF
    MDIMain.panHelp(0).FloodPercent = (xx / xCount) * 60 + 10
    xTable = dyn_Table("Table_Name")
    If IsNull(dyn_Table("EMPNBR_Alias")) Then xFld = "" Else xFld = dyn_Table("EMPNBR_Alias")
    If InStr(xFld, "_") = 0 Then xFldTitle = "" Else xFldTitle = Left(xFld, 3)
    If dyn_Table("Employee_Keyed") Then
        Call UpdateEMPNBR(xTable, xFld, xFldTitle, glbChgNewEmpnbr, glbLEE_ID)
        Select Case xTable
        Case "HR_ATTENDANCE", "HR_ATTENDANCE_HISTORY"
            Call UpdateEMPNBR(xTable, xFldTitle & "SUPER", xFldTitle, glbChgNewEmpnbr, glbLEE_ID)
        Case "HR_JOB_HISTORY", "HR_PERFORM_HISTORY"
            Call UpdateEMPNBR(xTable, xFldTitle & "REPTAU", xFldTitle, glbChgNewEmpnbr, glbLEE_ID)
            Call UpdateEMPNBR(xTable, xFldTitle & "REPTAU2", xFldTitle, glbChgNewEmpnbr, glbLEE_ID)
            Call UpdateEMPNBR(xTable, xFldTitle & "REPTAU3", xFldTitle, glbChgNewEmpnbr, glbLEE_ID)
        Case "HR_OCC_HEALTH_SAFETY"
            Call UpdateEMPNBR(xTable, xFldTitle & "EMPNOT", xFldTitle, glbChgNewEmpnbr, glbLEE_ID)
            Call UpdateEMPNBR(xTable, xFldTitle & "SUPERVISOR", xFldTitle, glbChgNewEmpnbr, glbLEE_ID)
        End Select
    End If
    dyn_Table.MoveNext
    xx = xx + 1
Loop
MDIMain.panHelp(0).FloodType = 0
glbLEE_ID = glbChgNewEmpnbr
Screen.MousePointer = DEFAULT
ChangeEmpnbr = True

Exit Function
ChangeEmpnbr_cmdUpdErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MDIMain.panHelp(0).FloodType = 0
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ChangeEmpnbr Error", xTable, "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If

End Function

Private Function WSQLQ(rsTA As ADODB.Recordset) As String

WSQLQ = WSQLQ & " WHERE " & Replace(glbSeleDeptUn, "ED_", "EL_")
If Len(rsTA("ED_DEPTNO")) > 0 Then WSQLQ = WSQLQ & " AND EL_DEPTNO = '" & rsTA("ED_DEPTNO") & "'"
If Len(rsTA("ED_DIV")) > 0 Then WSQLQ = WSQLQ & " AND EL_DIV = '" & rsTA("ED_DIV") & "' "
If Len(rsTA("ED_LOC")) > 0 Then WSQLQ = WSQLQ & " AND EL_LOC = '" & rsTA("ED_LOC") & "' "
If Len(rsTA("ED_ORG")) > 0 Then WSQLQ = WSQLQ & " AND EL_ORG = '" & rsTA("ED_ORG") & "' "
If (rsTA("ED_EMP")) > 0 Then WSQLQ = WSQLQ & " AND EL_EMP = '" & rsTA("ED_EMP") & "' "
If (rsTA("ED_PT")) > 0 Then WSQLQ = WSQLQ & " AND EL_PT = '" & rsTA("ED_PT") & "' "

End Function

Private Function WSQLQ1(rsTA As ADODB.Recordset) As String

WSQLQ1 = WSQLQ1 & " WHERE " & glbSeleDeptUn
If Len(rsTA("ED_DEPTNO")) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_DEPTNO = '" & rsTA("ED_DEPTNO") & "'"
If Len(rsTA("ED_DIV")) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_DIV = '" & rsTA("ED_DIV") & "' "
If Len(rsTA("ED_LOC")) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_LOC = '" & rsTA("ED_LOC") & "' "
If Len(rsTA("ED_ORG")) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_ORG = '" & rsTA("ED_ORG") & "' "
If (rsTA("ED_EMP")) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_EMP = '" & rsTA("ED_EMP") & "' "
If (rsTA("ED_PT")) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_PT = '" & rsTA("ED_PT") & "' "

End Function

Private Function UPDEML()

Dim Msg$, DgDef As Variant, Response%, noRecs&
Dim rsEML As New ADODB.Recordset
Dim rsTA As New ADODB.Recordset
Dim SQLQ, X%, strFld
Dim XUpdCount

UPDEML = False
'''On Error GoTo cmdInsErr

SQLQ = "SELECT ED_EMPNBR,ED_DIV,ED_DEPTNO,ED_LOC,ED_ORG,ED_EMP,ED_PT FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
rsTA.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
If rsTA.EOF = False And rsTA.BOF = False Then
    rsTA.MoveFirst
    XUpdCount = rsTA.RecordCount
End If


If XUpdCount > 0 Then

    SQLQ = "SELECT ID,EL_DIV,EL_DEPTNO,EL_LOC,EL_ORG,EL_EMP,EL_PT,EL_EML,EL_LDATE,EL_LTIME,EL_LUSER FROM HR_EMLSETUP" & WSQLQ(rsTA)
    rsEML.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEML.EOF Then
        ' UPDATE ED_EML VALUE IN HREMP
        SQLQ = "UPDATE HREMP SET"
        SQLQ = SQLQ & " ED_EML=" & rsEML("EL_EML")
        SQLQ = SQLQ & WSQLQ1(rsTA)
        gdbAdoIhr001.Execute SQLQ
     End If
End If

rsEML.Close
rsTA.Close
UPDEML = True

Exit Function
cmdInsErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
If glbErrNum& = -2147467259 Then
    MsgBox "The changes were not successful because it would create duplicate values"
    Exit Function
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EML Error", "HREMP", "UPDATE")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        RollBack
        Resume Next
    Else
        Unload Me
    End If
End If
End Function

Function CheckSINSSN(xEmpnbr, xSINSNN, TypeFlag)
Dim RsSIN As New ADODB.Recordset
Dim SQLQ
    CheckSINSSN = False
    SQLQ = "SELECT ED_EMPNBR,ED_SIN,ED_SSN,ED_SURNAME,ED_FNAME FROM HREMP "
    If TypeFlag = "SIN" Then
        SQLQ = SQLQ & "WHERE ED_SIN = '" & xSINSNN & "' "
    Else
        SQLQ = SQLQ & "WHERE ED_SSN = '" & xSINSNN & "' "
    End If
    SQLQ = SQLQ & "And ED_EMPNBR <> " & xEmpnbr
    RsSIN.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not RsSIN.EOF Then
        CheckSINSSN = True
    End If
    Dim xEmpList
    xEmpList = ""
    Do Until RsSIN.EOF
        xEmpList = xEmpList & RsSIN("ED_EMPNBR") & " - " & RsSIN("ED_SURNAME") & ", " & RsSIN("ED_FNAME") & vbNewLine
        RsSIN.MoveNext
    Loop
    If TypeFlag = "SIN" Then
        fDupSIN = xEmpList
    Else
        fDupSSN = xEmpList
    End If
        
    RsSIN.Close
    
    flgDupSINSSN_Term = False
    If glbCompSerial = "S/N - 2350W" Or glbCompSerial = "S/N - 2382W" Then
        'Listowel - Ticket #14040
        'Ticket #19937 2382W for Samuel -  Franks 05/06/2011
        SQLQ = "SELECT ED_EMPNBR,ED_SIN,ED_SSN,ED_SURNAME,ED_FNAME FROM TERM_HREMP "
        If TypeFlag = "SIN" Then
            SQLQ = SQLQ & "WHERE ED_SIN = '" & xSINSNN & "' "
        Else
            SQLQ = SQLQ & "WHERE ED_SSN = '" & xSINSNN & "' "
        End If
        SQLQ = SQLQ & "And ED_EMPNBR <> " & xEmpnbr
        RsSIN.Open SQLQ, gdbAdoIhr001X, adOpenStatic
        If Not RsSIN.EOF Then
            CheckSINSSN = True
            flgDupSINSSN_Term = True
        End If
        Dim xTerm_Emplist
        xTerm_Emplist = ""
        
        If Not RsSIN.EOF Then
            xTerm_Emplist = "Terminated Employee(s):" & vbNewLine
        End If
        
        Do Until RsSIN.EOF
            xTerm_Emplist = xTerm_Emplist & RsSIN("ED_EMPNBR") & " - " & RsSIN("ED_SURNAME") & ", " & RsSIN("ED_FNAME") & vbNewLine
            RsSIN.MoveNext
        Loop
        If TypeFlag = "SIN" Then
            fDupSIN_Term = xTerm_Emplist
        Else
            fDupSSN_Term = xTerm_Emplist
        End If
            
        RsSIN.Close
    
    End If
    
End Function

Public Sub Display_Value()
Dim rsDAT_Other As New ADODB.Recordset
Dim SQLQ As String

Call Set_Control("R", Me, rsDATA)
If Not rsDATA.EOF Then
    'Commented out by Bryan 23/Sep/05 Ticket#9362
'    If glbCompSerial = "S/N - 2259W" Then  'For County of Oxford    Ticket # 9119
'        If (Not IsNull(rsDATA!ED_LTHIRE)) Then
'            lblTitle(15).Caption = lStr("Last Hire Date") & ":"
'            lblDOH = rsDATA!ED_LTHIRE
'        Else
'            lblTitle(15).Caption = lStr("Original Hire Date") & ":"
'            If Not IsNull(rsDATA!ED_DOH) Then lblDOH = rsDATA!ED_DOH
'        End If
'    Else
        lbltitle(15).Caption = lStr("Original Hire Date") '& ":"
        If Not IsNull(rsDATA!ED_DOH) Then lblDOH = rsDATA!ED_DOH
'    End If
    If glbWFC Then
        If Not IsNull(rsDATA!ED_CANDIDATE) Then txtCandidate.Text = rsDATA!ED_CANDIDATE Else txtCandidate.Text = ""
        'Ticket #28637 Franks 05/18/2016 - begin
        If rsDAT_Other.State <> 0 Then rsDAT_Other.Close
        If glbtermopen Then
            SQLQ = "SELECT * FROM Term_HREMP_OTHER"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "SELECT * FROM HREMP_OTHER"
            SQLQ = SQLQ & " where ER_EMPNBR = " & glbLEE_ID
        End If
        rsDAT_Other.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsDAT_Other.EOF Then
            If Not IsNull(rsDAT_Other("ER_NETWORKLOGIN")) Then medNetworkLogin.Text = rsDAT_Other("ER_NETWORKLOGIN") Else medNetworkLogin.Text = ""
            If Not IsNull(rsDAT_Other("ER_VENDORNO")) Then medVendorNo.Text = rsDAT_Other("ER_VENDORNO") Else medVendorNo.Text = ""
        Else
            medNetworkLogin.Text = ""
            medVendorNo.Text = ""
        End If
        rsDAT_Other.Close
        'Ticket #28637 Franks 05/18/2016 - end
        
        'If clpCode(4).Text = "MISS" Or clpCode(4).Text = "TROY" Then
        If (clpCode(4).Text = "MISS" Or clpCode(4).Text = "TROY") And Not glbtermopen Then 'Ticket #29836 Franks 02/24/2017 - not for term
            lbltitle(55).FontBold = True
            lbltitle(56).FontBold = True
        Else
            lbltitle(55).FontBold = False
            lbltitle(56).FontBold = False
        End If
    End If
End If
Call SET_UP_MODE
'Hemu - 11/21/2003 Begin - For the first time it prompts to associate G/L with Dept
'                  even when the Dept. Code has not changed, this was because the
'                   RDept value was empty for the first time
RDept = clpDept.Text
Call cmdModify_Click
'Hemu - 11/21/2003

If glbCompSerial = "S/N - 2420W" Then 'Macaulay Ticket #25016 Franks 04/01/2014
    Call DisAltPayrollIDs
End If

End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNewEE Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fglbNewEE = True
End Property

Public Property Get RelateMode() As RelateModeEnum
'If glbtermopen Then
'    RelateMode = RelateTermEmp
'Else
    RelateMode = RelateEMP
'End If
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Basic
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = True
End Property

Public Property Get Deleteble() As Boolean
Deleteble = True
End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
'''On Error GoTo Set_Up_Mode_Err
If fglbNewEE Then
    UpdateState = NewRecord
    TF = True
ElseIf rsDATA.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)

If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/20/2014
    TF = getFamilyDayUpdateRight(UpdateRight, glbLEE_ID)
Else
    If Not UpdateRight Then TF = False
End If

Call ST_UPD_MODE(TF)
Set_Up_Mode_Err:
Exit Sub
End Sub

Private Function VadimControl(Action)
Dim ctlName As Control
Dim lblName As Label
Dim X
VadimControl = False
For X = 1 To 4
    If Vadim_PayType_TABLName = clpCode(X).TablName Then
        Set lblName = lbltitle(22 + X)
        Set ctlName = clpCode(X)
    End If
Next
If lblName Is Nothing Or ctlName Is Nothing Then VadimControl = True: Exit Function
If Action = "Show" Then
    lblName.FontBold = True
ElseIf Action = "Check" Then
    If Len(ctlName) = 0 Then
        'Ticket #24164 - Re-ordering
        'By default as when the ctlName is blank the .Container is not working so since all the Vadim fields are
        'currently on Organizational tab, setting that at default. This helps to avoid error 5.
        tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        
        If ctlName.Container = "frPersonal" Then
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        ElseIf ctlName.Container = "frOrganizational" Then
            tbDemographics.SelectedItem = tbDemographics.Tabs(2)
        ElseIf ctlName.Container = "frMiscellaneous" Then
            tbDemographics.SelectedItem = tbDemographics.Tabs(3)
        End If

        MsgBox lStr(lblName & " is a required field.")
        ctlName.SetFocus
        Exit Function
    End If
End If
VadimControl = True


End Function

Sub UpdateCurrentPosition()
Dim rsEmpJob As New ADODB.Recordset
If Not glbVadim Then Exit Sub
rsEmpJob.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & glbLEE_ID & " AND JH_PAYROLL_ID='" & oPayrollID & "'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
If Not rsEmpJob.EOF Then
    rsEmpJob("JH_DIV") = clpDiv
    rsEmpJob("JH_DEPTNO") = clpDept
    rsEmpJob("JH_GLNO") = clpGLNum
    rsEmpJob("JH_SECTION") = clpCode(4)
    rsEmpJob("JH_PAYROLL_ID") = txtPayrollID
    If glbCompSerial = "S/N - 2363W" Then 'CITY OF K LAKES
        rsEmpJob("JH_PAYROLL_CATEGORY") = clpCode(2)
    End If
    If glbCompSerial = "S/N - 2362W" Then 'CITY OF SARNIA
        rsEmpJob("JH_PAYROLL_CATEGORY") = clpDiv
    End If
    rsEmpJob.Update
End If
rsEmpJob.Close
End Sub

Private Sub UpdateCurrentPosition_KerrysPlace()
Dim rsEmpJob As New ADODB.Recordset
Dim SQLQ As String

SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & glbLEE_ID
'SQLQ = SQLQ & " AND JH_PAYROLL_ID='" & oPayrollID & "'"
rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
Do While Not rsEmpJob.EOF
    'Matching the original value then only change
    If rsEmpJob("JH_DIV") = SavDiv Then
        rsEmpJob("JH_DIV") = clpDiv.Text
    End If
    'Matching the original value then only change
    If rsEmpJob("JH_DEPTNO") = SavDept Then
        rsEmpJob("JH_DEPTNO") = clpDept.Text
    End If
    'rsEmpJob("JH_GLNO") = clpGLNum
    'rsEmpJob("JH_SECTION") = clpCode(4)
    'rsEmpJob("JH_PAYROLL_ID") = txtPayrollID
    rsEmpJob.Update
    
    rsEmpJob.MoveNext
Loop
rsEmpJob.Close
Set rsEmpJob = Nothing

End Sub

Private Function Check_EMPLOYEE_Number(mEmpNum)
Dim SQLQ As String, countr As Integer, SQLQ2 As String
Dim Desc As String
Dim blnFirst As Boolean
Dim DMaxNum, DMaxNumX
Dim snapEmp As New ADODB.Recordset

'''On Error GoTo Emp_Err

blnFirst = True
DMaxNum = 0
SQLQ = "Select ED_EMPNBR,ED_SurName,ED_FName from qry_HREMP "
SQLQ = SQLQ & "where ED_EMPNBR >=" & mEmpNum
SQLQ = SQLQ & " order by ED_EMPNBR "
snapEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic

If snapEmp.EOF Then
    DMaxNum = mEmpNum
Else
Do While Not snapEmp.EOF
    If snapEmp("ED_EMPNBR") <> mEmpNum Then
        If snapEmp("ED_EMPNBR") > mEmpNum Then
            DMaxNum = mEmpNum
            Exit Do
        End If
    Else
        If blnFirst Then
            blnFirst = False
            Check_EMPLOYEE_Number = mEmpNum & " assigned to " & snapEmp("ED_SURNAME") & "," & snapEmp("ED_FNAME")
        End If
        mEmpNum = mEmpNum + 1
    
    End If
    snapEmp.MoveNext
Loop
End If
If DMaxNum = 0 Then DMaxNum = mEmpNum

'Data1.DatabaseName = glbIHRDB   'laura nov 28, 1997
Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "HRPARCO"
 
SQLQ2 = "Select HRPARCO.* from HRPARCO"

Data1.RecordSource = SQLQ2
Data1.Refresh

'Data1.Recordset.Edit
Data1.Recordset("PC_NEXT_AVAILABLE_NBR") = DMaxNum
Data1.Recordset.UpdateBatch
Data1.Refresh

Exit Function

Emp_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Employees", "HREMP", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Function getProductLineCodeforLinamar(xOrgCode)
    Dim rsTABL As New ADODB.Recordset
    Dim xNewCode
    xNewCode = xOrgCode
    rsTABL.Open "SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDRG' AND TB_KEY='" & xOrgCode & "'", gdbAdoIhr001, adOpenForwardOnly
    If rsTABL.EOF Or rsTABL.BOF Then
        xNewCode = "ALL" & Mid(xOrgCode, 4)
    End If
    getProductLineCodeforLinamar = xNewCode
End Function

Function GetBonusRptDesc(TablKey)
    Dim rsTABL As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT * FROM WFC_Bonus_Loc_Department WHERE Dept_No = '" & TablKey & "' "
    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsTABL.EOF And rsTABL.BOF Then
        GetBonusRptDesc = ""
    Else
        GetBonusRptDesc = rsTABL("Dept_Name")
    End If
    rsTABL.Close
End Function

Public Sub imgEmail_Click()
Dim xEmail, xToEmail
'''On Error GoTo Email_Err
        If Not UserEmailExist Then
            Exit Sub
        End If
        'xEmail = GetCurEmpEmail
        If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
            xToEmail = GetComPreferEmail("EMAIL_ONNEWHIRE", lblEENum.Caption) 'glbLEE_ID)
            If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                xToEmail = GetComPreferEmail("EMAIL_ONNEWHIRE")
            End If
        Else
            'Ticket #20317 - More Emails for everyone
            xToEmail = GetComPreferEmail("EMAIL_ONNEWHIRE", glbLEE_ID)
            If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                xToEmail = GetComPreferEmail("EMAIL_ONNEWHIRE")
            End If
        End If
        If Len(xToEmail) > 0 Then
            frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONNEWHIRE")
            'frmSendEmail.txtCC.Text = xEmail
            'frmSendEmail.txtSubject.Text = "info:HR Employee New Hire Notice"
            'Ticket #18578
            frmSendEmail.txtSubject.Text = "info:HR Employee New Hire Notice - " & lblEEName.Caption
            frmSendEmail.txtBody.Text = MailBody
            frmSendEmail.Show 1
        End If
        'Else
        '    If Len(glbLEE_SName) = 0 Then
        '        MsgBox "There is no email on Status/Dates screen for employee. "
        '    Else
        '        MsgBox "There is no email on Status/Dates screen for employee " & glbLEE_SName & ", " & glbLEE_FName & ". "
        '    End If
        'End If

    Exit Sub

Email_Err:
    If Err.Number = 364 Then
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail", "SMTP", "SENDEMAIL")
    Resume Next

End Sub

Public Sub imgEmail_OtherChanges_Click()
Dim xEmail, xToEmail
'''On Error GoTo OtherChanges_Err
    If Not UserEmailExist Then
        Exit Sub
    End If
    'xEmail = GetCurEmpEmail
    
    'Ticket #20317 - More Emails for everyone
    xToEmail = GetComPreferEmail("EMAIL_ONNEWHIRE", glbLEE_ID)
    If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
        xToEmail = GetComPreferEmail("EMAIL_ONNEWHIRE")
    End If
    'If Len(xToEmail) > 0 Then
        frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONNEWHIRE")
        'frmSendEmail.txtCC.Text = xEmail
        frmSendEmail.txtSubject.Text = "info:HR Employee Information Change Notice"
        frmSendEmail.txtBody.Text = MailBody
        frmSendEmail.Show 1
    'End If
Exit Sub

OtherChanges_Err:
    If Err.Number = 364 Then
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Employee Change EMail", "SMTP", "SENDEMAIL")
    Resume Next

End Sub

Private Function chk_EMPNBR() As Boolean
'''On Error GoTo Eh

    chk_EMPNBR = False
    
    If glbSysGen = True Then
        Dim rs As New ADODB.Recordset
        Dim SQLQ As String
        Dim strMsg As String
        Dim retval As Long
        
        SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
        rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
        If rs.EOF = False And rs.BOF = False Then
            'Ticket #24164 - Re-ordering
            tbDemographics.SelectedItem = tbDemographics.Tabs(1)
        
            strMsg = "This Employee Number is already assigned to " & rs("ED_SURNAME") & "," & rs("ED_FNAME") & vbCrLf
            strMsg = strMsg & "Do you want to assign a new number?"
            retval = MsgBox(strMsg, vbQuestion + vbYesNo, "Duplicate Employee Number")
            If retval = vbYes Then
                glbLEE_ID = CLng(glbNextEmpl)
                lblEEID = glbLEE_ID
                Dim rsPA As New ADODB.Recordset
                rsPA.Open "select PC_NEXT_AVAILABLE_NBR,PC_FEDTAX,PC_PROVTAX from HRPARCO", gdbAdoIhr001, adOpenStatic, adLockPessimistic
                glbNextEmpl = glbNextEmpl + 1
                If glbCompSerial = "S/N - 2241W" Then '  Granite Club
                    Call Check_EMPLOYEE_Number(glbNextEmpl)
                Else
                    rsPA("PC_NEXT_AVAILABLE_NBR") = glbNextEmpl
                    rsPA.Update
                End If
                rsPA.Close
            ElseIf retval = vbNo Then
                Exit Function
            End If
        End If
        rs.Close
    End If
    
    chk_EMPNBR = True
    
exH:
    Exit Function
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chk_EMPNBR", "HREMP", "SELECT")
    Resume exH
   
End Function

Private Sub No_Security_Rights_on_Forms()
    Dim rsTN As New ADODB.Recordset, X
        
    strNoAccessForms = ""
    rsTN.Open "SELECT * FROM HRNEWHIRE WHERE NewHire<>0 ORDER BY ID", glbAdoIHRDB
    Do While Not rsTN.EOF
        If rsTN!MenuItem = "Demograghics" Then
            If Not gSec_Inq_Basic Then
                strNoAccessForms = strNoAccessForms & "Demograghics"
            End If
        End If
        If rsTN!MenuItem = "Status/Dates" Then
            If Not gSec_Inq_Basic Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Status/Dates"
            End If
        End If
        If rsTN!MenuItem = "Contacts" Then
            If Not gSec_Inq_Basic Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Contacts"
            End If
        End If
        If rsTN!MenuItem = "Dependents" Then
            If Not gSec_Inq_Dependents Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Dependents"
            End If
        End If
        If rsTN!MenuItem = "G/L Distribution" Then
            If Not gSec_Inq_GLDist Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "G/L Distribution"
            End If
        End If
        If rsTN!MenuItem = "Employee Flags" Then
            If Not gSec_Inq_EMP_FLAGS Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Employee Flags"
            End If
        End If
        If rsTN!MenuItem = "Payroll/Banking" Then
            If Not gSec_Inq_Banking Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Payroll/Banking"
            End If
        End If
        If rsTN!MenuItem = "Other Information" Then
            If Not gSec_Inq_OtherInformation Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Other Information"
            End If
        End If
        If rsTN!MenuItem = "Employee ADP Data" Then
            If Not gSec_Inq_ADP_Data Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Employee ADP Data"
            End If
        End If
        If rsTN!MenuItem = "Skills" Then
            If Not gSec_Inq_Skills Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Skills"
            End If
        End If
        If rsTN!MenuItem = "Formal Education" Then
            If Not gSec_Inq_Formal_Education Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Formal Education"
            End If
        End If
        If rsTN!MenuItem = "Education/Seminars" Then
            If Not gSec_Inq_Education_Seminars Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Education/Seminars"
            End If
        End If
        If rsTN!MenuItem = "Associations" Then
            If Not gSec_Inq_Associations Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Associations"
            End If
        End If
        If rsTN!MenuItem = "Languages" Then
            If Not gSec_Inq_EMP_LANG Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Languages"
            End If
        End If
        If rsTN!MenuItem = "Succession Planning" Then
            If Not gSec_Inq_SUCCESSION Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Succession Planning"
            End If
        End If
        If rsTN!MenuItem = "Benefits/Beneficiary" Then
            If Not gSec_Inq_Benefits Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Benefits/Beneficiary"
            End If
        End If
        If rsTN!MenuItem = "Dollar Entitlement" Then
            If Not gSec_Inq_Other_Entitlements Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Dollar Entitlement"
            End If
        End If
        If rsTN!MenuItem = "Other Earnings" Then
            If Not gSec_Inq_Earnings Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Other Earnings"
            End If
        End If
        If rsTN!MenuItem = "Position" Then
            If Not gSec_Inq_Position Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Position"
            End If
        End If
        If rsTN!MenuItem = "Salary" Then
            If Not gSec_Inq_Salary Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Salary"
            End If
        End If
        If rsTN!MenuItem = "Performance" Then
            If Not gSec_Inq_Performance Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Performance"
            End If
        End If
        If rsTN!MenuItem = "Attendance" Then
            If Not gSec_Inq_Attendance Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Attendance"
            End If
        End If
        If rsTN!MenuItem = "Vacation and Sick Entitlements" Then
            If Not gSec_Inq_Entitlements Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Vacation and Sick Entitlements"
            End If
        End If
        If rsTN!MenuItem = "Hourly Entitlements" Then
            If Not gSec_Inq_Hrly_Entitlements Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Hourly Entitlements"
            End If
        End If
        If rsTN!MenuItem = "Incident Data" Then
            If Not gSec_Inq_Health_Safety Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Incident Data"
            End If
        End If
        If rsTN!MenuItem = "Injury/Location" Then
            If Not gSec_Inq_Health_Safety Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Injury/Location"
            End If
        End If
        If rsTN!MenuItem = "Root Cause" Then
            If Not gSec_Inq_Health_Safety Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Root Cause"
            End If
        End If
        If rsTN!MenuItem = "Corrective Action" Then
            If Not gSec_Inq_Health_Safety Then
                strNoAccessForms = strNoAccessForms & "Corrective Action"
            End If
        End If
        If rsTN!MenuItem = "Claims/Medical Information" Then
            If Not gSec_Inq_Health_Safety Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Claims/Medical Information"
            End If
        End If
        If rsTN!MenuItem = "Contact Information" Then
            If Not gSec_Inq_Health_Safety Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Contact Information"
            End If
        End If
        If rsTN!MenuItem = "WSIB Cost Information" Then
            If Not gSec_Inq_Health_Safety Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "WSIB Cost Information"
            End If
        End If
        If rsTN!MenuItem = "Comments" Then
            If Not gSec_Inq_Comments Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Comments"
            End If
        End If
        If rsTN!MenuItem = "Employee Follow-up" Then
            If Not gSec_Inq_Follow_Ups Then
                strNoAccessForms = strNoAccessForms & IIf(Len(strNoAccessForms) <> 0, ", ", "") & "Employee Follow-up"
            End If
        End If
        
        rsTN.MoveNext
    Loop
    rsTN.Close
End Sub

Private Sub UPDOvertime_Overview()
Dim rsHREmp As New ADODB.Recordset
Dim rsOvtMst As New ADODB.Recordset
Dim rsOvtEmp As New ADODB.Recordset
Dim SQLQ As String
Dim flgUpdated As Boolean

'''On Error GoTo Err_UPDOvertime_Overview

    'Ticket #22847- To see if employee record updated, and if employee needs to be in the Overtime Bank master based on the rule
    flgUpdated = False

    SQLQ = "SELECT * FROM HR_OVERTIME_MASTER ORDER BY OM_ORG"
    rsOvtMst.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsOvtMst.EOF
        Set rsHREmp = Nothing
        SQLQ = "SELECT ED_EMPNBR, ED_EMP,ED_PT,ED_ORG,ED_LOC,ED_REGION,ED_ADMINBY,ED_SECTION FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID & " "
        'SQLQ = SQLQ & " AND ED_ORG = '" & rsOvtMst("OM_ORG") & "'"
        rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsHREmp.EOF Then
            'Ticket #15753 - from OM_LOC to OM_SECTION
            If ((IsNull(rsOvtMst("OM_EMP")) Or rsOvtMst("OM_EMP") = "") Or (rsHREmp("ED_EMP") = rsOvtMst("OM_EMP"))) And _
                ((IsNull(rsOvtMst("OM_PT")) Or rsOvtMst("OM_PT") = "") Or (rsHREmp("ED_PT") = rsOvtMst("OM_PT"))) And _
                ((IsNull(rsOvtMst("OM_LOC")) Or rsOvtMst("OM_LOC") = "") Or (rsHREmp("ED_LOC") = rsOvtMst("OM_LOC"))) And _
                ((IsNull(rsOvtMst("OM_REGION")) Or rsOvtMst("OM_REGION") = "") Or (rsHREmp("ED_REGION") = rsOvtMst("OM_REGION"))) And _
                ((IsNull(rsOvtMst("OM_ADMINBY")) Or rsOvtMst("OM_ADMINBY") = "") Or (rsHREmp("ED_ADMINBY") = rsOvtMst("OM_ADMINBY"))) And _
                ((IsNull(rsOvtMst("OM_SECTION")) Or rsOvtMst("OM_SECTION") = "") Or (rsHREmp("ED_SECTION") = rsOvtMst("OM_SECTION"))) And _
                ((IsNull(rsOvtMst("OM_ORG")) Or rsOvtMst("OM_ORG") = "") Or (rsHREmp("ED_ORG") = rsOvtMst("OM_ORG"))) Then
                
                    Set rsOvtEmp = Nothing
                    rsOvtEmp.Open "SELECT * FROM HR_OVERTIME_BANK WHERE OT_EMPNBR=" & rsHREmp("ED_EMPNBR"), gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If rsOvtEmp.EOF Then
                        rsOvtEmp.AddNew
                        rsOvtEmp("OT_PBANK") = 0
                    End If
                    rsOvtEmp("OT_COMPNO") = "001"
                    rsOvtEmp("OT_EMPNBR") = rsHREmp("ED_EMPNBR")
                    rsOvtEmp("OT_BANK") = Get_OvertimeBank(rsHREmp("ED_EMPNBR"), rsOvtMst("OM_EFDATE"), rsOvtMst("OM_ETDATE")) * Val(rsOvtMst("OM_MULTIPLIER"))
                    rsOvtEmp("OT_BANKT") = Get_OvertimeTaken(rsHREmp("ED_EMPNBR"), rsOvtMst("OM_EFDATE"), rsOvtMst("OM_ETDATE"))
                    rsOvtEmp("OT_MBANK") = rsOvtMst("OM_MAX_BANK_HRS")
                    rsOvtEmp("OT_EFDATE") = rsOvtMst("OM_EFDATE")    'Format("1/1/" & Year(Now()), "mm/dd/yyyy")
                    rsOvtEmp("OT_ETDATE") = rsOvtMst("OM_ETDATE")    'Format("12/31/" & Year(Now()), "mm/dd/yyyy")
                    rsOvtEmp("OT_LDATE") = Date
                    rsOvtEmp("OT_LTIME") = Time$
                    rsOvtEmp("OT_LUSER") = glbUserID
                    rsOvtEmp.Update
                    
                    rsOvtEmp.Close
                    
                    'Ticket #22847- Employee Overtime Bank record updated - employee part of the rule.
                    flgUpdated = True
                    
                    Exit Do
            End If
        End If
        rsHREmp.Close
        
        rsOvtMst.MoveNext
    Loop
    rsOvtMst.Close
    
    'Ticket #22847 - Check if Employee was part of the Overtime Bank Master rule
    If flgUpdated = False Then
        'Employee was not part of the Overtime Bank Master rule.
        'Delete the Overtime Bank record if it exists for this employee.
        gdbAdoIhr001.Execute "DELETE HR_OVERTIME_BANK WHERE OT_EMPNBR = " & glbLEE_ID
    End If
    
    Exit Sub

Err_UPDOvertime_Overview:
glbFrmCaption$ = "Overtime Bank Overview update on new hire"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UPUPDOvertime_Overview", "", "EML")
If gintRollBack% = False Then
    Resume Next
End If

End Sub
Private Sub Update_Overtime_Bank(xCode)
Dim rsOvtEmp As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim SQLQ As String

'Recalculate the Overtime Bank
SQLQ = "SELECT * FROM HR_OVERTIME_BANK WHERE OT_EMPNBR = " & glbLEE_ID
rsOvtEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not rsOvtEmp.EOF Then
    'If xUnion = "" Then
        'Delete the record from Overtime Bank
    '    gdbAdoIhr001.Execute "DELETE HR_OVERTIME_BANK WHERE OT_EMPNBR = " & glbLEE_ID
    'Else
        Call ReCalcOvt("OT_EMPNBR = " & glbLEE_ID)
    'End If
End If
rsOvtEmp.Close

End Sub

Private Sub Update_Position_Hours_DWP(xDHrs, xWhrs, xPHrs)
    Dim rsEmpJob As New ADODB.Recordset
    
    rsEmpJob.Open "SELECT JH_EMPNBR,JH_DHRS,JH_WHRS,JH_PHRS,JH_LDATE,JH_LUSER,JH_LTIME FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmpJob.EOF Then
        rsEmpJob("JH_DHRS") = xDHrs
        rsEmpJob("JH_WHRS") = xWhrs
        rsEmpJob("JH_PHRS") = xPHrs
        rsEmpJob("JH_LDATE") = Date
        rsEmpJob("JH_LUSER") = glbUserID
        rsEmpJob("JH_LTIME") = Time$
    
        rsEmpJob.Update
    End If
    rsEmpJob.Close
    Set rsEmpJob = Nothing
End Sub

Private Sub WFC_NGS_SmokerUptWithEmail() 'Ticket #22409 Frank 08/08/2012
Dim rsEmpee As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim strFields As String
Dim SQLQ As String
Dim xUptFlag As Boolean
Dim xNGSStart
Dim xDate1, xDate2
Dim xLDate
Dim a As Integer, Msg As String
Dim Title$, DgDef, Response%
Dim xYear, xMonth
Dim xEmail
Dim xToEmail As String

If (OSMOKER = ComSmoker.Text) Then
    Exit Sub
End If

''''On Error GoTo AUDIT_ERR
If Not glbNGS_OnFlag Then
    Exit Sub
End If
If NewHireForms.count > 0 Then
    Exit Sub 'modificaiton olnly
End If
If glbtermopen Then Exit Sub

glbEmpDiv = clpDiv.Text
SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2 FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic 'ED_VADIM1
If rsEmpee.EOF Then
    Exit Sub
Else
    'If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
    If IsNull(rsEmpee("ED_ORG")) Then glbUNION = "" Else glbUNION = rsEmpee("ED_ORG")
    If IsNull(rsEmpee("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = rsEmpee("ED_VADIM1")
    If IsNull(rsEmpee("ED_VADIM2")) Then glbWFCPayGroup = "" Else glbWFCPayGroup = rsEmpee("ED_VADIM2")
End If
rsEmpee.Close
'No NGS Sub Group, skip
If Len(glbWFCNGSSubGroup) = 0 Then Exit Sub

xLDate = Date

xNGSStart = ""
SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1 FROM HREMP_OTHER WHERE ER_EMPNBR = " & glbLEE_ID & ""
rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmpOther.EOF Then
    If IsDate(rsEmpOther("ER_OTHERDATE1")) Then
        xNGSStart = rsEmpOther("ER_OTHERDATE1")
    End If
End If
rsEmpOther.Close

'NGS field changes --------------------------------------
'OSMOKER, ComSmoker
If Not (OSMOKER = ComSmoker.Text) Then
    'Ticket #24164 - Re-ordering
    tbDemographics.SelectedItem = tbDemographics.Tabs(3)

    'Msg$ = "You are changing the Smoker status from '" & OSMOKER & "' to '" & ComSmoker.Text & "'. "
    'Msg$ = Msg$ & "If you continue with the save, the system will send an email to you and the emails addresses setup under Company Preferences. "
    'Msg$ = Msg$ & "Please enter a reason for the change."
    'Msg$ = Msg$ & Chr(10) & "Do you wish to continue?"
    'Ticket #23317 Franks 02/27/2013
    Msg$ = "You are changing the Smoker status from '" & OSMOKER & "' to '" & ComSmoker.Text & "'. "
    Msg$ = Msg$ & "Please enter a Reason and click on Yes to make the Smoker Status change."
    Msg$ = Msg$ & "To exit without changing the Smoker Status, click on No."
    
    glbChgTermReason = ""
    glbSpouseSIN = ""
    
    frmMsgTerm.PenTermDate = "WFC_SomkerChgReason"
    frmMsgTerm.lblNote1.Caption = Msg$
    frmMsgTerm.Show 1
    If glbSpouseSIN = "N" Then
        ComSmoker.Text = OSMOKER 'change it back
    End If
    If glbSpouseSIN = "Y" Then 'update and send email
        If gsEMAIL_ONPERFORMANCE Then
            'If Not UserEmailExist Then
            '    Exit Sub
            'End If
            'xEmail = GetCurUserEmail

            'Ticket #20317 - Send email to More Emails list as well.
            xToEmail = GetComPreferEmail("EMAIL_ONPERFORMANCE", glbLEE_ID)
            If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                xToEmail = GetComPreferEmail("EMAIL_ONPERFORMANCE")
            End If
            frmSendEmail.txtTo.Text = xToEmail
            frmSendEmail.txtCC.Text = GetCurUserEmail

            frmSendEmail.txtSubject.Text = "info:HR Smoker Status Change Notice - " & txtSurname.Text & ", " & txtFName.Text
            MailBody = "The Smoker status has been changed from '" & OSMOKER & "' to '" & ComSmoker.Text & "'. "
            MailBody = MailBody & vbCrLf & "The reason is " & glbChgTermReason
            frmSendEmail.txtBody.Text = MailBody
            
            MDIMain.panHelp(0).FloodType = 0
            MDIMain.panHelp(0).Caption = "Sending email..."
            frmSendEmail.Tag = ""
            frmSendEmail.cmdSend_Click
            Do
                DoEvents
            Loop Until frmSendEmail.Tag <> ""   ' MC - dkostka - 05/03/01 - Changed from = "DONE" to <> ""
            ' AC - dkostka - 05/03/01 - Added checking to make sure the email went through,
            '   otherwise refuse to terminate the employee.
            If frmSendEmail.Tag = "DONE" Then
                Unload frmSendEmail
                AbortTerm = False
            Else
                Unload frmSendEmail
                AbortTerm = True
            End If
            MDIMain.panHelp(0).Caption = ""
            MDIMain.panHelp(0).FloodType = 1
        
        End If
    End If
    glbChgTermReason = ""
    glbSpouseSIN = ""

End If

End Sub

Private Sub WFC_SmokerChange()
'Ticket #23301 Franks 02/20/2013
'"   If the SMOKER changes, write <<today's date>> + " " + <<user id> + " Change" in the Date/Smoker Affidavit field. Ie: 02/19/13 3142 Change.
Dim rsEmpee As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim strFields As String
Dim SQLQ As String
Dim xUptFlag As Boolean
Dim xNGSStart
Dim xDate1, xDate2
Dim xLDate
Dim a As Integer, Msg As String
Dim Title$, DgDef, Response%
Dim xYear, xMonth
Dim xMsg As String

If (OSMOKER = ComSmoker.Text) Then
    Exit Sub
End If

If NewHireForms.count > 0 Then
    Exit Sub 'modificaiton olnly
End If
If glbtermopen Then Exit Sub

If Not (OSMOKER = ComSmoker.Text) Then
    'Ticket #23317 Franks 02/27/2013
    xMsg = "User Change on " & Format(Date, "mm-dd-yy")
    medTYPEVEHICLE.Text = Left(xMsg, 30)
    xMsg = "User ID - " & glbUserID & ""
    medPARKPERMIT2.Text = Left(xMsg, 30)
End If

End Sub

Private Sub WFC_NGS_SmokerUpdate()
Dim rsEmpee As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim strFields As String
Dim SQLQ As String
Dim xUptFlag As Boolean
Dim xNGSStart
Dim xDate1, xDate2
Dim xLDate
Dim a As Integer, Msg As String
Dim Title$, DgDef, Response%
Dim xYear, xMonth

If (OSMOKER = ComSmoker.Text) Then
    Exit Sub
End If

''''On Error GoTo AUDIT_ERR
If Not glbNGS_OnFlag Then
    Exit Sub
End If
If NewHireForms.count > 0 Then
    Exit Sub 'modificaiton olnly
End If
If glbtermopen Then Exit Sub

glbEmpDiv = clpDiv.Text
SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2 FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic 'ED_VADIM1
If rsEmpee.EOF Then
    Exit Sub
Else
    'If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
    If IsNull(rsEmpee("ED_ORG")) Then glbUNION = "" Else glbUNION = rsEmpee("ED_ORG")
    If IsNull(rsEmpee("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = rsEmpee("ED_VADIM1")
    If IsNull(rsEmpee("ED_VADIM2")) Then glbWFCPayGroup = "" Else glbWFCPayGroup = rsEmpee("ED_VADIM2")
End If
rsEmpee.Close
'No NGS Sub Group, skip
If Len(glbWFCNGSSubGroup) = 0 Then Exit Sub

xLDate = Date

xNGSStart = ""
SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1 FROM HREMP_OTHER WHERE ER_EMPNBR = " & glbLEE_ID & ""
rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmpOther.EOF Then
    If IsDate(rsEmpOther("ER_OTHERDATE1")) Then
        xNGSStart = rsEmpOther("ER_OTHERDATE1")
    End If
End If
rsEmpOther.Close
''No NGS Effective Date, skip
'If Len(xNGSStart) = 0 Then Exit Sub
''If the Effective date is later than today, then LDate = Effective date
'If IsDate(xNGSStart) Then
'    If CVDate((xNGSStart)) > CVDate(Date) Then
'        xLDate = CVDate(xNGSStart)
'    End If
'End If

'NGS field changes --------------------------------------
'OSMOKER, ComSmoker
If Not (OSMOKER = ComSmoker.Text) Then
    If ComSmoker.Text = "No" Then
        'Ticket #24164 - Re-ordering
        tbDemographics.SelectedItem = tbDemographics.Tabs(3)
    
        'If Yes, update "Type of Vehicle" with a 4 digit year. If the system month is between 1 and 10,
        'use the current year of the system date. If the system month is 11 or 12,
        'the year is the next year.
        Msg$ = "Do you have a signed Smoker Affidavit on file?"
        Title$ = "Smoker Status change"
        DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
        Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
        If Response% = IDNO Then    ' Evaluate response
            Exit Sub
        End If
        xYear = Year(Date)
        xMonth = month(Date)
        If xMonth = 11 Or xMonth = 12 Then
            xYear = xYear + 1
        End If
        medTYPEVEHICLE.Text = xYear
    End If
    If ComSmoker.Text = "Yes" Then
        'If Smoker changes from No to Yes
        'Clear Type of Vehicle.
        medTYPEVEHICLE.Text = ""
    End If
End If

End Sub

Private Sub AUDIT_SAMUEL_TRANS() 'Ticket #20885 Franks 12/01/2011
Dim rsEmpee As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim strFields As String
Dim SQLQ As String
Dim xUptFlag As Boolean
Dim xNGSStart
Dim xDate1, xDate2
Dim xLDate
Dim xEmpID

'''On Error GoTo AUDIT_ERR

If NewHireForms.count > 0 Then
    Exit Sub 'modificaiton olnly
End If

glbEmpDiv = clpDiv.Text
If glbtermopen Then 'Ticket #20305 Franks 05/17/2011
    SQLQ = "SELECT ED_EMPNBR, ED_ADMINBY, ED_DIV, ED_SECTION, ED_REGION FROM Term_HREMP WHERE TERM_SEQ = " & glbTERM_Seq & " "
Else
    SQLQ = "SELECT ED_EMPNBR, ED_ADMINBY, ED_DIV, ED_SECTION, ED_REGION FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
End If
rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic 'ED_VADIM1
If rsEmpee.EOF Then
    Exit Sub
Else
    'If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
    If IsNull(rsEmpee("ED_ADMINBY")) Then glbEmpAdminBy = "" Else glbEmpAdminBy = rsEmpee("ED_ADMINBY")
    If IsNull(rsEmpee("ED_SECTION")) Then glbEmpSection = "" Else glbEmpSection = rsEmpee("ED_SECTION")
    If IsNull(rsEmpee("ED_REGION")) Then glbEmpRegion = "" Else glbEmpRegion = rsEmpee("ED_REGION")
End If
rsEmpee.Close
''No NGS Sub Group, skip
'If Len(glbWFCNGSSubGroup) = 0 Then Exit Sub

xLDate = Date

If glbtermopen Then
    xEmpID = glbTERM_ID
Else
    xEmpID = glbLEE_ID
End If
'field changes --------------------------------------
If Not (OSection = clpCode(4).Text) Then Call SamuelAuditAdd(xEmpID, "M", "Demographics", lStr("Section"), OSection, clpCode(4).Text, xLDate)
If Not (oRegion = clpCode(2).Text) Then Call SamuelAuditAdd(xEmpID, "M", "Demographics", lStr("Region"), oRegion, clpCode(2).Text, xLDate)

Exit Sub
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT_SAMUEL RECORD", "AUDIT_SAMUEL FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me
    

End Sub

Private Sub AUDIT_GWL_TRANS()
Dim rsEmpee As New ADODB.Recordset
Dim SQLQ As String
Dim xEmpID
Dim xForm As String
Dim xTranType
Dim xChgType
Dim xEDate, xDate1, xDate2
Dim xLDate
Dim xBenGroup

'''On Error GoTo AUDIT_ERR

    If Not glbIsGWL Then Exit Sub
    If NewHireForms.count > 0 Then
        Exit Sub 'modificaiton olnly
    End If

    If glbtermopen Then 'Ticket #20305 Franks 05/17/2011
        SQLQ = "SELECT ED_EMPNBR, ED_BENEFIT_GROUP, ED_DOH, ED_DIV, ED_ORG FROM Term_HREMP WHERE TERM_SEQ = " & glbTERM_Seq & " "
    Else
        SQLQ = "SELECT ED_EMPNBR, ED_BENEFIT_GROUP, ED_DOH, ED_DIV, ED_ORG FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
    End If
    rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic 'ED_VADIM1
    If rsEmpee.EOF Then
        Exit Sub
    Else
        If IsNull(rsEmpee("ED_BENEFIT_GROUP")) Then xBenGroup = "" Else xBenGroup = rsEmpee("ED_BENEFIT_GROUP")
    End If
    'rsEmpee.Close
    'No Benefit Group Code, skip
    If Len(xBenGroup) = 0 Then Exit Sub

    If glbtermopen Then
        xEmpID = glbTERM_ID
    Else
        xEmpID = glbLEE_ID
    End If
    
    If NewHireForms.count > 0 Then
        xTranType = "A"
        xChgType = "New Hire"
    Else
        xTranType = "R"
        xChgType = "Personal Info"
    End If

    xEDate = Date
    xForm = "Demographics"
    xLDate = Date
    If IsDate(rsEmpee("ED_DOH")) Then
        If CVDate(rsEmpee("ED_DOH")) > CVDate(xLDate) Then
            xLDate = rsEmpee("ED_DOH")
        End If
    End If
    rsEmpee.Close
    
    'GWL field changes --------------------------------------
    If Not (OSNAME = txtSurname.Text) Then Call GWLAuditAdd(xEmpID, xTranType, xChgType, xEDate, xForm, "Surname", OSNAME, txtSurname.Text, xLDate)
    If Not (OFNAME = txtFName.Text) Then Call GWLAuditAdd(xEmpID, xTranType, xChgType, xEDate, xForm, "First Name", OFNAME, txtFName.Text, xLDate)
    'Birth Date change - begin
    If IsDate(ODOB) Then xDate1 = CVDate(ODOB) Else xDate1 = ""
    If IsDate(dlpDOB.Text) Then xDate2 = CVDate(dlpDOB.Text) Else xDate2 = ""
    If Not (xDate1 = xDate2) Then Call GWLAuditAdd(xEmpID, xTranType, xChgType, xEDate, xForm, "Birth Date", xDate1, xDate2, xLDate)
    'Birth Date change - end '
    If Not (OSEX = txtGender.Text) Then Call GWLAuditAdd(xEmpID, xTranType, xChgType, xEDate, xForm, "Gender", OSEX, txtGender.Text, xLDate)
    
    If Not (OADD1 = txtAdd1.Text) Then Call GWLAuditAdd(xEmpID, xTranType, xChgType, xEDate, xForm, "Address 1", OADD1, txtAdd1.Text, xLDate)
    If Not (OADD2 = txtAdd2.Text) Then Call GWLAuditAdd(xEmpID, xTranType, xChgType, xEDate, xForm, "Address 2", OADD2, txtAdd2.Text, xLDate)
    If Not (OCITY = txtCity.Text) Then Call GWLAuditAdd(xEmpID, xTranType, xChgType, xEDate, xForm, "City", OCITY, txtCity, xLDate)
    If Not (oProv = clpProv.Text) Then Call GWLAuditAdd(xEmpID, xTranType, xChgType, xEDate, xForm, "Province", oProv, clpProv.Text, xLDate)
    If Not (oCountry = txtCountry.Text) Then Call GWLAuditAdd(xEmpID, xTranType, xChgType, xEDate, xForm, "Country", oCountry, txtCountry.Text, xLDate)
    If Not (OPCODE = medPCode.Text) Then Call GWLAuditAdd(xEmpID, xTranType, xChgType, xEDate, xForm, "Zip Code", OPCODE, medPCode.Text, xLDate)

    Exit Sub

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING GWL AUDIT RECORD", "GWL AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Sub

Private Sub AUDIT_ADP_PAYATWORK_TRAN() 'Ticket #24557 Franks 07/07/2015
Dim rsEmpee As New ADODB.Recordset
Dim SQLQ As String
Dim xEmpID
Dim xForm As String
Dim xTranType
Dim xChgType
Dim xEDate, xDate1, xDate2
Dim xLDate
Dim xBenGroup
Dim xChgFlag As Boolean
Dim xPayrollID, xADPBranch, xADPDept

    If NewHireForms.count > 0 Then Exit Sub 'modificaiton olnly
    If glbtermopen Then Exit Sub 'active employees olnly
    
    xChgFlag = False
    If Not ADPBranchOld(0) = clpCode(1).Text Then xChgFlag = True 'main Branch
    If Not ADPBranchOld(1) = clpCode(11).Text Then xChgFlag = True 'Alt Payroll ID - Branch 1
    If Not ADPBranchOld(2) = clpCode(12).Text Then xChgFlag = True 'Alt Payroll ID - Branch 2
    If Not ADPBranchOld(3) = clpCode(13).Text Then xChgFlag = True 'Alt Payroll ID - Branch 3
    If Not ADPDeptOld(0) = clpSalDist.Text Then xChgFlag = True 'main dept
    If Not ADPDeptOld(1) = clpSalDis2(0).Text Then xChgFlag = True 'Alt Payroll ID - dept 1
    If Not ADPDeptOld(2) = clpSalDis2(1).Text Then xChgFlag = True 'Alt Payroll ID - dept 2
    If Not ADPDeptOld(3) = clpSalDis2(2).Text Then xChgFlag = True 'Alt Payroll ID - dept 3
    If xChgFlag = False Then Exit Sub 'not any change
    
    SQLQ = "SELECT ED_EMPNBR, ED_BENEFIT_GROUP, ED_DOH, ED_DIV, ED_ORG FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
    
    rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic 'ED_VADIM1
    If rsEmpee.EOF Then
        Exit Sub
    End If

    If glbtermopen Then
        xEmpID = glbTERM_ID
    Else
        xEmpID = glbLEE_ID
    End If
    
    'xTranType = "R"
    xChgType = "Personal Info"

    xEDate = Date
    xForm = "Demographics"
    xLDate = Date
    'If IsDate(rsEmpee("ED_DOH")) Then
    '    If CVDate(rsEmpee("ED_DOH")) > CVDate(xLDate) Then
    '        xLDate = rsEmpee("ED_DOH")
    '    End If
    'End If
    rsEmpee.Close
    
    'ADP Branch or Dept field changes --------------------------------------
    xADPBranch = "ADPBranch" 'lStr("Location")
    xADPDept = "ADPDept" 'lStr("Salary Distribution")
    xPayrollID = txtPayrollID.Text
    xTranType = "M"
    If Len(xPayrollID) > 0 Then
        If Not ADPBranchOld(0) = clpCode(1).Text Then 'main Branch
            Call ADPPayAtWorkAuditAdd(xEmpID, xPayrollID, xTranType, xChgType, xEDate, xForm, xADPBranch, ADPBranchOld(0), clpCode(1).Text, xLDate)
        End If
        If Not ADPDeptOld(0) = clpSalDist.Text Then  'main dept
            Call ADPPayAtWorkAuditAdd(xEmpID, xPayrollID, xTranType, xChgType, xEDate, xForm, xADPDept, ADPDeptOld(0), clpSalDist.Text, xLDate)
        End If
    End If
    
    xPayrollID = medAltPayID(0).Text 'Alt Payroll ID 1
    xTranType = "1"
    If Len(xPayrollID) > 0 Then
        If Len(ADPBranchOld(1)) > 0 And Len(ADPDeptOld(1)) > 0 Then 'for change only
            If Not ADPBranchOld(1) = clpCode(11).Text Then  ' Branch
                Call ADPPayAtWorkAuditAdd(xEmpID, xPayrollID, xTranType, xChgType, xEDate, xForm, xADPBranch, ADPBranchOld(1), clpCode(11).Text, xLDate)
            End If
            If Not ADPDeptOld(1) = clpSalDis2(0).Text Then   ' dept
                Call ADPPayAtWorkAuditAdd(xEmpID, xPayrollID, xTranType, xChgType, xEDate, xForm, xADPDept, ADPDeptOld(1), clpSalDis2(0).Text, xLDate)
            End If
        End If
    End If
    xPayrollID = medAltPayID(1).Text 'Alt Payroll ID 2
    xTranType = "2"
    If Len(xPayrollID) > 0 Then
        If Len(ADPBranchOld(2)) > 0 And Len(ADPDeptOld(2)) > 0 Then 'for change only
            If Not ADPBranchOld(2) = clpCode(12).Text Then  ' Branch
                Call ADPPayAtWorkAuditAdd(xEmpID, xPayrollID, xTranType, xChgType, xEDate, xForm, xADPBranch, ADPBranchOld(2), clpCode(12).Text, xLDate)
            End If
            If Not ADPDeptOld(2) = clpSalDis2(1).Text Then   ' dept
                Call ADPPayAtWorkAuditAdd(xEmpID, xPayrollID, xTranType, xChgType, xEDate, xForm, xADPDept, ADPDeptOld(2), clpSalDis2(1).Text, xLDate)
            End If
        End If
    End If
    xPayrollID = medAltPayID(2).Text 'Alt Payroll ID 3
    xTranType = "3"
    If Len(xPayrollID) > 0 Then
        If Len(ADPBranchOld(3)) > 0 And Len(ADPDeptOld(3)) > 0 Then 'for change only
            If Not ADPBranchOld(3) = clpCode(13).Text Then  ' Branch
                Call ADPPayAtWorkAuditAdd(xEmpID, xPayrollID, xTranType, xChgType, xEDate, xForm, xADPBranch, ADPBranchOld(3), clpCode(13).Text, xLDate)
            End If
            If Not ADPDeptOld(3) = clpSalDis2(2).Text Then   ' dept
                Call ADPPayAtWorkAuditAdd(xEmpID, xPayrollID, xTranType, xChgType, xEDate, xForm, xADPDept, ADPDeptOld(3), clpSalDis2(2).Text, xLDate)
            End If
        End If
    End If

    Exit Sub

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING PAYATWORK AUDIT RECORD", "AUDIT_ADP_PAYATWORK_TRAN", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Sub

Sub ADPPayAtWorkAuditAdd(xEmpNo, xPayrollID, xTranType, xChgType, xEffDate, xForm, xItem, xOldVal, xNewVal, Optional xLDate, Optional xTermSEQ)
Dim rsADPAudit As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM HR_PAYATWORK_CHG_TRAN WHERE 1=2"
    rsADPAudit.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsADPAudit.AddNew
    rsADPAudit("MT_EMPNBR") = xEmpNo
    rsADPAudit("MT_PAYROLL_ID") = Left(xPayrollID, 15)
    rsADPAudit("MT_TRAN_TYPE") = xTranType
    rsADPAudit("MT_CHG_TYPE") = xChgType
    rsADPAudit("MT_EFF_DATE") = xEffDate
    rsADPAudit("MT_FORM") = xForm
    rsADPAudit("MT_ITEM") = xItem
    rsADPAudit("MT_NEW_VALUE") = xNewVal
    rsADPAudit("MT_OLD_VALUE") = xOldVal
    rsADPAudit("MT_UPLOAD") = "N"
    rsADPAudit("MT_LUSER") = glbUserID
    If IsMissing(xLDate) Then
        rsADPAudit("MT_LDATE") = Date
    Else
        rsADPAudit("MT_LDATE") = xLDate
    End If
    rsADPAudit("MT_LTIME") = Time$
    If IsMissing(xTermSEQ) Then
        rsADPAudit("MT_TERM_SEQ") = 0
    Else
        rsADPAudit("MT_TERM_SEQ") = xTermSEQ
    End If

    rsADPAudit.Update
    rsADPAudit.Close
    'to create audit report
    'create GWL Audit view using MT_EMPNBR + TermSEQ as key
    'create HREMP UNION TERM_HREMP and use ED_EMPNBR + 0 as key for Active, ED_EMPNBR + TermSEQ as key for term
    'then this report can show employee name for both active and term

End Sub

Private Sub AUDIT_NGS_TRANS()
Dim rsEmpee As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim strFields As String
Dim SQLQ As String
Dim xUptFlag As Boolean
Dim xNGSStart
Dim xDate1, xDate2
Dim xLDate
Dim xEmpID

'''On Error GoTo AUDIT_ERR
If Not glbNGS_OnFlag Then
    Exit Sub
End If
If NewHireForms.count > 0 Then
    Exit Sub 'modificaiton olnly
End If

'Ticket #20305 Franks 05/17/2011, user may change address for term employee
'If glbtermopen Then Exit Sub

glbEmpDiv = clpDiv.Text
If glbtermopen Then 'Ticket #20305 Franks 05/17/2011
    SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2 FROM Term_HREMP WHERE TERM_SEQ = " & glbTERM_Seq & " "
Else
    SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2 FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
End If
rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic 'ED_VADIM1
If rsEmpee.EOF Then
    Exit Sub
Else
    'If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
    If IsNull(rsEmpee("ED_ORG")) Then glbUNION = "" Else glbUNION = rsEmpee("ED_ORG")
    If IsNull(rsEmpee("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = rsEmpee("ED_VADIM1")
    If IsNull(rsEmpee("ED_VADIM2")) Then glbWFCPayGroup = "" Else glbWFCPayGroup = rsEmpee("ED_VADIM2")
End If
rsEmpee.Close
'No NGS Sub Group, skip
If Len(glbWFCNGSSubGroup) = 0 Then Exit Sub

xLDate = Date

xNGSStart = ""
If glbtermopen Then 'Ticket #20305 Franks 05/17/2011
    SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1 FROM Term_HREMP_OTHER WHERE TERM_SEQ = " & glbTERM_Seq & " "
Else
    SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1 FROM HREMP_OTHER WHERE ER_EMPNBR = " & glbLEE_ID & ""
End If
rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmpOther.EOF Then
    If IsDate(rsEmpOther("ER_OTHERDATE1")) Then
        xNGSStart = rsEmpOther("ER_OTHERDATE1")
    End If
End If
rsEmpOther.Close

'Ticket #20385 Franks 05/31/2011
'Change IHR to write to the NGS Audit Table if the employee has a NGS Sub Group
'regardless of entering a Start Date.
''No NGS Effective Date, skip
'If Len(xNGSStart) = 0 Then Exit Sub
''If the Effective date is later than today, then LDate = Effective date
'If IsDate(xNGSStart) Then
'    If CVDate((xNGSStart)) > CVDate(Date) Then
'        xLDate = CVDate(xNGSStart)
'    End If
'End If

If glbtermopen Then
    xEmpID = glbTERM_ID
Else
    xEmpID = glbLEE_ID
End If
'NGS field changes --------------------------------------
If Not (OSNAME = txtSurname.Text) Then Call NGSAuditAdd(xEmpID, "M", "Demographics", "Surname", OSNAME, txtSurname.Text, xLDate)
If Not (OFNAME = txtFName.Text) Then Call NGSAuditAdd(xEmpID, "M", "Demographics", "First Name", OFNAME, txtFName.Text, xLDate)
If Not (OSIN = medSIN.Text) Then Call NGSAuditAdd(xEmpID, "M", "Demographics", "SSN", OSIN, medSIN.Text, xLDate)
'Birth Date change - begin
If IsDate(ODOB) Then xDate1 = CVDate(ODOB) Else xDate1 = ""
If IsDate(dlpDOB.Text) Then xDate2 = CVDate(dlpDOB.Text) Else xDate2 = ""
If Not (xDate1 = xDate2) Then Call NGSAuditAdd(xEmpID, "M", "Demographics", "Birth Date", xDate1, xDate2, xLDate)
'Birth Date change - end '
If Not (OSEX = txtGender.Text) Then Call NGSAuditAdd(xEmpID, "M", "Demographics", "Gender", OSEX, txtGender.Text, xLDate)
If Not (OSMOKER = ComSmoker.Text) Then Call NGSAuditAdd(xEmpID, "M", "Demographics", "Smoker", OSMOKER, ComSmoker.Text, xLDate)
If Not (OMSTAT = txtMStatus.Text) Then Call NGSAuditAdd(xEmpID, "M", "Demographics", "Marital Status", OMSTAT, txtMStatus.Text, xLDate)
If Not (OADD1 = txtAdd1.Text) Then Call NGSAuditAdd(xEmpID, "M", "Demographics", "Address 1", OADD1, txtAdd1.Text, xLDate)
If Not (OADD2 = txtAdd2.Text) Then Call NGSAuditAdd(xEmpID, "M", "Demographics", "Address 2", OADD2, txtAdd2.Text, xLDate)
If Not (OCITY = txtCity.Text) Then Call NGSAuditAdd(xEmpID, "M", "Demographics", "City", OCITY, txtCity, xLDate)
If Not (oProv = clpProv.Text) Then Call NGSAuditAdd(xEmpID, "M", "Demographics", "Province", oProv, clpProv.Text, xLDate)
If Not (oCountry = txtCountry.Text) Then Call NGSAuditAdd(xEmpID, "M", "Demographics", "Country", oCountry, txtCountry.Text, xLDate)
If Not (OPCODE = medPCode.Text) Then Call NGSAuditAdd(xEmpID, "M", "Demographics", "Zip Code", OPCODE, medPCode.Text, xLDate)
If Not (OPHONE = medTelephone.Text) Then Call NGSAuditAdd(xEmpID, "M", "Demographics", "Telephone", OPHONE, medTelephone.Text, xLDate)

Exit Sub
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING NGS AUDIT RECORD", "NGS AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me
    

End Sub

Private Sub UpdateGPPayCodeOxford()
Dim X%, Y%, SQLQ, Msg
Dim wInComdOld As Boolean
Dim wInComdNew As Boolean
    'The the following GP Integration function must be turn on
    'Employee Master
    If Not glbGP Then
        Exit Sub
    End If
    If NewHireForms.count > 0 Then
        'No position and salary, can't create the Pay Code records
        Exit Sub
    End If
    If isTransferGP("Great Plains", "Employee Master") Then
        If OGLNum <> clpGLNum.Text Then
            'check if there are Pay codes associated with this union code
            wInComdOld = GPBDPayCode(OGLNum)
            wInComdNew = GPBDPayCode(clpGLNum.Text)
            If wInComdNew Or wInComdOld Then
                  If wInComdNew Then 'Add, update, delete
                    'Msg = "Do you want add/update the Employee's Pay Codes with the Benefit/Deduction " & Chr(10)
                    Msg = "Do you want add/update the Employee's Pay Codes with the Income Codes " & Chr(10)
                    Msg = Msg & "defined in the Income Code Matrix under menu item Great Plains? "
                    If MsgBox(Msg, 36, "info:HR") = 6 Then
                        Call UpdateGPBenefitDeduction(glbLEE_ID, clpGLNum.Text, OGLNum)
                        DoEvents
                        frmGPPayCodeList.Show 1
                    End If
                Else
                    If wInComdOld Then 'delete the old income codes only
                        Msg = "Do you want delete the Employee's Pay Codes with the " & lStr("G/L #") & " " & Chr(10)
                        Msg = Msg & "Codes '" & OGLNum & "' defined in the Income Code Matrix under menu item Great Plains? "
                        If MsgBox(Msg, 36, "info:HR") = 6 Then
                            Call UpdateGPBenefitDeduction(glbLEE_ID, clpGLNum.Text, OGLNum)
                            DoEvents
                            frmGPPayCodeList.Show 1
                        End If
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub clpCode_Change(Index As Integer)

If glbCompSerial = "S/N - 2241W" And clpCode(3).Text = "94" Then    'Granite Club - Ticket #19375
    If Index = 3 Then
        clpCode(4).MaxLength = 10
        clpCode(4).Width = 2700
        clpCode(4).Text = "granite"
    End If
End If
End Sub

Sub UpdateOxfordCurrentPosition()
Dim rsEmpJob As New ADODB.Recordset

rsEmpJob.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenStatic, adLockOptimistic
If Not rsEmpJob.EOF Then
    rsEmpJob("JH_DIV") = clpDiv
    rsEmpJob("JH_DEPTNO") = clpDept
    rsEmpJob("JH_GLNO") = clpGLNum
    rsEmpJob("JH_SECTION") = clpCode(4)
    rsEmpJob("JH_PAYROLL_ID") = txtPayrollID
    rsEmpJob.Update
End If
rsEmpJob.Close
Set rsEmpJob = Nothing
End Sub

Sub funGoToRehire(xSIN) 'Ticket #19937 Samuel -  Franks 05/06/2011
Dim rslTerm As New ADODB.Recordset
Dim SQLQ As String
    xSIN = Replace(xSIN, "-", "")
    xSIN = Replace(xSIN, " ", "")
    SQLQ = "SELECT Term_HRTRMEMP.Employee_Number, Term_HRTRMEMP.TERM_SEQ, Term_HRTRMEMP.Term_ID, Term_HRTRMEMP.Term_DOT, Term_HRTRMEMP.Term_DOR, Term_HRTRMEMP.Term_Reason, Term_HRTRMEMP.Term_Rehire,"
    SQLQ = SQLQ & " ED_DEPTNO ,ED_EMPNBR AS EMPNBR, "
    SQLQ = SQLQ & "ED_SURNAME, ED_FNAME,"
    SQLQ = SQLQ & "ED_EMPNBR, ED_PAYROLL_ID,ED_COUNTRY,"
    SQLQ = SQLQ & "ED_INTEL, ED_EMP, ED_DOH, ED_PT, ED_ORG,ED_ADMINBY "
    SQLQ = SQLQ & ",ED_DIV, ED_SIN, ED_SSN "
    SQLQ = SQLQ & "FROM Term_HREMP INNER JOIN "
    SQLQ = SQLQ & "Term_HRTRMEMP ON Term_HRTRMEMP.TERM_SEQ = Term_HREMP.TERM_SEQ "
    SQLQ = SQLQ & " Where ED_SIN = '" & xSIN & "' "
    SQLQ = SQLQ & "ORDER BY Term_HRTRMEMP.TERM_SEQ DESC "
    rslTerm.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rslTerm.EOF Then
        'get the Terminated Employee Info
        glbTermOK = True   'Added 98/05/09 by Andy
        glbTERM_Seq = rslTerm("TERM_SEQ")
        glbTERM_ID = rslTerm("Employee_Number")
        glbTermDate = rslTerm("Term_DOT")
        If Not IsNull(rslTerm("ED_FNAME")) Then
            glbLEE_FName = rslTerm("ED_FNAME")
        Else
            glbLEE_FName = "*ERROR*"
        End If
        If Not IsNull(rslTerm("ED_SURNAME")) Then
            glbLEE_SName = rslTerm("ED_SURNAME")
        Else
            glbLEE_SName = "*ERROR*"
        End If
        If IsNull(rslTerm("ED_ORG")) Then
            glbUNIONTe = ""
        Else
            glbUNIONTe = rslTerm("ED_ORG")
        End If
        glbTerm_FName = glbLEE_FName
        glbTerm_SName = glbLEE_SName
        
        
        'Close the current form
        Call cmdCancel_Click
        DoEvents
        locUploadWithoutCheck = True
        'cancel newhire
        Do While NewHireForms.count > 0
            NewHireForms.Remove 1
        Loop

        Unload Me
        DoEvents
        
    End If
    rslTerm.Close
End Sub

Private Sub UpdateGPMainBenDed()
Dim X%, Y%, SQLQ, Msg
Dim wInComdOld As Boolean
Dim wInComdNew As Boolean

    If Not glbGP Then
        Exit Sub
    End If
    If NewHireForms.count > 0 Then
        'No position and salary, can't create the Pay Code records
        Exit Sub
    End If
    'If isTransferGP("Great Plains", "Emp_PayCode_Benefit_To_GP") Or isTransferGP("Great Plains", "Emp_PayCode_Salary_To_GP") Then
    If isTransferGP("Great Plains", "Employee Master") Then
        If OGLNum <> clpGLNum.Text Then
            'check if there are Pay codes associated with this union code
            wInComdOld = GPBDPayCode(OGLNum)
            wInComdNew = GPBDPayCode(clpGLNum.Text)
            If wInComdNew Or wInComdOld Then
                If wInComdNew Then 'Add, update, delete
                    'Msg = "Do you want add/update the Employee's Pay Codes with the Benefit/Deduction " & Chr(10)
                    Msg = "Do you want add/update the Employee's Pay Codes with the Income Codes " & Chr(10)
                    Msg = Msg & "defined in the Income Code Matrix under menu item Great Plains? "
                    If MsgBox(Msg, 36, "info:HR") = 6 Then
                        Call UpdateGPBenefitDeduction(glbLEE_ID, clpGLNum.Text, OGLNum)
                        DoEvents
                        frmGPPayCodeList.Show 1
                    End If
                Else
                    If wInComdOld Then 'delete the old income codes only
                        Msg = "Do you want delete the Employee's Pay Codes with the Benefit/Deduction " & Chr(10)
                        Msg = Msg & "Codes '" & OGLNum & "' defined in the Income Code Matrix under menu item Great Plains? "
                        If MsgBox(Msg, 36, "info:HR") = 6 Then
                            Call UpdateGPBenefitDeduction(glbLEE_ID, clpGLNum.Text, OGLNum)
                            'Call Employee_GP_BenefitDeduction_Integration(glbLEE_ID, glbUserID, True)
                            DoEvents
                            frmGPPayCodeList.Show 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Function GetTransDiv(xNo As Integer) 'Ticket #21106 Franks 11/07/2011
Dim rsTran As New ADODB.Recordset
Dim SQLQ As String
Dim xFirstVal As String
Dim xFinal As String
'''On Error GoTo Err_Line
    If xNo = 1 Then
        'Parent - Region
        'Child - Section
        xFirstVal = clpCode(2).Text 'Region
        xFinal = "'*'"
        If Not Len(xFirstVal) = 0 Then
            SQLQ = "SELECT * FROM "
            SQLQ = SQLQ & "HRTABL_LINKS "
            SQLQ = SQLQ & "WHERE (1=1) "
            SQLQ = SQLQ & "AND TB_FIRST_TABL = 'EDRG' AND TB_FIRSTCODE ='" & xFirstVal & "' "
            SQLQ = SQLQ & "AND TB_SECOND_TABL = 'EDSE' "
            If rsTran.State <> 0 Then rsTran.Close
            rsTran.Open SQLQ, gdbAdoIhr001, adOpenStatic
            Do While Not rsTran.EOF
                xFinal = xFinal & ",'" & rsTran("TB_SECONDCODE") & "'"
                rsTran.MoveNext
            Loop
            rsTran.Close
        End If
        GetTransDiv = xFinal
    End If
    If xNo = 2 Then 'Ticket #22423 Franks 08/30/2012
        'Parent - Location
        'Child - Region
        xFirstVal = clpCode(1).Text 'Location
        xFinal = "'*'"
        If Not Len(xFirstVal) = 0 Then
            SQLQ = "SELECT * FROM "
            SQLQ = SQLQ & "HRTABL_LINKS "
            SQLQ = SQLQ & "WHERE (1=1) "
            SQLQ = SQLQ & "AND TB_FIRST_TABL = 'EDLC' AND TB_FIRSTCODE ='" & xFirstVal & "' "
            SQLQ = SQLQ & "AND TB_SECOND_TABL = 'EDRG' "
            If rsTran.State <> 0 Then rsTran.Close
            rsTran.Open SQLQ, gdbAdoIhr001, adOpenStatic
            Do While Not rsTran.EOF
                xFinal = xFinal & ",'" & rsTran("TB_SECONDCODE") & "'"
                rsTran.MoveNext
            Loop
            rsTran.Close
        End If
        GetTransDiv = xFinal
    End If
Exit Function
Err_Line:
    'Debug.Print Err.Description
End Function

Private Sub CheckReptAuth() 'Ticket #20885 Franks 11/10/2011 for Samuel
Dim xFlag1 As Boolean
Dim xFlag2 As Boolean
Dim xMsg As String
    xFlag1 = False
    xFlag2 = False
    'Region Change
    If Len(oRegion) > 0 And Len(clpCode(2).Text) > 0 Then
        If oRegion <> clpCode(2).Text Then
            'check if this employee is a Reporting Authority
            If IsReportAuth(glbLEE_ID) Then
                xFlag1 = True
            End If
        End If
    End If
    'Section Change
    If Len(OSection) > 0 And Len(clpCode(4).Text) > 0 Then
        If OSection <> clpCode(4).Text Then
            'check if this employee is a Reporting Authority
            If IsReportAuth(glbLEE_ID) Then
                xFlag2 = True
            End If
        End If
    End If
    If xFlag1 Or xFlag2 Then
        xMsg = "This employee has been assigned as a Reporting Authority on other employee files."
        xMsg = xMsg & "Does this change in "
        If xFlag1 And xFlag2 Then
            xMsg = xMsg & lStr("Region") & "/" & lStr("Section")
        Else
            If xFlag1 Then
                xMsg = xMsg & lStr("Region")
            Else
                xMsg = xMsg & lStr("Section")
            End If
        End If
        xMsg = xMsg & " affect the Reporting Authority structures?"
        frmMsgYesNoUn.lblMsg.Caption = xMsg
        frmMsgYesNoUn.lblMsg.Alignment = 0
        frmMsgYesNoUn.Show 1
        If glbMsgCustomVal = 1 Or glbMsgCustomVal = 3 Then
            'create a report to show the employee list
            Call CreateEmpList4ReportAuth(glbLEE_ID)
            'show the report - begin
            Me.vbxCrystal.Reset
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList2.rpt"
            If Len(glbstrSelCri) >= 0 Then
                Me.vbxCrystal.SelectionFormula = " {HR_EMPLIST_WRK.TT_WRKEMP}='" & glbUserID & "'"
            End If
            'Me.vbxCrystal.Formulas(0) = "rTitle='Employee List for Reporting Authority " & lblEEName & "'"
            'Ticket #21669 Franks 03/01/2012
            xMsg = Replace(lblEEName, "'", "''")
            Me.vbxCrystal.Formulas(0) = "rTitle='Employee List for Reporting Authority " & xMsg & "'"
            Me.vbxCrystal.Connect = RptODBC_SQL
            Me.vbxCrystal.WindowTitle = "Employee List for Reporting Authority " & lblEEName
            Me.vbxCrystal.Destination = 0
            Me.vbxCrystal.Action = 1
            Me.vbxCrystal.Reset
            'show the report - end
        End If
    End If
    
End Sub

Private Sub WFCScreenSetup()
    
    'Ticket #15396
    lbltitle(32).FontBold = True    'Payroll ID
    'lblTitle(0).FontBold = True     'Salutation
    lbltitle(24).FontBold = True    'Region
    lbltitle(12).FontBold = True    'G/L #
    lbltitle(13).FontBold = True    'Division
    lbltitle(23).FontBold = True    'Location
    lbltitle(25).FontBold = True    'Admin By
    
    lblRptNo.Visible = True
    imgIcon.Visible = True
    txtDeptBonusCtr.Visible = True
    lblDeptBonusDesc.Visible = True
    'Ticket #21544 Franks 02/08/2012 - begin
    clpDiv.Visible = False
    frmWFCDIV.Left = clpDiv.Left ' 1515
    frmWFCDIV.Top = clpDiv.Top ' 8100
    frmWFCDIV.BorderStyle = 0
    frmWFCDIV.Visible = True
    'Ticket #21544 Franks 02/08/2012 - end
    
    ''Ticket #22481 Franks 08/27/2012
    'lblWFCNote1.Left = 8880 ' 6240
    'lblWFCNote1.Top = 330
    'lblWFCNote1.Visible = True
    
    'Ticket #22556 Franks 09/18/2012,do not use lblWFCNote1, use imgHelp
    'imgHelp.Visible = True
    'Ticket #22591 Franks 10/12/2012 - use a label instead of imgHelp
    Label1.Visible = True
    
    'Ticket #22553 Franks 09/18/2012 - move it to banking screen for WFC Payforce
    lbltitle(37).Visible = False
    medCOMBINATION.Visible = False
    
    
    'Release 8.0 - Ticket #24361: Re-ordering the Smoker field. Move the rest down
    lbltitle(35).Top = lbltitle(36).Top
    medLICPLATE2.Top = medLOCKER.Top
    
    lbltitle(40).Top = lbltitle(34).Top
    medPARKPERMIT2.Top = medLICPLATE1.Top
    
    lbltitle(38).Top = lbltitle(39).Top
    medTYPEVEHICLE.Top = medPARKPERMIT1.Top
    
    lbltitle(16).Top = lbltitle(33).Top
    ComSmoker.Top = medDRIVERLIC.Top
    lbltitle(16).Left = lbltitle(38).Left
    ComSmoker.Left = medTYPEVEHICLE.Left
    ComSmoker.TabIndex = 46
    
    medPARKPERMIT1.TabIndex = 47
    medTYPEVEHICLE.TabIndex = 48
    medLICPLATE1.TabIndex = 49
    medPARKPERMIT2.TabIndex = 50
    medLOCKER.TabIndex = 51
    medLICPLATE2.TabIndex = 52
    
    lbltitle(46).Visible = True
    txtCandidate.Visible = True
    
    'Ticket #28637 Franks 05/18/2016 - begin
    'lbltitle(55).Top = 1710
    lbltitle(55).FontBold = False
    'medNetworkLogin.Top = 1650
    lbltitle(55).Visible = True
    medNetworkLogin.Visible = True
    'medNetworkLogin.DataField = "ER_NETWORKLOGIN"
    
    'lbltitle(56).Top = 1710
    'medVendorNo.Top = 1650
    lbltitle(56).Visible = True
    lbltitle(56).FontBold = False
    medVendorNo.Visible = True
    'medVendorNo.DataField = "ER_VENDORNO"
    'Ticket #28637 Franks 05/18/2016 - end
    
    cmdEditDiv.Visible = True 'Ticket #28808 Franks 06/20/2016
End Sub

Private Sub SamuelRegionChg()
Dim Msg$
Dim xOldVal
Dim SQLQ As String
    If glbtermopen Then Exit Sub
    
    If Not (oRegion = clpCode(2).Text) Then
        If Len(oRegion) > 0 Then
                Screen.MousePointer = DEFAULT
                Msg$ = lStr("Region") & " has been changed from " & oRegion & " to " & clpCode(2).Text & vbNewLine
                Msg$ = Msg$ & "Please enter the new " & lStr("Last Day") & " Date " & vbNewLine & vbNewLine
                If IsNull(rsDATA("ED_LDAY")) Then xOldVal = "" Else xOldVal = rsDATA("ED_LDAY")
                glbChgTermDate = xOldVal '""
                frmMsgTerm.PenTermDate = "SamDateRegion"
                frmMsgTerm.lblNote1.Caption = Msg$
                frmMsgTerm.lblNote1.Top = 300: frmMsgTerm.lblNote1.Visible = True
                frmMsgTerm.lbltitle(0).Top = 1080: frmMsgTerm.dlpTermDate.Top = 1080
                frmMsgTerm.dlpTermDate = glbChgTermDate
                frmMsgTerm.Show 1
                If Not glbChgTermDate = xOldVal Then
                    'update First Day with glbChgTermDate
                    If IsDate(glbChgTermDate) Then
                        SQLQ = "UPDATE HREMP SET ED_LDAY = " & Date_SQL(glbChgTermDate) & " WHERE ED_EMPNBR = " & glbLEE_ID
                        gdbAdoIhr001.Execute SQLQ
                        'Ticket #22912 Franks 12/06/2012 - begin
                        xFutureDateRegion = glbChgTermDate
                        If CVDate(glbChgTermDate) > Date Then
                            xFutureChgRegion = True
                        End If
                        'Ticket #22912 Franks 12/06/2012 - end
                    End If
                End If
                glbChgTermDate = ""
        End If
    End If
End Sub

Private Sub SamuelSectionChg()
Dim Msg$
Dim xOldVal
Dim SQLQ As String
    If glbtermopen Then Exit Sub
    
    If Not (OSection = clpCode(4).Text) Then
        If Len(OSection) > 0 Then
                Screen.MousePointer = DEFAULT
                Msg$ = lStr("Section") & " has been changed from " & OSection & " to " & clpCode(4).Text & vbNewLine
                Msg$ = Msg$ & "Please enter the new " & lStr("First Day") & " Date " & vbNewLine & vbNewLine
                If IsNull(rsDATA("ED_FDAY")) Then xOldVal = "" Else xOldVal = rsDATA("ED_FDAY")
                glbChgTermDate = xOldVal '""
                frmMsgTerm.PenTermDate = "SamDateSection"
                frmMsgTerm.lblNote1.Caption = Msg$
                frmMsgTerm.lblNote1.Top = 300: frmMsgTerm.lblNote1.Visible = True
                frmMsgTerm.lbltitle(0).Top = 1080: frmMsgTerm.dlpTermDate.Top = 1080
                frmMsgTerm.dlpTermDate = glbChgTermDate
                frmMsgTerm.Show 1
                If Not glbChgTermDate = xOldVal Then
                    'update First Day with glbChgTermDate
                    If IsDate(glbChgTermDate) Then
                        SQLQ = "UPDATE HREMP SET ED_FDAY = " & Date_SQL(glbChgTermDate) & " WHERE ED_EMPNBR = " & glbLEE_ID
                        gdbAdoIhr001.Execute SQLQ
                        'Ticket #22912 Franks 12/06/2012 - begin
                        xFutureDateSection = glbChgTermDate
                        If CVDate(glbChgTermDate) > Date Then
                            xFutureChgSection = True
                        End If
                        'Ticket #22912 Franks 12/06/2012 - end
                    End If
                End If
                glbChgTermDate = ""
        End If
    End If
End Sub

Private Sub SamuelFutureAudit() 'Ticket #22912 Franks 12/06/2012
    If xFutureChgDeptNo Then
        Call SamuelFutureAudAdd(glbLEE_ID, "DEPTNO")
    End If
    If xFutureChgSection Then
        Call SamuelFutureAudAdd(glbLEE_ID, "SECTION")
    End If
    If xFutureChgRegion Then
        Call SamuelFutureAudAdd(glbLEE_ID, "REGION")
    End If
End Sub

Private Sub SamuelFutureAudAdd(xEmpNo, xType)
Dim SQLQ
Dim FTE
Dim USDate
Dim rsBENF As New ADODB.Recordset
Dim NomalCCost
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xPT, xDiv
    rsTA.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsTA.AddNew
    
    rsTB.Open "SELECT * FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
    If Not rsTB.EOF Then
        If IsNull(rsTB("ED_PT")) Then xPT = "" Else xPT = rsTB("ED_PT")
        If IsNull(rsTB("ED_DIV")) Then xDiv = "" Else xDiv = rsTB("ED_DIV")
        If Not IsNull(rsTB("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsTB("ED_PAYROLL_ID")
    Else
        xPT = ""
        xDiv = ""
    End If
    rsTA("AU_EMPNBR") = xEmpNo
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_NEWEMP") = "N"
    rsTA("AU_PTUPL") = xPT
    rsTA("AU_DIVUPL") = xDiv
    If Not rsTB.EOF Then ' Ticket #23341 Franks 02/26/2013
        rsTA("AU_ADMINBY") = rsTB("ED_ADMINBY")
    Else
        rsTA("AU_ADMINBY") = clpCode(3).Text
    End If
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_LDATE") = Date 'by Default
    If xType = "DEPTNO" And IsDate(xFutureDateDeptNo) Then
        rsTA("AU_DEPTNO") = clpDept.Text ' rsTB("ED_DEPTNO")
        rsTA("AU_DEPTEDATE") = xFutureDateDeptNo
        rsTA("AU_LDATE") = xFutureDateDeptNo
    End If
    If xType = "SECTION" And IsDate(xFutureDateSection) Then
        rsTA("AU_SECTION") = clpCode(4).Text ' rsTB("ED_SECTION")
        rsTA("AU_FDAY") = xFutureDateSection
        rsTA("AU_LDATE") = xFutureDateSection
    End If
    If xType = "REGION" And IsDate(xFutureDateRegion) Then
        rsTA("AU_REGION") = clpCode(2).Text ' rsTB("ED_REGION")
        rsTA("AU_LDAY") = xFutureDateRegion
        rsTA("AU_LDATE") = xFutureDateRegion
    End If
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "M"
    rsTA.Update
    rsTA.Close
    rsTB.Close
End Sub
Private Sub WFC_EmpFieldsForUSBen(rsTmpDATA As ADODB.Recordset, xEmpNo)   'Ticket #23247 Franks 04/22/2013
Dim rsNGS As New ADODB.Recordset
Dim SQLQ As String
    If Len(glbTrsStatus) > 0 Then rsTmpDATA("ED_EMP") = glbTrsStatus
    If Len(glbTrsUnion) > 0 Then
        rsTmpDATA("ED_ORG") = glbTrsUnion
        clpCode(0).Text = glbTrsUnion
    End If
    'check with Status
    SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & glbTrsDIV & "' "
    SQLQ = SQLQ & "AND NG_ORG = '" & glbTrsUnion & "' "
    SQLQ = SQLQ & "AND NG_PLAN_CODE = '" & glbTrsStatus & "' "
    If rsNGS.State <> 0 Then rsNGS.Close
    rsNGS.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsNGS.EOF Then
        'check "-" status, such as "-ACT2", convert "-ACT2" to "ACT2" with column NEG_STATUS
        SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & glbTrsDIV & "' "
        SQLQ = SQLQ & "AND NG_ORG = '" & glbTrsUnion & "' "
        SQLQ = SQLQ & "AND LEFT(NG_PLAN_CODE,1) = '-' " 'for "-" code only
        SQLQ = SQLQ & "AND NOT ((CASE LEFT(NG_PLAN_CODE,1) WHEN '-' THEN REPLACE(NG_PLAN_CODE,'-', '') ELSE '' END) = '" & glbTrsStatus & "') "
        If rsNGS.State <> 0 Then rsNGS.Close
        rsNGS.Open SQLQ, gdbAdoIhr001, adOpenStatic
        'If Not rsNGS.EOF Then
        '    Debug.Print ""
        'End If
        If rsNGS.EOF Then
            'No Status
            SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & glbTrsDIV & "' "
            SQLQ = SQLQ & "AND NG_ORG = '" & glbTrsUnion & "' "
            SQLQ = SQLQ & "AND (NG_PLAN_CODE IS NULL OR NG_PLAN_CODE = '') "
            If rsNGS.State <> 0 Then rsNGS.Close
            rsNGS.Open SQLQ, gdbAdoIhr001, adOpenStatic
        End If
    End If
    If Not rsNGS.EOF Then
        'Status/date screen
        If Not IsNull(rsNGS("NG_BENEFIT_GROUP")) Then
            xBenGrpCode = rsNGS("NG_BENEFIT_GROUP")
            rsTmpDATA("ED_BENEFIT_GROUP") = rsNGS("NG_BENEFIT_GROUP")
        End If
        If Not IsNull(rsNGS("NG_PAY_GROUP")) Then
            xWFCPayGroup = rsNGS("NG_PAY_GROUP")
            rsTmpDATA("ED_VADIM2") = rsNGS("NG_PAY_GROUP") 'Pay Group
        End If
        If Not IsNull(rsNGS("NG_SUB_GROUP")) Then
            xWFCNGSCode = rsNGS("NG_SUB_GROUP")
            If glbTrsStatus = "COOP" Or glbTrsStatus = "STUD" Then 'Ticket #25352 Franks 04/16/2014
                '"   If Employment Status = COOP or STUD, don't update the Benefit Master Code or NGS Sub Group
            Else
                rsTmpDATA("ED_VADIM1") = rsNGS("NG_SUB_GROUP") 'NGS Sub Group
            End If
        End If
    End If

End Sub

Private Sub WFCHRSoftDispValues()
Dim rsCanid As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim xTemp

xHRSoftPTCode = ""
xHRSoftJob = ""
xETHNICITY = ""
xRACE = ""
If glbCandidate > 0 Then
    SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE SF_CANDIDATE = " & glbCandidate & " "
    SQLQ = SQLQ & "AND SF_UPT_DEMO = 0 "
    If rsCanid.State <> 0 Then rsCanid.Close
    rsCanid.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsCanid.EOF Then
        Exit Sub
    End If

    'Basic info -------------- begin
    txtCandidate.Text = glbCandidate
    If Not IsNull(rsCanid("SF_SURNAME")) Then txtSurname.Text = rsCanid("SF_SURNAME")
    If Not IsNull(rsCanid("SF_FNAME")) Then txtFName.Text = rsCanid("SF_FNAME")
    If Not IsNull(rsCanid("SF_ADDR1")) Then txtAdd1.Text = rsCanid("SF_ADDR1")
    If Not IsNull(rsCanid("SF_CITY")) Then txtCity.Text = rsCanid("SF_CITY")
    If Not IsNull(rsCanid("SF_HR_PROV")) Then clpProv.Text = rsCanid("SF_HR_PROV")
    If Not IsNull(rsCanid("SF_PCODE")) Then medPCode.Text = rsCanid("SF_PCODE")
    If Not IsNull(rsCanid("SF_COUNTRY")) Then
        txtCountry.Text = rsCanid("SF_COUNTRY")
        comCountry.Text = rsCanid("SF_COUNTRY")
    End If
    If Not IsNull(rsCanid("SF_WORKCOUNTRY")) Then
        txtCountryOfEmp.Text = rsCanid("SF_WORKCOUNTRY")
        comCountryOfEmp.Text = rsCanid("SF_WORKCOUNTRY")
    End If
    If Not IsNull(rsCanid("SF_STARTDATE")) Then
        dlpDate(0).Text = rsCanid("SF_STARTDATE")
        dlpDeptEDate.Text = rsCanid("SF_STARTDATE")
        dlpDivEDate.Text = rsCanid("SF_STARTDATE")
    End If
    If Not IsNull(rsCanid("SF_GENDER")) Then
        If rsCanid("SF_GENDER") = "Male" Then optGender(0).Value = True
        If rsCanid("SF_GENDER") = "Female" Then optGender(1).Value = True
    Else
        optGender(0).Value = False
        optGender(1).Value = False
    End If
    If Not IsNull(rsCanid("SF_PHONE")) Then medTelephone.Text = rsCanid("SF_PHONE")
    'If Not IsNull(rsCanid("SF_BUSNBR")) Then medCellPhone.Text = rsCanid("SF_BUSNBR")
    'Ticket #25562 Franks 06/16/2014
    If Not IsNull(rsCanid("SF_BUSNBR")) Then medPageNbr.Text = rsCanid("SF_BUSNBR")
    If Not IsNull(rsCanid("SF_PTCODE")) Then xHRSoftPTCode = rsCanid("SF_PTCODE")
    If Not IsNull(rsCanid("SF_POSITIONCODE")) Then xHRSoftJob = rsCanid("SF_POSITIONCODE")
    
    If Not IsNull(rsCanid("SF_ETHNICITY")) Then
        If rsCanid("SF_ETHNICITY") = "Hispanic or Latino" Then xETHNICITY = "HL"
        If rsCanid("SF_ETHNICITY") = "Not Hispanic or Latino" Then xETHNICITY = "NHL"
    End If
    If Not IsNull(rsCanid("SF_RACE")) Then
        xRACE = getHRTABLCodeFromDesc("EDRC", Left(rsCanid("SF_RACE"), 25))
    End If
    
    If Not IsNull(rsCanid("SF_HIRETYPE")) Then
        'Ticket #24652 Franks 12/02/2013
        '"   When we did a New Hire after a Rehire, the system used the SSN from the rehire on the new hire record. SSN is not in HRsoft.
        If Not rsCanid("SF_HIRETYPE") = "NEW" Then
            If Len(glbSIN) > 0 Then medSIN.Text = glbSIN
        End If
    End If
    'If Not IsNull(rsCanid("SF_PAYROLL_ID")) Then txtPayrollID.Text = rsCanid("SF_PAYROLL_ID")
End If

End Sub

Private Function getDupEmpByPlantBadgeID(xEmpNo, xPlant, xBadgeID)
'WFC can't enter a duplicate Badge ID within the same plant.
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim retval
    retval = 0
    If Len(xPlant) > 0 And Len(xBadgeID) > 0 Then
        SQLQ = "SELECT ED_EMPNBR, ED_BADGEID FROM HREMP WHERE ED_SECTION = '" & xPlant & "' "
        SQLQ = SQLQ & "AND ED_BADGEID = '" & xBadgeID & "' "
        SQLQ = SQLQ & "AND NOT ED_EMPNBR = " & xEmpNo & " "
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            retval = rsTemp("ED_EMPNBR")
        End If
        rsTemp.Close
    End If
    getDupEmpByPlantBadgeID = retval
End Function

Private Function getDupEmpByPlantPayrollID(xEmpNo, xPayID, xPlant)
'WFC can't enter a duplicate Payroll ID within the same plant.
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim retval
    retval = 0
    If Len(xPlant) > 0 And Len(xPayID) > 0 Then
        SQLQ = "SELECT ED_EMPNBR, ED_BADGEID FROM HREMP WHERE ED_SECTION = '" & xPlant & "' "
        SQLQ = SQLQ & "AND ED_PAYROLL_ID = '" & xPayID & "' "
        SQLQ = SQLQ & "AND NOT ED_EMPNBR = " & xEmpNo & " "
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            retval = rsTemp("ED_EMPNBR")
        End If
        rsTemp.Close
    End If
    getDupEmpByPlantPayrollID = retval
End Function

Private Sub UptWFCEMPOTHER() '#28637 Franks 05/18/2016
Dim rsDAT_Other As New ADODB.Recordset
Dim SQLQ As String

If rsDAT_Other.State <> 0 Then rsDAT_Other.Close

If glbtermopen Then
    SQLQ = "SELECT * FROM Term_HREMP_OTHER"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "SELECT * FROM HREMP_OTHER"
    SQLQ = SQLQ & " where ER_EMPNBR = " & glbLEE_ID
End If
rsDAT_Other.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If rsDAT_Other.EOF Then
    rsDAT_Other.AddNew
    rsDAT_Other("ER_COMPNO") = "001"
    rsDAT_Other("ER_EMPNBR") = glbLEE_ID
    If glbtermopen Then
        rsDAT_Other("TERM_SEQ") = glbTERM_Seq
    End If
End If

'Network Login
If Len(medNetworkLogin.Text) = 0 Then rsDAT_Other("ER_NETWORKLOGIN") = Null Else rsDAT_Other("ER_NETWORKLOGIN") = Left(medNetworkLogin.Text, 40)
'Vendor Number
If Len(medVendorNo.Text) = 0 Then rsDAT_Other("ER_VENDORNO") = Null Else rsDAT_Other("ER_VENDORNO") = Left(medVendorNo.Text, 40)

rsDAT_Other.Update
rsDAT_Other.Close

End Sub

Private Sub Update_WorkVisa_Info(xEmpNo)
Dim rsEmpOther As New ADODB.Recordset
Dim SQLQ As String
    
    SQLQ = "SELECT * FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo & ""
    rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsEmpOther.EOF Then
        rsEmpOther.AddNew
        rsEmpOther("ER_COMPNO") = "001"
        rsEmpOther("ER_EMPNBR") = xEmpNo
    End If
    
    If glbLinamar Then 'Ticket #28875 Franks 07/13/2016
        'don't need it
    Else
        rsEmpOther("ER_VISAPERMITNO") = glbWorkVisaNo
    End If
    rsEmpOther("ER_VISAPERMITDATE") = CVDate(glbWorkExpDate)
    rsEmpOther("ER_LDATE") = Date
    rsEmpOther("ER_LTIME") = Time$
    rsEmpOther("ER_LUSER") = glbUserID
    rsEmpOther.Update
    rsEmpOther.Close
    
End Sub

Private Function GetTabState() As Boolean
    GetTabState = False
    If GetKeyState(VK_TAB) And -256 Then
        GetTabState = True
    End If
End Function

Private Sub UptAltPayrollIDs() 'Macaulay Ticket #25016 Franks 04/01/2014
Dim rsDAT_Other As New ADODB.Recordset
Dim SQLQ As String

If rsDAT_Other.State <> 0 Then rsDAT_Other.Close

If glbtermopen Then
    SQLQ = "SELECT * FROM Term_HREMP_OTHER"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "SELECT * FROM HREMP_OTHER"
    SQLQ = SQLQ & " where ER_EMPNBR = " & glbLEE_ID
End If
rsDAT_Other.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If rsDAT_Other.EOF Then
    rsDAT_Other.AddNew
    rsDAT_Other("ER_COMPNO") = "001"
    rsDAT_Other("ER_EMPNBR") = glbLEE_ID
    If glbtermopen Then
        rsDAT_Other("TERM_SEQ") = glbTERM_Seq
    End If
End If
'Payroll id
If Len(medAltPayID(0).Text) = 0 Then rsDAT_Other("ER_PAYROLL_ID1") = Null Else rsDAT_Other("ER_PAYROLL_ID1") = Left(medAltPayID(0).Text, 15)
If Len(medAltPayID(1).Text) = 0 Then rsDAT_Other("ER_PAYROLL_ID2") = Null Else rsDAT_Other("ER_PAYROLL_ID2") = Left(medAltPayID(1).Text, 15)
If Len(medAltPayID(2).Text) = 0 Then rsDAT_Other("ER_PAYROLL_ID3") = Null Else rsDAT_Other("ER_PAYROLL_ID3") = Left(medAltPayID(2).Text, 15)
'If Len(medAltPayID(3).Text) = 0 Then rsDAT_Other("ER_PAYROLL_ID4") = Null Else rsDAT_Other("ER_PAYROLL_ID4") = Left(medAltPayID(3).Text, 15)
'If Len(medAltPayID(4).Text) = 0 Then rsDAT_Other("ER_PAYROLL_ID5") = Null Else rsDAT_Other("ER_PAYROLL_ID5") = Left(medAltPayID(4).Text, 15)
'Company code
If Len(clpCode(8).Text) = 0 Then rsDAT_Other("ER_PAYR_COMP1") = Null Else rsDAT_Other("ER_PAYR_COMP1") = Left(clpCode(8).Text, 15)
If Len(clpCode(9).Text) = 0 Then rsDAT_Other("ER_PAYR_COMP2") = Null Else rsDAT_Other("ER_PAYR_COMP2") = Left(clpCode(9).Text, 15)
If Len(clpCode(10).Text) = 0 Then rsDAT_Other("ER_PAYR_COMP3") = Null Else rsDAT_Other("ER_PAYR_COMP3") = Left(clpCode(10).Text, 15)
'If Len(clpCode(11).Text) = 0 Then rsDAT_Other("ER_PAYR_COMP4") = Null Else rsDAT_Other("ER_PAYR_COMP4") = Left(clpCode(11).Text, 15)
'If Len(clpCode(12).Text) = 0 Then rsDAT_Other("ER_PAYR_COMP5") = Null Else rsDAT_Other("ER_PAYR_COMP5") = Left(clpCode(12).Text, 15)

'Ticket #24557 Franks 09/05/2012
'Location
If Len(clpCode(11).Text) = 0 Then rsDAT_Other("ER_LOC1") = Null Else rsDAT_Other("ER_LOC1") = Left(clpCode(11).Text, 10)
If Len(clpCode(12).Text) = 0 Then rsDAT_Other("ER_LOC2") = Null Else rsDAT_Other("ER_LOC2") = Left(clpCode(12).Text, 10)
If Len(clpCode(13).Text) = 0 Then rsDAT_Other("ER_LOC3") = Null Else rsDAT_Other("ER_LOC3") = Left(clpCode(13).Text, 10)
'Salary Distribution
If Len(clpSalDis2(0).Text) = 0 Then rsDAT_Other("ER_SALDIST1") = Null Else rsDAT_Other("ER_SALDIST1") = Left(clpSalDis2(0).Text, 10)
If Len(clpSalDis2(1).Text) = 0 Then rsDAT_Other("ER_SALDIST2") = Null Else rsDAT_Other("ER_SALDIST2") = Left(clpSalDis2(1).Text, 10)
If Len(clpSalDis2(2).Text) = 0 Then rsDAT_Other("ER_SALDIST3") = Null Else rsDAT_Other("ER_SALDIST3") = Left(clpSalDis2(2).Text, 10)
'Hire Date
If Not IsDate(dlpDate(1).Text) Then rsDAT_Other("ER_DOH1") = Null Else rsDAT_Other("ER_DOH1") = CVDate(dlpDate(1).Text)
If Not IsDate(dlpDate(2).Text) Then rsDAT_Other("ER_DOH2") = Null Else rsDAT_Other("ER_DOH2") = CVDate(dlpDate(2).Text)
If Not IsDate(dlpDate(3).Text) Then rsDAT_Other("ER_DOH3") = Null Else rsDAT_Other("ER_DOH3") = CVDate(dlpDate(3).Text)
'Region
If Len(clpCode(14).Text) = 0 Then rsDAT_Other("ER_REGION1") = Null Else rsDAT_Other("ER_REGION1") = Left(clpCode(14).Text, 4)
If Len(clpCode(15).Text) = 0 Then rsDAT_Other("ER_REGION2") = Null Else rsDAT_Other("ER_REGION2") = Left(clpCode(15).Text, 4)
If Len(clpCode(16).Text) = 0 Then rsDAT_Other("ER_REGION3") = Null Else rsDAT_Other("ER_REGION3") = Left(clpCode(16).Text, 4)
'Status
If Len(clpCode(17).Text) = 0 Then rsDAT_Other("ER_EMP1") = Null Else rsDAT_Other("ER_EMP1") = Left(clpCode(17).Text, 4)
If Len(clpCode(18).Text) = 0 Then rsDAT_Other("ER_EMP2") = Null Else rsDAT_Other("ER_EMP2") = Left(clpCode(18).Text, 4)
If Len(clpCode(19).Text) = 0 Then rsDAT_Other("ER_EMP3") = Null Else rsDAT_Other("ER_EMP3") = Left(clpCode(19).Text, 4)
'Terminated Date
If Not IsDate(dlpTermDate(0).Text) Then rsDAT_Other("ER_DOT1") = Null Else rsDAT_Other("ER_DOT1") = CVDate(dlpTermDate(0).Text)
If Not IsDate(dlpTermDate(1).Text) Then rsDAT_Other("ER_DOT2") = Null Else rsDAT_Other("ER_DOT2") = CVDate(dlpTermDate(1).Text)
If Not IsDate(dlpTermDate(2).Text) Then rsDAT_Other("ER_DOT3") = Null Else rsDAT_Other("ER_DOT3") = CVDate(dlpTermDate(2).Text)

rsDAT_Other.Update
rsDAT_Other.Close
End Sub

Private Sub DisAltPayrollIDs()
Dim rsDAT_Other As New ADODB.Recordset
Dim SQLQ As String

If rsDAT_Other.State <> 0 Then rsDAT_Other.Close

If glbtermopen Then
    SQLQ = "SELECT * FROM Term_HREMP_OTHER"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "SELECT * FROM HREMP_OTHER"
    SQLQ = SQLQ & " where ER_EMPNBR = " & glbLEE_ID
End If
rsDAT_Other.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not rsDAT_Other.EOF Then
    'Payroll id
    If IsNull(rsDAT_Other("ER_PAYROLL_ID1")) Then medAltPayID(0).Text = "" Else medAltPayID(0).Text = rsDAT_Other("ER_PAYROLL_ID1")
    If IsNull(rsDAT_Other("ER_PAYROLL_ID2")) Then medAltPayID(1).Text = "" Else medAltPayID(1).Text = rsDAT_Other("ER_PAYROLL_ID2")
    If IsNull(rsDAT_Other("ER_PAYROLL_ID3")) Then medAltPayID(2).Text = "" Else medAltPayID(2).Text = rsDAT_Other("ER_PAYROLL_ID3")
    'If IsNull(rsDAT_Other("ER_PAYROLL_ID4")) Then medAltPayID(3).Text = "" Else medAltPayID(3).Text = rsDAT_Other("ER_PAYROLL_ID4")
    'If IsNull(rsDAT_Other("ER_PAYROLL_ID5")) Then medAltPayID(4).Text = "" Else medAltPayID(4).Text = rsDAT_Other("ER_PAYROLL_ID5")
    'Company code
    If IsNull(rsDAT_Other("ER_PAYR_COMP1")) Then clpCode(8).Text = "" Else clpCode(8).Text = rsDAT_Other("ER_PAYR_COMP1")
    If IsNull(rsDAT_Other("ER_PAYR_COMP2")) Then clpCode(9).Text = "" Else clpCode(9).Text = rsDAT_Other("ER_PAYR_COMP2")
    If IsNull(rsDAT_Other("ER_PAYR_COMP3")) Then clpCode(10).Text = "" Else clpCode(10).Text = rsDAT_Other("ER_PAYR_COMP3")
    'If IsNull(rsDAT_Other("ER_PAYR_COMP4")) Then clpCode(11).Text = "" Else clpCode(11).Text = rsDAT_Other("ER_PAYR_COMP4")
    'If IsNull(rsDAT_Other("ER_PAYR_COMP5")) Then clpCode(12).Text = "" Else clpCode(12).Text = rsDAT_Other("ER_PAYR_COMP5")
    
    'Ticket #24557 Franks 09/05/2012
    If IsNull(rsDAT_Other("ER_LOC1")) Then clpCode(11).Text = "" Else clpCode(11).Text = rsDAT_Other("ER_LOC1")
    If IsNull(rsDAT_Other("ER_LOC2")) Then clpCode(12).Text = "" Else clpCode(12).Text = rsDAT_Other("ER_LOC2")
    If IsNull(rsDAT_Other("ER_LOC3")) Then clpCode(13).Text = "" Else clpCode(13).Text = rsDAT_Other("ER_LOC3")
    'SalDist
    If IsNull(rsDAT_Other("ER_SALDIST1")) Then clpSalDis2(0).Text = "" Else clpSalDis2(0).Text = rsDAT_Other("ER_SALDIST1")
    If IsNull(rsDAT_Other("ER_SALDIST2")) Then clpSalDis2(1).Text = "" Else clpSalDis2(1).Text = rsDAT_Other("ER_SALDIST2")
    If IsNull(rsDAT_Other("ER_SALDIST3")) Then clpSalDis2(2).Text = "" Else clpSalDis2(2).Text = rsDAT_Other("ER_SALDIST3")
    'Hire Date
    If IsNull(rsDAT_Other("ER_DOH1")) Then dlpDate(1).Text = "" Else dlpDate(1).Text = rsDAT_Other("ER_DOH1")
    If IsNull(rsDAT_Other("ER_DOH2")) Then dlpDate(2).Text = "" Else dlpDate(2).Text = rsDAT_Other("ER_DOH2")
    If IsNull(rsDAT_Other("ER_DOH3")) Then dlpDate(3).Text = "" Else dlpDate(3).Text = rsDAT_Other("ER_DOH3")
    'Region
    If IsNull(rsDAT_Other("ER_REGION1")) Then clpCode(14).Text = "" Else clpCode(14).Text = rsDAT_Other("ER_REGION1")
    If IsNull(rsDAT_Other("ER_REGION2")) Then clpCode(15).Text = "" Else clpCode(15).Text = rsDAT_Other("ER_REGION2")
    If IsNull(rsDAT_Other("ER_REGION3")) Then clpCode(16).Text = "" Else clpCode(16).Text = rsDAT_Other("ER_REGION3")
    'Status
    If IsNull(rsDAT_Other("ER_EMP1")) Then clpCode(17).Text = "" Else clpCode(17).Text = rsDAT_Other("ER_EMP1")
    If IsNull(rsDAT_Other("ER_EMP2")) Then clpCode(18).Text = "" Else clpCode(18).Text = rsDAT_Other("ER_EMP2")
    If IsNull(rsDAT_Other("ER_EMP3")) Then clpCode(19).Text = "" Else clpCode(19).Text = rsDAT_Other("ER_EMP3")
    'Terminated Date
    If IsNull(rsDAT_Other("ER_DOT1")) Then dlpTermDate(0).Text = "" Else dlpTermDate(0).Text = rsDAT_Other("ER_DOT1")
    If IsNull(rsDAT_Other("ER_DOT2")) Then dlpTermDate(1).Text = "" Else dlpTermDate(1).Text = rsDAT_Other("ER_DOT2")
    If IsNull(rsDAT_Other("ER_DOT3")) Then dlpTermDate(2).Text = "" Else dlpTermDate(2).Text = rsDAT_Other("ER_DOT3")
    
Else
    'Payroll id
    medAltPayID(0).Text = ""
    medAltPayID(1).Text = ""
    medAltPayID(2).Text = ""
    'medAltPayID(3).Text = ""
    'medAltPayID(4).Text = ""
    clpCode(8).Text = ""
    clpCode(9).Text = ""
    clpCode(10).Text = ""
    'clpCode(11).Text = ""
    'clpCode(12).Text = ""
    'Loc
    clpCode(11).Text = ""
    clpCode(12).Text = ""
    clpCode(13).Text = ""
    'Salary Distribution
    clpSalDis2(0).Text = ""
    clpSalDis2(1).Text = ""
    clpSalDis2(2).Text = ""
    'Hire Date
    dlpDate(1).Text = ""
    dlpDate(2).Text = ""
    dlpDate(3).Text = ""
    'Region
    clpCode(14).Text = ""
    clpCode(15).Text = ""
    clpCode(16).Text = ""
    'Status
    clpCode(17).Text = ""
    clpCode(18).Text = ""
    clpCode(19).Text = ""
    'Term Date
    dlpTermDate(0).Text = ""
    dlpTermDate(1).Text = ""
    dlpTermDate(2).Text = ""
End If
rsDAT_Other.Close
End Sub

Private Sub Track_Courses_Renewal_Update(Optional xDelete, Optional xPosType, Optional xOldPosCode)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsPosCourse As New ADODB.Recordset
    Dim rsEmpJob As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim SQLQ As String
    Dim flgRequired As Boolean
    Dim PosCode As String
    
    On Error GoTo Track_Courses_Renewal_Err
    
    PosCode = xOldPosCode
    
    '???If chkTrackCrsRenewal And IsMissing(xDelete) Then  'Previous Position course being tracked
    '???    'Turn-ON the tracking
    '???    Call Update_Employee_Job_Training_List(PosCode, "Previous")
        
    '???Else
        'Turn-OFF the tracking
        '-------------------------------------------------------------------------------------------------
        'remove Renewal Date from the Continuing Education screen
        'retrieve Unique for each Position courses first
        SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_DATCOMP,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
        SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND ES_JOB = '" & PosCode & "'"
        SQLQ = SQLQ & " AND ES_CRSCODE IN (SELECT TR_CRSCODE FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
        'SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0))"
        SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0)"
        'Courses that do not belong to the Department of the Employee
        SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT PC_CRSCODE FROM HR_JOB_COURSE WHERE PC_JOB = '" & PosCode & "' AND PC_DEPTNO IS NOT NULL AND PC_DEPTNO <> '" & clpDept.Text & "'))"
        SQLQ = SQLQ & " ORDER BY ES_RENEW"
        rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsContEdu.EOF Then
            rsContEdu.MoveFirst
            
            Do While Not rsContEdu.EOF
                SQLQ = "SELECT * FROM HR_TRAIN"
                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                SQLQ = SQLQ & " AND TR_JOB = '" & PosCode & "'"
                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsContEdu("ES_CRSCODE") & "'"
                rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsHRTrain.EOF Then
                    If (rsHRTrain("TR_RENEW") = rsContEdu("ES_RENEW")) And (IsNull(rsHRTrain("TR_COURSE_TAKEN")) Or (rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP"))) Then
                        rsContEdu("ES_RENEW") = Null
                        rsContEdu("ES_LDATE") = Date
                        rsContEdu("ES_LUSER") = glbUserID
                        rsContEdu("ES_LTIME") = Time$
                        rsContEdu.Update
                    
                        If Not IsNull(rsContEdu("ES_DATCOMP")) And Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                            'Since the course was completed - mark the Follow Up as
                            'Completed instead of deleting it.
                            SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                            If Not IsMissing(xPosType) Then
                                SQLQ = SQLQ & " AND TR_POS_TYPE = '" & xPosType & "'"
                            End If
                            'SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0))"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsContEdu("ES_CRSCODE") & "')"
                            gdbAdoIhr001.Execute SQLQ
                        Else
                            'Delete "Unique for each Position" courses from Follow Up records
                            'as no Course completion record found
                            SQLQ = "DELETE FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                            If Not IsMissing(xPosType) Then
                                SQLQ = SQLQ & " AND TR_POS_TYPE = '" & xPosType & "'"
                            End If
                            'SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0))"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsContEdu("ES_CRSCODE") & "')"
                            gdbAdoIhr001.Execute SQLQ
                        
                            'Clear the Follow Up ID in the Position record
                            'if the course code is TRAIN
                            If rsContEdu("ES_CRSCODE") = "TRAIN" Then
                                'Search HR_JOB_HISTORY table for this Position record
                                'and update with Follow Up Id
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("JH_FOLLOWUP_ID") = Null
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                            End If
                        End If
                    End If
                Else
                    'Delete "Unique for each Position" courses from Follow Up records
                    'as no Course completion record found
                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                    If Not IsMissing(xPosType) Then
                        SQLQ = SQLQ & " AND TR_POS_TYPE = '" & xPosType & "'"
                    End If
                    'SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0))"
                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsContEdu("ES_CRSCODE") & "')"
                    gdbAdoIhr001.Execute SQLQ
                
                    'Clear the Follow Up ID in the Primary Position record
                    'if the course code is TRAIN
                    If rsContEdu("ES_CRSCODE") = "TRAIN" Then
                        'Search HR_JOB_HISTORY table for this Position record
                        'and update with Follow Up Id
                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsTJob.EOF Then
                            rsTJob("JH_FOLLOWUP_ID") = Null
                            rsTJob.Update
                        End If
                        rsTJob.Close
                        Set rsTJob = Nothing
                    End If
                End If
                rsHRTrain.Close
                Set rsHRTrain = Nothing
                
                rsContEdu.MoveNext
            Loop
        Else
            'Delete "Unique for each Position" courses from Follow Up records
            'as no Course completion record found
            SQLQ = "DELETE FROM HR_FOLLOW_UP"
            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID IN (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
            If Not IsMissing(xPosType) Then
                SQLQ = SQLQ & " AND TR_POS_TYPE = '" & xPosType & "'"
            End If
            'SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0))"
            SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0)"
            'SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsContEdu("ES_CRSCODE") & "')"
            'Courses that do not belong to the Department of the Employee
            SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT PC_CRSCODE FROM HR_JOB_COURSE WHERE PC_JOB = '" & PosCode & "' AND PC_DEPTNO IS NOT NULL AND PC_DEPTNO <> '" & clpDept.Text & "'))"
            gdbAdoIhr001.Execute SQLQ
        End If
        rsContEdu.Close
        Set rsContEdu = Nothing
        
        'from the Training List
        SQLQ = "DELETE FROM HR_TRAIN"
        SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND TR_JOB = '" & PosCode & "'"
        If Not IsMissing(xPosType) Then
            SQLQ = SQLQ & " AND TR_POS_TYPE = '" & xPosType & "'"
        End If
        SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0)"
        'Courses that do not belong to the Department of the Employee
        SQLQ = SQLQ & " AND TR_CRSCODE IN (SELECT PC_CRSCODE FROM HR_JOB_COURSE WHERE PC_JOB = '" & PosCode & "' AND PC_DEPTNO IS NOT NULL AND PC_DEPTNO <> '" & clpDept.Text & "')"
        gdbAdoIhr001.Execute SQLQ
        '-------------------------------------------------------------------------------------------------
        
        'Rest of this position's required courses which are not 'unique for each position'
        'Retrieve the Required Courses for this position - Non Unqiue for each Position courses
        SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE PC_JOB = '" & PosCode & "'"
        SQLQ = SQLQ & " AND PC_CRSCODE NOT IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0)"
        
        'Ticket #25609 - Training Plan by Department
        'Only courses NOT matching employee's Department if the Course has Department Code assigned
        SQLQ = SQLQ & " AND ((PC_DEPTNO IS NOT NULL) AND (PC_DEPTNO <> '" & clpDept.Text & "'))"
        
        rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsReqCourse.EOF Then
            
            rsReqCourse.MoveFirst
            flgRequired = False
            
            Do While Not rsReqCourse.EOF
                'Initialise if this course is required by any other position
                flgRequired = False
                
                'Check if the each required courses for this position is also required by other positions
                'Select all current positions in HR_JOB_HISTORY and HR_TEMP_WORK, and
                'Previous Positions with Tracking ON - for this employee
                SQLQ = "SELECT JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & glbLEE_ID & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
                '???and not the position currently selected
                'SQLQ = SQLQ & " AND (JH_ID <> " & RSDATA!JH_ID & ")"
                SQLQ = SQLQ & " UNION "
                SQLQ = SQLQ & " SELECT TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL FROM HR_TEMP_WORK WHERE TW_EMPNBR = " & glbLEE_ID & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
                rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsEmpJob.EOF Then
                    rsEmpJob.MoveFirst
                    
                    Do While Not rsEmpJob.EOF
                        'Check in the Required Courses table if the retrieved required course is required by other retrieved position
                        SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE PC_JOB = '" & rsEmpJob("TW_JOB") & "'"
                        SQLQ = SQLQ & " AND PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                        rsPosCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsPosCourse.EOF Then
                            'Check if this course has Current and/or Previous Renewal Period
                            If rsEmpJob("TW_CURRENT") And (IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0) Then
                                'Current Position - no Current Renewal Period
                                'Check if another position required this course
                                GoTo Next_Position
                            ElseIf rsEmpJob("TW_TRK_CRS_RENEWAL") And (IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) Or rsReqCourse("PC_RENEW_CRS_PRV") = 0) Then
                                'Previous Position - no Previous Renewal Period
                                'Check if another position required this course
                                GoTo Next_Position
                            End If
                            
                            'Required by another position. Do not delete this Course
                            flgRequired = False 'Changed
                                              
                            rsPosCourse.Close
                            Set rsPosCourse = Nothing
                                              
                            'Move to the next Course
                            GoTo Next_RequiredCourse
                        End If
Next_Position:
                        rsPosCourse.Close
                        Set rsPosCourse = Nothing
                        
                        rsEmpJob.MoveNext
                    Loop
                End If
Next_RequiredCourse:
                rsEmpJob.Close
                Set rsEmpJob = Nothing

                If flgRequired Then
                    'Call procedure to update Renewal Date and Position Code, Follow Up effective date
                    'Do not do anything now. At the end of this loop go through each of the
                    'courses and update the Renewal Dates and Position Codes and create the follow up records.
                Else
                    'This course is not required by any other position this employee is holding
                    'or the Current and/or Previous Renewal Period is missing which means
                    'any other position Current/Previous requiring this course will not be
                    'able to renew it without the appropriate renewal period.
                    
                    'Clear the Renewal date for this course and for this employee from
                    'Continuing Education screen
                    SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                    SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                    SQLQ = SQLQ & " AND ES_JOB = '" & PosCode & "'"
                    SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                    'Ticket #26211 Franks 10/29/2014
                    'SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                    SQLQ = SQLQ & " AND ES_RENEW IN (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    'Ticket #26211 Franks 10/29/2014
                    'SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                    SQLQ = SQLQ & " AND ES_DATCOMP IN (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsContEdu.EOF Then
                        rsContEdu("ES_RENEW") = Null
                        rsContEdu("ES_LDATE") = Date
                        rsContEdu("ES_LUSER") = glbUserID
                        rsContEdu("ES_LTIME") = Time$
                        rsContEdu.Update
                        
                        If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                            'Since the course was completed - mark the Follow Up as
                            'Completed instead of deleting it.
                            SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            gdbAdoIhr001.Execute SQLQ
                        Else
                            'Delete the Follow Up record for this training record
                            'as no Course completion record found
                            SQLQ = "DELETE FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            gdbAdoIhr001.Execute SQLQ
                        
                            'Clear the Follow Up ID in the Temp/Cross Training Position record
                            'if the course code is TRAIN
                            If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                'Search HR_JOB_HISTORY table for this Position record
                                'and update with Follow Up Id
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("JH_FOLLOWUP_ID") = Null
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                            End If
                        End If
                    Else
                        'Delete the Follow Up record for this training record
                        'as no Course completion record found
                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & PosCode & "'"
                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                        gdbAdoIhr001.Execute SQLQ
                    
                        'Clear the Follow Up Id in the Temp/Cross Training Position record
                        'if the course code is TRAIN
                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                            'Search HR_JOB_HISTORY table for this Position record
                            'and update with Follow Up Id
                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                rsTJob("JH_FOLLOWUP_ID") = Null
                                rsTJob.Update
                            End If
                            rsTJob.Close
                            Set rsTJob = Nothing
                        End If
                    End If
                    rsContEdu.Close
                    Set rsContEdu = Nothing
                    
                    'Delete this Training List record as the course is not required by other positions
                    SQLQ = "DELETE FROM HR_TRAIN"
                    SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                    SQLQ = SQLQ & " AND TR_JOB = '" & PosCode & "'"
                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                    gdbAdoIhr001.Execute SQLQ
                End If
                rsReqCourse.MoveNext
            Loop
        End If
        rsReqCourse.Close
        Set rsReqCourse = Nothing
        
        'Call procedure to update Renewal Dates and Position Codes and create/update the follow up records.
        'For the remaining required courses for this position which are required by other positions.
        Call Update_Remaining_Tracked_Courses(PosCode)
    '???End If
    
Exit Sub

Track_Courses_Renewal_Err:
If Err = 3018 Then
    Err = 0
    Resume Next
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
If Len(SQLQ) = 0 Then
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Track_Courses_Renewal_Update", "HR_JOB_HISTORY", "Courses_Renewal_Update")
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, SQLQ, "HR_JOB_HISTORY", "Courses_Renewal_Update")
End If
Call RollBack '26July99 js
End Sub

Private Sub Update_Remaining_Tracked_Courses(xJob)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsPosCourse As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsCourseCode As New ADODB.Recordset
    Dim rsEmpJobs As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim SQLQ As String
    Dim xDWMY As String
    Dim xRenewalDt
    Dim xComments As String
    
    On Error GoTo Remaining_Tracked_Courses_Err
    
    'Retrieve the Required Courses for this position - Non Unqiue for each Position courses
    'SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE PC_JOB = '" & clpJob.Text & "'"
    'SQLQ = SQLQ & " AND PC_CRSCODE NOT IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_UNIQUE_FOR_POS<>0)"
    'rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    'If Not rsReqCourse.EOF Then
    '    rsReqCourse.MoveFirst
        
    '    Do While Not rsReqCourse.EOF
            'Select all current positions in HR_JOB_HISTORY and HR_TEMP_WORK, and
            'Previous Positions with Tracking ON - for this employee
            'The records will be ordered by Current, Temporary, Previous tracked
            SQLQ = "SELECT JH_ID AS TW_ID, JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
            SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
            '???and not the position currently selected
            'SQLQ = SQLQ & " AND (JH_ID <> " & RSDATA!JH_ID & ")"
            SQLQ = SQLQ & " UNION "
            SQLQ = SQLQ & " SELECT TW_ID, TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
            SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
            'SQLQ = SQLQ & " ORDER BY POS_TYPE ASC,TW_CURRENT DESC"
            SQLQ = SQLQ & " ORDER BY TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
            rsEmpJobs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsEmpJobs.EOF Then
                rsEmpJobs.MoveFirst
                
                Do While Not rsEmpJobs.EOF
                    If (rsEmpJobs("POS_TYPE") = "CURRENT" Or rsEmpJobs("POS_TYPE") = "TEMPORARY") And rsEmpJobs("TW_CURRENT") Then
                        'Changed
                        'Call Update_Employee_Job_Training_List(rsEmpJobs("TW_JOB"), "Current", rsEmpJobs("TW_SDATE"))
                        Call Update_Employee_Job_Training_List(rsEmpJobs("TW_JOB"), IIf(rsEmpJobs("POS_TYPE") = "CURRENT", "Current", "Temporary"), rsEmpJobs("TW_SDATE"))
                    Else
                        Call Update_Employee_Job_Training_List(rsEmpJobs("TW_JOB"), "Previous", rsEmpJobs("TW_SDATE"), rsEmpJobs("TW_ENDDATE"))
                    End If
                    GoTo next_EmpJob
                    
                    'Find out which position requires this course and update the training list accordingly.
                    SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE PC_JOB = '" & rsEmpJobs("TW_JOB") & "'"
                    SQLQ = SQLQ & " AND PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                    rsPosCourse.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsPosCourse.EOF Then
                        'If Primary - CURRENT or TEMPORARY - Current
                        'Get Current Renewal Period from Course Code Master to calculate Renewal Date
                        If (rsEmpJobs("POS_TYPE") = "CURRENT" Or rsEmpJobs("POS_TYPE") = "TEMPORARY") And rsEmpJobs("TW_CURRENT") Then
                            SQLQ = "SELECT ES_CRSCODE,ES_RENEW_CRS_CUR,ES_CUR_PRD_DWMY,ES_RENEW_FOLLOWUP, ES_FLWUP_PRD_DWMY FROM HR_COURSECODE_MASTER"
                            SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            rsCourseCode.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsCourseCode.EOF Then
                                'Update Training List record with new renewal date, position, position start date, type of position
                                'Update Follow Up record and Continuing Education record as well
                                SQLQ = "SELECT * FROM HR_TRAIN"
                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                                '???SQLQ = SQLQ & " AND TR_JOB = '" & clpJob.Text & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsHRTrain.EOF Then
                                    'Keep the original Renewal Date for record retrieval from
                                    'Continuing Education screen
                                    xRenewalDt = rsHRTrain("TR_RENEW")
                                    
                                    If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                        Select Case rsCourseCode("ES_FLWUP_PRD_DWMY")
                                            Case "D"
                                                xDWMY = "d"
                                            Case "W"
                                                xDWMY = "ww"
                                            Case "M"
                                                xDWMY = "m"
                                            Case "Y"
                                                xDWMY = "yyyy"
                                        End Select
                                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsCourseCode("ES_RENEW_FOLLOWUP"), CVDate(rsEmpJobs("TW_SDATE")))
                                    Else
                                        Select Case rsCourseCode("ES_CUR_PRD_DWMY")
                                            Case "D"
                                                xDWMY = "d"
                                            Case "W"
                                                xDWMY = "ww"
                                            Case "M"
                                                xDWMY = "m"
                                            Case "Y"
                                                xDWMY = "yyyy"
                                        End Select
                                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsCourseCode("ES_RENEW_CRS_CUR"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                    End If
                                    rsHRTrain("TR_JOB") = rsEmpJobs("TW_JOB")
                                    rsHRTrain("TR_SDATE") = rsEmpJobs("TW_SDATE")
                                    rsHRTrain("TR_POS_TYPE") = Left(rsEmpJobs("POS_TYPE"), 1)
                                    ''If Renewal date is greater than today's date then clear the Course Taken Date
                                    'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                                    '    rsHRTrain("TR_COURSE_TAKEN") = Null
                                    'End If
                                    rsHRTrain("TR_LDATE") = Date
                                    rsHRTrain("TR_LUSER") = glbUserID
                                    rsHRTrain("TR_LTIME") = Time$
                                    
                                    'Ticket #24300
                                    'rsHRTrain.Update
                                    
                                    'If follow up id is null then find the id
                                    If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                        xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                        SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                        SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                                        SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                        SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsFollowUp.EOF Then
                                            rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                        End If
                                        rsFollowUp.Close
                                        Set rsFollowUp = Nothing
                                    End If
                                    rsReqCourse.Close
                                    Set rsReqCourse = Nothing
                                    
                                    'Ticket #24300
                                    If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                        'Add a Follow Up record for this Training course
                                        'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                        rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsCourseCode("ES_CRSCODE"), rsEmpJobs("TW_JOB"))
                                        
                                        rsHRTrain.Update
                                    Else
                                        rsHRTrain.Update
                                    
                                        'Update Follow Up record - Effective Date
                                        SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsFollowUp.EOF Then
                                            rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                            rsFollowUp("EF_COMMENTS") = "Course: " & rsCourseCode("ES_CRSCODE") & " - " & GetTABLDesc("ESCD", rsCourseCode("ES_CRSCODE")) & " for Position: " & rsEmpJobs("TW_JOB")
                                            rsFollowUp("EF_LDATE") = Date
                                            rsFollowUp("EF_LUSER") = glbUserID
                                            rsFollowUp("EF_LTIME") = Time$
                                            rsFollowUp.Update
                                        End If
                                        rsFollowUp.Close
                                        Set rsFollowUp = Nothing
                                    End If
                                    
                                    'Update the Continuing Education record for this course and this employee
                                    'with Renewal Date and Job Code
                                    SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                    SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                    '???SQLQ = SQLQ & " AND ES_JOB = '" & clpJob.Text & "'"
                                    SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                    SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(xRenewalDt)
                                    rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsContEdu.EOF Then
                                        rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                        rsContEdu("ES_JOB") = rsEmpJobs("TW_JOB")
                                        rsContEdu("ES_LDATE") = Date
                                        rsContEdu("ES_LUSER") = glbUserID
                                        rsContEdu("ES_LTIME") = Time$
                                        rsContEdu.Update
                                    End If
                                    rsContEdu.Close
                                    Set rsContEdu = Nothing
                                    
                                    'Update Temp/Cross Training Position record with Follow Up ID
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & rsEmpJobs("TW_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                                rsHRTrain.Close
                                Set rsHRTrain = Nothing
                            End If
                            rsCourseCode.Close
                            Set rsCourseCode = Nothing
                        
                        ElseIf (Not rsEmpJobs("TW_CURRENT")) And rsEmpJobs("TW_TRK_CRS_RENEWAL") Then
                            'If PREVIOUS
                            'Get Previous Renewal Period from Course Code Master to calculate the Renewal Date
                            SQLQ = "SELECT ES_CRSCODE,ES_RENEW_CRS_CUR,ES_CUR_PRD_DWMY,ES_RENEW_CRS_PRV,ES_PRV_PRD_DWMY,ES_RENEW_FOLLOWUP,ES_FLWUP_PRD_DWMY FROM HR_COURSECODE_MASTER"
                            SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            rsCourseCode.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsCourseCode.EOF Then
                                'Update Training List record with new renewal date, position, position start date, type of position
                                'Update Follow Up record and Continuing Education record as well
                                SQLQ = "SELECT * FROM HR_TRAIN"
                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                                '???SQLQ = SQLQ & " AND TR_JOB = '" & clpJob.Text & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsHRTrain.EOF Then
                                    'Keep the original Renewal Date for record retrieval from
                                    'Continuing Education screen
                                    xRenewalDt = rsHRTrain("TR_RENEW")
                                    
                                    If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                        Select Case rsCourseCode("ES_FLWUP_PRD_DWMY")
                                            Case "D"
                                                xDWMY = "d"
                                            Case "W"
                                                xDWMY = "ww"
                                            Case "M"
                                                xDWMY = "m"
                                            Case "Y"
                                                xDWMY = "yyyy"
                                        End Select
                                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsCourseCode("ES_RENEW_FOLLOWUP"), CVDate(rsEmpJobs("TW_SDATE")))
                                    Else
                                        Select Case rsCourseCode("ES_PRV_PRD_DWMY")
                                            Case "D"
                                                xDWMY = "d"
                                            Case "W"
                                                xDWMY = "ww"
                                            Case "M"
                                                xDWMY = "m"
                                            Case "Y"
                                                xDWMY = "yyyy"
                                        End Select
                                        If Not IsNull(rsCourseCode("ES_RENEW_CRS_CUR")) And rsCourseCode("ES_RENEW_CRS_CUR") <> 0 Then
                                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsCourseCode("ES_RENEW_CRS_PRV"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                        Else
                                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsCourseCode("ES_RENEW_CRS_PRV"), CVDate(rsEmpJobs("TW_ENDDATE")))
                                        End If
                                    End If
                                    rsHRTrain("TR_JOB") = rsEmpJobs("TW_JOB")
                                    rsHRTrain("TR_SDATE") = rsEmpJobs("TW_SDATE")
                                    rsHRTrain("TR_POS_TYPE") = "P"
                                    'If Renewal date is greater than today's date then clear the Course Taken Date
                                    'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                                    '    rsHRTrain("TR_COURSE_TAKEN") = Null
                                    'End If
                                    rsHRTrain("TR_LDATE") = Date
                                    rsHRTrain("TR_LUSER") = glbUserID
                                    rsHRTrain("TR_LTIME") = Time$
                                    
                                    'Ticket #24300
                                    'rsHRTrain.Update
                                    
                                    'If follow up id is null then find the id
                                    If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                        xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                        SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                        SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                                        SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                        SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsFollowUp.EOF Then
                                            rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                        End If
                                        rsFollowUp.Close
                                        Set rsFollowUp = Nothing
                                    End If
                                    rsReqCourse.Close
                                    Set rsReqCourse = Nothing
                                    
                                    'Ticket #24300
                                    If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                        'Add a Follow Up record for this Training course
                                        'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                        rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), rsEmpJobs("TW_JOB"))
                                        rsHRTrain.Update
                                    Else
                                        rsHRTrain.Update
                                    
                                        'Update Follow Up record - Effective Date
                                        SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsFollowUp.EOF Then
                                            rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                            rsFollowUp("EF_COMMENTS") = "Course: " & rsCourseCode("ES_CRSCODE") & " - " & GetTABLDesc("ESCD", rsCourseCode("ES_CRSCODE")) & " for Position: " & rsEmpJobs("TW_JOB")
                                            rsFollowUp("EF_LDATE") = Date
                                            rsFollowUp("EF_LUSER") = glbUserID
                                            rsFollowUp("EF_LTIME") = Time$
                                            rsFollowUp.Update
                                        End If
                                        rsFollowUp.Close
                                        Set rsFollowUp = Nothing
                                    End If
                                    
                                    'Update the Continuing Education record for this course and this employee
                                    'with Renewal Date and Job Code
                                    SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                    SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                    '???SQLQ = SQLQ & " AND ES_JOB = '" & clpJob.Text & "'"
                                    SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                    SQLQ = SQLQ & " AND ES_RENEW = " & Date_SQL(xRenewalDt)
                                    rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsContEdu.EOF Then
                                        rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                        rsContEdu("ES_JOB") = rsEmpJobs("TW_JOB")
                                        rsContEdu("ES_LDATE") = Date
                                        rsContEdu("ES_LUSER") = glbUserID
                                        rsContEdu("ES_LTIME") = Time$
                                        rsContEdu.Update
                                    End If
                                    rsContEdu.Close
                                    Set rsContEdu = Nothing
                                
                                    'Update Temp/Cross Training Position record with Follow Up ID
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & rsEmpJobs("TW_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                Else
                                    'Hemu - Testing
                                    'Data1.Recordset.Find "JH_ID=" & rsEmpJobs("TW_ID")
                                    'Call Display_Value
                                    'Call Update_Employee_Job_Training_List(rsEmpJobs("TW_JOB"), "Previous")
                                End If
                                rsHRTrain.Close
                                Set rsHRTrain = Nothing
                            End If
                            rsCourseCode.Close
                            Set rsCourseCode = Nothing
                        End If
                    End If
                    rsPosCourse.Close
                    Set rsPosCourse = Nothing
next_EmpJob:
                    rsEmpJobs.MoveNext
                Loop
            End If
            rsEmpJobs.Close
            Set rsEmpJobs = Nothing
            
    '        rsReqCourse.MoveNext
    '    Loop
    'End If
    'rsReqCourse.Close
    'Set rsReqCourse = Nothing

Exit Sub

Remaining_Tracked_Courses_Err:
If Err = 3018 Then
    Err = 0
    Resume Next
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
If Len(SQLQ) = 0 Then
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update_Remaining_Tracked_Courses", "HR_JOB_HISTORY", "Remaining_Tracked_Courses")
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, SQLQ, "HR_JOB_HISTORY", "Remaining_Tracked_Courses")
End If
Call RollBack '26July99 js
End Sub

Public Sub Update_Employee_Job_Training_List(xJob, xPosType, Optional xStartEndDate, Optional xEndDate)
    Dim rsHRTrain As New ADODB.Recordset
    Dim rsReqCourse As New ADODB.Recordset
    Dim rsFollowUp As New ADODB.Recordset
    Dim rsCourseMst As New ADODB.Recordset
    Dim rsContEdu As New ADODB.Recordset
    Dim rsTJob As New ADODB.Recordset
    Dim rsEmpJob As New ADODB.Recordset
    Dim xDWMY, xorgPosType, xorgJob As String
    Dim SQLQ  As String
    Dim flgUnqForPos, flgNoPrvRnwl, flgNoCurRnwl, flgCrsTakenBefore, flgProcCalled As Boolean
    Dim xPrvEndDate
    Dim xComments As String
    
    On Error GoTo Employee_Job_Training_Err

    'Note: If tracking is for the Previous Job then any courses for this job which does not have
    'Previous Renewal defined should be removed for this position or
    'If tracking is for Current Job then any courses for this job which does not have
    'Current Renewal defined should be removed for this position
    
    'if this procedure is called from another procedure and not an event
    If IsMissing(xStartEndDate) Then
        flgProcCalled = False
        xorgPosType = xPosType
        xorgJob = xJob
        xStartEndDate = ""
        xEndDate = ""
    Else
        flgProcCalled = True
    End If
    
    'Get the list of Required Courses for the Job
    SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE PC_JOB = '" & xJob & "'"
    
    'Ticket #25609 - Training Plan by Department
    'Only courses matching employee's Department if the Course has Department Code assigned
    SQLQ = SQLQ & " AND ((PC_DEPTNO IS NULL) OR (PC_DEPTNO = '" & GetEmpData(glbLEE_ID, "ED_DEPTNO") & "'))"
    
    rsReqCourse.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsReqCourse.EOF Then
        rsReqCourse.MoveFirst
        
        Do While Not rsReqCourse.EOF
            'Ticket #25609 - Training Plan by Department
            'Check if the Course has Department assigned. If so then check if the Department of the Course matches
            'employee's Department
            'If Not IsNull(rsReqCourse("PC_DEPTNO")) And rsReqCourse("PC_DEPTNO") <> "" Then
            '    If rsReqCourse("PC_DEPTNO") <> GetEmpData(glbLEE_ID, "ED_DEPTNO") Then
            '        'Skip this course as Employee does not belong to the department this Course is setup for
            '        GoTo Next_Required_Course
            '    End If
            'End If
        
            'Check if this required course is Unique for each Position.
            'If so, then it will have to be added in the Training List even
            'though the Course code already exists for this employee for other positions
            flgUnqForPos = False
            flgNoPrvRnwl = False
            flgNoCurRnwl = False
            SQLQ = "SELECT ES_CRSCODE,ES_UNIQUE_FOR_POS,ES_RENEW_CRS_CUR,ES_CUR_PRD_DWMY, ES_RENEW_CRS_PRV,ES_PRV_PRD_DWMY, ES_RENEW_FOLLOWUP, ES_FLWUP_PRD_DWMY FROM HR_COURSECODE_MASTER"
            SQLQ = SQLQ & " WHERE ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
            rsCourseMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsCourseMst.EOF Then
                flgUnqForPos = IIf(IsNull(rsCourseMst("ES_UNIQUE_FOR_POS")), False, rsCourseMst("ES_UNIQUE_FOR_POS"))
                
                If flgUnqForPos = False And rsCourseMst("ES_RENEW_FOLLOWUP") = 99 And rsCourseMst("ES_FLWUP_PRD_DWMY") = "Y" Then
                    'Skip this course
                    GoTo Next_Required_Course
                End If
            Else
                'Course not defined in the Course Code Master - skip this course
                GoTo Next_Required_Course
            End If
            'rsCourseMst.Close
            'Set rsCourseMst = Nothing
            
            'Follow Up Effective Date Period is mandatory. Check if it exists otherwise the logic below will give an error.
            If IsNull(rsReqCourse("PC_RENEW_FOLLOWUP")) Or rsReqCourse("PC_RENEW_FOLLOWUP") = "" Then
                'Follow Up Effective Date renewal Period missing
                GoTo Next_Required_Course
            End If
                        
            'Add the Required Courses in the Training List
            'if it does not already exists for this employee or Unique for each Position
            SQLQ = "SELECT * FROM HR_TRAIN"
            SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
            If flgUnqForPos <> 0 Then
                SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
                'If xPosType = "Previous" And chkTrackCrsRenewal And chkCurrent(0) Then
                '    SQLQ = SQLQ & " AND TR_POS_TYPE = 'C'"
                'Else
                '    If chkTrackCrsRenewal And chkCurrent(0) Then
                '        SQLQ = SQLQ & " AND TR_POS_TYPE = 'P'"
                '    Else
                '        SQLQ = SQLQ & " AND TR_POS_TYPE = '" & Left(xPosType, 1) & "'"
                '    End If
                'End If
            End If
            rsHRTrain.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsHRTrain.EOF Then
                'TRAINING RECORD DOES NOT EXISTS - ADD NEW ONE
                
                'Check first if this Course was taken before in the Continuing Education screen
                flgCrsTakenBefore = False
                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_JOB, ES_DATCOMP,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                If flgUnqForPos <> 0 Then
                    SQLQ = SQLQ & " AND ES_JOB = '" & xJob & "'"
                End If
                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                SQLQ = SQLQ & " AND (ES_RENEW = '' OR ES_RENEW IS NULL)"
                SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
                SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsContEdu.EOF Then
                    'Course Taken Before
                    rsContEdu.MoveFirst
                    flgCrsTakenBefore = True
                Else
                    'Course not taken before
                    flgCrsTakenBefore = False
                End If
                
                
                'May be Training List accidently deleted or messed up
                'if the Course is Previous and procedure not called from another procdure then
                'check if this course is required by another Primary or Temporary Current or Previous Position if so then
                'change the xJob to that Position and Start & Date Date to that Position Start Date & End Date
                If flgProcCalled = False And xPosType = "Previous" Then
                    'Check if Primary Current or Previous or Temp Current or other Previous required this Course
                    SQLQ = "SELECT JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
                    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
                    SQLQ = SQLQ & " AND JH_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    If flgUnqForPos <> 0 Then
                        SQLQ = SQLQ & " AND JH_JOB = '" & xJob & "'"
                    End If
                    SQLQ = SQLQ & " AND (JH_ID <> " & rsDATA!JH_ID & ")"
                    SQLQ = SQLQ & " UNION "
                    SQLQ = SQLQ & " SELECT TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
                    SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
                    SQLQ = SQLQ & " AND TW_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    If flgUnqForPos <> 0 Then
                        SQLQ = SQLQ & " AND TW_JOB = '" & xJob & "'"
                    End If
                    SQLQ = SQLQ & " ORDER BY TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
                    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsEmpJob.EOF Then
                        'The first record gets it
                        'the order is Primary Current, Temp Current and then Previous depending on most recent end date
                        rsEmpJob.MoveFirst
                        If Not IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                            If rsEmpJob("TW_TRK_CRS_RENEWAL") Then
                                '???If CVDate(rsEmpJob("TW_ENDDATE")) > CVDate(dlpENDDATE.Text) Then
                                    'Previous Position requires this course
                                    xJob = rsEmpJob("TW_JOB")
                                    xStartEndDate = rsEmpJob("TW_SDATE")
                                    xEndDate = rsEmpJob("TW_ENDDATE")
                                    xPosType = "Previous"
                                '???End If
                            Else
                                If rsEmpJob("POS_TYPE") = "CURRENT" Then
                                    If xJob <> rsEmpJob("TW_JOB") Then    'If Current becoming Previous
                                        xPosType = "Current"
                                        xJob = rsEmpJob("TW_JOB")
                                        xStartEndDate = rsEmpJob("TW_SDATE")
                                    End If
                                Else
                                    xPosType = "Temporary"
                                    xJob = rsEmpJob("TW_JOB")
                                    xStartEndDate = rsEmpJob("TW_SDATE")
                                End If
                            End If
                        Else
                            If rsEmpJob("POS_TYPE") = "CURRENT" Then
                                If xJob <> rsEmpJob("TW_JOB") Then    'If Current becoming Previous
                                    xPosType = "Current"
                                    xJob = rsEmpJob("TW_JOB")
                                    xStartEndDate = rsEmpJob("TW_SDATE")
                                End If
                            Else
                                xPosType = "Temporary"
                                xJob = rsEmpJob("TW_JOB")
                                xStartEndDate = rsEmpJob("TW_SDATE")
                            End If
                        End If
                    Else
                        xStartEndDate = ""  'Ticket #22951
                    End If
                    rsEmpJob.Close
                    Set rsEmpJob = Nothing
                Else
                    'if Current then do not do anything as Current record takes precedence
                End If
                
                'If the course is being added for the Previous Position and this course
                'does not have previous renewal period then do not add this course
                'If xPosType = "Current" Or (xPosType = "Previous" And (Not IsNull(rsReqCourse("PC_RENEW_CRS_PRV"))) And rsReqCourse("PC_RENEW_CRS_PRV") <> 0) Then
                
                'If Course was taken and it's Position is Current then
                'make sure Current Renewal Period is there otherwise do not add the course
                'If the course is being added for the Previous Position and this course
                'does not have previous renewal period then do not add this course
                'Changed
                If (flgCrsTakenBefore = True And (xPosType = "Current" Or xPosType = "Temporary") And (Not IsNull(rsReqCourse("PC_RENEW_CRS_CUR"))) And rsReqCourse("PC_RENEW_CRS_CUR") <> 0) Or _
                    (flgCrsTakenBefore = False And (xPosType = "Current" Or xPosType = "Temporary")) Or (flgCrsTakenBefore = True And xPosType = "Previous" And (Not IsNull(rsReqCourse("PC_RENEW_CRS_PRV"))) And rsReqCourse("PC_RENEW_CRS_PRV") <> 0) Or _
                    (flgCrsTakenBefore = False And xPosType = "Previous") Then
                    
                    'Add Training Record
                    rsHRTrain.AddNew
                    rsHRTrain("TR_COMPNO") = "001"
                    rsHRTrain("TR_EMPNBR") = glbLEE_ID
                    rsHRTrain("TR_CRSCODE") = rsReqCourse("PC_CRSCODE")
                    
                    If flgCrsTakenBefore = False Then
                        If Not IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) And rsReqCourse("PC_RENEW_CRS_CUR") <> 0 Then
                            'Current Course Renewal found
                            xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                            Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            If xPosType = "Current" Or xPosType = "Temporary" Or xPosType = "Previous" Then
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    '???rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
                                End If
                                
                            'For courses not taken and are now Previous, the renewal date is based
                            'on Follow Up Renewal Period and not Previous Renewal Period - above
                            'ElseIf xPosType = "Previous" Then
                            '    Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                            '        Case "D"
                            '            xDWMY = "d"
                            '        Case "W"
                            '            xDWMY = "ww"
                            '        Case "M"
                            '            xDWMY = "m"
                            '        Case "Y"
                            '            xDWMY = "yyyy"
                            '    End Select
                            '    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(dlpStartDate.Text))
                            End If
                        Else    'No Current Course Renewal Period
                            If xPosType = "Current" Or xPosType = "Temporary" Or xPosType = "Previous" Then
                                xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    '???rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
                                End If
                            'ElseIf xPosType = "Previous" Then
                            '    'For courses not taken and are now Previous, the renewal date is based
                            '    'on Follow Up Renewal Period and not Previous Renewal Period.
                            '    'If there is no current renewal then it's based on End Date only and
                            '    'Prev Renewal Period - for courses taken.
                            '    'Compute Renewal with Position End Date because there is no Current Renewal Period defined
                            '    Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                            '        Case "D"
                            '            xDWMY = "d"
                            '        Case "W"
                            '            xDWMY = "ww"
                            '        Case "M"
                            '            xDWMY = "m"
                            '        Case "Y"
                            '            xDWMY = "yyyy"
                            '    End Select
                            '    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(dlpENDDATE.Text))
                            End If
                        End If
                    Else    'Course Has Been Taken Before
                        'Course has been taken before, compute Renewal Date based on Course Taken Date
                        If xPosType = "Current" Or xPosType = "Temporary" Then
                            Select Case rsReqCourse("PC_CUR_PRD_DWMY")
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_CUR"), CVDate(rsContEdu("ES_DATCOMP")))
                            rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                        ElseIf xPosType = "Previous" Then
                            xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                            Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                                Case "D"
                                    xDWMY = "d"
                                Case "W"
                                    xDWMY = "ww"
                                Case "M"
                                    xDWMY = "m"
                                Case "Y"
                                    xDWMY = "yyyy"
                            End Select
                            If Not IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) And rsReqCourse("PC_RENEW_CRS_CUR") <> 0 Then
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(rsContEdu("ES_DATCOMP")))
                            Else
                                If IsMissing(xEndDate) Or xEndDate = "" Or IsNull(xEndDate) Then
                                    '???rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(dlpENDDATE.Text))
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(xEndDate))
                                End If
                            End If
                            rsHRTrain("TR_COURSE_TAKEN") = rsContEdu("ES_DATCOMP")  'Since adding the course back based on last Complete Date - put the last Complete Date as well
                        End If
                        
                        'Update Continuing Education with new Renewal Date
                        rsContEdu("ES_JOB") = xJob
                        rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                        rsContEdu("ES_LDATE") = Date
                        rsContEdu("ES_LUSER") = glbUserID
                        rsContEdu("ES_LTIME") = Time$
                        rsContEdu.Update
                    End If
                    
                    rsHRTrain("TR_JOB") = xJob
                    If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                        '???rsHRTrain("TR_SDATE") = dlpStartDate.Text
                    Else
                        rsHRTrain("TR_SDATE") = xStartEndDate
                    End If
                    If xPosType = "Current" Then
                        rsHRTrain("TR_POS_TYPE") = "C"
                    ElseIf xPosType = "Temporary" Then
                        rsHRTrain("TR_POS_TYPE") = "T"
                    ElseIf xPosType = "Previous" Then
                        rsHRTrain("TR_POS_TYPE") = "P"
                    End If
                    'rsHRTrain("TR_COURSE_TAKEN")   - Remains BLANK
                    rsHRTrain("TR_LDATE") = Date
                    rsHRTrain("TR_LTIME") = Time$
                    rsHRTrain("TR_LUSER") = glbUserID
                    
                    'Add a Follow Up record for this Training course
'                    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE 1 = 2"
'                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                    rsFollowUp.AddNew
'                    rsFollowUp("EF_COMPNO") = "001"
'                    rsFollowUp("EF_EMPNBR") = glbLEE_ID
'                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
'                    rsFollowUp("EF_FREAS_TABL") = "FURE"
'                    'Ticket #24257 - Do not update Admin By for them only
'                    If glbCompSerial <> "S/N - 2262W" Then
'                        rsFollowUp("EF_ADMINBY_TABL") = "EDAB"
'                        rsFollowUp("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
'                    End If
'                    rsFollowUp("EF_FREAS") = "EDUC"
'                    rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
'                    rsFollowUp("EF_LDATE") = Date
'                    rsFollowUp("EF_LTIME") = Time$
'                    rsFollowUp("EF_LUSER") = glbUserID
'                    rsFollowUp.Update
                    
                    'Ticket #24300
                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                    
                    rsHRTrain.Update
                    
                    'rsFollowUp.Close
                    'Set rsFollowUp = Nothing
                
                    'Update Position record with Follow Up ID
                    'if the course code is TRAIN
                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                        'Search HR_JOB_HISTORY table for this Position record
                        'and update with Follow Up Id
                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsTJob.EOF Then
                            rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                            rsTJob.Update
                        End If
                        rsTJob.Close
                        Set rsTJob = Nothing
                    End If
                End If
                rsContEdu.Close
                Set rsContEdu = Nothing
                
                If flgProcCalled = False Then
                    xPosType = xorgPosType
                    xJob = xorgJob
                End If
            Else
                'TRAINING RECORD FOUND
                
                'May be Training List accidently deleted or messed up
                'if the Course is Previous and procedure not called from another procdure then
                'check if this course is required by another Primary or Temporary Current or Previous Position if so then
                'change the xJob to that Position and Start & Date Date to that Position Start Date & End Date
                If flgProcCalled = False And xPosType = "Previous" Then
                    'Check if Primary Current or Previous or Temp Current or other Previous required this Course
                    SQLQ = "SELECT JH_EMPNBR AS TW_EMPNBR, 'CURRENT' AS POS_TYPE, JH_JOB AS TW_JOB, JH_CURRENT AS TW_CURRENT, JH_TRK_CRS_RENEWAL AS TW_TRK_CRS_RENEWAL, JH_SDATE AS TW_SDATE, JH_ENDDATE AS TW_ENDDATE FROM HR_JOB_HISTORY "
                    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID & " AND ((JH_CURRENT <> 0) OR (JH_TRK_CRS_RENEWAL <> 0))"
                    SQLQ = SQLQ & " AND JH_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    If flgUnqForPos <> 0 Then
                        SQLQ = SQLQ & " AND JH_JOB = '" & xJob & "'"
                    End If
                    SQLQ = SQLQ & " AND (JH_ID <> " & rsDATA!JH_ID & ")"
                    SQLQ = SQLQ & " UNION "
                    SQLQ = SQLQ & " SELECT TW_EMPNBR, 'TEMPORARY' AS POS_TYPE, TW_JOB, TW_CURRENT, TW_TRK_CRS_RENEWAL,TW_SDATE,TW_ENDDATE FROM HR_TEMP_WORK "
                    SQLQ = SQLQ & " WHERE TW_EMPNBR = " & glbLEE_ID & " AND ((TW_CURRENT <> 0) OR (TW_TRK_CRS_RENEWAL <> 0))"
                    SQLQ = SQLQ & " AND TW_JOB IN (SELECT PC_JOB FROM HR_JOB_COURSE WHERE PC_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                    If flgUnqForPos <> 0 Then
                        SQLQ = SQLQ & " AND TW_JOB = '" & xJob & "'"
                    End If
                    SQLQ = SQLQ & " ORDER BY TW_TRK_CRS_RENEWAL ASC,POS_TYPE ASC,TW_CURRENT DESC,TW_ENDDATE DESC"
                    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsEmpJob.EOF Then
                        'The first record gets it
                        'the order is Primary Current, Temp Current and then Previous depending on most recent end date
                        rsEmpJob.MoveFirst
                        If Not IsNull(rsEmpJob("TW_TRK_CRS_RENEWAL")) Then
                            If rsEmpJob("TW_TRK_CRS_RENEWAL") Then
                                '???If CVDate(rsEmpJob("TW_ENDDATE")) > CVDate(dlpENDDATE.Text) Then
                                    'Previous Position requires this course
                                    xJob = rsEmpJob("TW_JOB")
                                    xStartEndDate = rsEmpJob("TW_SDATE")
                                    xEndDate = rsEmpJob("TW_ENDDATE")
                                    xPosType = "Previous"
                                '???End If
                            Else
                                If rsEmpJob("POS_TYPE") = "CURRENT" Then
                                    If xJob <> rsEmpJob("TW_JOB") Then    'If Current becoming Previous
                                        xPosType = "Current"
                                        xJob = rsEmpJob("TW_JOB")
                                        xStartEndDate = rsEmpJob("TW_SDATE")
                                    End If
                                Else
                                    xPosType = "Temporary"
                                    xJob = rsEmpJob("TW_JOB")
                                    xStartEndDate = rsEmpJob("TW_SDATE")
                                End If
                            End If
                        Else
                            If rsEmpJob("POS_TYPE") = "CURRENT" Then
                                If xJob <> rsEmpJob("TW_JOB") Then    'If Current becoming Previous
                                    xPosType = "Current"
                                    xJob = rsEmpJob("TW_JOB")
                                    xStartEndDate = rsEmpJob("TW_SDATE")
                                End If
                            Else
                                xPosType = "Temporary"
                                xJob = rsEmpJob("TW_JOB")
                                xStartEndDate = rsEmpJob("TW_SDATE")
                            End If
                        End If
                    Else
                        xStartEndDate = ""  'Ticket #22951
                    End If
                    rsEmpJob.Close
                    Set rsEmpJob = Nothing
                Else
                    'if Current then do not do anything as Current record takes precedence
                End If
                
                
                
                'Training record for this course already exists so update the Renewal Date
                'Check which Type of Position is assigned to this course
                If rsHRTrain("TR_POS_TYPE") = "C" Then
                    'Currently the course is holding Primary Current Position Code
                    'Check which type of position requires this course
                    If xPosType = "Current" Then
                        'These courses are for new Current Primary Position so recalculate the
                        'Renewal Dates - based on Position Start Date or last Course Taken date
                        'See which Position Start Date is most recent
                        '???If CVDate(rsHRTrain("TR_SDATE")) < CVDate(IIf(IsMissing(xStartEndDate) Or xStartEndDate = "", dlpStartDate.Text, xStartEndDate)) Then
                            'Training List has older Position Start Date so update with new Position info.
                            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    '???rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
                                End If
                            Else
                                'Check if Current Renewal period is defined
                                If IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0 Then
                                    'No Current Renewal Period defined so delete this job from this current position.
                                    'It should not be in the training list for any current job
                                    flgNoCurRnwl = True
                                Else
                                    Select Case rsReqCourse("PC_CUR_PRD_DWMY")
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_CUR"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                End If
                            End If
                            If flgNoCurRnwl = False Then
                                'Update Continuing Education record as well
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                'Ticket #28673 Franks 05/27/2016 changed "=" to "IN"
                                'SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW IN (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                'Ticket #28673 Franks 05/27/2016 changed "=" to "IN"
                                'SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_DATCOMP IN (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_JOB") = xJob
                                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                            
                                rsHRTrain("TR_JOB") = xJob
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    '???rsHRTrain("TR_SDATE") = dlpStartDate.Text
                                Else
                                    rsHRTrain("TR_SDATE") = xStartEndDate
                                End If
                                rsHRTrain("TR_POS_TYPE") = "C"   'Current Primary
                                ''If Renewal date is greater than today's date then clear the Course Taken Date
                                'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                                '    rsHRTrain("TR_COURSE_TAKEN") = Null
                                'End If
                                rsHRTrain("TR_LDATE") = Date
                                rsHRTrain("TR_LUSER") = glbUserID
                                rsHRTrain("TR_LTIME") = Time$
                                
                                'If follow up id is null then find the id
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                    'SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND EF_FDATE IN (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    'Add a Follow Up record for this Training course
                                    'Ticket #24300
                                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                    
                                    rsHRTrain.Update
                                Else
                                    rsHRTrain.Update
                                
                                    'Update Follow Up record - Effective Date
                                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                        rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                        rsFollowUp("EF_LDATE") = Date
                                        rsFollowUp("EF_LUSER") = glbUserID
                                        rsFollowUp("EF_LTIME") = Time$
                                        rsFollowUp.Update
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Update Position record with Follow Up ID
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and update with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                
                                    'Clear the Follow Up Id on the other current position rec in the Temp Position table
                                    'Search HR_JOB_HISTORY table for this Position record
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            Else
                                'CURRENT - Current
                                'No Current renewal found for this course
                                
                                'Clear the Renewal date for this course and for this employee from
                                'Continuing Education screen
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                
                                    If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                                        'Since the course was completed - mark the Follow Up as
                                        'Completed instead of deleting it.
                                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    Else
                                        'Delete the Follow Up record for this training record
                                        'as no Course completion record found
                                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    
                                        'Clear the Follow Up Id on the Position record
                                        'if the course code is TRAIN
                                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                            'Search HR_JOB_HISTORY table for this Position record
                                            'and update with Follow Up Id
                                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                            If Not rsTJob.EOF Then
                                                rsTJob("JH_FOLLOWUP_ID") = Null
                                                rsTJob.Update
                                            End If
                                            rsTJob.Close
                                            Set rsTJob = Nothing
                                        End If
                                    End If
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                
                                    'Clear the Follow Up ID in the Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                
                                'Delete this Training List record as the course is not required by other positions
                                SQLQ = "DELETE FROM HR_TRAIN"
                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                gdbAdoIhr001.Execute SQLQ
                            End If
                        '???Else
                            'Do not do anything because Training List has most recent Position Start Date
                        '???End If
                    ElseIf xPosType = "Previous" Then
                        'CURRENT - Previous
                        'Current Job becoming Previous
                        'Previous Primary Position is being tracked but Current Primary Position has this course
                        'Check if the Position in HR_TRAIN is same this Position
                        '???If (rsHRTrain("TR_JOB") <> xJob) Or (rsHRTrain("TR_JOB") = xJob And CVDate(rsHRTrain("TR_SDATE")) <> CVDate(dlpStartDate.Text) And CVDate(rsHRTrain("TR_SDATE")) <> CVDate(IIf(IsMissing(xStartEndDate) Or xStartEndDate = "", dlpStartDate.Text, xStartEndDate))) Then
                            'Do not do anything because Current takes the priority
                        '???Else
                            'Renewal Date based on last Course Taken date if present
                            'otherwise Follow Up Effective Date Period
                            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    '???rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
                                End If
                            Else
                                'Change the renewal dates if Previous renewal is defined
                                If IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) Or rsReqCourse("PC_RENEW_CRS_PRV") = 0 Then
                                    'No Previous Renewal Period defined so delete this job from this previous position.
                                    'It should not be in the training list for any previous job
                                    flgNoPrvRnwl = True
                                Else
                                    Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                End If
                            End If
                            If flgNoPrvRnwl = False Then
                                'Update Continuing Education record as well
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                'Ticket #28673 Franks 05/27/2016 changed "=" to "IN"
                                'SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW IN (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                'Ticket #28673 Franks 05/27/2016 changed "=" to "IN"
                                'SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_DATCOMP IN (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_JOB") = xJob
                                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                            
                                'Previous Renewal period available
                                rsHRTrain("TR_JOB") = xJob
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    '???rsHRTrain("TR_SDATE") = dlpStartDate.Text
                                Else
                                    rsHRTrain("TR_SDATE") = xStartEndDate
                                End If
                                rsHRTrain("TR_POS_TYPE") = "P"   'Previous Primary
                                ''If Renewal date is greater than today's date then clear the Course Taken Date
                                'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                                '    rsHRTrain("TR_COURSE_TAKEN") = Null
                                'End If
                                rsHRTrain("TR_LDATE") = Date
                                rsHRTrain("TR_LUSER") = glbUserID
                                rsHRTrain("TR_LTIME") = Time$
                                
                                'If follow up id is null then find the id
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                    'SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND EF_FDATE IN (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Ticket #24300
                                'rsHRTrain.Update
                                
                                'Ticket #24300
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    'Add a Follow Up record for this Training course
                                    'Ticket #24300
                                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                    rsHRTrain.Update
                                Else
                                    rsHRTrain.Update
                                    'Update Follow Up record - Effective Date
                                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                        rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                        rsFollowUp("EF_LDATE") = Date
                                        rsFollowUp("EF_LUSER") = glbUserID
                                        rsFollowUp("EF_LTIME") = Time$
                                        rsFollowUp.Update
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                            
                                'Update Position record with Follow Up ID
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and update with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            Else
                                'CURRENT - Previous
                                'No Previous renewal found for this course
                                
                                'Clear the Renewal date for this course and for this employee from
                                'Continuing Education screen
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                
                                    If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                                        'Since the course was completed - mark the Follow Up as
                                        'Completed instead of deleting it.
                                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    Else
                                        'Delete the Follow Up record for this training record
                                        'as no Course completion record found
                                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    
                                        'Clear the Follow Up ID in the Position record
                                        'if the course code is TRAIN
                                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                            'Search HR_JOB_HISTORY table for this Position record
                                            'and update with Follow Up Id
                                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                            If Not rsTJob.EOF Then
                                                rsTJob("JH_FOLLOWUP_ID") = Null
                                                rsTJob.Update
                                            End If
                                            rsTJob.Close
                                            Set rsTJob = Nothing
                                        End If
                                    End If
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                
                                    'Clear the Follow Up ID in the Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                
                                'Delete this Training List record as the course is not required by other positions
                                SQLQ = "DELETE FROM HR_TRAIN"
                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                gdbAdoIhr001.Execute SQLQ
                            End If
                        '???End If
                    End If
                ElseIf rsHRTrain("TR_POS_TYPE") = "T" Then
                    'Currently the Temporary Current Position is holding this course
                    'Check which type of position requires this course now
                    If xPosType = "Current" Then
                        'These courses are for new Current Primary Position so recalculate the
                        'Renewal Dates - based on Position Start Date or last Course Taken date
                        'See which Position Start Date is most recent
                        'If CVDate(rsHRTrain("TR_SDATE")) <= CVDate(IIf(IsMissing(xStartEndDate) Or xStartEndDate = "", dlpStartDate.Text, xStartEndDate)) Then
                            'Training List has older Position Start Date so update with new Position info.
                            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    '???rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
                                End If
                            Else
                                'Check if Current Renewal period is defined
                                If IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0 Then
                                    'No Current Renewal Period defined so delete this job from this current position.
                                    'It should not be in the training list for any current job
                                    flgNoCurRnwl = True
                                Else
                                    Select Case rsReqCourse("PC_CUR_PRD_DWMY")
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_CUR"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                End If
                            End If
                            If flgNoCurRnwl = False Then
                                'Update Continuing Education record as well
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                'Ticket #28673 Franks 05/27/2016 changed "=" to "IN"
                                'SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW IN (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                  SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                'Ticket #28673 Franks 05/27/2016 changed "=" to "IN"
                                'SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_DATCOMP IN (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_JOB") = xJob
                                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                                        
                                rsHRTrain("TR_JOB") = xJob
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    '???rsHRTrain("TR_SDATE") = dlpStartDate.Text
                                Else
                                    rsHRTrain("TR_SDATE") = xStartEndDate
                                End If
                                rsHRTrain("TR_POS_TYPE") = "C"   'Current Primary
                                ''If Renewal date is greater than today's date then clear the Course Taken Date
                                'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                                '    rsHRTrain("TR_COURSE_TAKEN") = Null
                                'End If
                                rsHRTrain("TR_LDATE") = Date
                                rsHRTrain("TR_LUSER") = glbUserID
                                rsHRTrain("TR_LTIME") = Time$
                                
                                'If follow up id is null then find the id
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                    'SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND EF_FDATE IN (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Ticket #24300
                                'rsHRTrain.Update
                                
                                'Ticket #24300
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    'Add a Follow Up record for this Training course
                                    'Ticket #24300
                                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                    rsHRTrain.Update
                                Else
                                    rsHRTrain.Update
                                    
                                    'Update Follow Up record - Effective Date
                                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                        rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                        rsFollowUp("EF_LDATE") = Date
                                        rsFollowUp("EF_LUSER") = glbUserID
                                        rsFollowUp("EF_LTIME") = Time$
                                        rsFollowUp.Update
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Update Position record with Follow Up ID
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and update with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                    
                                    'Clear the Follow Up Id on the position in the Temp/Cross Training Position table
                                    'Search HR_TEMP_WORK table for this Position record
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            Else
                                'TEMPORARY - Current
                                'No Current renewal found for this course
                                
                                'Clear the Renewal date for this course and for this employee from
                                'Continuing Education screen
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                
                                    If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                                        'Since the course was completed - mark the Follow Up as
                                        'Completed instead of deleting it.
                                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    Else
                                        'Delete the Follow Up record for this training record
                                        'as no Course completion record found
                                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                        
                                        'Clear the Follow Up ID in the Position record
                                        'if the course code is TRAIN
                                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                            'Search HR_TEMP_WORK table for this Position record
                                            'and clear with Follow Up Id
                                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                            If Not rsTJob.EOF Then
                                                rsTJob("TW_FOLLOWUP_ID") = Null
                                                rsTJob.Update
                                            End If
                                            rsTJob.Close
                                            Set rsTJob = Nothing
                                        End If
                                    End If
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                
                                    'Clear the Follow Up ID in the Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_TEMP_WORK table for this Position record
                                        'and update with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("TW_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                
                                'Delete this Training List record as the course is not required by other positions
                                SQLQ = "DELETE FROM HR_TRAIN"
                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                gdbAdoIhr001.Execute SQLQ
                            End If
                        'Else
                            'Do not do anything because Training List has most recent Position Start Date
                        'End If
                    ElseIf xPosType = "Previous" Then
                        'TEMPORARY - Previous
                        'Do not do anything because Training List record is of the Current
                        'Temporary/Cross Training Position
                    
'                        'Previous Primary Position is being tracked but Temp. Current Position is holding this course
'                        'Check if the Position in HR_TRAIN is same this Position
'                        If rsHRTrain("TR_JOB") <> xJob Then
'                            'Do not do anything because Current takes the  priority
'                        Else
'                            'Change the renewal dates if Previous renewal is defined
'                            If IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) Or rsReqCourse("PC_RENEW_CRS_PRV") = 0 Then
'                                'No Previous Renewal Period defined so delete this job from this previous position.
'                                'It should not be in the training list for any previous job
'                                flgNoPrvRnwl = True
'                            Else
'                                'Renewal Date based on last Course Taken date if present
'                                'otherwise Follow Up Effective Date Period
'                                If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
'                                    Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
'                                        Case "D"
'                                            xDWMY = "d"
'                                        Case "W"
'                                            xDWMY = "ww"
'                                        Case "M"
'                                            xDWMY = "m"
'                                        Case "Y"
'                                            xDWMY = "yyyy"
'                                    End Select
'                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
'                                Else
'                                    Select Case rsReqCourse("PC_PRV_PRD_DWMY")
'                                        Case "D"
'                                            xDWMY = "d"
'                                        Case "W"
'                                            xDWMY = "ww"
'                                        Case "M"
'                                            xDWMY = "m"
'                                        Case "Y"
'                                            xDWMY = "yyyy"
'                                    End Select
'                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
'                                End If
'                            End If
'                            If flgNoPrvRnwl = False Then
'                                'Previous Renewal period available
'                                rsHRTrain("TR_JOB") = xJob
'                                rsHRTrain("TR_SDATE") = dlpStartDate.Text
'                                rsHRTrain("TR_POS_TYPE") = "P"   'Previous Primary
'                                ''If Renewal date is greater than today's date then clear the Course Taken Date
'                                'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
'                                '    rsHRTrain("TR_COURSE_TAKEN") = Null
'                                'End If
'                                rsHRTrain("TR_LDATE") = Date
'                                rsHRTrain("TR_LUSER") = glbUserID
'                                rsHRTrain("TR_LTIME") = Time$
'                                rsHRTrain.Update
'
'                                'Update Follow Up record - Effective Date
'                                SQLQ = "SELECT * FROM HR_FOLLOW_UP"
'                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
'                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                If Not rsFollowUp.EOF Then
'                                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
'                                    rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
'                                    rsFollowUp("EF_LDATE") = Date
'                                    rsFollowUp("EF_LUSER") = glbUserID
'                                    rsFollowUp("EF_LTIME") = Time$
'                                    rsFollowUp.Update
'                                End If
'                                rsFollowUp.Close
'                                Set rsFollowUp = Nothing
'
'                                'Update Position record with Follow Up ID
'                                'if the course code is TRAIN
'                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
'                                    'Search HR_JOB_HISTORY table for this Position record
'                                    'and update with Follow Up Id
'                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("TW_ID")
'                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                    If Not rsTJob.EOF Then
'                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
'                                        rsTJob.Update
'                                    End If
'                                    rsTJob.Close
'                                    Set rsTJob = Nothing
'
'                                    'Clear the Follow Up Id on the position in the Temp/Cross Training Position table
'                                    'Search HR_TEMP_WORK table for this Position record
'                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
'                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                    If Not rsTJob.EOF Then
'                                        rsTJob("TW_FOLLOWUP_ID") = Null
'                                        rsTJob.Update
'                                    End If
'                                    rsTJob.Close
'                                    Set rsTJob = Nothing
'                                End If
'                            Else
'                                'No Previous renewal found for this course
'
'                                'Clear the Renewal date for this course and for this employee from
'                                'Continuing Education screen
'                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
'                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
'                                SQLQ = SQLQ & " AND ES_JOB = '" & xJob & "'"
'                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
'                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & xJob & "'"
'                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & xJob & "'"
'                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                If Not rsContEdu.EOF Then
'                                    rsContEdu("ES_RENEW") = Null
'                                    rsContEdu("ES_LDATE") = Date
'                                    rsContEdu("ES_LUSER") = glbUserID
'                                    rsContEdu("ES_LTIME") = Time$
'                                    rsContEdu.Update
'
'                                    If Not IsNull(rsContEdu("ES_DATCOMP")) Then
'                                        'Since the course was completed - mark the Follow Up as
'                                        'Completed instead of deleting it.
'                                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP"))
'                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & xJob & "'"
'                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                                        gdbAdoIhr001.Execute SQLQ
'                                    Else
'                                        'Delete the Follow Up record for this training record
'                                        'as no Course completion record found
'                                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
'                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & xJob & "'"
'                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                                        gdbAdoIhr001.Execute SQLQ
'
'                                        'Clear the Follow Up ID in the Position record
'                                        'if the course code is TRAIN
'                                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
'                                            'Search HR_TEMP_WORK table for this Position record
'                                            'and clear with Follow Up Id
'                                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
'                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                            If Not rsTJob.EOF Then
'                                                rsTJob("TW_FOLLOWUP_ID") = Null
'                                                rsTJob.Update
'                                            End If
'                                            rsTJob.Close
'                                            Set rsTJob = Nothing
'                                        End If
'                                    End If
'                                Else
'                                    'Delete the Follow Up record for this training record
'                                    'as no Course completion record found
'                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
'                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & xJob & "'"
'                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
'                                    gdbAdoIhr001.Execute SQLQ
'
'                                    'Clear the Follow Up ID in the Position record
'                                    'if the course code is TRAIN
'                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
'                                        'Search HR_TEMP_WORK table for this Position record
'                                        'and clear with Follow Up Id
'                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
'                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                                        If Not rsTJob.EOF Then
'                                            rsTJob("TW_FOLLOWUP_ID") = Null
'                                            rsTJob.Update
'                                        End If
'                                        rsTJob.Close
'                                        Set rsTJob = Nothing
'                                    End If
'                                End If
'                                rsContEdu.Close
'                                Set rsContEdu = Nothing
'
'                                'Delete this Training List record as the course is not required by other positions
'                                SQLQ = "DELETE FROM HR_TRAIN"
'                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
'                                SQLQ = SQLQ & " AND TR_JOB = '" & xJob & "'"
'                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
'                                gdbAdoIhr001.Execute SQLQ
'                            End If
'                        End If
                    End If
                ElseIf rsHRTrain("TR_POS_TYPE") = "P" Then
                    'Previous Primary or Temporary position is holding this course
                    If xPosType = "Current" Then
                        'This course is required by new Current Primary Position so recalculate the
                        'Renewal Dates
                        If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                            'Check if Current Renewal period is defined
                            If IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0 Then
                                'No Current Renewal Period defined so delete this job from this current position.
                                'It should not be in the training list for any current job
                                flgNoCurRnwl = True
                                
                                'Check if the Course was taken before ever
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_LDATE,ES_LTIME,ES_LUSER FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                'SQLQ = SQLQ & " AND ES_JOB = '" & xJob & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND (ES_RENEW = '' OR ES_RENEW IS NULL)"
                                SQLQ = SQLQ & " AND (ES_DATCOMP IS NOT NULL)"
                                SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'Course Taken Before
                                    flgNoCurRnwl = True
                                Else
                                    'Course not taken before
                                    flgNoCurRnwl = False
                                    xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                                    Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                        '???rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
                                    Else
                                        rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
                                    End If
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                            Else
                                xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    '???rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
                                End If
                            End If
                        Else
                            'Check if Current Renewal period is defined
                            If IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0 Then
                                'No Current Renewal Period defined so delete this job from this current position.
                                'It should not be in the training list for any current job
                                flgNoCurRnwl = True
                            Else
                                Select Case rsReqCourse("PC_CUR_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_CUR"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                            End If
                        End If
                        If flgNoCurRnwl = False Then
                            'Update Continuing Education record as well
                            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                            SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            'Ticket #28673 Franks 05/27/2016 changed "=" to "IN"
                            'SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND ES_RENEW IN (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            'Ticket #28673 Franks 05/27/2016 changed "=" to "IN"
                            'SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND ES_DATCOMP IN (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsContEdu.EOF Then
                                'rsContEdu("ES_RENEW") = Null
                                rsContEdu("ES_JOB") = xJob
                                rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                rsContEdu("ES_LDATE") = Date
                                rsContEdu("ES_LUSER") = glbUserID
                                rsContEdu("ES_LTIME") = Time$
                                rsContEdu.Update
                            End If
                            rsContEdu.Close
                            Set rsContEdu = Nothing
                            
                            rsHRTrain("TR_JOB") = xJob
                            If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                '???rsHRTrain("TR_SDATE") = dlpStartDate.Text
                            Else
                                rsHRTrain("TR_SDATE") = xStartEndDate
                            End If
                            rsHRTrain("TR_POS_TYPE") = "C"   'Current Primary
                            ''If Renewal date is greater than today's date then clear the Course Taken Date
                            'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                            '    rsHRTrain("TR_COURSE_TAKEN") = Null
                            'End If
                            rsHRTrain("TR_LDATE") = Date
                            rsHRTrain("TR_LUSER") = glbUserID
                            rsHRTrain("TR_LTIME") = Time$
                            
                            'If follow up id is null then find the id
                            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                'SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND EF_FDATE IN (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                            
                            'Ticket #24300
                            'rsHRTrain.Update
                            
                            'Ticket #24300
                            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                'Add a Follow Up record for this Training course
                                'Ticket #24300
                                'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                rsHRTrain.Update
                            Else
                                rsHRTrain.Update
                                
                                'Update Follow Up record - Effective Date
                                SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                    rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                    rsFollowUp("EF_LDATE") = Date
                                    rsFollowUp("EF_LUSER") = glbUserID
                                    rsFollowUp("EF_LTIME") = Time$
                                    rsFollowUp.Update
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                        
                            'Update Position record with Follow Up ID
                            'if the course code is TRAIN
                            If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                'Clear the Follow Up from the Previous Job in Primary/Temp Position
                                'Search HR_JOB_HISTORY table for this Position record
                                'and clear with Follow Up Id
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("JH_FOLLOWUP_ID") = Null
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                                
                                'Since Previous in HR_TRAIN can be Primary or Temp Position
                                'Search HR_TEMP_WORK table for this Position record
                                'and clear with Follow Up Id
                                SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("TW_FOLLOWUP_ID") = Null
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                                
                                'Search HR_JOB_HISTORY table for this Position record
                                'and update with Follow Up Id
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                            End If
                        Else
                            'No Current renewal found for this course - Correct logic - confirmed with email -March 09, 2009 1:18 PM
                                                        
                            'Clear the Renewal date for this course and for this employee from
                            'Continuing Education screen
                            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                            SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            'SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & xJob & "'"
                            'SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            If Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            End If
                            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsContEdu.EOF Then
                                rsContEdu("ES_RENEW") = Null
                                rsContEdu("ES_LDATE") = Date
                                rsContEdu("ES_LUSER") = glbUserID
                                rsContEdu("ES_LTIME") = Time$
                                rsContEdu.Update
                                
                                If Not IsNull(rsContEdu("ES_DATCOMP")) And Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                    'Since the course was completed - mark the Follow Up as
                                    'Completed instead of deleting it.
                                    SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                
                                    'Clear the Follow Up ID in the Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                        
                                        'Since Previous in HR_TRAIN can be Primary or Temp Position
                                        'Search HR_TEMP_WORK table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("TW_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                            Else
                                'Delete the Follow Up record for this training record
                                'as no Course completion record found
                                SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                gdbAdoIhr001.Execute SQLQ
                                
                                'Since Previous in HR_TRAIN can be Primary or Temp Position
                                'Clear the Follow Up ID in the Temp/Cross Training Position record
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                    
                                    'Since Previous in HR_TRAIN can be Primary or Temp Position
                                    'Search HR_TEMP_WORK table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            End If
                            rsContEdu.Close
                            Set rsContEdu = Nothing
                            
                            'Delete this Training List record as the course is not required by other positions
                            SQLQ = "DELETE FROM HR_TRAIN"
                            SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                            SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            gdbAdoIhr001.Execute SQLQ
                        End If
                    ElseIf xPosType = "Previous" Then
                        'PREVIOUS - Previous
                        'Track for the most recent previous position requiring this course
                        'These courses are for new Previous Primary Position so recalculate the
                        'Renewal Dates
                        '???xPrvEndDate = Get_Position_End_Date(rsHRTrain("TR_JOB"), rsHRTrain("TR_SDATE"))
                        If Not IsDate(xPrvEndDate) Then xPrvEndDate = rsHRTrain("TR_SDATE")
                        'If CVDate(rsHRTrain("TR_SDATE")) < CVDate(IIf(IsMissing(xStartEndDate), dlpStartDate.Text, xStartEndDate)) Then
                        '???If (dlpENDDATE.Text = "") And (IsNull(xEndDate) Or xEndDate = "" Or IsMissing(xEndDate)) Then
                        '???Else
                        '???If CVDate(xPrvEndDate) < CVDate(IIf(IsMissing(xEndDate) Or xEndDate = "" Or IsNull(xEndDate), dlpENDDATE.Text, xEndDate)) Then
                            'Training List has older Position Start Date so update with new Position info.
                            If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                xDWMY = "d" 'Ticket #27989 Franks 0129/2016 - default to "d", PC_FLWUP_PRD_DWMY can be blank then xDWMY is blank too
                                Select Case rsReqCourse("PC_FLWUP_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    '???rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(dlpStartDate.Text))
                                Else
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_FOLLOWUP"), CVDate(xStartEndDate))
                                End If
                            Else
                                'Check if Previous Renewal period is defined
                                If IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) Or rsReqCourse("PC_RENEW_CRS_PRV") = 0 Then
                                    'No Previous Renewal Period defined so delete this job from this previous position.
                                    'It should not be in the training list for any previous job
                                    flgNoPrvRnwl = True
                                Else
                                    Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                                        Case "D"
                                            xDWMY = "d"
                                        Case "W"
                                            xDWMY = "ww"
                                        Case "M"
                                            xDWMY = "m"
                                        Case "Y"
                                            xDWMY = "yyyy"
                                    End Select
                                    rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                End If
                            End If
                            If flgNoPrvRnwl = False Then
                                'Update Continuing Education record as well
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                'Ticket #28673 Franks 05/27/2016 changed "=" to "IN"
                                'SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW IN (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                'Ticket #28673 Franks 05/27/2016 changed "=" to "IN"
                                'SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_DATCOMP IN (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    'rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_JOB") = xJob
                                    rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                            
                                'Previous Renewal period available
                                rsHRTrain("TR_JOB") = xJob
                                If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                    '???rsHRTrain("TR_SDATE") = dlpStartDate.Text
                                Else
                                    rsHRTrain("TR_SDATE") = xStartEndDate
                                End If
                                rsHRTrain("TR_POS_TYPE") = "P"   'Previous Primary
                                ''If Renewal date is greater than today's date then clear the Course Taken Date
                                'If CVDate(rsHRTrain("TR_RENEW")) >= CVDate(Now) Then
                                '    rsHRTrain("TR_COURSE_TAKEN") = Null
                                'End If
                                rsHRTrain("TR_LDATE") = Date
                                rsHRTrain("TR_LUSER") = glbUserID
                                rsHRTrain("TR_LTIME") = Time$
                                
                                'If follow up id is null then find the id
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                    SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                                    SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                    'SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND EF_FDATE IN (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Ticket #24300
                                'rsHRTrain.Update
                                
                                'Ticket #24300
                                If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                    'Add a Follow Up record for this Training course
                                    'Ticket #24300
                                    'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                    rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                    rsHRTrain.Update
                                Else
                                    rsHRTrain.Update
                                    
                                    'Update Follow Up record - Effective Date
                                    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsFollowUp.EOF Then
                                        rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                        rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                        rsFollowUp("EF_LDATE") = Date
                                        rsFollowUp("EF_LUSER") = glbUserID
                                        rsFollowUp("EF_LTIME") = Time$
                                        rsFollowUp.Update
                                    End If
                                    rsFollowUp.Close
                                    Set rsFollowUp = Nothing
                                End If
                                
                                'Update Position record with Follow Up ID
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Clear the Follow Up from the Previous Job in Primary/Temp Position
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                    
                                    'Since Previous in HR_TRAIN can be Primary or Temp Position
                                    'Search HR_TEMP_WORK table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                    
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and update with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            Else
                                'No Previous renewal found for this course
                                
                                'Clear the Renewal date for this course and for this employee from
                                'Continuing Education screen
                                SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                                SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND ES_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsContEdu.EOF Then
                                    rsContEdu("ES_RENEW") = Null
                                    rsContEdu("ES_LDATE") = Date
                                    rsContEdu("ES_LUSER") = glbUserID
                                    rsContEdu("ES_LTIME") = Time$
                                    rsContEdu.Update
                                    
                                    If Not IsNull(rsContEdu("ES_DATCOMP")) Then
                                        'Since the course was completed - mark the Follow Up as
                                        'Completed instead of deleting it.
                                        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    Else
                                        'Delete the Follow Up record for this training record
                                        'as no Course completion record found
                                        SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                        SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                        SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                        gdbAdoIhr001.Execute SQLQ
                                    
                                        'Clear the Follow Up ID in the Position record
                                        'if the course code is TRAIN
                                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                            'Search HR_JOB_HISTORY table for this Position record
                                            'and clear with Follow Up Id
                                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                            If Not rsTJob.EOF Then
                                                rsTJob("JH_FOLLOWUP_ID") = Null
                                                rsTJob.Update
                                            End If
                                            rsTJob.Close
                                            Set rsTJob = Nothing
                                            
                                            'Since Previous in HR_TRAIN can be Primary or Temp Position
                                            'Search HR_TEMP_WORK table for this Position record
                                            'and clear with Follow Up Id
                                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                            If Not rsTJob.EOF Then
                                                rsTJob("TW_FOLLOWUP_ID") = Null
                                                rsTJob.Update
                                            End If
                                            rsTJob.Close
                                            Set rsTJob = Nothing
                                        End If
                                    End If
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                                                
                                    'Since Previous in HR_TRAIN can be Primary or Temp Position
                                    'Clear the Follow Up ID in the Temp/Cross Training Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                        
                                        'Since Previous in HR_TRAIN can be Primary or Temp Position
                                        'Search HR_TEMP_WORK table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("TW_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                                rsContEdu.Close
                                Set rsContEdu = Nothing
                                
                                'Delete this Training List record as the course is not required by other positions
                                SQLQ = "DELETE FROM HR_TRAIN"
                                SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                                gdbAdoIhr001.Execute SQLQ
                            End If
                        
                        '???Else
                            'Do not do anything because Training List has most recent Position Start Date
                        '???End If
                        '???End If
                    End If
                ElseIf IsNull(rsHRTrain("TR_POS_TYPE")) Or rsHRTrain("TR_POS_TYPE") = "" Then
                    'Check if the course was taken before. If taken then use the normal Training List logic based
                    'on the renewal date if the course should continue to exist or not
                    If IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                        'COURSE NEVER TAKEN BEFORE
                        'It's an independent course and never taken before, update with this Position's information
                        'even though there is no renewal period for the type of position this is
                        rsHRTrain("TR_JOB") = xJob
                        If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                            '???rsHRTrain("TR_SDATE") = dlpStartDate.Text
                        Else
                            rsHRTrain("TR_SDATE") = xStartEndDate
                        End If
                        If xPosType = "Current" Then
                            rsHRTrain("TR_POS_TYPE") = "C"
                        ElseIf xPosType = "Temporary" Then
                            rsHRTrain("TR_POS_TYPE") = "T"
                        ElseIf xPosType = "Previous" Then
                            rsHRTrain("TR_POS_TYPE") = "P"
                        End If
    
                        'Do not overwrite the Renewal Date entered for this independent course
                        'rsHRTrain("TR_RENEW")) =
                        rsHRTrain("TR_LDATE") = Date
                        rsHRTrain("TR_LUSER") = glbUserID
                        rsHRTrain("TR_LTIME") = Time$
                        
                        'If follow up id is null then find the id
                        If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                            xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                            SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                            SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                            SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(rsHRTrain("TR_RENEW"))
                            rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsFollowUp.EOF Then
                                rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                            End If
                            rsFollowUp.Close
                            Set rsFollowUp = Nothing
                        End If
                        
                        'Ticket #24300
                        'rsHRTrain.Update
                        
                        'Ticket #24300
                        If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                            'Add a Follow Up record for this Training course
                            'Ticket #24300
                            'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                            rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                            rsHRTrain.Update
                        Else
                            rsHRTrain.Update
                        
                            'Update Follow Up record - Comments with Position
                            SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                            SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsFollowUp.EOF Then
                                'No change to renewal date
                                'rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                rsFollowUp("EF_LDATE") = Date
                                rsFollowUp("EF_LUSER") = glbUserID
                                rsFollowUp("EF_LTIME") = Time$
                                rsFollowUp.Update
                            End If
                            rsFollowUp.Close
                            Set rsFollowUp = Nothing
                        End If
                    
                        'Update Position record with Follow Up ID
                        'if the course code is TRAIN
                        If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                            'Clear the Follow Up from the Previous Job in Primary/Temp Position
                            'Search HR_JOB_HISTORY table for this Position record
                            'and clear with Follow Up Id
                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                rsTJob("JH_FOLLOWUP_ID") = Null
                                rsTJob.Update
                            End If
                            rsTJob.Close
                            Set rsTJob = Nothing
                            
                            'Since Previous in HR_TRAIN can be Primary or Temp Position
                            'Search HR_TEMP_WORK table for this Position record
                            'and clear with Follow Up Id
                            SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                rsTJob("TW_FOLLOWUP_ID") = Null
                                rsTJob.Update
                            End If
                            rsTJob.Close
                            Set rsTJob = Nothing
                            
                            'Search HR_JOB_HISTORY table for this Position record
                            'and update with Follow Up Id
                            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                            rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsTJob.EOF Then
                                rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                rsTJob.Update
                            End If
                            rsTJob.Close
                            Set rsTJob = Nothing
                        End If
                    Else
                        'COURSE TAKEN BEFORE
                        'Which kind of Position is this
                        If xPosType = "Current" Or xPosType = "Temporary" Then
                            'Check if Current Renewal period is defined
                            If IsNull(rsReqCourse("PC_RENEW_CRS_CUR")) Or rsReqCourse("PC_RENEW_CRS_CUR") = 0 Then
                                'No Current Renewal Period defined so delete this job from this current position.
                                'It should not be in the training list for any current job
                                flgNoCurRnwl = True
                            Else
                                Select Case rsReqCourse("PC_CUR_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_CUR"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                flgNoCurRnwl = False
                            End If
                        ElseIf xPosType = "Previous" Then
                            'Check if Previous Renewal period is defined
                            If IsNull(rsReqCourse("PC_RENEW_CRS_PRV")) Or rsReqCourse("PC_RENEW_CRS_PRV") = 0 Then
                                'No Previous Renewal Period defined so delete this job from this previous position.
                                'It should not be in the training list for any previous job
                                flgNoCurRnwl = True
                            Else
                                Select Case rsReqCourse("PC_PRV_PRD_DWMY")
                                    Case "D"
                                        xDWMY = "d"
                                    Case "W"
                                        xDWMY = "ww"
                                    Case "M"
                                        xDWMY = "m"
                                    Case "Y"
                                        xDWMY = "yyyy"
                                End Select
                                rsHRTrain("TR_RENEW") = DateAdd(xDWMY, rsReqCourse("PC_RENEW_CRS_PRV"), CVDate(rsHRTrain("TR_COURSE_TAKEN")))
                                flgNoCurRnwl = False
                            End If
                        End If
                        
                        If flgNoCurRnwl = False Then
                            'Renewal Period Found - updated existing records
                            'Update Continuing Education record as well
                            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                            SQLQ = SQLQ & " AND (ES_JOB = '' OR ES_JOB IS NULL)"    'No Job - Independent course
                            SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            'Ticket #28673 Franks 05/27/2016 changed "=" to "IN"
                            'SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                            SQLQ = SQLQ & " AND ES_RENEW IN (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            'Ticket #28673 Franks 05/27/2016 changed "=" to "IN"
                            'SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                            SQLQ = SQLQ & " AND ES_DATCOMP IN (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsContEdu.EOF Then
                                'rsContEdu("ES_RENEW") = Null
                                rsContEdu("ES_JOB") = xJob
                                rsContEdu("ES_RENEW") = rsHRTrain("TR_RENEW")
                                rsContEdu("ES_LDATE") = Date
                                rsContEdu("ES_LUSER") = glbUserID
                                rsContEdu("ES_LTIME") = Time$
                                rsContEdu.Update
                            End If
                            rsContEdu.Close
                            Set rsContEdu = Nothing
                            
                            'Renewal Period available
                            rsHRTrain("TR_JOB") = xJob
                            If IsMissing(xStartEndDate) Or xStartEndDate = "" Then
                                '???rsHRTrain("TR_SDATE") = dlpStartDate.Text
                            Else
                                rsHRTrain("TR_SDATE") = xStartEndDate
                            End If
                            If xPosType = "Current" Then
                                rsHRTrain("TR_POS_TYPE") = "C"
                            ElseIf xPosType = "Temporary" Then
                                rsHRTrain("TR_POS_TYPE") = "T"
                            ElseIf xPosType = "Previous" Then
                                rsHRTrain("TR_POS_TYPE") = "P"
                            End If
                            
                            rsHRTrain("TR_LDATE") = Date
                            rsHRTrain("TR_LUSER") = glbUserID
                            rsHRTrain("TR_LTIME") = Time$
                            
                            'If follow up id is null then find the id
                            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                xComments = "Course: " & rsHRTrain("TR_CRSCODE") & " "
                                SQLQ = "SELECT EF_FOLLOWUP_ID FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
                                SQLQ = SQLQ & " AND EF_COMMENTS LIKE '" & xComments & "%' "
                                'SQLQ = SQLQ & " AND EF_FDATE = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND EF_FDATE IN (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & rsHRTrain("TR_JOB") & "'"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                            
                            'Ticket #24300
                            'rsHRTrain.Update
                            
                            'Ticket #24300
                            If IsNull(rsHRTrain("TR_FOLLOWUP_ID")) Then
                                'Add a Follow Up record for this Training course
                                'Ticket #24300
                                'rsHRTrain("TR_FOLLOWUP_ID") = rsFollowUp("EF_FOLLOWUP_ID")
                                rsHRTrain("TR_FOLLOWUP_ID") = Add_Train_FollowUp(glbLEE_ID, rsHRTrain("TR_RENEW"), rsReqCourse("PC_CRSCODE"), xJob)
                                rsHRTrain.Update
                            Else
                                rsHRTrain.Update
                            
                                'Update Follow Up record - Effective Date
                                SQLQ = "SELECT * FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsFollowUp.EOF Then
                                    rsFollowUp("EF_FDATE") = rsHRTrain("TR_RENEW")
                                    rsFollowUp("EF_COMMENTS") = "Course: " & rsReqCourse("PC_CRSCODE") & " - " & GetTABLDesc("ESCD", rsReqCourse("PC_CRSCODE")) & " for Position: " & xJob
                                    rsFollowUp("EF_LDATE") = Date
                                    rsFollowUp("EF_LUSER") = glbUserID
                                    rsFollowUp("EF_LTIME") = Time$
                                    rsFollowUp.Update
                                End If
                                rsFollowUp.Close
                                Set rsFollowUp = Nothing
                            End If
                        
                            'Update Position record with Follow Up ID
                            'if the course code is TRAIN
                            If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                'Clear the Follow Up from the Previous Job in Primary/Temp Position
                                'Search HR_JOB_HISTORY table for this Position record
                                'and clear with Follow Up Id
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("JH_FOLLOWUP_ID") = Null
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                                
                                'Since Previous in HR_TRAIN can be Primary or Temp Position
                                'Search HR_TEMP_WORK table for this Position record
                                'and clear with Follow Up Id
                                SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("TW_FOLLOWUP_ID") = Null
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                                
                                'Search HR_JOB_HISTORY table for this Position record
                                'and update with Follow Up Id
                                SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_ID = " & Data1.Recordset("JH_ID")
                                rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                If Not rsTJob.EOF Then
                                    rsTJob("JH_FOLLOWUP_ID") = rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Update
                                End If
                                rsTJob.Close
                                Set rsTJob = Nothing
                            End If
                        Else
                            'No Renewal Period found for this course
                                                        
                            'Clear the Renewal date for this course and for this employee from
                            'Continuing Education screen
                            SQLQ = "SELECT ES_EMPNBR, ES_CRSCODE,ES_DATCOMP,ES_RENEW,ES_JOB,ES_LDATE,ES_LUSER,ES_LTIME FROM HREDSEM"
                            SQLQ = SQLQ & " WHERE ES_EMPNBR = " & glbLEE_ID
                            SQLQ = SQLQ & " AND (ES_JOB = '' OR ES_JOB IS NULL)"    'Independent course
                            SQLQ = SQLQ & " AND ES_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            'SQLQ = SQLQ & " AND ES_RENEW = (SELECT TR_RENEW FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND TR_JOB = '" & xJob & "'"
                            'SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            If Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                SQLQ = SQLQ & " AND ES_DATCOMP = (SELECT TR_COURSE_TAKEN FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                            End If
                            rsContEdu.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                            If Not rsContEdu.EOF Then
                                rsContEdu("ES_RENEW") = Null
                                rsContEdu("ES_LDATE") = Date
                                rsContEdu("ES_LUSER") = glbUserID
                                rsContEdu("ES_LTIME") = Time$
                                rsContEdu.Update
                                
                                If Not IsNull(rsContEdu("ES_DATCOMP")) And Not IsNull(rsHRTrain("TR_COURSE_TAKEN")) Then
                                    'Since the course was completed - mark the Follow Up as
                                    'Completed instead of deleting it.
                                    SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_FDATE = " & Date_SQL(rsContEdu("ES_DATCOMP")) & ", EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "'"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                Else
                                    'Delete the Follow Up record for this training record
                                    'as no Course completion record found
                                    SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                    SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                    SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                    gdbAdoIhr001.Execute SQLQ
                                
                                    'Clear the Follow Up ID in the Position record
                                    'if the course code is TRAIN
                                    If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                        'Search HR_JOB_HISTORY table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("JH_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                        
                                        'Since Previous in HR_TRAIN can be Primary or Temp Position
                                        'Search HR_TEMP_WORK table for this Position record
                                        'and clear with Follow Up Id
                                        SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                        rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                        If Not rsTJob.EOF Then
                                            rsTJob("TW_FOLLOWUP_ID") = Null
                                            rsTJob.Update
                                        End If
                                        rsTJob.Close
                                        Set rsTJob = Nothing
                                    End If
                                End If
                            Else
                                'Delete the Follow Up record for this training record
                                'as no Course completion record found
                                SQLQ = "DELETE FROM HR_FOLLOW_UP"
                                SQLQ = SQLQ & " WHERE EF_FOLLOWUP_ID = (SELECT TR_FOLLOWUP_ID FROM HR_TRAIN WHERE TR_EMPNBR = " & glbLEE_ID & " AND (TR_JOB = '' OR TR_JOB IS NULL)"
                                SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "')"
                                gdbAdoIhr001.Execute SQLQ
                                
                                'Since Previous in HR_TRAIN can be Primary or Temp Position
                                'Clear the Follow Up ID in the Temp/Cross Training Position record
                                'if the course code is TRAIN
                                If rsReqCourse("PC_CRSCODE") = "TRAIN" Then
                                    'Search HR_JOB_HISTORY table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("JH_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                    
                                    'Since Previous in HR_TRAIN can be Primary or Temp Position
                                    'Search HR_TEMP_WORK table for this Position record
                                    'and clear with Follow Up Id
                                    SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE TW_FOLLOWUP_ID = " & rsHRTrain("TR_FOLLOWUP_ID")
                                    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                                    If Not rsTJob.EOF Then
                                        rsTJob("TW_FOLLOWUP_ID") = Null
                                        rsTJob.Update
                                    End If
                                    rsTJob.Close
                                    Set rsTJob = Nothing
                                End If
                            End If
                            rsContEdu.Close
                            Set rsContEdu = Nothing
                            
                            'Delete this Training List record as the course is not required by other positions
                            SQLQ = "DELETE FROM HR_TRAIN"
                            SQLQ = SQLQ & " WHERE TR_EMPNBR = " & glbLEE_ID
                            SQLQ = SQLQ & " AND (TR_JOB = '' OR TR_JOB IS NULL)"    'Independent course
                            SQLQ = SQLQ & " AND TR_CRSCODE = '" & rsReqCourse("PC_CRSCODE") & "'"
                            gdbAdoIhr001.Execute SQLQ
                        End If
                        
                    End If
                End If
                
                If flgProcCalled = False Then
                    xPosType = xorgPosType
                    xJob = xorgJob
                End If
                
            End If
            rsHRTrain.Close
            Set rsHRTrain = Nothing
            
Next_Required_Course:
            rsCourseMst.Close
            Set rsCourseMst = Nothing

            rsReqCourse.MoveNext
        Loop
    End If
    rsReqCourse.Close
    Set rsReqCourse = Nothing

Exit Sub

Employee_Job_Training_Err:
If Err = 3018 Then
    Err = 0
    Resume Next
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
If Len(SQLQ) = 0 Then
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update_Employee_Job_Training_List", "HR_JOB_HISTORY", "Update_Emp_Job_Training_List")
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, SQLQ, "HR_JOB_HISTORY", "Update_Emp_Job_Training_List")
End If
Call RollBack '26July99 js
End Sub

Private Sub WFCCandidateDele(xCandidate)
Dim SQLQ As String
    If IsNumeric(xCandidate) Then
        SQLQ = "DELETE FROM HRSF_XML_IMPORT WHERE SF_CANDIDATE = " & xCandidate & " "
        gdbAdoIhr001.Execute SQLQ
    End If
End Sub
    
Private Sub MacaulayAltPayIDScreen() 'Ticket #24557 Franks 09/03/2014
    lbltitle(49).Caption = lStr("Location")
    lbltitle(50).Caption = lStr("Salary Distribution")
    lbltitle(54).Caption = lStr("Region")
End Sub

Private Sub ScreenSetupNew() 'Ticket #25323 Franks 12/16/2014
    If glbCompSerial = "S/N - 2460W" Then 'OPL
        lbltitle(32).FontBold = True    'Payroll ID
        lbltitle(24).FontBold = True    'Region
    End If
End Sub
