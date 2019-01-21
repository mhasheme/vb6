VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmComPrefer 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Company Preference"
   ClientHeight    =   9915
   ClientLeft      =   4395
   ClientTop       =   4080
   ClientWidth     =   11205
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
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9915
   ScaleWidth      =   11205
   WindowState     =   2  'Maximized
   Begin VB.Frame scrFrameGen 
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   120
      TabIndex        =   72
      Top             =   240
      Visible         =   0   'False
      Width           =   10455
      Begin VB.TextBox txtSMTPServer 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   10
         Tag             =   "SMTP Server"
         Top             =   3000
         Width           =   2475
      End
      Begin VB.TextBox txtSMTPUsername 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   11
         Tag             =   "SMTP Username"
         Top             =   3375
         Width           =   2475
      End
      Begin VB.TextBox txtSMTPPassword 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   12
         Tag             =   "SMTP Password"
         Top             =   3735
         Width           =   2475
      End
      Begin VB.TextBox txtSMTPPort 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   13
         Tag             =   "SMTP Port"
         Top             =   4110
         Width           =   2475
      End
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   696
         Width           =   2925
         _Version        =   65536
         _ExtentX        =   5159
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Attachment"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   8
         Left            =   6840
         TabIndex        =   68
         Top             =   2280
         Visible         =   0   'False
         Width           =   3405
         _Version        =   65536
         _ExtentX        =   6006
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Not Use - Use Great Plains Holding File"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   4
         Top             =   1272
         Width           =   2085
         _Version        =   65536
         _ExtentX        =   3678
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Secured Password"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   10
         Left            =   6840
         TabIndex        =   5
         Top             =   1920
         Visible         =   0   'False
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Not in Use - Benefit History (was added back in 2007)"
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
      End
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   15
         Left            =   360
         TabIndex        =   3
         Top             =   984
         Width           =   3405
         _Version        =   65536
         _ExtentX        =   6006
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Show Compa-Ratio"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   0
         Top             =   120
         Width           =   2685
         _Version        =   65536
         _ExtentX        =   4736
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Single Sign-On in info:HR"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   17
         Left            =   360
         TabIndex        =   1
         Top             =   408
         Width           =   3165
         _Version        =   65536
         _ExtentX        =   5583
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Single Sign-On in Web Systems"
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
      Begin INFOHR_Controls.DateLookup dlpWSRotEffDate 
         DataSource      =   " "
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Tag             =   "40-Effective Date"
         Top             =   1920
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin VB.ComboBox comWSRotWks 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "00-Work Schedule Rotation Weeks"
         Top             =   1600
         Width           =   600
      End
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   26
         Left            =   360
         TabIndex        =   8
         Top             =   2280
         Visible         =   0   'False
         Width           =   3165
         _Version        =   65536
         _ExtentX        =   5583
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Encrypt Database Connection"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   27
         Left            =   360
         TabIndex        =   9
         Top             =   2640
         Width           =   3165
         _Version        =   65536
         _ExtentX        =   5583
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  SMTP Connection Information"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   28
         Left            =   360
         TabIndex        =   14
         Top             =   4560
         Visible         =   0   'False
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Flex Logic"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   32
         Left            =   360
         TabIndex        =   15
         Top             =   4800
         Visible         =   0   'False
         Width           =   3525
         _Version        =   65536
         _ExtentX        =   6218
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Disable Compensatory Time Entries"
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
      Begin VB.Image imgHelpCompTime 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   0
         Picture         =   "fspreference.frx":0000
         Stretch         =   -1  'True
         Top             =   4800
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgHelpFlexLogic 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   0
         Picture         =   "fspreference.frx":0442
         Stretch         =   -1  'True
         Top             =   4560
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblUsername 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SMTP Username:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   99
         Top             =   3420
         Width           =   1530
      End
      Begin VB.Label lblServer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SMTP Server:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   98
         Top             =   3030
         Width           =   1605
      End
      Begin VB.Label lblPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SMTP Password:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   97
         Top             =   3750
         Width           =   1530
      End
      Begin VB.Label lblPort 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SMTP Port:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   96
         Top             =   4155
         Width           =   1290
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   17
         Left            =   1080
         TabIndex        =   94
         Top             =   1965
         Width           =   780
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Work Schedule Rotation Weeks"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   1080
         TabIndex        =   93
         Top             =   1660
         Width           =   3675
      End
      Begin VB.Label lblSSOWebSysInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "( Controlled from Application Setting )"
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
         Left            =   3600
         TabIndex        =   74
         Top             =   420
         Width           =   2595
      End
      Begin VB.Image imgHelpSSOinfoHR 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   0
         Picture         =   "fspreference.frx":0884
         Stretch         =   -1  'True
         Top             =   120
         Width           =   255
      End
      Begin VB.Image imgHelp 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   0
         Picture         =   "fspreference.frx":0CC6
         Stretch         =   -1  'True
         Top             =   1272
         Width           =   255
      End
      Begin VB.Label lblNote 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9480
         TabIndex        =   69
         Top             =   9960
         Width           =   6615
      End
   End
   Begin VB.Frame scrFrameEmailNoti2 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   120
      TabIndex        =   105
      Top             =   5040
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton cmdShowHS 
         Caption         =   "More Emails"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   108
         Top             =   945
         Width           =   1695
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Height          =   485
         Index           =   14
         Left            =   1680
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   107
         Tag             =   "00-Email Address"
         Top             =   840
         Width           =   6660
      End
      Begin VB.CommandButton cmdScreenLeft 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9120
         Picture         =   "fspreference.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   106
         Tag             =   "Previous Security Screen"
         Top             =   120
         Width           =   705
      End
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   31
         Left            =   240
         TabIndex        =   109
         Top             =   600
         Width           =   3285
         _Version        =   65536
         _ExtentX        =   5794
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Email Notification on H&&S Incident "
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
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Left            =   600
         TabIndex        =   110
         Top             =   885
         Width           =   1095
      End
   End
   Begin VB.Frame scrFrameEmailNoti 
      BorderStyle     =   0  'None
      Height          =   11415
      Left            =   120
      TabIndex        =   75
      Top             =   2040
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton cmdPageRight 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9120
         Picture         =   "fspreference.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   104
         Tag             =   "Grant All Basic"
         Top             =   120
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Height          =   485
         Index           =   13
         Left            =   1560
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Tag             =   "00-Email Address"
         Top             =   9600
         Visible         =   0   'False
         Width           =   6660
      End
      Begin VB.CommandButton cmdShowEmpFlags 
         Caption         =   "More Emails"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   53
         Top             =   9705
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdShowPerformance 
         Caption         =   "More Emails"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   102
         Top             =   5870
         Width           =   1695
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Height          =   485
         Index           =   7
         Left            =   1560
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   101
         Tag             =   "00-Email Address"
         Top             =   5760
         Width           =   6660
      End
      Begin VB.CommandButton cmdShowRequestSubmission 
         Caption         =   "More Emails"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   56
         Top             =   10575
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Height          =   485
         Index           =   12
         Left            =   1560
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   55
         Tag             =   "00-Email Address"
         Top             =   10470
         Visible         =   0   'False
         Width           =   6660
      End
      Begin VB.CommandButton cmdShowNewApplicant 
         Caption         =   "More Emails"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   50
         Top             =   8950
         Width           =   1695
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Height          =   485
         Index           =   11
         Left            =   1560
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   49
         Tag             =   "00-Email Address"
         Top             =   8840
         Width           =   6660
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Height          =   485
         Index           =   10
         Left            =   1560
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Tag             =   "00-Email Address"
         Top             =   8040
         Width           =   6660
      End
      Begin VB.CommandButton cmdShowAddressChg 
         Caption         =   "More Emails"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   47
         Top             =   8150
         Width           =   1695
      End
      Begin VB.CommandButton cmdShowRequestApproval 
         Caption         =   "More Emails"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   44
         Top             =   7385
         Width           =   1695
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Height          =   485
         Index           =   6
         Left            =   1560
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Tag             =   "00-Email Address"
         Top             =   7275
         Width           =   6660
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Height          =   485
         Index           =   5
         Left            =   1560
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Tag             =   "00-Email Address"
         Top             =   5040
         Width           =   6660
      End
      Begin VB.CommandButton cmdShowLeaveChanges 
         Caption         =   "More Emails"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   37
         Top             =   5150
         Width           =   1695
      End
      Begin VB.CommandButton cmdShowRehire 
         Caption         =   "More Emails"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   34
         Top             =   4430
         Width           =   1695
      End
      Begin VB.CommandButton cmdShowTermination 
         Caption         =   "More Emails"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   31
         Top             =   3710
         Width           =   1695
      End
      Begin VB.CommandButton cmdShowBenefits 
         Caption         =   "More Emails"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   28
         Top             =   2990
         Width           =   1695
      End
      Begin VB.CommandButton cmdShowSalary 
         Caption         =   "More Emails"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   25
         Top             =   2270
         Width           =   1695
      End
      Begin VB.CommandButton cmdShowNewhire 
         Caption         =   "More Emails"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   19
         Top             =   830
         Width           =   1695
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Height          =   485
         Index           =   4
         Left            =   1560
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Tag             =   "00-Email Address"
         Top             =   4320
         Width           =   6660
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Height          =   485
         Index           =   0
         Left            =   1560
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Tag             =   "00-Email Address"
         Top             =   720
         Width           =   6660
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Height          =   485
         Index           =   3
         Left            =   1560
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Tag             =   "00-Email Address"
         Top             =   3600
         Width           =   6660
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Height          =   485
         Index           =   2
         Left            =   1560
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Tag             =   "00-Email Address"
         Top             =   2880
         Width           =   6660
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Height          =   485
         Index           =   1
         Left            =   1560
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Tag             =   "00-Email Address"
         Top             =   2160
         Width           =   6660
      End
      Begin VB.CommandButton cmdShowDependent 
         Caption         =   "More Emails"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   41
         Top             =   6600
         Width           =   1695
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Height          =   485
         Index           =   8
         Left            =   1560
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Tag             =   "00-Email Address"
         Top             =   6480
         Width           =   6660
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
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
         Height          =   485
         Index           =   9
         Left            =   1560
         MaxLength       =   400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Tag             =   "00-Email Address"
         Top             =   1440
         Width           =   6660
      End
      Begin VB.CommandButton cmdShowPosition 
         Caption         =   "More Emails"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   22
         Top             =   1550
         Width           =   1695
      End
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   3285
         _Version        =   65536
         _ExtentX        =   5794
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Email Notification on New Hire"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Email Notification on Salary"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   3165
         _Version        =   65536
         _ExtentX        =   5583
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Email Notification on Benefits"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   29
         Top             =   3360
         Width           =   3285
         _Version        =   65536
         _ExtentX        =   5794
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Email Notification on Termination"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   3405
         _Version        =   65536
         _ExtentX        =   6006
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Email Sending Function"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   32
         Top             =   4080
         Width           =   3285
         _Version        =   65536
         _ExtentX        =   5794
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Email Notification on Rehire"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   35
         Top             =   4800
         Width           =   3645
         _Version        =   65536
         _ExtentX        =   6429
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Email Notification on Leave Changes"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   42
         Top             =   7005
         Width           =   5325
         _Version        =   65536
         _ExtentX        =   9393
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Email Notification (Additional) on ESS - Request Approval"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   38
         Top             =   5520
         Width           =   5085
         _Version        =   65536
         _ExtentX        =   8969
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Email Notification on Performance"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   39
         Top             =   6240
         Width           =   5085
         _Version        =   65536
         _ExtentX        =   8969
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Email Notification on Dependent"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Email Notification on Position"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   57
         Top             =   11040
         Visible         =   0   'False
         Width           =   3765
         _Version        =   65536
         _ExtentX        =   6641
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Dependent Eligible Date 30 days Email"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   45
         Top             =   7800
         Width           =   3765
         _Version        =   65536
         _ExtentX        =   6641
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Email Notification on Address Change"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   48
         Top             =   8590
         Width           =   3765
         _Version        =   65536
         _ExtentX        =   6641
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Email Notification on New Applicant"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   29
         Left            =   120
         TabIndex        =   54
         Top             =   10200
         Visible         =   0   'False
         Width           =   5565
         _Version        =   65536
         _ExtentX        =   9816
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Email Notification (Additional) on ESS - Request Submission"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   30
         Left            =   120
         TabIndex        =   51
         Top             =   9350
         Visible         =   0   'False
         Width           =   5085
         _Version        =   65536
         _ExtentX        =   8969
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Email Notification on Employee Flags"
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
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Left            =   480
         TabIndex        =   103
         Top             =   9645
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Index           =   19
         Left            =   480
         TabIndex        =   100
         Top             =   10515
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Left            =   480
         TabIndex        =   95
         Top             =   8885
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Left            =   480
         TabIndex        =   92
         Top             =   8085
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Index           =   9
         Left            =   480
         TabIndex        =   91
         Top             =   5805
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Left            =   480
         TabIndex        =   90
         Top             =   7320
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Left            =   480
         TabIndex        =   89
         Top             =   5085
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Left            =   480
         TabIndex        =   88
         Top             =   4365
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Left            =   480
         TabIndex        =   87
         Top             =   3645
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Left            =   480
         TabIndex        =   86
         Top             =   2925
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Left            =   480
         TabIndex        =   85
         Top             =   2205
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Left            =   480
         TabIndex        =   84
         Top             =   765
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Left            =   480
         TabIndex        =   83
         Top             =   6525
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Left            =   480
         TabIndex        =   82
         Top             =   1485
         Width           =   1095
      End
   End
   Begin VB.Frame scrFrameFileLoc 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      TabIndex        =   76
      Top             =   3600
      Visible         =   0   'False
      Width           =   10455
      Begin VB.TextBox txtFriesensPath 
         Appearance      =   0  'Flat
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
         Left            =   1800
         MaxLength       =   150
         TabIndex        =   63
         Tag             =   "00-MS Word Report Forms Path"
         Top             =   1470
         Visible         =   0   'False
         Width           =   6660
      End
      Begin VB.TextBox txtTMPatch 
         Appearance      =   0  'Flat
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
         Left            =   1800
         MaxLength       =   150
         TabIndex        =   59
         Tag             =   "00-Save Excel Reports in this path"
         Top             =   360
         Width           =   6660
      End
      Begin VB.TextBox txtWSIBForm7Path 
         Appearance      =   0  'Flat
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
         Left            =   1800
         MaxLength       =   150
         TabIndex        =   65
         Tag             =   "00-WSIB Form 7 Path"
         Top             =   2040
         Visible         =   0   'False
         Width           =   6660
      End
      Begin VB.TextBox txtFUEmailSendPath 
         Appearance      =   0  'Flat
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
         Left            =   1800
         MaxLength       =   150
         TabIndex        =   67
         Tag             =   "00-Follow Up Email Sending Path"
         Top             =   2610
         Visible         =   0   'False
         Width           =   6660
      End
      Begin VB.TextBox txtEmpPhotoPath 
         Appearance      =   0  'Flat
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
         Left            =   1800
         MaxLength       =   150
         TabIndex        =   61
         Tag             =   "00-Save Employees Photos in this path"
         Top             =   885
         Width           =   6660
      End
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   58
         Top             =   120
         Width           =   3285
         _Version        =   65536
         _ExtentX        =   5794
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Excel Reports In Other Folder"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   62
         Top             =   1200
         Visible         =   0   'False
         Width           =   3285
         _Version        =   65536
         _ExtentX        =   5794
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Friesens's MS Word Reports"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   19
         Left            =   360
         TabIndex        =   64
         Top             =   1770
         Visible         =   0   'False
         Width           =   3285
         _Version        =   65536
         _ExtentX        =   5794
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  WSIB Form 7"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   21
         Left            =   360
         TabIndex        =   66
         Top             =   2340
         Visible         =   0   'False
         Width           =   2925
         _Version        =   65536
         _ExtentX        =   5159
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Follow Up Email Sending Log"
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
      Begin Threed.SSCheck chkPreFunction 
         Height          =   255
         Index           =   23
         Left            =   360
         TabIndex        =   60
         Top             =   645
         Width           =   3285
         _Version        =   65536
         _ExtentX        =   5794
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "  Employee Photo In Other Folder"
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
      Begin VB.Image imgHelpPhoto 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   0
         Picture         =   "fspreference.frx":198C
         Stretch         =   -1  'True
         Top             =   645
         Width           =   255
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Path"
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
         Index           =   6
         Left            =   720
         TabIndex        =   81
         Top             =   1515
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Path"
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
         Left            =   720
         TabIndex        =   80
         Top             =   405
         Width           =   330
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Path"
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
         Left            =   720
         TabIndex        =   79
         Top             =   2085
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Path"
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
         Left            =   720
         TabIndex        =   78
         Top             =   2655
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Path"
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
         Left            =   720
         TabIndex        =   77
         Top             =   930
         Width           =   330
      End
   End
   Begin VB.VScrollBar scrControl 
      Height          =   8300
      LargeChange     =   315
      Left            =   10800
      Max             =   100
      SmallChange     =   315
      TabIndex        =   71
      Top             =   360
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   120
      Max             =   50
      SmallChange     =   4
      TabIndex        =   73
      Top             =   8640
      Width           =   10575
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   70
      Top             =   9615
      Width           =   11205
      _Version        =   65536
      _ExtentX        =   19764
      _ExtentY        =   529
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
         Left            =   4800
         Top             =   120
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
      Begin MSAdodcLib.Adodc Data1 
         Height          =   375
         Left            =   7200
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
   End
End
Attribute VB_Name = "frmComPrefer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean
Dim oWSRotWks As Integer
Dim fODBConnEncrypt As Boolean

Private Sub Display_Values()
Dim rsSR As New ADODB.Recordset
Dim X%, SQLQ

SQLQ = "select * from HRPREFERENCE "
rsSR.Open SQLQ, gdbAdoIhr001, adOpenStatic

Call ResetAll

Do Until rsSR.EOF
    If UCase(rsSR("HP_FUN_NAME")) = UCase("ATTACHMENT") Then chkPreFunction(0) = rsSR("HP_ENABLED")
    If UCase(rsSR("HP_FUN_NAME")) = UCase("EMAIL_SENDING") Then chkPreFunction(1) = rsSR("HP_ENABLED")
    If UCase(rsSR("HP_FUN_NAME")) = UCase("COMPA-RATIO") Then chkPreFunction(15) = rsSR("HP_ENABLED")
    If UCase(rsSR("HP_FUN_NAME")) = UCase("EMAIL_ONNEWHIRE") Then
        chkPreFunction(2) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtEmail(0) = ""
        Else
            txtEmail(0) = rsSR("HP_EMAIL")
        End If
    End If
    'Ticket #21444 Franks 02/09/2012
    If UCase(rsSR("HP_FUN_NAME")) = UCase("EMAIL_ONPOSITION") Then
        chkPreFunction(20) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtEmail(9) = ""
        Else
            txtEmail(9) = rsSR("HP_EMAIL")
        End If
    End If
    If UCase(rsSR("HP_FUN_NAME")) = UCase("EMAIL_ONSALARY") Then
        chkPreFunction(3) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtEmail(1) = ""
        Else
            txtEmail(1) = rsSR("HP_EMAIL")
        End If
    End If
    If UCase(rsSR("HP_FUN_NAME")) = UCase("EMAIL_ONBENEFIT") Then
        chkPreFunction(4) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtEmail(2) = ""
        Else
            txtEmail(2) = rsSR("HP_EMAIL")
        End If
    End If
    If UCase(rsSR("HP_FUN_NAME")) = UCase("EMAIL_ONTERM") Then
        chkPreFunction(5) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtEmail(3) = ""
        Else
            txtEmail(3) = rsSR("HP_EMAIL")
        End If
    End If
    If UCase(rsSR("HP_FUN_NAME")) = UCase("EMAIL_ONREHIRE") Then
        chkPreFunction(6) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtEmail(4) = ""
        Else
            txtEmail(4) = rsSR("HP_EMAIL")
        End If
    End If
    If UCase(rsSR("HP_FUN_NAME")) = UCase("EMAIL_ONLEAVECHANGES") Then
        chkPreFunction(12) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtEmail(5) = ""
        Else
            txtEmail(5) = rsSR("HP_EMAIL")
        End If
    End If
    '7.9
    If UCase(rsSR("HP_FUN_NAME")) = UCase("EMAIL_ONPERFORMANCE") Then
        chkPreFunction(14) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtEmail(7) = ""
        Else
            txtEmail(7) = rsSR("HP_EMAIL")
        End If
    End If
    '7.9
    If UCase(rsSR("HP_FUN_NAME")) = UCase("EMAIL_ONDEPENDENT") Then
        chkPreFunction(18) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtEmail(8) = ""
        Else
            txtEmail(8) = rsSR("HP_EMAIL")
        End If
    End If
    
    'Ticket #18223
    If UCase(rsSR("HP_FUN_NAME")) = UCase("EMAIL_ONREQUESTAPPROVAL") Then
        chkPreFunction(13) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtEmail(6) = ""
        Else
            txtEmail(6) = rsSR("HP_EMAIL")
        End If
    End If
    
    If UCase(rsSR("HP_FUN_NAME")) = UCase("TRAININGMATRIX") Then
        chkPreFunction(7) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtTMPatch = ""
        Else
            txtTMPatch = rsSR("HP_EMAIL")
        End If
    End If
    If UCase(rsSR("HP_FUN_NAME")) = UCase("GP_HOLDING") Then chkPreFunction(8) = rsSR("HP_ENABLED")
    If UCase(rsSR("HP_FUN_NAME")) = UCase("SECURED_PSW") Then chkPreFunction(9) = rsSR("HP_ENABLED")
    If UCase(rsSR("HP_FUN_NAME")) = UCase("BENEFIT_HISTORY") Then chkPreFunction(10) = rsSR("HP_ENABLED")
    
    'Friesens - Ticket #17029
    If glbCompSerial = "S/N - 2279W" Then
        If UCase(rsSR("HP_FUN_NAME")) = UCase("FRIESENSWORDPATH") Then
            chkPreFunction(11) = rsSR("HP_ENABLED")
            If IsNull(rsSR("HP_EMAIL")) Then
                txtFriesensPath = ""
            Else
                txtFriesensPath = rsSR("HP_EMAIL")
            End If
        End If
    End If
    
    '7.9 Enhancement
    If UCase(rsSR("HP_FUN_NAME")) = UCase("SSO_INFOHR") Then chkPreFunction(16) = rsSR("HP_ENABLED")
    If UCase(rsSR("HP_FUN_NAME")) = UCase("SSO_WEBSYS") Then chkPreFunction(17) = rsSR("HP_ENABLED")
    
    'Ticket #20038 - WSIB Form 7 Path to Save
    If glbWSIBModule Then
        If UCase(rsSR("HP_FUN_NAME")) = UCase("WSIBFORM7PATH") Then
            chkPreFunction(19) = rsSR("HP_ENABLED")
            If IsNull(rsSR("HP_EMAIL")) Then
                txtWSIBForm7Path = ""
            Else
                txtWSIBForm7Path = rsSR("HP_EMAIL")
            End If
        End If
    End If
    
    'Ticket #21023 - Oshawa Community Health Centre - Follow Up Email Sending
    If glbCompSerial = "S/N - 2396W" Then
        If UCase(rsSR("HP_FUN_NAME")) = UCase("FOLLOWUPEMAILLOGPATH") Then
            chkPreFunction(21) = rsSR("HP_ENABLED")
            If IsNull(rsSR("HP_EMAIL")) Then
                txtFUEmailSendPath = ""
            Else
                txtFUEmailSendPath = rsSR("HP_EMAIL")
            End If
        End If
    End If
    
    If glbWFC Then 'Ticket #22061 Franks 05/24/2012
        If UCase(rsSR("HP_FUN_NAME")) = UCase("DEPENDENT30DAYSEMAIL") Then
            chkPreFunction(22) = rsSR("HP_ENABLED")
        End If
    End If

    '8.0 - Ticket #22682 - Employee Photo Path
    If UCase(rsSR("HP_FUN_NAME")) = UCase("EMPLOYEEPHOTOPATH") Then
        chkPreFunction(23) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtEmpPhotoPath = ""
        Else
            txtEmpPhotoPath = rsSR("HP_EMAIL")
        End If
    End If
    
    '8.0 - Ticket #22682 - Employee Address Change
    If UCase(rsSR("HP_FUN_NAME")) = UCase("EMAIL_ONADDRESSCHANGE") Then
        chkPreFunction(24) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtEmail(10) = ""
        Else
            txtEmail(10) = rsSR("HP_EMAIL")
        End If
    End If
    
    'Ticket #24485 - Work Schedule Rotation Weeks
    If UCase(rsSR("HP_FUN_NAME")) = UCase("WS_ROTATIONWEEKS") Then
        comWSRotWks.Text = rsSR("HP_NUM")
        oWSRotWks = rsSR("HP_NUM")  'keep the original value
        If Not IsDate(rsSR("HP_DATE")) Then
            dlpWSRotEffDate.Text = ""
        Else
            dlpWSRotEffDate.Text = rsSR("HP_DATE")
        End If
    End If
    
    '8.0 - Ticket #25273 - New Applicant - ATS
    If UCase(rsSR("HP_FUN_NAME")) = UCase("AT_EMAIL_ONNEWAPPLICANT") Then
        chkPreFunction(25) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtEmail(11) = ""
        Else
            txtEmail(11) = rsSR("HP_EMAIL")
        End If
    End If
    
    '8.0 - Ticket #24352 - Database Connection Encryption
    If UCase(rsSR("HP_FUN_NAME")) = UCase("DB_CONNNECT_ENCRYPT") Then chkPreFunction(26) = rsSR("HP_ENABLED")
    
    '8.1 - Ticket #26529 - SMTP Connection information
    If UCase(rsSR("HP_FUN_NAME")) = UCase("SMTP_INFORMATION") Then
        chkPreFunction(27) = rsSR("HP_ENABLED")
        'SMTP Server Name
        If IsNull(rsSR("HP_SERVER")) Then
            txtSMTPServer.Text = ""
        Else
            txtSMTPServer.Text = rsSR("HP_SERVER")
        End If
        'SMTP User Name
        If IsNull(rsSR("HP_USERNAME")) Then
            txtSMTPUsername.Text = ""
        Else
            txtSMTPUsername.Text = rsSR("HP_USERNAME")
        End If
        'SMTP Password
        If IsNull(rsSR("HP_PASSWORD")) Then
            txtSMTPPassword.Text = ""
        Else
            txtSMTPPassword.Text = rsSR("HP_PASSWORD")
        End If
        'SMTP Port
        If IsNull(rsSR("HP_PORT")) Then
            txtSMTPPort.Text = ""
        Else
            txtSMTPPort.Text = rsSR("HP_PORT")
        End If
    End If
    
    'Ticket #26576 - WDGPHU - Flex Logic
    If UCase(rsSR("HP_FUN_NAME")) = UCase("FLEX_LOGIC") Then chkPreFunction(28) = rsSR("HP_ENABLED")
    
    'Ticket #30305 - Disable Compensatory Time Entries
    If UCase(rsSR("HP_FUN_NAME")) = UCase("DISABLE_COMPTIME") Then chkPreFunction(32) = rsSR("HP_ENABLED")
    
    'Ticket #27060 - S.U.C.C.E.S.S.
    If UCase(rsSR("HP_FUN_NAME")) = UCase("EMAIL_ONREQUESTSUBMISSION") Then
        chkPreFunction(29) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtEmail(12) = ""
        Else
            txtEmail(12) = rsSR("HP_EMAIL")
        End If
    End If
    
    'Ticket #26934 - Oshawa Community Health Centre - Employee Flags
    If UCase(rsSR("HP_FUN_NAME")) = UCase("EMAIL_ONEMPLOYEEFLAGS") Then
        chkPreFunction(30) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtEmail(13) = ""
        Else
            txtEmail(13) = rsSR("HP_EMAIL")
        End If
    End If
    
    'Ticket #28815 - Opened for all so copied the WFC routine and making this general
    If UCase(rsSR("HP_FUN_NAME")) = UCase("EMAIL_ONHSINCIDENT") Then 'Ticket #28664 Franks 05/30/2016
        chkPreFunction(31) = rsSR("HP_ENABLED")
        If IsNull(rsSR("HP_EMAIL")) Then
            txtEmail(14) = ""
        Else
            txtEmail(14) = rsSR("HP_EMAIL")
        End If
    End If
    
    rsSR.MoveNext
Loop
rsSR.Close

End Sub

Private Sub ResetAll()
Dim X%

For X% = 0 To 20 '18
    chkPreFunction(X%).Value = 0
Next X%

'8.0 - Ticket #22682 - Employee Address Change
chkPreFunction(24).Value = 0

'8.0 - Ticket #25273 - New Applicant - ATS
chkPreFunction(25).Value = 0

'8.0 - Ticket #24352 - Database Connection Encryption
chkPreFunction(26).Value = 0

'8.1 - Ticket #26529 - SMTP Connection information
chkPreFunction(27).Value = 0

'Ticket #26576 - WDGPHU - Flex Logic
chkPreFunction(28).Value = 0

'Ticket #30305 - Disable Compensatory Time Entries
chkPreFunction(32).Value = 0

'Ticket #27060 - S.U.C.C.E.S.S.
chkPreFunction(29).Value = 0

'Ticket #26934 - Oshawa Community Health Centre - Employee Flags
chkPreFunction(30).Value = 0

End Sub

Private Function chkComp()
Dim X%, Response As Integer, Msg As String
Dim flgWSRuleExists As Boolean

On Error GoTo MchkComp_Err
    
    chkComp = False
    'Check if the Doc DB and table exist
    If (glbSQL Or glbOracle) And scrFrameGen.Visible = True Then
        If chkPreFunction(0).Value Then
            If gdbAdoIhr001_DOC.State = adStateOpen Then gdbAdoIhr001_DOC.Close
            gdbAdoIhr001_DOC.CommandTimeout = 600
            gdbAdoIhr001_DOC.Open glbAdoIHRDB_DOC
            Dim rsDoc As New ADODB.Recordset
            rsDoc.Open "SELECT RE_ID FROM HRDOC_EMP", gdbAdoIhr001_DOC, adOpenStatic
            rsDoc.Close
        End If
    End If
    
    '8.0 - Ticket #22682 - Employee Photo Path
    If chkPreFunction(23).Value Then
        If Len(Trim(txtEmpPhotoPath.Text)) = 0 Then
            MsgBox "'Employee Photo In Other Folder' path is blank, it will be updated with info:HR folder path.", vbInformation, "Missing Employee Photo folder path"
            txtEmpPhotoPath.Text = glbIHRREPORTS
            txtEmpPhotoPath.SetFocus
        End If
    End If
    
    'Ticket #24485 - # of WS Rotatation Week changed
    If (oWSRotWks <> comWSRotWks.Text) And scrFrameGen.Visible = True Then
        'Check if WS Rule exists then only do the rest of the checking - Jerry said
        flgWSRuleExists = False
        flgWSRuleExists = WorkSchedule_Rule_Exists
        
        'Just making sure if any data entered it is valid
        If Len(dlpWSRotEffDate.Text) > 0 Then
            If Not IsDate(dlpWSRotEffDate.Text) Then
                MsgBox "Effective Date is not a valid date."
                dlpWSRotEffDate.SetFocus
                Exit Function
            End If
        End If
        
        If flgWSRuleExists Then
            'Check if the Effective Date is entered
            If Len(dlpWSRotEffDate.Text) < 1 Then
                MsgBox "Effective Date is required when 'Number of Work Schedule Rotation Weeks' changes."
                dlpWSRotEffDate.SetFocus
                Exit Function
            End If
            
            If Not IsDate(dlpWSRotEffDate.Text) Then
                MsgBox "Effective Date is not a valid date."
                dlpWSRotEffDate.SetFocus
                Exit Function
            End If
            
            'To Date cannot be greater than the Effective date
            If WorkScheduleExists(dlpWSRotEffDate.Text) Then
                MsgBox "There cannot be any Work Schedules overlapping Effective Date. User will have to change the To Date of all those Work Schedules before making this change.", vbInformation, "Employee's Work Schedules overlaps Effective Date"
                dlpWSRotEffDate.SetFocus
                Exit Function
            End If
            
            '??? This checking has been added but Jerry has to discuss this with the client as the employees are required
            'to book their vacation/time off way in advance. In cases where there is a change in the Work Schedule plan,
            'Jerry has to find out what client does to the submitted/approved requests of future period.???
            'Check if any Unapproved Requests (ESS) exists. If it does then do not allow to save this.
            If AnyUnapprovedRequestExists(dlpWSRotEffDate.Text) Then
                'MsgBox "There are Vacation or Time Off Requests pending Approval. Cannot save this Work Schedule.", vbInformation, "Vacation/Time Off Requests pending Approval"
                MsgBox "There are Unapproved or Rejected Vacation/Time Requests outstanding for employees. These requests may be invalid based on the new 'Number of Work Schedule Rotation Weeks'. Please have the employees verify their open requests before making this change."
                dlpWSRotEffDate.SetFocus
                Exit Function
            End If
            
            
            'Users will not be able to create new work schedules when this change is Saved.
            'Verify this change.
            Response = MsgBox("By saving this change in 'Number of Work Schedule Rotation Weeks' you will not be able to create any NEW Work Schedules for employees prior to Effective Date." & vbCrLf & vbCrLf & "Are you sure you want to proceed with this change now?", vbQuestion + vbYesNo, "Creating New Work Schedules for Employee(s)")
            If Response <> 6 Then
                Exit Function
            End If
        End If
    Else
        If scrFrameGen.Visible = True Then
            'Just making sure if any data entered it is valid
            If Len(dlpWSRotEffDate.Text) > 0 Then
                If Not IsDate(dlpWSRotEffDate.Text) Then
                    MsgBox "Effective Date is not a valid date."
                    dlpWSRotEffDate.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    
    
    '8.1 - Ticket #26529 - SMTP Connection information
    If chkPreFunction(27).Value And scrFrameGen.Visible = True Then
        'SMTP information cannot be blank if SMPT Information is checked.
        
        'SMTP Server Name
        If Len(Trim(txtSMTPServer.Text)) = 0 Then
            MsgBox "SMTP Server cannot be blank if 'SMTP Information' is checked."
            txtSMTPServer.SetFocus
        End If
        'SMTP User Name - Can be blank for Annonymous account
        'If Len(Trim(txtSMTPUsername.Text)) = 0 Then
        '    MsgBox "SMTP User Name cannot be blank if 'SMTP Information' is checked."
        '    txtSMTPUsername.SetFocus
        'End If
        'SMTP Password - Can be blank for Annonymous account
        'If Len(Trim(txtSMTPPassword)) = 0 Then
        '    MsgBox "SMTP Password cannot be blank if 'SMTP Information' is checked."
        '    txtSMTPPassword.SetFocus
        'End If
        'SMTP Port - Can be blank
        If Not IsNumeric(txtSMTPPort.Text) Then
            txtSMTPPort.Text = ""
        End If
    End If
    
    chkComp = True
    Exit Function
    
MchkComp_Err:
    Msg = "The Attachment Database has not been setup yet." & Chr(10)
    Msg = Msg & "Please ask your IT person to set it up."
    MsgBox Msg
End Function

Sub cmdModify_Click()

On Error GoTo Mod_Err

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '08June99 js

End Sub

Sub cmdOK_Click()

Dim DtTm As Variant, rc As Integer, X%, Msg, Response

If Not chkComp() Then Exit Sub

'Ticket #24352 - PIPEDA - Check if the value has changed and display the appropriate message.
If chkPreFunction(26).Value <> fODBConnEncrypt And chkPreFunction(26).Value Then
    MsgBox "When you turn ON the Encryption of Database Connection, make sure you know the Password to maintain the Data Source screen.", vbOKOnly + vbInformation, "Data Source can only be maintained with Password"
End If
fODBConnEncrypt = chkPreFunction(26).Value

'Ticket #28819 - Some users with the limited permission to the registry are getting error that
'"Access to Key Software\HR Systems\ODBC SetupDRIVERNAME Denied" when they are not even using this PIPEDA function.
'So adding condition to only update registry if you are using the PIPEDA function.
If chkPreFunction(26).Visible Then
    'Ticket #24352 - PIPEDA
    'Follow the Encryption of Database Connection is turned-ON. Only show the Data Source screen if they enter the Password
    If chkPreFunction(26).Value Then
        'Encrypt Database Connection is turned-ON. Delete the Database Connection information from ODBC Setup keys and
        'create/update the License key under Options
        gsDB_CONNECT_ENCRYPT = True
        
        Call AddRemove_ODBCSetup_RegistryKey("AddLic")
        
    Else
        'Encrypt Database Connection is turned-OFF. Add the Database Connection information on ODBC Setup back and
        'remove the License Key under Options
        gsDB_CONNECT_ENCRYPT = False
        
        Call AddRemove_ODBCSetup_RegistryKey("RemoveLic")
    End If
End If

Call UpdSecAccess

Screen.MousePointer = DEFAULT

fglbNew = False

    
Exit Sub

Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRPREFERENCE", "Update")
Resume Next
Unload Me

End Sub

Sub cmdCancel_Click()

On Error GoTo Can_Err

Call Display_Values

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRPREFERENCE", "Cancel")
Resume Next

End Sub

Private Sub UpdSecAccess()
Dim SQLQ

SQLQ = "DELETE FROM HRPREFERENCE "
gdbAdoIhr001.Execute SQLQ

Call AddSecAccess

End Sub

Private Sub AddSecAccess()
Dim SQLQ, sqlI
sqlI = "INSERT INTO HRPREFERENCE(HP_CO,HP_FUN_NAME,HP_ENABLED,HP_EMAIL) "
sqlI = sqlI & " VALUES('001',"

SQLQ = sqlI & "'ATTACHMENT'," & IIf(chkPreFunction(0), 1, 0) & ",Null)"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EMAIL_SENDING'," & IIf(chkPreFunction(1), 1, 0) & ",Null)"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EMAIL_ONNEWHIRE'," & IIf(chkPreFunction(2), 1, 0) & ",'" & txtEmail(0) & "')"
gdbAdoIhr001.Execute SQLQ
'Ticket #21444 Franks 02/09/2012
SQLQ = sqlI & "'EMAIL_ONPOSITION'," & IIf(chkPreFunction(20), 1, 0) & ",'" & txtEmail(9) & "')"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EMAIL_ONSALARY'," & IIf(chkPreFunction(3), 1, 0) & ",'" & txtEmail(1) & "')"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EMAIL_ONBENEFIT'," & IIf(chkPreFunction(4), 1, 0) & ",'" & txtEmail(2) & "')"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EMAIL_ONTERM'," & IIf(chkPreFunction(5), 1, 0) & ",'" & txtEmail(3) & "')"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EMAIL_ONREHIRE'," & IIf(chkPreFunction(6), 1, 0) & ",'" & txtEmail(4) & "')"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'EMAIL_ONLEAVECHANGES'," & IIf(chkPreFunction(12), 1, 0) & ",'" & txtEmail(5) & "')"
gdbAdoIhr001.Execute SQLQ
'7.9
SQLQ = sqlI & "'EMAIL_ONPERFORMANCE'," & IIf(chkPreFunction(14), 1, 0) & ",'" & txtEmail(7) & "')"
gdbAdoIhr001.Execute SQLQ
'7.9
SQLQ = sqlI & "'EMAIL_ONDEPENDENT'," & IIf(chkPreFunction(18), 1, 0) & ",'" & txtEmail(8) & "')"
gdbAdoIhr001.Execute SQLQ

SQLQ = sqlI & "'COMPA-RATIO'," & IIf(chkPreFunction(15), 1, 0) & ",Null)"
gdbAdoIhr001.Execute SQLQ

'Ticket #18223
SQLQ = sqlI & "'EMAIL_ONREQUESTAPPROVAL'," & IIf(chkPreFunction(13), 1, 0) & ",'" & txtEmail(6) & "')"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'TRAININGMATRIX'," & IIf(chkPreFunction(7), 1, 0) & ",'" & txtTMPatch & "')"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'GP_HOLDING'," & IIf(chkPreFunction(8), 1, 0) & ",NULL)"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'SECURED_PSW'," & IIf(chkPreFunction(9), 1, 0) & ",NULL)"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'BENEFIT_HISTORY'," & IIf(chkPreFunction(10), 1, 0) & ",NULL)"
gdbAdoIhr001.Execute SQLQ
'Friesens - Ticket #17029
If glbCompSerial = "S/N - 2279W" Then
    SQLQ = sqlI & "'FRIESENSWORDPATH'," & IIf(chkPreFunction(11), 1, 0) & ",'" & txtFriesensPath.Text & "')"
    gdbAdoIhr001.Execute SQLQ
End If

'7.9 Enhacement
SQLQ = sqlI & "'SSO_INFOHR'," & IIf(chkPreFunction(16), 1, 0) & ",Null)"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'SSO_WEBSYS'," & IIf(chkPreFunction(17), 1, 0) & ",Null)"
gdbAdoIhr001.Execute SQLQ

gsAttachment_DB = IIf(chkPreFunction(0), 1, 0)

'7.9
gsCompaRatio = IIf(chkPreFunction(15), 1, 0)

'Ticket #20038 - WSIB Form 7 Path to Save
If glbWSIBModule Then
    SQLQ = sqlI & "'WSIBFORM7PATH'," & IIf(chkPreFunction(19), 1, 0) & ",'" & txtWSIBForm7Path.Text & "')"
    gdbAdoIhr001.Execute SQLQ
End If

'Ticket #21023 - Oshawa Community Health Centre - Follow Up Email Sending
If glbCompSerial = "S/N - 2396W" Then
    SQLQ = sqlI & "'FOLLOWUPEMAILLOGPATH'," & IIf(chkPreFunction(21), 1, 0) & ",'" & txtFUEmailSendPath.Text & "')"
    gdbAdoIhr001.Execute SQLQ
End If

If glbWFC Then 'Ticket #22061 Franks 05/24/2012
    SQLQ = sqlI & "'DEPENDENT30DAYSEMAIL'," & IIf(chkPreFunction(22), 1, 0) & ",'" & txtFUEmailSendPath.Text & "')"
    gdbAdoIhr001.Execute SQLQ
End If

'8.0 - Ticket #22682 - Employee Photo Path
SQLQ = sqlI & "'EMPLOYEEPHOTOPATH'," & IIf(chkPreFunction(23), 1, 0) & ",'" & txtEmpPhotoPath.Text & "')"
gdbAdoIhr001.Execute SQLQ

'8.0 - Ticket #22682 - Employee Address Change
SQLQ = sqlI & "'EMAIL_ONADDRESSCHANGE'," & IIf(chkPreFunction(24), 1, 0) & ",'" & txtEmail(10) & "')"
gdbAdoIhr001.Execute SQLQ

'Ticket #24485 - Work Schedule Rotation Weeks
sqlI = "INSERT INTO HRPREFERENCE(HP_CO,HP_FUN_NAME,HP_ENABLED,HP_EMAIL,HP_NUM,HP_DATE) "
sqlI = sqlI & " VALUES('001',"
SQLQ = sqlI & "'WS_ROTATIONWEEKS',1,Null," & comWSRotWks.Text & "," & IIf(IsDate(dlpWSRotEffDate.Text), Date_SQL(dlpWSRotEffDate.Text), Date_SQL(Date)) & ")"
gdbAdoIhr001.Execute SQLQ
oWSRotWks = comWSRotWks.Text

sqlI = "INSERT INTO HRPREFERENCE(HP_CO,HP_FUN_NAME,HP_ENABLED,HP_EMAIL) "
sqlI = sqlI & " VALUES('001',"

'8.0 - Ticket #25273 - New Applicant - ATS
SQLQ = sqlI & "'AT_EMAIL_ONNEWAPPLICANT'," & IIf(chkPreFunction(25), 1, 0) & ",'" & txtEmail(11) & "')"
gdbAdoIhr001.Execute SQLQ

'8.0 - Ticket #24352 - Database Connection Encryption
SQLQ = sqlI & "'DB_CONNNECT_ENCRYPT'," & IIf(chkPreFunction(26), 1, 0) & ",Null)"
gdbAdoIhr001.Execute SQLQ

'Ticket #26576 - WDGPHU - Flex Logic
SQLQ = sqlI & "'FLEX_LOGIC'," & IIf(chkPreFunction(28), 1, 0) & ",Null)"
gdbAdoIhr001.Execute SQLQ

'Ticket #30305 - Disable Compensatory Time Entries
SQLQ = sqlI & "'DISABLE_COMPTIME'," & IIf(chkPreFunction(32), 1, 0) & ",Null)"
gdbAdoIhr001.Execute SQLQ

'8.1 - Ticket #26529 - SMTP Connection information
sqlI = "INSERT INTO HRPREFERENCE(HP_CO,HP_FUN_NAME,HP_ENABLED,HP_EMAIL,HP_SERVER,HP_USERNAME,HP_PASSWORD,HP_PORT) "
sqlI = sqlI & " VALUES('001',"
If Trim(Len(txtSMTPPort.Text)) = 0 Then
    SQLQ = sqlI & "'SMTP_INFORMATION'," & IIf(chkPreFunction(27), 1, 0) & ",Null,'" & txtSMTPServer.Text & "','" & txtSMTPUsername.Text & "','" & txtSMTPPassword.Text & "',Null)"
Else
    SQLQ = sqlI & "'SMTP_INFORMATION'," & IIf(chkPreFunction(27), 1, 0) & ",Null,'" & txtSMTPServer.Text & "','" & txtSMTPUsername.Text & "','" & txtSMTPPassword.Text & "'," & Trim(txtSMTPPort.Text) & ")"
End If
gdbAdoIhr001.Execute SQLQ

sqlI = "INSERT INTO HRPREFERENCE(HP_CO,HP_FUN_NAME,HP_ENABLED,HP_EMAIL) "
sqlI = sqlI & " VALUES('001',"

'Ticket #27060 - S.U.C.C.E.S.S.
SQLQ = sqlI & "'EMAIL_ONREQUESTSUBMISSION'," & IIf(chkPreFunction(29), 1, 0) & ",'" & txtEmail(12) & "')"
gdbAdoIhr001.Execute SQLQ

'Ticket #26934 - Oshawa Community Health Centre - Employee Flags
SQLQ = sqlI & "'EMAIL_ONEMPLOYEEFLAGS'," & IIf(chkPreFunction(30), 1, 0) & ",'" & txtEmail(13) & "')"
gdbAdoIhr001.Execute SQLQ

'Ticket #28815 - Opened for all so copied the WFC routine and making this general
'Ticket #28664 Franks 05/30/2016 WFC - H&S Incident
SQLQ = sqlI & "'EMAIL_ONHSINCIDENT'," & IIf(chkPreFunction(31), 1, 0) & ",'" & txtEmail(14) & "')"
gdbAdoIhr001.Execute SQLQ

End Sub

Private Sub chkPreFunction_Click(Index As Integer, Value As Integer)
    If Index = 1 Then
        If Not chkPreFunction(1).Value Then
            chkPreFunction(2).Value = False
            chkPreFunction(3).Value = False
            chkPreFunction(4).Value = False
            chkPreFunction(5).Value = False
            chkPreFunction(6).Value = False
            chkPreFunction(12).Value = False
            chkPreFunction(14).Value = False
            chkPreFunction(18).Value = False
            'chkPreFunction(13).Value = False
            chkPreFunction(20).Value = False
            
            '8.0 - Ticket #22682 - Employee Address Change
            chkPreFunction(24).Value = False
        
            '8.0 - Ticket #25273 - New Applicant - ATS
            chkPreFunction(25).Value = False
            
            'Ticket #26934 - Oshawa Community Health Centre - Employee Flags
            chkPreFunction(30).Value = False
        End If
    End If
End Sub

Private Sub cmdPageRight_Click()
    scrFrameEmailNoti.Visible = False
    scrFrameEmailNoti2.Visible = True
    txtEmail(14).SetFocus
End Sub

Private Sub cmdScreenLeft_Click()
    scrFrameEmailNoti.Visible = True
    scrFrameEmailNoti2.Visible = False
End Sub

Private Sub cmdShowAddressChg_Click()
    glbEmalType = "Address Changes"
    frmComPreEmail.Show 1
End Sub

Private Sub cmdShowBenefits_Click()
    glbEmalType = "Benefits"
    frmComPreEmail.Show 1
End Sub

Private Sub cmdShowDependent_Click()
    glbEmalType = "Dependent"
    frmComPreEmail.Show 1
End Sub

Private Sub cmdShowEmpFlags_Click()
    glbEmalType = "Employee Flags"
    frmComPreEmail.Show 1
End Sub

Private Sub cmdShowHS_Click()
    glbEmalType = "H&S Incident"
    frmComPreEmail.Show 1
End Sub

Private Sub cmdShowLeaveChanges_Click()
    glbEmalType = "Leave Changes"
    frmComPreEmail.Show 1
End Sub

Private Sub cmdShowNewApplicant_Click()
    glbEmalType = "New Applicant"
    frmComPreEmail.Show 1
End Sub

Private Sub cmdShowNewhire_Click()
    glbEmalType = "New Hire"
    frmComPreEmail.Show 1
End Sub

Private Sub cmdShowPerformance_Click()
    glbEmalType = "Performance"
    frmComPreEmail.Show 1
End Sub

Private Sub cmdShowPosition_Click()
    glbEmalType = "Position"
    frmComPreEmail.Show 1
End Sub

Private Sub cmdShowRehire_Click()
    glbEmalType = "Rehire"
    frmComPreEmail.Show 1
End Sub

Private Sub cmdShowRequestApproval_Click()
    glbEmalType = "ESS-Request Approval"
    frmComPreEmail.Show 1
End Sub

Private Sub cmdShowRequestSubmission_Click()
    glbEmalType = "ESS-Request Submit"
    frmComPreEmail.Show 1
End Sub

Private Sub cmdShowSalary_Click()
    glbEmalType = "Salary"
    frmComPreEmail.Show 1
End Sub

Private Sub cmdShowTermination_Click()
    glbEmalType = "Termination"
    frmComPreEmail.Show 1
End Sub

Private Sub Form_Activate()
    glbOnTop = "FRMcomprefer"
    Call SET_UP_MODE
End Sub

Private Sub Form_Load()

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
 glbOnTop = "FRMcomprefer"
' part of question is two fields in this table
' relaying client's desire for 2 or 3 decimal salary
'retentions... only presently modifying one.

On Error GoTo Ld_Err

Screen.MousePointer = HOURGLASS

chkPreFunction(14).Caption = "  Email Notification on " & lStr("Performance")

'SSO - Web System
chkPreFunction(17).Enabled = False

'Load WS Rotation Weeks
comWSRotWks.Clear
comWSRotWks.AddItem "1"
comWSRotWks.AddItem "2"
comWSRotWks.AddItem "3"
comWSRotWks.AddItem "4"
comWSRotWks.ListIndex = 0   'Default

Call Display_Values

If Not (glbSQL Or glbOracle) Then
    chkPreFunction(0).Enabled = False
End If

'Friesens - Ticket #17029
If glbCompSerial = "S/N - 2279W" Then
    lblTitle(6).Visible = True
    chkPreFunction(11).Visible = True
    txtFriesensPath.Visible = True
End If

If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
    cmdShowNewhire.Enabled = True
    cmdShowSalary.Enabled = True
    cmdShowBenefits.Enabled = True
    cmdShowTermination.Enabled = True
    cmdShowRehire.Enabled = True
    cmdShowLeaveChanges.Enabled = True  'Ticket #18235
End If

'Ticket #18223 - Four Villages CHC
If glbCompSerial = "S/N - 2425W" Then
    lblTitle(8).Enabled = True
    chkPreFunction(13).Enabled = True
    txtEmail(6).Enabled = True
    cmdShowRequestApproval.Enabled = True
End If

'Ticket #20038 - WSIB Form 7 Path to Save
If glbWSIBModule Then
    If chkPreFunction(11).Visible = False Or chkPreFunction(8).Visible = False Then
        If chkPreFunction(11).Visible = False Then
            lblTitle(11).Top = lblTitle(6).Top '10965
            chkPreFunction(19).Top = chkPreFunction(11).Top '10680
            txtWSIBForm7Path.Top = txtFriesensPath.Top ' 10935
        'chkPreFunction(8) - not use any more - 'Ticket #21444 Franks 02/09/2012
        'ElseIf chkPreFunction(8).Visible = False Then
        '    lblTitle(11).Top = 11565
        '    chkPreFunction(19).Top = 11280
        '    txtWSIBForm7Path.Top = 11535
        End If
    End If
    lblTitle(11).Visible = True
    chkPreFunction(19).Visible = True
    txtWSIBForm7Path.Visible = True
End If

'Ticket #21023 - Oshawa Community Health Centre - Follow Up Email Sending
If glbCompSerial = "S/N - 2396W" Then
    chkPreFunction(21).Visible = True
    lblTitle(13).Visible = True
    chkPreFunction(21).Caption = lStr("Follow-ups Email Sending Log")
    txtFUEmailSendPath.Visible = True
End If

'Ticket #26934 - Oshawa Community Health Centre - Employee Flags
If glbCompSerial = "S/N - 2396W" Then
    lblTitle(20).Visible = True
    chkPreFunction(30).Visible = True
    txtEmail(13).Visible = True
    cmdShowEmpFlags.Visible = True
End If

If glbWFC Then 'Ticket #22061 Franks 05/24/2012
    chkPreFunction(22).Visible = True
    'Ticket #22409 Frank 08/08/2012
    chkPreFunction(14).Caption = "  Email Notification on Smoker Status Change"
End If

'Ticket #28815 - Opened for all so copied the WFC routine and making this general
'If glbWFC Then 'Ticket #28664 Franks 05/30/2016
    cmdPageRight.Visible = True
'End If

'8.0 - Ticket #22682 - Employee Photo Path by default
If Len(Trim(txtEmpPhotoPath.Text)) = 0 And chkPreFunction(23) Then
    'txtEmpPhotoPath.Text = glbIHRREPORTS
End If

'Ticket #24352 - PIPEDA - Only provide the Encryption functionality to the clients who have asked for it.
'Ticket #26336 - Hicks Morley Hamilton Stewart Storie LLP
If glbCompSerial = "S/N - 2468W" Then
    chkPreFunction(26).Visible = True
Else
    chkPreFunction(26).Visible = False
End If

'Ticket #24352 - PIPEDA - Store the existing value so it can be compared later on to see if the value has changed.
If chkPreFunction(26).Value Then
    fODBConnEncrypt = True
Else
    fODBConnEncrypt = False
End If

'Ticket #26576 - WDGPHU - Flex Logic
If glbCompSerial = "S/N - 2411W" Then
    chkPreFunction(28).Visible = True
    imgHelpFlexLogic.Visible = True
Else
    chkPreFunction(28).Visible = False
    imgHelpFlexLogic.Visible = False
End If

'Ticket #30305 - Disable Compensatory Time Entries
If glbCompSerial = "S/N - 2466W" Then
    chkPreFunction(32).Visible = True
    imgHelpCompTime.Visible = True
Else
    chkPreFunction(32).Visible = False
    imgHelpCompTime.Visible = False
End If

'Ticket #27060 - S.U.C.C.E.S.S.
If glbCompSerial = "S/N - 2422W" Then
    lblTitle(19).Visible = True
    chkPreFunction(29).Visible = True
    txtEmail(12).Visible = True
    cmdShowRequestSubmission.Visible = True
End If


Screen.MousePointer = DEFAULT

Exit Sub

Ld_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Company", "HRPREFERENCE", "Select")
Resume Next

End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

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

'Ticket #22682 - Release 8.0 - Break down Company Preference
Call Show_Selected_Frame

'Ticket #16967
'chkPreFunction(8).Visible = glbGP

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fglbNew = True
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_CompanyPreference
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean

If gSec_Upd_Company Then           'May99 js
    Updateble = True
Else
    Updateble = True
End If

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

If fglbNew Then
    UpdateState = NewRecord
    TF = True
'ElseIf Data1.Recordset.EOF Then
'    UpdateState = NoRecord
'    TF = False
Else
    UpdateState = OPENING
    TF = True
End If

Call set_Buttons(UpdateState)

If Not UpdateRight Then TF = False

Call ST_UPD_MODE(TF)

End Sub

Private Sub imgHelp_Click()
Dim MsgStr As String
    MsgStr = "If 'Secured Password' is checked, the user must follow these rules: "
    MsgStr = MsgStr & Chr(10) & "  1. Password must be at least eight characters."
    MsgStr = MsgStr & Chr(10) & "  2. Password must be alpha-numeric."
    MsgStr = MsgStr & Chr(10) & "  3. New password cannot equal any of the last three old passwords."
    MsgStr = MsgStr & Chr(10) & "  4. On expiration day, user must enter current password and new password before login is complete."
    MsgBox MsgStr, vbInformation
End Sub

Private Sub imgHelpFlexLogic_Click()
    Dim MsgStr As String
    MsgStr = "When 'Flex Logic' is turned ON, the user will not be able to maintain any Flex Codes (FX*) related Attendance records in info:HR except FX+Y. If you have any questions, please call info:HR support."
    MsgBox MsgStr, vbInformation
End Sub

Private Sub imgHelpCompTime_Click()
    'Ticket #30305 - Disable Compensatory Time Entries
    Dim MsgStr As String
    MsgStr = "When 'Disable Compensatory Time Entries' is turned ON, the user will not be able to enter or update any Compensatory Time (OT* and CT*) related Attendance records in info:HR." & vbCrLf & vbCrLf & "If you have any questions, please call info:HR support."
    MsgBox MsgStr, vbInformation
End Sub

Private Sub imgHelpPhoto_Click()
    Dim MsgStr As String
    MsgStr = "If this option isn't checked, photos will be in the info:HR database. Changes and additions will need to be maintained via the Mass Update 'Maintain Photos' option."
    MsgBox MsgStr, vbInformation
End Sub

Private Sub imgHelpSSOinfoHR_Click()
    Dim MsgStr As String
    MsgStr = "'Single Sign-On in info:HR' assumes that the Windows login ID matches the info:HR's User Id under the Security Master. Passwords do not need to match between info:HR and Active Directory. If you have any questions, please call info:HR support."
    MsgBox MsgStr, vbInformation
End Sub

Private Sub scrControl_Change()
    'Ticket #22682 - Release 8.0 - Break down Company Preferences
    If scrFrameGen.Visible Then
        scrFrameGen.Top = 120 - scrControl.Value
    ElseIf scrFrameEmailNoti.Visible Then
        scrFrameEmailNoti.Top = 120 - scrControl.Value
    ElseIf scrFrameFileLoc.Visible Then
        scrFrameFileLoc.Top = 120 - scrControl.Value
    End If
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then

    If scrFrameGen.Visible Then
        'Vertical scroll bar
        If Me.Height >= 5300 Then
            scrControl.Value = 0
            
            'Ticket #22682 - Release 8.0 - Break down Company Preferences
            If scrFrameGen.Visible Then
                scrFrameGen.Top = 120
            ElseIf scrFrameEmailNoti.Visible Then
                scrFrameFileLoc.Top = 120
            ElseIf scrFrameFileLoc.Visible Then
                scrFrameFileLoc.Top = 120
            End If
            scrControl.Visible = False
        Else
            scrControl.Left = Me.Width - 540    '400
            scrControl.Visible = True
            scrControl.Height = Me.Height - 1000
            If Me.Height < 5200 Then
                scrControl.Max = 4000
            Else
                scrControl.Max = 1000
            End If
            'scrControl.Left = Me.Width - scrControl.Width - 120
            'If Me.Height - scrControl.Top - 780 > 0 Then
            '    scrControl.Height = Me.Height - scrControl.Top - 780
            'End If
        End If
    ElseIf scrFrameFileLoc.Visible Then
        'Vertical scroll bar
        If Me.Height >= 3000 Then
            scrControl.Value = 0
            
            'Ticket #22682 - Release 8.0 - Break down Company Preferences
            If scrFrameGen.Visible Then
                scrFrameGen.Top = 120
            ElseIf scrFrameEmailNoti.Visible Then
                scrFrameFileLoc.Top = 120
            ElseIf scrFrameFileLoc.Visible Then
                scrFrameFileLoc.Top = 120
            End If
            scrControl.Visible = False
        Else
            scrControl.Left = Me.Width - 540    '400
            scrControl.Visible = True
            scrControl.Height = Me.Height - 1000
            If Me.Height < 2000 Then
                scrControl.Max = 1000
            Else
                scrControl.Max = 700
            End If
            'scrControl.Left = Me.Width - scrControl.Width - 120
            'If Me.Height - scrControl.Top - 780 > 0 Then
            '    scrControl.Height = Me.Height - scrControl.Top - 780
            'End If
        End If
    Else
        'Vertical scroll bar
        If Me.Height >= 10005 Then
            scrControl.Value = 0
            
            'Ticket #22682 - Release 8.0 - Break down Company Preferences
            If scrFrameGen.Visible Then
                scrFrameGen.Top = 120
            ElseIf scrFrameEmailNoti.Visible Then
                scrFrameFileLoc.Top = 120
            ElseIf scrFrameFileLoc.Visible Then
                scrFrameFileLoc.Top = 120
            End If
            scrControl.Visible = False
        Else
            scrControl.Left = Me.Width - 540    '400
            scrControl.Visible = True
            scrControl.Height = Me.Height - 1000
            If Me.Height < 8700 Then
                scrControl.Max = 4000
            Else
                scrControl.Max = 2000
            End If
            'scrControl.Left = Me.Width - scrControl.Width - 120
            'If Me.Height - scrControl.Top - 780 > 0 Then
            '    scrControl.Height = Me.Height - scrControl.Top - 780
            'End If
        End If
    End If
    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 250
    'scrFrame.Height = Me.ScaleHeight - (scrHScroll.Height - 200)  '
    If Me.Width >= 11295 Then
        scrHScroll.Value = 0
        scrHScroll.Visible = False
    Else
        scrHScroll.Visible = True
        If Me.Width < 7000 Then
            scrHScroll.Max = 100
        Else
            scrHScroll.Max = 50
        End If
        scrHScroll.Top = Me.Height - 700
        scrHScroll.Width = Me.Width - 250
    End If
    
    'scrFrame.Width = Me.Width - 800 '1860
End If
End Sub

Private Sub scrHScroll_Change()
    'Ticket #22682 - Release 8.0 - Break down Company Preferences
    If scrFrameGen.Visible Then
        scrFrameGen.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    ElseIf scrFrameEmailNoti.Visible Then
        scrFrameEmailNoti.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    ElseIf scrFrameFileLoc.Visible Then
        scrFrameFileLoc.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    End If
End Sub

Private Sub txtEmail_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Show_Selected_Frame()

scrFrameGen.Visible = False
scrFrameEmailNoti.Visible = False
scrFrameFileLoc.Visible = False

If Me.Caption = "Company Preference - General" Then
    scrFrameGen.Visible = True
    scrFrameGen.Top = 240
    scrFrameGen.Height = 5200   '2775
ElseIf Me.Caption = "Company Preference - Email Notifications" Then
    scrFrameEmailNoti.Visible = True
    scrFrameEmailNoti.Top = 240
    scrFrameEmailNoti.Height = 10695    '9015 '8300
    scrFrameEmailNoti.Left = 120
    'Ticket #28664 Franks 05/30/2016 for WFC
    scrFrameEmailNoti2.Top = 240
    scrFrameEmailNoti2.Height = 10695
    scrFrameEmailNoti2.Left = 120
ElseIf Me.Caption = "Company Preference - File Locations" Then
    scrFrameFileLoc.Visible = True
    scrFrameFileLoc.Top = 240
    scrFrameFileLoc.Height = 3135
End If
End Sub

Private Function WorkSchedule_Rule_Exists()
    Dim rsWSRule As New ADODB.Recordset
    Dim SQLQ As String
    
    SQLQ = "SELECT * FROM HRWORKSCHDRULE ORDER BY WR_ID"
    rsWSRule.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsWSRule.EOF Then
        WorkSchedule_Rule_Exists = True
    Else
        WorkSchedule_Rule_Exists = False
    End If
    rsWSRule.Close
    Set rsWSRule = Nothing
    
End Function
