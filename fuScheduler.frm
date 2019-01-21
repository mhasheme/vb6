VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmUScheduler 
   Appearance      =   0  'Flat
   Caption         =   "Work Schedule Mass Update"
   ClientHeight    =   10020
   ClientLeft      =   1305
   ClientTop       =   2625
   ClientWidth     =   12180
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10020
   ScaleWidth      =   12180
   Tag             =   "Other Earnings Mass Update"
   WindowState     =   2  'Maximized
   Begin VB.Frame frWeek1 
      Caption         =   "Week 1"
      Enabled         =   0   'False
      Height          =   3015
      Left            =   240
      TabIndex        =   83
      Top             =   5040
      Width           =   2535
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "SD_MON_HRS"
         Height          =   285
         Index           =   1
         Left            =   1275
         TabIndex        =   15
         Tag             =   "11-Work Hours of the Day"
         Top             =   720
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "SD_TUE_HRS"
         Height          =   285
         Index           =   2
         Left            =   1275
         TabIndex        =   16
         Tag             =   "11-Work Hours of the Day"
         Top             =   1080
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "SD_WED_HRS"
         Height          =   285
         Index           =   3
         Left            =   1275
         TabIndex        =   17
         Tag             =   "11-Work Hours of the Day"
         Top             =   1440
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "SD_FRI_HRS"
         Height          =   285
         Index           =   5
         Left            =   1275
         TabIndex        =   19
         Tag             =   "11-Work Hours of the Day"
         Top             =   2160
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "SD_THU_HRS"
         Height          =   285
         Index           =   4
         Left            =   1275
         TabIndex        =   18
         Tag             =   "11-Work Hours of the Day"
         Top             =   1800
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "SD_SAT_HRS"
         Height          =   285
         Index           =   6
         Left            =   1275
         TabIndex        =   20
         Tag             =   "11-Work Hours of the Day"
         Top             =   2520
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "SD_SUN_HRS"
         Height          =   285
         Index           =   0
         Left            =   1275
         TabIndex        =   14
         Tag             =   "11-Work Hours of the Day"
         Top             =   360
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Saturday"
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
         TabIndex        =   90
         Top             =   2535
         Width           =   630
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Thursday"
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
         Left            =   120
         TabIndex        =   89
         Top             =   1815
         Width           =   660
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Friday"
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
         TabIndex        =   88
         Top             =   2175
         Width           =   420
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Wednesday"
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
         Left            =   120
         TabIndex        =   87
         Top             =   1455
         Width           =   855
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tuesday"
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
         TabIndex        =   86
         Top             =   1095
         Width           =   615
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monday"
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
         TabIndex        =   85
         Top             =   735
         Width           =   570
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sunday"
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
         Left            =   120
         TabIndex        =   84
         Top             =   375
         Width           =   540
      End
   End
   Begin VB.Frame frWeek2 
      Caption         =   "Week 2"
      Enabled         =   0   'False
      Height          =   3015
      Left            =   2880
      TabIndex        =   75
      Top             =   5040
      Width           =   2535
      Begin MSMask.MaskEdBox medHours2 
         DataField       =   "SD_SUN_HRS2"
         Height          =   285
         Index           =   0
         Left            =   1275
         TabIndex        =   21
         Tag             =   "11-Work Hours of the Day"
         Top             =   360
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours2 
         DataField       =   "SD_MON_HRS2"
         Height          =   285
         Index           =   1
         Left            =   1275
         TabIndex        =   22
         Tag             =   "11-Work Hours of the Day"
         Top             =   720
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours2 
         DataField       =   "SD_TUE_HRS2"
         Height          =   285
         Index           =   2
         Left            =   1275
         TabIndex        =   23
         Tag             =   "11-Work Hours of the Day"
         Top             =   1080
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours2 
         DataField       =   "SD_WED_HRS2"
         Height          =   285
         Index           =   3
         Left            =   1275
         TabIndex        =   24
         Tag             =   "11-Work Hours of the Day"
         Top             =   1440
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours2 
         DataField       =   "SD_FRI_HRS2"
         Height          =   285
         Index           =   5
         Left            =   1275
         TabIndex        =   26
         Tag             =   "11-Work Hours of the Day"
         Top             =   2160
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours2 
         DataField       =   "SD_THU_HRS2"
         Height          =   285
         Index           =   4
         Left            =   1275
         TabIndex        =   25
         Tag             =   "11-Work Hours of the Day"
         Top             =   1800
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours2 
         DataField       =   "SD_SAT_HRS2"
         Height          =   285
         Index           =   6
         Left            =   1275
         TabIndex        =   27
         Tag             =   "11-Work Hours of the Day"
         Top             =   2520
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label lblTitle2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Saturday"
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
         Left            =   120
         TabIndex        =   82
         Top             =   2535
         Width           =   630
      End
      Begin VB.Label lblTitle2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Thursday"
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
         TabIndex        =   81
         Top             =   1815
         Width           =   660
      End
      Begin VB.Label lblTitle2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Friday"
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
         Left            =   120
         TabIndex        =   80
         Top             =   2175
         Width           =   420
      End
      Begin VB.Label lblTitle2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Wednesday"
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
         TabIndex        =   79
         Top             =   1455
         Width           =   855
      End
      Begin VB.Label lblTitle2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tuesday"
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
         Left            =   120
         TabIndex        =   78
         Top             =   1095
         Width           =   615
      End
      Begin VB.Label lblTitle2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monday"
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
         TabIndex        =   77
         Top             =   735
         Width           =   570
      End
      Begin VB.Label lblTitle2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sunday"
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
         TabIndex        =   76
         Top             =   375
         Width           =   540
      End
   End
   Begin VB.Frame frWeek3 
      Caption         =   "Week 3"
      Enabled         =   0   'False
      Height          =   3015
      Left            =   5520
      TabIndex        =   67
      Top             =   5040
      Width           =   2535
      Begin MSMask.MaskEdBox medHours3 
         DataField       =   "SD_SUN_HRS3"
         Height          =   285
         Index           =   0
         Left            =   1275
         TabIndex        =   28
         Tag             =   "11-Work Hours of the Day"
         Top             =   360
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours3 
         DataField       =   "SD_MON_HRS3"
         Height          =   285
         Index           =   1
         Left            =   1275
         TabIndex        =   29
         Tag             =   "11-Work Hours of the Day"
         Top             =   720
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours3 
         DataField       =   "SD_TUE_HRS3"
         Height          =   285
         Index           =   2
         Left            =   1275
         TabIndex        =   30
         Tag             =   "11-Work Hours of the Day"
         Top             =   1080
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours3 
         DataField       =   "SD_WED_HRS3"
         Height          =   285
         Index           =   3
         Left            =   1275
         TabIndex        =   31
         Tag             =   "11-Work Hours of the Day"
         Top             =   1440
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours3 
         DataField       =   "SD_FRI_HRS3"
         Height          =   285
         Index           =   5
         Left            =   1275
         TabIndex        =   33
         Tag             =   "11-Work Hours of the Day"
         Top             =   2160
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours3 
         DataField       =   "SD_THU_HRS3"
         Height          =   285
         Index           =   4
         Left            =   1275
         TabIndex        =   32
         Tag             =   "11-Work Hours of the Day"
         Top             =   1800
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours3 
         DataField       =   "SD_SAT_HRS3"
         Height          =   285
         Index           =   6
         Left            =   1275
         TabIndex        =   34
         Tag             =   "11-Work Hours of the Day"
         Top             =   2520
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label lblTitle3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Saturday"
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
         Left            =   120
         TabIndex        =   74
         Top             =   2535
         Width           =   630
      End
      Begin VB.Label lblTitle3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Thursday"
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
         TabIndex        =   73
         Top             =   1815
         Width           =   660
      End
      Begin VB.Label lblTitle3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Friday"
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
         Left            =   120
         TabIndex        =   72
         Top             =   2175
         Width           =   420
      End
      Begin VB.Label lblTitle3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Wednesday"
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
         TabIndex        =   71
         Top             =   1455
         Width           =   855
      End
      Begin VB.Label lblTitle3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tuesday"
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
         Left            =   120
         TabIndex        =   70
         Top             =   1095
         Width           =   615
      End
      Begin VB.Label lblTitle3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monday"
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
         TabIndex        =   69
         Top             =   735
         Width           =   570
      End
      Begin VB.Label lblTitle3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sunday"
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
         TabIndex        =   68
         Top             =   375
         Width           =   540
      End
   End
   Begin VB.Frame frWeek4 
      Caption         =   "Week 4"
      Enabled         =   0   'False
      Height          =   3015
      Left            =   8160
      TabIndex        =   59
      Top             =   5040
      Width           =   2535
      Begin MSMask.MaskEdBox medHours4 
         DataField       =   "SD_SUN_HRS4"
         Height          =   285
         Index           =   0
         Left            =   1275
         TabIndex        =   35
         Tag             =   "11-Work Hours of the Day"
         Top             =   360
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours4 
         DataField       =   "SD_MON_HRS4"
         Height          =   285
         Index           =   1
         Left            =   1275
         TabIndex        =   36
         Tag             =   "11-Work Hours of the Day"
         Top             =   720
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours4 
         DataField       =   "SD_TUE_HRS4"
         Height          =   285
         Index           =   2
         Left            =   1275
         TabIndex        =   37
         Tag             =   "11-Work Hours of the Day"
         Top             =   1080
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours4 
         DataField       =   "SD_WED_HRS4"
         Height          =   285
         Index           =   3
         Left            =   1275
         TabIndex        =   38
         Tag             =   "11-Work Hours of the Day"
         Top             =   1440
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours4 
         DataField       =   "SD_FRI_HRS4"
         Height          =   285
         Index           =   5
         Left            =   1275
         TabIndex        =   40
         Tag             =   "11-Work Hours of the Day"
         Top             =   2160
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours4 
         DataField       =   "SD_THU_HRS4"
         Height          =   285
         Index           =   4
         Left            =   1275
         TabIndex        =   39
         Tag             =   "11-Work Hours of the Day"
         Top             =   1800
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours4 
         DataField       =   "SD_SAT_HRS4"
         Height          =   285
         Index           =   6
         Left            =   1275
         TabIndex        =   41
         Tag             =   "11-Work Hours of the Day"
         Top             =   2520
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label lblTitle4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Saturday"
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
         Left            =   120
         TabIndex        =   66
         Top             =   2535
         Width           =   630
      End
      Begin VB.Label lblTitle4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Thursday"
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
         TabIndex        =   65
         Top             =   1815
         Width           =   660
      End
      Begin VB.Label lblTitle4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Friday"
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
         Left            =   120
         TabIndex        =   64
         Top             =   2175
         Width           =   420
      End
      Begin VB.Label lblTitle4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Wednesday"
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
         TabIndex        =   63
         Top             =   1455
         Width           =   855
      End
      Begin VB.Label lblTitle4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tuesday"
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
         Left            =   120
         TabIndex        =   62
         Top             =   1095
         Width           =   615
      End
      Begin VB.Label lblTitle4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monday"
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
         TabIndex        =   61
         Top             =   735
         Width           =   570
      End
      Begin VB.Label lblTitle4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sunday"
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
         TabIndex        =   60
         Top             =   375
         Width           =   540
      End
   End
   Begin VB.TextBox memComments 
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
      Height          =   855
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   42
      Tag             =   "00-Enter Comments"
      Top             =   8400
      Width           =   8805
   End
   Begin INFOHR_Controls.DateLookup dlpEffectiveDate 
      Height          =   285
      Left            =   2010
      TabIndex        =   10
      Tag             =   "40-Work Schedule From Date"
      Top             =   4320
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   2010
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1725
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   2010
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2055
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   2010
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1395
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   2010
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1050
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   2010
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   720
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   2010
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   390
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   2010
      TabIndex        =   7
      Tag             =   "00-Enter Administered By Code"
      Top             =   2715
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   2010
      TabIndex        =   8
      Tag             =   "00-Enter Section Code"
      Top             =   3060
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   2010
      TabIndex        =   6
      Tag             =   "00-Enter Region Code"
      Top             =   2385
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   2010
      TabIndex        =   9
      Tag             =   "10-Enter Employee Number"
      Top             =   3390
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   6835
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin MSMask.MaskEdBox medHrsDay 
      Height          =   285
      Left            =   2340
      TabIndex        =   12
      Tag             =   "10-Usual work hours per day"
      Top             =   4680
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   5
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
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   9600
      Top             =   7800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin INFOHR_Controls.DateLookup dlpToDate 
      DataField       =   "SD_TDATE"
      Height          =   285
      Left            =   6720
      TabIndex        =   11
      Tag             =   "41-Work Schedule To Date"
      Top             =   4320
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   503
      MultiSelect     =   -1  'True
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpChangeDate 
      Height          =   285
      Left            =   6720
      TabIndex        =   13
      Tag             =   "41-Work Schedule To Date"
      Top             =   4680
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   503
      MultiSelect     =   -1  'True
      TextBoxWidth    =   1215
   End
   Begin VB.Label lblMsg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "*** 'Number of Work Schedule Rotation Weeks' not setup on the Company Preference screen."
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   92
      Top             =   9600
      Visible         =   0   'False
      Width           =   8010
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Change Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   32
      Left            =   5520
      TabIndex        =   91
      Top             =   4725
      Width           =   1125
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   5520
      TabIndex        =   58
      Top             =   4335
      Width           =   705
   End
   Begin VB.Label lblHrsDay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Default Hours/Day"
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
      Left            =   240
      TabIndex        =   57
      Top             =   4725
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
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
      Left            =   240
      TabIndex        =   56
      Top             =   8160
      Width           =   990
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
      Left            =   330
      TabIndex        =   55
      Top             =   3105
      Width           =   540
   End
   Begin VB.Label lblRegion 
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
      Left            =   330
      TabIndex        =   54
      Top             =   2430
      Width           =   510
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
      Left            =   330
      TabIndex        =   53
      Top             =   2760
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
      Left            =   330
      TabIndex        =   52
      Top             =   1095
      Width           =   615
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
      Left            =   330
      TabIndex        =   51
      Top             =   1770
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
      Left            =   330
      TabIndex        =   50
      Top             =   1440
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
      Left            =   330
      TabIndex        =   49
      Top             =   765
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
      Left            =   330
      TabIndex        =   48
      Top             =   435
      Width           =   555
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Index           =   0
      Left            =   330
      TabIndex        =   47
      Top             =   3435
      Width           =   1290
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
      Left            =   330
      TabIndex        =   46
      Top             =   2100
      Width           =   630
   End
   Begin VB.Label lblCostEmp 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Schedule Update"
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
      Left            =   60
      TabIndex        =   45
      Top             =   4000
      Width           =   3525
   End
   Begin VB.Label lblEffDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   44
      Top             =   4365
      Width           =   1050
   End
   Begin VB.Label lblSelCri 
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
      Left            =   150
      TabIndex        =   43
      Top             =   90
      Width           =   1575
   End
End
Attribute VB_Name = "frmUScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XUpdCount, Actn
Dim xWeek1Changed As Boolean
Dim xWeek2Changed As Boolean
Dim xWeek3Changed As Boolean
Dim xWeek4Changed As Boolean

Private Function chkScheduler()
Dim dd&
Dim Msg$, DgDef As Variant, Response%
Dim X%
Dim flgHrsEntered As Boolean

chkScheduler = False

On Error GoTo chkScheduler_Err

If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    MsgBox lStr("If Division Entered - it must be known")
     clpDiv.SetFocus
    Exit Function
End If

If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
     clpDept.SetFocus
    Exit Function
End If

For X% = 0 To 5
    If Len(clpCode(X%).Text) > 0 And clpCode(X%).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(X%).SetFocus
        Exit Function
    End If
Next X%

If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    MsgBox lStr("Category code must be valid")
    clpPT.SetFocus
    Exit Function
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

If Len(dlpEffectiveDate.Text) < 1 Then
    MsgBox "From Date is mandatory"
    dlpEffectiveDate.SetFocus
    Exit Function
Else
    If Not IsDate(dlpEffectiveDate.Text) Then
        MsgBox "From Date must be valid"
        dlpEffectiveDate.SetFocus
        Exit Function
    End If
End If

If Len(dlpToDate.Text) < 1 Then
    MsgBox "To Date is required."
    dlpToDate.SetFocus
    Exit Function
End If

If Not IsDate(dlpToDate.Text) Then
    MsgBox "To Date is not a valid date."
    dlpToDate.SetFocus
    Exit Function
End If

If CVDate(dlpEffectiveDate.Text) > CVDate(dlpToDate.Text) Then
    MsgBox "From From Date cannot be greater than To Date"
    dlpEffectiveDate.SetFocus
    Exit Function
End If

If Len(medHrsDay) > 0 Then
    If Not IsNumeric(medHrsDay) Then
        MsgBox "Default Hours/Day is not valid"
        medHrsDay.SetFocus
        Exit Function
    End If
End If

flgHrsEntered = False
'Ticket #24485 - # of WS Rotation Weeks
If frWeek1.Enabled Then
    For X = 0 To 6
        If Len(medHours(X)) > 0 Then
            If Not IsNumeric(medHours(X)) Then
                MsgBox lblTitle(X + 2).Caption & " hours is not valid"
                medHours(X).SetFocus
                Exit Function
            End If
            flgHrsEntered = True
        End If
    Next
End If
If frWeek2.Enabled Then
    For X = 0 To 6
        If Len(medHours2(X)) > 0 Then
            If Not IsNumeric(medHours2(X)) Then
                MsgBox lblTitle2(X).Caption & " hours is not valid"
                medHours2(X).SetFocus
                Exit Function
            End If
            flgHrsEntered = True
        End If
    Next
End If
If frWeek3.Enabled Then
    For X = 0 To 6
        If Len(medHours3(X)) > 0 Then
            If Not IsNumeric(medHours3(X)) Then
                MsgBox lblTitle3(X).Caption & " hours is not valid"
                medHours3(X).SetFocus
                Exit Function
            End If
            flgHrsEntered = True
        End If
    Next
End If
If frWeek4.Enabled Then
    For X = 0 To 6
        If Len(medHours4(X)) > 0 Then
            If Not IsNumeric(medHours4(X)) Then
                MsgBox lblTitle4(X).Caption & " hours is not valid"
                medHours4(X).SetFocus
                Exit Function
            End If
            flgHrsEntered = True
        End If
    Next
End If

If flgHrsEntered = False And Actn <> "D" Then
    MsgBox "There are no valid work hours to update Work Schedule with."
    Exit Function
End If

chkScheduler = True

Exit Function

chkScheduler_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkScheduler", "HR_SCHEDULER", "Update")
Resume Next

End Function

Public Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdNew_Click()
Dim Msg$, DgDef As Variant, Response%
Dim dd&
Dim Title$
Dim recCount As Integer

If Not gSec_Upd_Work_Schedule Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Actn = "A"

If Not chkScheduler() Then Exit Sub

Title$ = "Mass Add Work Schedule"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to Add Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Add
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Work Schedule Record " Else Msg$ = Msg$ & " Work Schedule Records "
    Msg$ = Msg$ & "will be Added. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Employee record found to add the Work Schedule record or more recent Work Schedule already exists."
    Exit Sub
End If

Screen.MousePointer = HOURGLASS

If Not modInsRecs() Then
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

Screen.MousePointer = DEFAULT
If XUpdCount > 0 Then
    MsgBox Str(XUpdCount) & " Records Added Successfully."
Else
    MsgBox "No Records Added."
End If

End Sub

Public Sub cmdDelete_Click()
Dim a As Integer
Dim SQLQ As String, rc%, DtTm As Variant, X%
Dim DgDef, Title$, Msg$, Response%
Dim recCount As Integer

If Not gSec_Upd_Work_Schedule Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Actn = "D"

If Not chkScheduler() Then Exit Sub

Title$ = "Mass Work Schedule Delete"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are You Sure You Want To Delete ALL records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Delete
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Work Schedule Record " Else Msg$ = Msg$ & " Work Schedule Records "
    Msg$ = Msg$ & "will be Deleted. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, Title)     ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Work Schedule record found to delete."
    Exit Sub
End If

If Not modDelRecs Then Exit Sub

Screen.MousePointer = DEFAULT
If XUpdCount > 0 Then
    MsgBox Str(XUpdCount) & " Records Deleted Successfully."
Else
    MsgBox "No Records Deleted."
End If
Screen.MousePointer = DEFAULT
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "Work Schedule", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Public Sub cmdModify_Click()
Dim Msg$, DgDef As Variant, Response%
Dim dd&
Dim Title$
Dim recCount As Integer

If Not gSec_Upd_Work_Schedule Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Actn = "M"

If Not chkScheduler() Then Exit Sub

Title$ = "Mass Update Work Schedule"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to Update Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Modify
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Work Schedule Record " Else Msg$ = Msg$ & " Work Schedule Records "
    Msg$ = Msg$ & "will be Updated. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Work Schedule record found to update."
    Exit Sub
End If

Screen.MousePointer = HOURGLASS

If Not modUptRecs() Then Exit Sub

Screen.MousePointer = DEFAULT
If XUpdCount > 0 Then
    MsgBox Str(XUpdCount) & " Records Updated Successfully."
Else
    MsgBox "No Records Updated."
End If

End Sub

Private Sub dlpEffectiveDate_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMUSCHEDULER"
End Sub

Private Sub Form_Load()
glbOnTop = "FRMUSCHEDULER"

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Screen.MousePointer = HOURGLASS

Call setRptCaption(Me)

If glbCompSerial = "S/N - 2227W" Then clpCode(4).MaxLength = 6

'Ticket #24485 - Enable/Disable the Works Schedule Rotation Weeks based on the Company Preference selection
Call Enable_Disable_RotationWeeks(gsWS_ROTATIONWEEKS)

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Function modUptRecs()

Dim Msg$, DgDef As Variant, Response%, noRecs&
Dim rsSchedule As New ADODB.Recordset
Dim SQLQ, X%, strFld

modUptRecs = False

On Error GoTo cmdUpdErr

SQLQ = "SELECT * FROM HR_SCHEDULER WHERE "
SQLQ = SQLQ & " SD_EDATE = " & Date_SQL(dlpEffectiveDate.Text)
SQLQ = SQLQ & " AND SD_TDATE = " & Date_SQL(dlpToDate.Text)
SQLQ = SQLQ & " AND SD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ & ")"
rsSchedule.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
XUpdCount = rsSchedule.RecordCount
Do While Not rsSchedule.EOF
    
    'Check if any Unapproved Requests (ESS) exists. If it does then do not allow to save this.
    'If UnapprovedRequestExists(rsSchedule("SD_EMPNBR")) Then
     If UnapprovedRequestExistsFromToDt(rsSchedule("SD_EMPNBR"), dlpEffectiveDate.Text) Then
        Response% = MsgBox("There are Vacation or Time Off Requests pending Approval for some employee(s). Transaction cannot be processed. Do you want a list of Vacation Requests Pending Approval?" & vbCrLf & vbCrLf & "Click 'Yes' to view the list of employees; click 'No' to abort the Mass Update process.", vbQuestion + vbYesNo, "Vacation/Time Off Requests pending Approval")
        
        If Response% = IDYES Then
            'Run the report
            Call ShowReport_UnapprovedRequests
            Screen.MousePointer = DEFAULT
            Exit Function
        End If
        Screen.MousePointer = DEFAULT
        Exit Function
    ElseIf UnapprovedRequestExistsFromToDt(rsSchedule("SD_EMPNBR"), dlpToDate.Text) Then
        Response% = MsgBox("There are Vacation or Time Off Requests pending Approval for some employee(s). Transaction cannot be processed. Do you want a list of Vacation Requests Pending Approval?" & vbCrLf & vbCrLf & "Click 'Yes' to view the list of employees; click 'No' to abort the Mass Update process.", vbQuestion + vbYesNo, "Vacation/Time Off Requests pending Approval")
        
        If Response% = IDYES Then
            'Run the report
            Call ShowReport_UnapprovedRequests
            Screen.MousePointer = DEFAULT
            Exit Function
        End If
        Screen.MousePointer = DEFAULT
        Exit Function
    ElseIf UnapprovedRequestExistsFromToDt(rsSchedule("SD_EMPNBR"), dlpEffectiveDate.Text, dlpToDate.Text) Then
        Response% = MsgBox("There are Vacation or Time Off Requests pending Approval for some employee(s). Transaction cannot be processed. Do you want a list of Vacation Requests Pending Approval?" & vbCrLf & vbCrLf & "Click 'Yes' to view the list of employees; click 'No' to abort the Mass Update process.", vbQuestion + vbYesNo, "Vacation/Time Off Requests pending Approval")
        
        If Response% = IDYES Then
            'Run the report
            Call ShowReport_UnapprovedRequests
            Screen.MousePointer = DEFAULT
            Exit Function
        End If
        Screen.MousePointer = DEFAULT
        Exit Function
    End If
    
    'Any Requests exists from ChangedDate?
    If UnapprovedRequestExistsChangeDate(rsSchedule("SD_EMPNBR"), dlpChangeDate.Text) Then
        Response% = MsgBox("There are Vacation or Time Off Requests pending Approval for some employee(s) from the 'Change Date'. Transaction cannot be processed. Do you want a list of Vacation Requests Pending Approval?" & vbCrLf & vbCrLf & "Click 'Yes' to view the list of employees; click 'No' to abort the Mass Update process.", vbQuestion + vbYesNo, "Vacation/Time Off Requests pending Approval")
        If Response% = IDYES Then
            'Run the report
            Call ShowReport_UnapprovedRequests(dlpChangeDate.Text)
            Screen.MousePointer = DEFAULT
            Exit Function
        End If
        Screen.MousePointer = DEFAULT
        Exit Function
    End If
    
    rsSchedule.MoveNext
Loop
rsSchedule.Close
Set rsSchedule = Nothing

'More checking
'Ticket #24485 - To Date cannot be greater than Effective Date. This is to freeze any new WS entry when another
'rotation weeks will be coming into play in near future.
If IsDate(gsWS_ROTATIONWEEKSEFFDATE) Then
    If CVDate(dlpEffectiveDate.Text) < CVDate(gsWS_ROTATIONWEEKSEFFDATE) Then  'Not already in the new WS Rotation Weeks
        If CVDate(dlpToDate.Text) >= CVDate(gsWS_ROTATIONWEEKSEFFDATE) Then
            MsgBox "To Date cannot be greater or equal to 'Effective Date' of the new '# of Work Schedule Rotation Weeks." & vbCrLf & vbCrLf & "- Number of Work Schedule Rotation Weeks = " & gsWS_ROTATIONWEEKS & vbCrLf & "- Effective Date = " & gsWS_ROTATIONWEEKSEFFDATE
            dlpToDate.SetFocus
            Screen.MousePointer = DEFAULT
            Exit Function
        End If
    End If
End If

If Len(dlpChangeDate.Text) < 1 Then
    MsgBox "Change Date is required."
    dlpChangeDate.SetFocus
    Screen.MousePointer = DEFAULT
    Exit Function
End If

If Not IsDate(dlpChangeDate.Text) Then
    MsgBox "Change Date is not a valid date."
    dlpChangeDate.SetFocus
    Screen.MousePointer = DEFAULT
    Exit Function
End If

If CVDate(dlpChangeDate.Text) < CVDate(dlpEffectiveDate.Text) Then
    MsgBox "Change Date cannot be prior to From Date"
    dlpChangeDate.SetFocus
    Screen.MousePointer = DEFAULT
    Exit Function
End If

If CVDate(dlpChangeDate.Text) > CVDate(dlpToDate.Text) Then
    MsgBox "Change Date cannot be greater than To Date"
    dlpChangeDate.SetFocus
    Screen.MousePointer = DEFAULT
    Exit Function
End If

'Confirm Hours change after indicating the impact it may cause
Response% = MsgBox("Any changes to the Employee(s) Work Schedule(s) hours will not be reflected in any submitted or approved Vacation/Time Requests. Prior to making this change, future-dated Vacation/Time Requests should be deleted." & vbCrLf & vbCrLf & "Are you sure you want to proceed?", vbQuestion + vbYesNo, "Submitted/Approved Vacation/Time Off Requests")
If Response% <> 6 Then
    Screen.MousePointer = DEFAULT
    Exit Function
End If


'Update WS
SQLQ = "SELECT * FROM HR_SCHEDULER WHERE "
SQLQ = SQLQ & " SD_EDATE = " & Date_SQL(dlpEffectiveDate.Text)
SQLQ = SQLQ & " AND SD_TDATE = " & Date_SQL(dlpToDate.Text)
SQLQ = SQLQ & " AND SD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ & ")"
rsSchedule.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
XUpdCount = rsSchedule.RecordCount

Do While Not rsSchedule.EOF
    'Ticket #24485 - # of WS Rotation Weeks
    
    'Enable/Disable Rotation Weeks first based on the # of Rotation Weeks saved
    Call Enable_Disable_RotationWeeks(rsSchedule("SD_ROTWKS"))
    
    If frWeek1.Enabled Then
        If IsNumeric(medHours(0)) Then
            rsSchedule("SD_SUN_HRS") = medHours(0)
        End If
        If IsNumeric(medHours(1)) Then
            rsSchedule("SD_MON_HRS") = medHours(1)
        End If
        If IsNumeric(medHours(2)) Then
            rsSchedule("SD_TUE_HRS") = medHours(2)
        End If
        If IsNumeric(medHours(3)) Then
            rsSchedule("SD_WED_HRS") = medHours(3)
        End If
        If IsNumeric(medHours(4)) Then
            rsSchedule("SD_THU_HRS") = medHours(4)
        End If
        If IsNumeric(medHours(5)) Then
            rsSchedule("SD_FRI_HRS") = medHours(5)
        End If
        If IsNumeric(medHours(6)) Then
            rsSchedule("SD_SAT_HRS") = medHours(6)
        End If
    End If
    If frWeek2.Enabled Then
        If IsNumeric(medHours2(0)) Then
            rsSchedule("SD_SUN_HRS2") = medHours2(0)
        End If
        If IsNumeric(medHours2(1)) Then
            rsSchedule("SD_MON_HRS2") = medHours2(1)
        End If
        If IsNumeric(medHours2(2)) Then
            rsSchedule("SD_TUE_HRS2") = medHours2(2)
        End If
        If IsNumeric(medHours2(3)) Then
            rsSchedule("SD_WED_HRS2") = medHours2(3)
        End If
        If IsNumeric(medHours2(4)) Then
            rsSchedule("SD_THU_HRS2") = medHours2(4)
        End If
        If IsNumeric(medHours2(5)) Then
            rsSchedule("SD_FRI_HRS2") = medHours2(5)
        End If
        If IsNumeric(medHours2(6)) Then
            rsSchedule("SD_SAT_HRS2") = medHours2(6)
        End If
    End If
    If frWeek3.Enabled Then
        If IsNumeric(medHours3(0)) Then
            rsSchedule("SD_SUN_HRS3") = medHours3(0)
        End If
        If IsNumeric(medHours3(1)) Then
            rsSchedule("SD_MON_HRS3") = medHours3(1)
        End If
        If IsNumeric(medHours3(2)) Then
            rsSchedule("SD_TUE_HRS3") = medHours3(2)
        End If
        If IsNumeric(medHours3(3)) Then
            rsSchedule("SD_WED_HRS3") = medHours3(3)
        End If
        If IsNumeric(medHours3(4)) Then
            rsSchedule("SD_THU_HRS3") = medHours3(4)
        End If
        If IsNumeric(medHours3(5)) Then
            rsSchedule("SD_FRI_HRS3") = medHours3(5)
        End If
        If IsNumeric(medHours3(6)) Then
            rsSchedule("SD_SAT_HRS3") = medHours3(6)
        End If
    End If
    If frWeek4.Enabled Then
        If IsNumeric(medHours4(0)) Then
            rsSchedule("SD_SUN_HRS4") = medHours4(0)
        End If
        If IsNumeric(medHours4(1)) Then
            rsSchedule("SD_MON_HRS4") = medHours4(1)
        End If
        If IsNumeric(medHours4(2)) Then
            rsSchedule("SD_TUE_HRS4") = medHours4(2)
        End If
        If IsNumeric(medHours4(3)) Then
            rsSchedule("SD_WED_HRS4") = medHours4(3)
        End If
        If IsNumeric(medHours4(4)) Then
            rsSchedule("SD_THU_HRS4") = medHours4(4)
        End If
        If IsNumeric(medHours4(5)) Then
            rsSchedule("SD_FRI_HRS4") = medHours4(5)
        End If
        If IsNumeric(medHours4(6)) Then
            rsSchedule("SD_SAT_HRS4") = medHours4(6)
        End If
    End If
    
    rsSchedule("SD_COMMENTS") = memComments.Text
    rsSchedule("SD_LDATE") = Date
    rsSchedule("SD_LTIME") = Time$
    rsSchedule("SD_LUSER") = glbUserID
    rsSchedule.Update
    
    'Ticket #24485 - # of WS Rotation Weeks
    'Update Work Schedule Details as well
    Call Update_WorkSchedule_Detail(rsSchedule("SD_EMPNBR"))
    
    
    rsSchedule.MoveNext
Loop
rsSchedule.Close
Set rsSchedule = Nothing

modUptRecs = True

Exit Function

cmdUpdErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Mass change", "HR_SCHEDULER", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If

End Function

Private Function modInsRecs()

Dim Msg$, DgDef As Variant, Response%, noRecs&
Dim rsTA As New ADODB.Recordset
Dim rsSchedule As New ADODB.Recordset
Dim SQLQ, strFld

modInsRecs = False

On Error GoTo cmdInsErr

SQLQ = "SELECT * FROM HR_SCHEDULER WHERE "
SQLQ = SQLQ & " ((SD_EDATE <= " & Date_SQL(dlpEffectiveDate.Text)
SQLQ = SQLQ & " AND SD_TDATE >= " & Date_SQL(dlpEffectiveDate.Text) & ")"
SQLQ = SQLQ & " OR (SD_EDATE >= " & Date_SQL(dlpEffectiveDate.Text) & "))"
SQLQ = SQLQ & " AND SD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ & ")"
rsSchedule.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
XUpdCount = rsSchedule.RecordCount
Do While Not rsSchedule.EOF
    
    'Validation on Adding schedule
    'Check if the same Effective Date schedule already exists. If it does then do not allow to save this.
    If ScheduleAlreadyExists(rsSchedule("SD_EMPNBR"), dlpEffectiveDate.Text) Then
    'If ScheduleAlreadyExistsFromToDt(rsSchedule("SD_EMPNBR"), dlpEffectiveDate.Text) Then
        Response% = MsgBox("Work Schedule for this Effective Date already exists for some employee(s). Do you want to skip those employees?" & vbCrLf & vbCrLf & "Click 'Yes' to skip; click 'No' to abort the Mass Update process.", vbQuestion + vbYesNo, "Work Schedule already exists")
        If Response% = IDNO Then
            Exit Function
        End If
    End If
    If ScheduleAlreadyExistsFromToDt(rsSchedule("SD_EMPNBR"), dlpEffectiveDate.Text) Then
        Response% = MsgBox("Work Schedule for this From Date already exists for some employee(s). Do you want to skip those employees?" & vbCrLf & vbCrLf & "Click 'Yes' to skip; click 'No' to abort the Mass Update process.", vbQuestion + vbYesNo, "Work Schedule already exists")
        If Response% = IDNO Then
            Exit Function
        End If
    End If
    
    'Check if later Effective Date schedule already exists. If it does then do not allow to save this.
    If LaterScheduleExists(rsSchedule("SD_EMPNBR"), dlpEffectiveDate.Text) Then
        Response% = MsgBox("A more recent Work Schedule already exists for some employees(s). Do you want to skip those employees?" & vbCrLf & vbCrLf & "Click 'Yes' to skip; click 'No' to abort the Mass Update process.", vbQuestion + vbYesNo, "Recent Work Schedule already exists")
        If Response% = IDNO Then
            Exit Function
        End If
    'ElseIf LaterScheduleExistsFromToDt(rsSchedule("SD_EMPNBR"), dlpToDate.Text) Then
    '    Response% = MsgBox("A more recent Work Schedule already exists for some employees(s). Do you want to skip those employees?" & vbCrLf & vbCrLf & "Click 'Yes' to skip; click 'No' to abort the Mass Update process.", vbInformation + vbYesNo, "Recent Work Schedule already exists")
    '    If Response% = IDNO Then
    '        Exit Function
    '    End If
    End If
    
    'Check if any Unapproved Requests (ESS) exists. If it does then do not allow to save this.
    If UnapprovedRequestExists(rsSchedule("SD_EMPNBR")) Then
        Response% = MsgBox("There are Vacation or Time Off Requests pending Approval for some employee(s). Transaction cannot be processed. Do you want a list of Vacation Requests Pending Approval?" & vbCrLf & vbCrLf & "Click 'Yes' to view the list of employees; click 'No' to abort the Mass Update process.", vbQuestion + vbYesNo, "Vacation/Time Off Requests pending Approval")
        
        If Response% = IDYES Then
            'Run the report
            Call ShowReport_UnapprovedRequests
            Screen.MousePointer = DEFAULT
            Exit Function
        End If
        Screen.MousePointer = DEFAULT
        Exit Function
    End If
    
    rsSchedule.MoveNext
Loop
rsSchedule.Close
Set rsSchedule = Nothing

'More checking before adding WS
'Ticket #24485 - To Date cannot be greater than Effective Date. This is to freeze any new WS entry when another
'rotation weeks will be coming into play in near future.
If IsDate(gsWS_ROTATIONWEEKSEFFDATE) Then
    If CVDate(dlpEffectiveDate.Text) < CVDate(gsWS_ROTATIONWEEKSEFFDATE) Then
        MsgBox "From Date cannot be prior to 'Effective Date' of the new '# of Work Schedule Rotation Weeks." & vbCrLf & vbCrLf & "- Number of Work Schedule Rotation Weeks = " & gsWS_ROTATIONWEEKS & vbCrLf & "- Effective Date = " & gsWS_ROTATIONWEEKSEFFDATE
        dlpEffectiveDate.SetFocus
        Screen.MousePointer = DEFAULT
        Exit Function
    End If
End If

'Adding a new WS - Change date should be same as From Date.
dlpChangeDate.Text = dlpEffectiveDate.Text


'Add WS
XUpdCount = 0
SQLQ = "SELECT ED_COMPNO,ED_EMPNBR FROM HREMP " & WSQLQ
rsTA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
Do While Not rsTA.EOF
    'Skip employees already with this Work Schedule and employee who have later date Work Schedule
    SQLQ = "SELECT * FROM HR_SCHEDULER WHERE "
    SQLQ = SQLQ & " ((SD_EDATE <= " & Date_SQL(dlpEffectiveDate.Text)
    SQLQ = SQLQ & " AND SD_TDATE >= " & Date_SQL(dlpEffectiveDate.Text) & ")"
    SQLQ = SQLQ & " OR (SD_EDATE >= " & Date_SQL(dlpEffectiveDate.Text) & "))"
    
    SQLQ = SQLQ & " AND SD_EMPNBR = " & rsTA("ED_EMPNBR")
    rsSchedule.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsSchedule.EOF Then
        XUpdCount = XUpdCount + 1
        rsSchedule.AddNew
        rsSchedule("SD_COMPNO") = rsTA("ED_COMPNO")
        rsSchedule("SD_EMPNBR") = rsTA("ED_EMPNBR")
        rsSchedule("SD_EDATE") = dlpEffectiveDate.Text
        rsSchedule("SD_TDATE") = dlpToDate.Text
        
        'Ticket #24485 - # of WS Rotation Weeks
        rsSchedule("SD_ROTWKS") = gsWS_ROTATIONWEEKS
        
        'Ticket #24485 - # of WS Rotation Weeks
        If frWeek1.Enabled Then
            If IsNumeric(medHours(0)) Then
                rsSchedule("SD_SUN_HRS") = medHours(0)
            End If
            If IsNumeric(medHours(1)) Then
                rsSchedule("SD_MON_HRS") = medHours(1)
            End If
            If IsNumeric(medHours(2)) Then
                rsSchedule("SD_TUE_HRS") = medHours(2)
            End If
            If IsNumeric(medHours(3)) Then
                rsSchedule("SD_WED_HRS") = medHours(3)
            End If
            If IsNumeric(medHours(4)) Then
                rsSchedule("SD_THU_HRS") = medHours(4)
            End If
            If IsNumeric(medHours(5)) Then
                rsSchedule("SD_FRI_HRS") = medHours(5)
            End If
            If IsNumeric(medHours(6)) Then
                rsSchedule("SD_SAT_HRS") = medHours(6)
            End If
        End If
        
        If frWeek2.Enabled Then
            If IsNumeric(medHours2(0)) Then
                rsSchedule("SD_SUN_HRS2") = medHours2(0)
            End If
            If IsNumeric(medHours2(1)) Then
                rsSchedule("SD_MON_HRS2") = medHours2(1)
            End If
            If IsNumeric(medHours2(2)) Then
                rsSchedule("SD_TUE_HRS2") = medHours2(2)
            End If
            If IsNumeric(medHours2(3)) Then
                rsSchedule("SD_WED_HRS2") = medHours2(3)
            End If
            If IsNumeric(medHours2(4)) Then
                rsSchedule("SD_THU_HRS2") = medHours2(4)
            End If
            If IsNumeric(medHours2(5)) Then
                rsSchedule("SD_FRI_HRS2") = medHours2(5)
            End If
            If IsNumeric(medHours2(6)) Then
                rsSchedule("SD_SAT_HRS2") = medHours2(6)
            End If
        End If
        
        If frWeek3.Enabled Then
            If IsNumeric(medHours3(0)) Then
                rsSchedule("SD_SUN_HRS3") = medHours3(0)
            End If
            If IsNumeric(medHours3(1)) Then
                rsSchedule("SD_MON_HRS3") = medHours3(1)
            End If
            If IsNumeric(medHours3(2)) Then
                rsSchedule("SD_TUE_HRS3") = medHours3(2)
            End If
            If IsNumeric(medHours3(3)) Then
                rsSchedule("SD_WED_HRS3") = medHours3(3)
            End If
            If IsNumeric(medHours3(4)) Then
                rsSchedule("SD_THU_HRS3") = medHours3(4)
            End If
            If IsNumeric(medHours3(5)) Then
                rsSchedule("SD_FRI_HRS3") = medHours3(5)
            End If
            If IsNumeric(medHours3(6)) Then
                rsSchedule("SD_SAT_HRS3") = medHours3(6)
            End If
        End If
        
        If frWeek4.Enabled Then
            If IsNumeric(medHours4(0)) Then
                rsSchedule("SD_SUN_HRS4") = medHours4(0)
            End If
            If IsNumeric(medHours4(1)) Then
                rsSchedule("SD_MON_HRS4") = medHours4(1)
            End If
            If IsNumeric(medHours4(2)) Then
                rsSchedule("SD_TUE_HRS4") = medHours4(2)
            End If
            If IsNumeric(medHours4(3)) Then
                rsSchedule("SD_WED_HRS4") = medHours4(3)
            End If
            If IsNumeric(medHours4(4)) Then
                rsSchedule("SD_THU_HRS4") = medHours4(4)
            End If
            If IsNumeric(medHours4(5)) Then
                rsSchedule("SD_FRI_HRS4") = medHours4(5)
            End If
            If IsNumeric(medHours4(6)) Then
                rsSchedule("SD_SAT_HRS4") = medHours4(6)
            End If
        End If
        
        rsSchedule("SD_COMMENTS") = memComments.Text
        rsSchedule("SD_LDATE") = Date
        rsSchedule("SD_LTIME") = Time$
        rsSchedule("SD_LUSER") = glbUserID
        rsSchedule.Update
        
        'Ticket #24485 - # of WS Rotation Weeks
        'Add in Work Schedule Details as well
        Call Add_Workschedule_Detail(rsTA("ED_EMPNBR"))
    Else
        'Skip adding schedule for this employee
    End If
    rsSchedule.Close
    Set rsSchedule = Nothing
    
    rsTA.MoveNext
Loop
rsTA.Close
Set rsTA = Nothing

modInsRecs = True

Exit Function

cmdInsErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
If glbErrNum& = -2147467259 Then
    MsgBox "The changes were not successful because it would create duplicate values."
    Exit Function
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Mass Add", "HR_SCHEDULER", "Insert")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        RollBack
        Resume Next
    Else
        Unload Me
    End If
End If
End Function

Private Function WSQLQ() As String
Dim countr As Integer

WSQLQ = WSQLQ & " WHERE " & glbSeleDeptUn

Call glbCri_DeptUN(clpDept.Text)

If Len(clpDept.Text) > 0 Then WSQLQ = WSQLQ & " AND ED_DEPTNO = '" & clpDept.Text & "'"

If Len(clpDiv.Text) > 0 Then WSQLQ = WSQLQ & " AND ED_DIV = '" & clpDiv.Text & "' "
If Len(clpDiv.Text) > 0 Then glbstrSelCri = glbstrSelCri & " AND " & "({HREMP.ED_DIV} in ['" & Replace(clpDiv.Text, ",", "','") & "'])"

If Len(clpCode(0).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_LOC = '" & clpCode(0).Text & "' "
If Len(clpCode(0).Text) > 0 Then glbstrSelCri = glbstrSelCri & " AND " & "({HREMP.ED_LOC} in ['" & Replace(clpCode(0).Text, ",", "','") & "'])"

If Len(clpCode(1).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_ORG = '" & clpCode(1).Text & "' "
If Len(clpCode(1).Text) > 0 Then glbstrSelCri = glbstrSelCri & " AND " & "({HREMP.ED_ORG} in ['" & Replace(clpCode(1).Text, ",", "','") & "'])"

If Len(clpCode(2).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_EMP = '" & clpCode(2).Text & "' "
If Len(clpCode(2).Text) > 0 Then glbstrSelCri = glbstrSelCri & " AND " & "({HREMP.ED_EMP} in ['" & Replace(clpCode(2).Text, ",", "','") & "'])"

If Len(clpCode(3).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_REGION = '" & clpCode(3).Text & "' "
If Len(clpCode(3).Text) > 0 Then glbstrSelCri = glbstrSelCri & " AND " & "({HREMP.ED_REGION} in ['" & Replace(clpCode(3).Text, ",", "','") & "'])"

If Len(clpCode(4).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_ADMINBY = '" & clpCode(4).Text & "' "
If Len(clpCode(4).Text) > 0 Then glbstrSelCri = glbstrSelCri & " AND " & "({HREMP.ED_ADMINBY} in ['" & Replace(clpCode(4).Text, ",", "','") & "'])"

If Len(clpCode(5).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_SECTION = '" & clpCode(5).Text & "' "
If Len(clpCode(5).Text) > 0 Then glbstrSelCri = glbstrSelCri & " AND " & "({HREMP.ED_SECTION} in ['" & Replace(clpCode(5).Text, ",", "','") & "'])"

If Len(clpPT) > 0 Then WSQLQ = WSQLQ & " AND ED_PT = '" & clpPT.Text & "' "
If Len(clpPT) > 0 Then glbstrSelCri = glbstrSelCri & " AND " & "{HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"

If Len(elpEEID.Text) > 0 Then WSQLQ = WSQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
If Len(elpEEID.Text) > 0 Then glbstrSelCri = glbstrSelCri & " AND " & "{HREMP.ED_EMPNBR} IN [ " & getEmpnbr(elpEEID.Text) & "] "
'Ticket #22221
WSQLQ = WSQLQ & " AND ED_EMP IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'EDEM' AND TB_WORKSCHED = 1)"

End Function

Private Function modDelRecs()
Dim BD As Integer
Dim SQLQ As String, countr As Integer
Dim Dat1 As Variant, Dat2 As Variant
Dim iOneWhere As Integer, NxtSQL As String, strReas$
Dim rsScheduler As New ADODB.Recordset
Dim Response%

modDelRecs = False
On Error GoTo cmdDel_Err


SQLQ = "SELECT * FROM HR_SCHEDULER WHERE "
SQLQ = SQLQ & " SD_EDATE = " & Date_SQL(dlpEffectiveDate.Text)
SQLQ = SQLQ & " AND SD_TDATE = " & Date_SQL(dlpToDate.Text)
SQLQ = SQLQ & " AND SD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ & ")"
rsScheduler.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic ', adOpenStatic
XUpdCount = rsScheduler.RecordCount
Do While Not rsScheduler.EOF
    
    'Check if any Unapproved Requests (ESS) exists. If it does then do not allow to delete this.
    If UnapprovedRequestExistsFromToDt(rsScheduler("SD_EMPNBR"), dlpEffectiveDate.Text) Then
        Response% = MsgBox("There are Vacation or Time Off Requests pending Approval for some employee(s). Transaction cannot be processed. Do you want a list of Vacation Requests Pending Approval?" & vbCrLf & vbCrLf & "Click 'Yes' to view the list of employees; click 'No' to abort the Mass Update process.", vbQuestion + vbYesNo, "Vacation/Time Off Requests pending Approval")
        If Response% = IDYES Then
            'Run the report
            Call ShowReport_UnapprovedRequests
            Screen.MousePointer = DEFAULT
            Exit Function
        End If
        Screen.MousePointer = DEFAULT
        Exit Function
    ElseIf UnapprovedRequestExistsFromToDt(rsScheduler("SD_EMPNBR"), dlpToDate.Text) Then
        Response% = MsgBox("There are Vacation or Time Off Requests pending Approval for some employee(s). Transaction cannot be processed. Do you want a list of Vacation Requests Pending Approval?" & vbCrLf & vbCrLf & "Click 'Yes' to view the list of employees; click 'No' to abort the Mass Update process.", vbQuestion + vbYesNo, "Vacation/Time Off Requests pending Approval")
        If Response% = IDYES Then
            'Run the report
            Call ShowReport_UnapprovedRequests
            Screen.MousePointer = DEFAULT
            Exit Function
        End If
        Screen.MousePointer = DEFAULT
        Exit Function
    ElseIf UnapprovedRequestExistsFromToDt(rsScheduler("SD_EMPNBR"), dlpEffectiveDate.Text, dlpToDate.Text) Then
        Response% = MsgBox("There are Vacation or Time Off Requests pending Approval for some employee(s). Transaction cannot be processed. Do you want a list of Vacation Requests Pending Approval?" & vbCrLf & vbCrLf & "Click 'Yes' to view the list of employees; click 'No' to abort the Mass Update process.", vbQuestion + vbYesNo, "Vacation/Time Off Requests pending Approval")
        If Response% = IDYES Then
            'Run the report
            Call ShowReport_UnapprovedRequests
            Screen.MousePointer = DEFAULT
            Exit Function
        End If
        Screen.MousePointer = DEFAULT
        Exit Function
    End If
    
    rsScheduler.MoveNext
Loop
rsScheduler.Close
Set rsScheduler = Nothing


Screen.MousePointer = HOURGLASS

SQLQ = "SELECT * FROM HR_SCHEDULER WHERE "
SQLQ = SQLQ & " SD_EDATE = " & Date_SQL(dlpEffectiveDate.Text)
SQLQ = SQLQ & " AND SD_TDATE = " & Date_SQL(dlpToDate.Text)
SQLQ = SQLQ & " AND SD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ & ")"
rsScheduler.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic ', adOpenStatic
XUpdCount = rsScheduler.RecordCount

While Not rsScheduler.EOF
    rsScheduler.Delete
    
    'Ticket #24485 - Work Schedule Rotation Weeks - Delete the WS Details since the template is deleted now
    Call Delete_Workschedule_Detail(rsScheduler("SD_EMPNBR"))
    
    rsScheduler.MoveNext
Wend
rsScheduler.Close
Set rsScheduler = Nothing

modDelRecs = True

Screen.MousePointer = DEFAULT

Exit Function

cmdDel_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_SCHEDULER", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frmUScheduler = Nothing 'carmen apr 2000
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
    'UpdateRight = gSec_Upd_Earnings
    UpdateRight = GetMassUpdateSecurities("Work_Schedule_MassUpdate", glbUserID)
End Property

Public Property Get Addable() As Boolean
    Addable = True
End Property

Public Property Get Updateble() As Boolean
    Updateble = True
End Property

Public Property Get Deleteble() As Boolean
    Deleteble = True
End Property

Public Property Get Printable() As Boolean
    Printable = False
End Property

Private Sub medHours_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medHrsDay_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medHrsDay_LostFocus()
    Dim X As Integer
    If IsNumeric(medHrsDay) Then
        'Ticket #24485 - # of WS Rotation Weeks
        If frWeek1.Enabled Then
            For X = 0 To 6
                If Not IsNumeric(medHours(X)) Then
                    medHours(X) = medHrsDay
                End If
            Next
        End If
        If frWeek2.Enabled Then
            For X = 0 To 6
                If Not IsNumeric(medHours2(X)) Then
                    medHours2(X) = medHrsDay
                End If
            Next
        End If
        If frWeek3.Enabled Then
            For X = 0 To 6
                If Not IsNumeric(medHours3(X)) Then
                    medHours3(X) = medHrsDay
                End If
            Next
        End If
        If frWeek4.Enabled Then
            For X = 0 To 6
                If Not IsNumeric(medHours4(X)) Then
                    medHours4(X) = medHrsDay
                End If
            Next
        End If
    End If
End Sub

Private Sub memComments_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Function ShowReport_UnapprovedRequests(Optional xChangeDate)
    Dim dtYYY%, dtMM%, dtDD%, X%
    Dim TempCri As String
    Dim TempCri1 As String
    
'    Dim rsVacReq As New ADODB.Recordset
'    Dim SQLQ As String
'
'    'Create a query to give a list of unapproved requests and with date range greater than 1 day
'    SQLQ = "SELECT * FROM HR_VACTIMEOFF_REQ WHERE "
'    SQLQ = SQLQ & " AND VT_DELFLAG=0"
'    SQLQ = SQLQ & " AND VT_FROM <> VT_TO"   'date range has more than 1 day
'    SQLQ = SQLQ & " AND (VT_APPROVED IS NULL OR VT_APPROVED = '' OR VT_APPROVED IN ('RESUBMITTED','APP/FWD'))"
'    'SQLQ = SQLQ & " AND VT_VACTIME=1 "  'vacation time only
'    rsVacReq.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If Not rsVacReq.EOF Then
'
'    End If
'    rsVacReq.Close
'    Set rsVacReq = Nothing
    
    'If a list of unapproved requests are required then show the following fields:
    'Number, Name,From Date, To Date, Requested Date, Processed Date, Supervisor Name, Status, Hours Requested
    'VT_EMPNBR,VT_EMPNAME,VT_FROM,VT_TO,VT_REQDATE,VT_PROCDATE,VT_SUPERNAME,VT_APPROVED(Status),VT_HRS(Hours Requested)?
    
    
    Screen.MousePointer = HOURGLASS
    
    'report filename
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZUnApprvdLst.rpt"
    
    If Not IsMissing(xChangeDate) Then
        If Len(xChangeDate) > 0 Then
            TempCri = "(({HR_VACTIMEOFF_REQ.VT_FROM} "
            dtYYY% = Year(xChangeDate)
            dtMM% = month(xChangeDate)
            dtDD% = Day(xChangeDate)
            TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
            
            TempCri1 = "({HR_VACTIMEOFF_REQ.VT_TO} "
            dtYYY% = Year(xChangeDate)
            dtMM% = month(xChangeDate)
            dtDD% = Day(xChangeDate)
            TempCri1 = " OR " & TempCri1 & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & "))) "
        End If
    
        TempCri = TempCri & TempCri1
        glbstrSelCri = glbstrSelCri & " AND " & TempCri

    End If
    
    'set location for database tables
    Me.vbxCrystal.Connect = RptODBC_SQL
    
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
    
    'window title if appropriate
    Me.vbxCrystal.WindowTitle = "Vacation/Time Off Requests Pending Approval"
    
    Me.vbxCrystal.Destination = 0
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    
    Screen.MousePointer = DEFAULT
    
End Function

Private Function getRecordCount_Modify()
    Dim SQLQ As String
    Dim rsSchedule As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Modify = 0
    recCount = 0

    SQLQ = "SELECT COUNT(SD_EMPNBR) AS TOT_REC FROM HR_SCHEDULER WHERE "
    SQLQ = SQLQ & " SD_EDATE = " & Date_SQL(dlpEffectiveDate.Text)
    SQLQ = SQLQ & " AND SD_TDATE = " & Date_SQL(dlpToDate.Text)
    SQLQ = SQLQ & " AND SD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ & ")"
    rsSchedule.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsSchedule.EOF Then
        recCount = rsSchedule("TOT_REC")
    Else
        recCount = 0
    End If
    rsSchedule.Close
    Set rsSchedule = Nothing
    
    getRecordCount_Modify = recCount

End Function

Private Function getRecordCount_Delete()
    Dim SQLQ As String
    Dim rsScheduler As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Delete = 0
    recCount = 0

    SQLQ = "SELECT COUNT(SD_EMPNBR) AS TOT_REC FROM HR_SCHEDULER WHERE "
    SQLQ = SQLQ & " SD_EDATE = " & Date_SQL(dlpEffectiveDate.Text)
    SQLQ = SQLQ & " AND SD_TDATE = " & Date_SQL(dlpToDate.Text)
    SQLQ = SQLQ & " AND SD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ & ")"
    rsScheduler.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic ', adOpenStatic
    If Not rsScheduler.EOF Then
        recCount = rsScheduler("TOT_REC")
    Else
        recCount = 0
    End If
    rsScheduler.Close
    Set rsScheduler = Nothing
    
    getRecordCount_Delete = recCount

End Function

Private Function getRecordCount_Add()
    Dim SQLQ As String
    Dim rsSchedule As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Add = 0
    recCount = 0

    SQLQ = "SELECT COUNT(ED_EMPNBR) AS TOT_REC FROM HREMP " & WSQLQ
    SQLQ = SQLQ & " AND (ED_EMPNBR NOT IN (SELECT SD_EMPNBR FROM HR_SCHEDULER WHERE "
    'SQLQ = SQLQ & " ((SD_EDATE >=" & Date_SQL(dlpEffectiveDate.Text)
    'SQLQ = SQLQ & " AND SD_TDATE <=" & Date_SQL(dlpEffectiveDate.Text) & ") OR "
    'SQLQ = SQLQ & " (SD_EDATE >=" & Date_SQL(dlpToDate.Text)
    'SQLQ = SQLQ & " AND SD_TDATE <=" & Date_SQL(dlpToDate.Text) & "))))"

    SQLQ = SQLQ & " ((SD_EDATE <= " & Date_SQL(dlpEffectiveDate.Text)
    SQLQ = SQLQ & " AND SD_TDATE >= " & Date_SQL(dlpEffectiveDate.Text) & ")"
    SQLQ = SQLQ & " OR (SD_EDATE >= " & Date_SQL(dlpEffectiveDate.Text) & "))))"
    
    
    rsSchedule.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsSchedule.EOF Then
        recCount = rsSchedule("TOT_REC")
    Else
        recCount = 0
    End If
    rsSchedule.Close
    Set rsSchedule = Nothing
    
    getRecordCount_Add = recCount

End Function

Private Function Add_Workschedule_Detail(xEmpNo)
    Dim rsSchDetail As New ADODB.Recordset
    Dim rsSchDetail2 As New ADODB.Recordset
    Dim SQLQ As String
    Dim xCurrDate As Date
    Dim xDay As Integer
    Dim xLstWSDay As Date
    Dim xWeekNo As Integer
    Dim xWeekDayCount As Integer
    Dim xLstWSWkDay As Integer
    Dim xLstWSDate As Date
    Dim xStartDay As String
    Dim WKSSQLQ As String
    Dim xWkDay As String
    
                
    SQLQ = "SELECT * FROM HR_SCHEDULER_DETAIL WHERE WS_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " AND WS_FDATE = " & Date_SQL(dlpEffectiveDate.Text)
    SQLQ = SQLQ & " AND WS_TDATE = " & Date_SQL(dlpToDate.Text)
    rsSchDetail.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsSchDetail.EOF Then
        'No Work Schedule detail found. Create details
        
        'Set the Start Date of the Period so for each date upto the To Date the Work Schedule can be prepared
        xCurrDate = CVDate(dlpEffectiveDate.Text) - 1
        
        'Get the day of the Week to start from
        xLstWSWkDay = 0
        xLstWSWkDay = Weekday(xCurrDate)
        
        'Compute the Day to start adding the WS details
        If xLstWSWkDay < 7 Then
            xStartDay = xLstWSWkDay
        Else
            xStartDay = 0
        End If
        
        
        Do While CVDate(xCurrDate) < CVDate(dlpToDate.Text)
            
            If frWeek1.Enabled Then
                For xDay = xStartDay To 6
                    'Add Work Schedule for the day in a week
                    rsSchDetail.AddNew
                    rsSchDetail("WS_EMPNBR") = xEmpNo
                    rsSchDetail("WS_FDATE") = dlpEffectiveDate.Text
                    rsSchDetail("WS_TDATE") = dlpToDate.Text
                    rsSchDetail("WS_CHGDATE") = dlpEffectiveDate.Text   'Change from Date will be same as From Date the first time
                
                    rsSchDetail("WS_WEEKNO") = 1
                    rsSchDetail("WS_DATE") = DateAdd("d", 1, xCurrDate)
                    rsSchDetail("WS_HRS") = medHours(xDay)
                
                    rsSchDetail("WS_LDATE") = Date
                    rsSchDetail("WS_LTIME") = Time$
                    rsSchDetail("WS_LUSER") = glbUserID
                    rsSchDetail.Update
                    
                    'ReSet the CurrDate
                    xCurrDate = CVDate(rsSchDetail("WS_DATE"))
                    
                    'Exit the Loop if xCurrDate exceeds the ToDate.
                    If CVDate(xCurrDate) > CVDate(dlpToDate.Text) Then Exit Do
                Next
                
                'Reset the start day for the next Week
                xStartDay = 0
            End If
            
            If frWeek2.Enabled Then
                For xDay = 0 To 6
                    'Add Work Schedule for the day in a week
                    rsSchDetail.AddNew
                    rsSchDetail("WS_EMPNBR") = xEmpNo
                    rsSchDetail("WS_FDATE") = dlpEffectiveDate.Text
                    rsSchDetail("WS_TDATE") = dlpToDate.Text
                    rsSchDetail("WS_CHGDATE") = dlpEffectiveDate.Text   'Change from Date will be same as From Date the first time
                
                    rsSchDetail("WS_WEEKNO") = 2
                    rsSchDetail("WS_DATE") = DateAdd("d", 1, xCurrDate)
                    rsSchDetail("WS_HRS") = medHours2(xDay)
                
                    rsSchDetail("WS_LDATE") = Date
                    rsSchDetail("WS_LTIME") = Time$
                    rsSchDetail("WS_LUSER") = glbUserID
                    rsSchDetail.Update
                    
                    'ReSet the CurrDate
                    xCurrDate = CVDate(rsSchDetail("WS_DATE"))
                
                    'Exit the Loop if xCurrDate exceeds the ToDate.
                    If CVDate(xCurrDate) > CVDate(dlpToDate.Text) Then Exit Do
                Next
                
                'Reset the start day for the next Week
                xStartDay = 0
            Else
                'If this week is not visible so the rest of the Weeks will not be visible too
                GoTo NextWorkSchSet
            End If
            
            If frWeek3.Enabled Then
                For xDay = 0 To 6
                    'Add Work Schedule for the day in a week
                    rsSchDetail.AddNew
                    rsSchDetail("WS_EMPNBR") = xEmpNo
                    rsSchDetail("WS_FDATE") = dlpEffectiveDate.Text
                    rsSchDetail("WS_TDATE") = dlpToDate.Text
                    rsSchDetail("WS_CHGDATE") = dlpEffectiveDate.Text   'Change from Date will be same as From Date the first time
                
                    rsSchDetail("WS_WEEKNO") = 3
                    rsSchDetail("WS_DATE") = DateAdd("d", 1, xCurrDate)
                    rsSchDetail("WS_HRS") = medHours3(xDay)
                
                    rsSchDetail("WS_LDATE") = Date
                    rsSchDetail("WS_LTIME") = Time$
                    rsSchDetail("WS_LUSER") = glbUserID
                    rsSchDetail.Update
                    
                    'ReSet the CurrDate
                    xCurrDate = CVDate(rsSchDetail("WS_DATE"))
                    
                    'Exit the Loop if xCurrDate exceeds the ToDate.
                    If CVDate(xCurrDate) > CVDate(dlpToDate.Text) Then Exit Do
                Next
                
                'Reset the start day for the next Week
                xStartDay = 0
            Else
                'If this week is not visible so the rest of the Weeks will not be visible too
                GoTo NextWorkSchSet
            End If
            
            If frWeek4.Enabled Then
                For xDay = 0 To 6
                    'Add Work Schedule for the day in a week
                    rsSchDetail.AddNew
                    rsSchDetail("WS_EMPNBR") = xEmpNo
                    rsSchDetail("WS_FDATE") = dlpEffectiveDate.Text
                    rsSchDetail("WS_TDATE") = dlpToDate.Text
                    rsSchDetail("WS_CHGDATE") = dlpEffectiveDate.Text   'Change from Date will be same as From Date the first time
                
                    rsSchDetail("WS_WEEKNO") = 4
                    rsSchDetail("WS_DATE") = DateAdd("d", 1, xCurrDate)
                    rsSchDetail("WS_HRS") = medHours4(xDay)
                
                    rsSchDetail("WS_LDATE") = Date
                    rsSchDetail("WS_LTIME") = Time$
                    rsSchDetail("WS_LUSER") = glbUserID
                    rsSchDetail.Update
                    
                    'ReSet the CurrDate
                    xCurrDate = CVDate(rsSchDetail("WS_DATE"))
                    
                    'Exit the Loop if xCurrDate exceeds the ToDate.
                    If CVDate(xCurrDate) > CVDate(dlpToDate.Text) Then Exit Do
                Next
                
                'Reset the start day for the next Week
                xStartDay = 0
            End If
            
NextWorkSchSet:
        Loop
    End If
    rsSchDetail.Close
    Set rsSchDetail = Nothing
    
End Function

Private Function Update_WorkSchedule_Detail(xEmpNo)
    Dim rsSchDetail As New ADODB.Recordset
    Dim rsSchDetail2 As New ADODB.Recordset
    Dim SQLQ As String
    Dim xCurrDate As Date
    Dim xDay As Integer
    Dim xLstWSDay As Date
    Dim xWeekNo As Integer
    Dim xWeekDayCount As Integer
    Dim xLstWSWkDay As Integer
    Dim xLstWSDate As Date
    Dim xStartDay As String
    Dim WKSSQLQ As String
    Dim xWkDay As String


    SQLQ = "SELECT * FROM HR_SCHEDULER_DETAIL WHERE WS_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " AND WS_FDATE = " & Date_SQL(dlpEffectiveDate.Text)
    SQLQ = SQLQ & " AND WS_TDATE = " & Date_SQL(dlpToDate.Text)
    rsSchDetail.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsSchDetail.EOF Then

        'Check if Hours has changed
        If Has_Hours_Changed(xEmpNo) Then
        
            'Hours has changed updated WS Details
            If xWeek1Changed Or xWeek2Changed Or xWeek3Changed Or xWeek4Changed Then
                WKSSQLQ = ""
                
                'Retrieve the WS Detail with dates from the Changed Date to change the hours for
                SQLQ = "SELECT * FROM HR_SCHEDULER_DETAIL WHERE WS_EMPNBR = " & xEmpNo
                SQLQ = SQLQ & " AND WS_FDATE = " & Date_SQL(dlpEffectiveDate.Text)
                SQLQ = SQLQ & " AND WS_TDATE = " & Date_SQL(dlpToDate.Text)
                SQLQ = SQLQ & " AND WS_DATE >= " & Date_SQL(dlpChangeDate.Text)
                If xWeek1Changed Then
                    WKSSQLQ = WKSSQLQ & " AND (WS_WEEKNO = 1"
                End If
                If xWeek2Changed Then
                    If Len(WKSSQLQ) > 0 Then
                        WKSSQLQ = WKSSQLQ & " OR WS_WEEKNO = 2"
                    Else
                        WKSSQLQ = WKSSQLQ & " AND (WS_WEEKNO = 2"
                    End If
                End If
                If xWeek3Changed Then
                    If Len(WKSSQLQ) > 0 Then
                        WKSSQLQ = WKSSQLQ & " OR WS_WEEKNO = 3"
                    Else
                        WKSSQLQ = WKSSQLQ & " AND (WS_WEEKNO = 3"
                    End If
                End If
                If xWeek4Changed Then
                    If Len(WKSSQLQ) > 0 Then
                        WKSSQLQ = WKSSQLQ & " OR WS_WEEKNO = 4"
                    Else
                        WKSSQLQ = WKSSQLQ & " AND (WS_WEEKNO = 4"
                    End If
                End If
                If Len(WKSSQLQ) > 0 Then
                    WKSSQLQ = WKSSQLQ & ")"
                    SQLQ = SQLQ & WKSSQLQ
                End If
                SQLQ = SQLQ & " ORDER BY WS_WEEKNO, WS_DATE"
                rsSchDetail2.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsSchDetail2.EOF Then
                    rsSchDetail2.MoveFirst
                    Do While Not rsSchDetail2.EOF
                        'Get the Week No and Day of WS Date
                        xWeekNo = rsSchDetail2("WS_WEEKNO")
                        xWkDay = Weekday(rsSchDetail2("WS_DATE"))
                                            
                        'Convert the xWkDay retrieved to match the screen Week Day #
                        xWkDay = xWkDay - 1
                                            
                        'Get the Hours from the screen based on the current WS Detail record to compare
                        If xWeekNo = 1 Then
                            If medHours(xWkDay) <> rsSchDetail2("WS_HRS") Then
                                'Hours do not match, change the WS Details Hours and Changed Date.
                                rsSchDetail2("WS_CHGDATE") = dlpChangeDate.Text
                                rsSchDetail2("WS_HRS") = medHours(xWkDay)
                            
                                rsSchDetail2("WS_LDATE") = Date
                                rsSchDetail2("WS_LTIME") = Time$
                                rsSchDetail2("WS_LUSER") = glbUserID
                                rsSchDetail2.Update
                            End If
                        ElseIf xWeekNo = 2 Then
                            If medHours2(xWkDay) <> rsSchDetail2("WS_HRS") Then
                                'Hours do not match, change the WS Details Hours and Changed Date.
                                rsSchDetail2("WS_CHGDATE") = dlpChangeDate.Text
                                rsSchDetail2("WS_HRS") = medHours2(xWkDay)
                            
                                rsSchDetail2("WS_LDATE") = Date
                                rsSchDetail2("WS_LTIME") = Time$
                                rsSchDetail2("WS_LUSER") = glbUserID
                                rsSchDetail2.Update
                            End If
                        ElseIf xWeekNo = 3 Then
                            If medHours3(xWkDay) <> rsSchDetail2("WS_HRS") Then
                                'Hours do not match, change the WS Details Hours and Changed Date.
                                rsSchDetail2("WS_CHGDATE") = dlpChangeDate.Text
                                rsSchDetail2("WS_HRS") = medHours3(xWkDay)
                            
                                rsSchDetail2("WS_LDATE") = Date
                                rsSchDetail2("WS_LTIME") = Time$
                                rsSchDetail2("WS_LUSER") = glbUserID
                                rsSchDetail2.Update
                            End If
                        ElseIf xWeekNo = 4 Then
                            If medHours4(xWkDay) <> rsSchDetail2("WS_HRS") Then
                                'Hours do not match, change the WS Details Hours and Changed Date.
                                rsSchDetail2("WS_CHGDATE") = dlpChangeDate.Text
                                rsSchDetail2("WS_HRS") = medHours4(xWkDay)
                            
                                rsSchDetail2("WS_LDATE") = Date
                                rsSchDetail2("WS_LTIME") = Time$
                                rsSchDetail2("WS_LUSER") = glbUserID
                                rsSchDetail2.Update
                            End If
                        End If
                        
                        rsSchDetail2.MoveNext
                    Loop
                End If
                rsSchDetail2.Close
                Set rsSchDetail2 = Nothing
            End If
        End If
    End If
    rsSchDetail.Close
    Set rsSchDetail = Nothing
End Function

Private Function Has_Hours_Changed(xEmpNo) As Boolean
    Dim rsSch As New ADODB.Recordset
    Dim SQLQ As String
    Dim xDay As Integer
    
    'Initialise
    xWeek1Changed = False
    xWeek2Changed = False
    xWeek3Changed = False
    xWeek4Changed = False
    
    Has_Hours_Changed = False
    
    'Check if the hours has changed in the WS Master table. Check with old To Date in case the To Date has changed
    SQLQ = "SELECT * FROM HR_SCHEDULER WHERE SD_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " AND SD_EDATE = " & Date_SQL(dlpEffectiveDate.Text)
    SQLQ = SQLQ & " AND SD_TDATE = " & Date_SQL(dlpToDate.Text)
    rsSch.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsSch.EOF Then
        'Compare Week 1
        For xDay = 0 To 6
            If rsSch("SD_SUN_HRS") <> medHours(xDay) Then
                xWeek1Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_MON_HRS") <> medHours(xDay) Then
                xWeek1Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_TUE_HRS") <> medHours(xDay) Then
                xWeek1Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_WED_HRS") <> medHours(xDay) Then
                xWeek1Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_THU_HRS") <> medHours(xDay) Then
                xWeek1Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_FRI_HRS") <> medHours(xDay) Then
                xWeek1Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_SAT_HRS") <> medHours(xDay) Then
                xWeek1Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
        Next
        
        'Compare Week 2
        For xDay = 0 To 6
            If rsSch("SD_SUN_HRS2") = medHours2(xDay) Then
                xWeek2Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_MON_HRS2") <> medHours2(xDay) Then
                xWeek2Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_TUE_HRS2") <> medHours2(xDay) Then
                xWeek2Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_WED_HRS2") <> medHours2(xDay) Then
                xWeek2Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_THU_HRS2") <> medHours2(xDay) Then
                xWeek2Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_FRI_HRS2") <> medHours2(xDay) Then
                xWeek2Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_SAT_HRS2") <> medHours2(xDay) Then
                xWeek2Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
        Next
        
        'Compare Week 3
        For xDay = 0 To 6
            If rsSch("SD_SUN_HRS3") = medHours3(xDay) Then
                xWeek3Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_MON_HRS3") <> medHours3(xDay) Then
                xWeek3Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_TUE_HRS3") <> medHours3(xDay) Then
                xWeek3Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_WED_HRS3") <> medHours3(xDay) Then
                xWeek3Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_THU_HRS3") <> medHours3(xDay) Then
                xWeek3Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_FRI_HRS3") <> medHours3(xDay) Then
                xWeek3Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_SAT_HRS3") <> medHours3(xDay) Then
                xWeek3Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
        Next
        
        'Compare Week 4
        For xDay = 0 To 6
            If rsSch("SD_SUN_HRS4") = medHours4(xDay) Then
                xWeek4Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_MON_HRS4") <> medHours4(xDay) Then
                xWeek4Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_TUE_HRS4") <> medHours4(xDay) Then
                xWeek4Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_WED_HRS4") <> medHours4(xDay) Then
                xWeek4Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_THU_HRS4") <> medHours4(xDay) Then
                xWeek4Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_FRI_HRS4") <> medHours4(xDay) Then
                xWeek4Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_SAT_HRS4") <> medHours4(xDay) Then
                xWeek4Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
        Next
    End If
    rsSch.Close
    Set rsSch = Nothing
    
End Function

Private Sub Enable_Disable_RotationWeeks(xRotationWks)

    lblMsg.Visible = False
    If Not IsNull(xRotationWks) Then
        Select Case xRotationWks
            Case 1
                frWeek1.Enabled = True
                frWeek2.Enabled = False
                frWeek3.Enabled = False
                frWeek4.Enabled = False
            
            Case 2
                frWeek1.Enabled = True
                frWeek2.Enabled = True
                frWeek3.Enabled = False
                frWeek4.Enabled = False
            
            Case 3
                frWeek1.Enabled = True
                frWeek2.Enabled = True
                frWeek3.Enabled = True
                frWeek4.Enabled = False
            
            Case 4
                frWeek1.Enabled = True
                frWeek2.Enabled = True
                frWeek3.Enabled = True
                frWeek4.Enabled = True
            
            Case Else
                frWeek1.Enabled = False
                frWeek2.Enabled = False
                frWeek3.Enabled = False
                frWeek4.Enabled = False
                lblMsg.Visible = True
        End Select
    Else
        frWeek1.Enabled = False
        frWeek2.Enabled = False
        frWeek3.Enabled = False
        frWeek4.Enabled = False
        lblMsg.Visible = True
    End If
End Sub

Private Function Delete_Workschedule_Detail(xEmpNo)
    Dim SQLQ As String
    
    SQLQ = "DELETE FROM HR_SCHEDULER_DETAIL WHERE WS_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " AND WS_FDATE = " & Date_SQL(dlpEffectiveDate.Text)
    SQLQ = SQLQ & " AND WS_TDATE = " & Date_SQL(dlpToDate.Text)
    gdbAdoIhr001.Execute SQLQ
End Function
