VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEScheduler 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Work Hours Schedule"
   ClientHeight    =   10890
   ClientLeft      =   -150
   ClientTop       =   765
   ClientWidth     =   11670
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
   ScaleHeight     =   10890
   ScaleWidth      =   11670
   WindowState     =   2  'Maximized
   Begin VB.Frame frWeek4 
      Caption         =   "Week 4"
      Enabled         =   0   'False
      Height          =   3015
      Left            =   8160
      TabIndex        =   58
      Top             =   4080
      Width           =   2535
      Begin MSMask.MaskEdBox medHours4 
         DataField       =   "SD_SUN_HRS4"
         Height          =   285
         Index           =   0
         Left            =   1275
         TabIndex        =   23
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
         TabIndex        =   24
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
         TabIndex        =   25
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
         TabIndex        =   26
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   29
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
         TabIndex        =   79
         Top             =   375
         Width           =   540
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
         TabIndex        =   78
         Top             =   735
         Width           =   570
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
         TabIndex        =   77
         Top             =   1095
         Width           =   615
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
         TabIndex        =   76
         Top             =   1455
         Width           =   855
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
         TabIndex        =   75
         Top             =   2175
         Width           =   420
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
         TabIndex        =   74
         Top             =   1815
         Width           =   660
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
         TabIndex        =   73
         Top             =   2535
         Width           =   630
      End
   End
   Begin VB.Frame frWeek3 
      Caption         =   "Week 3"
      Enabled         =   0   'False
      Height          =   3015
      Left            =   5520
      TabIndex        =   57
      Top             =   4080
      Width           =   2535
      Begin MSMask.MaskEdBox medHours3 
         DataField       =   "SD_SUN_HRS3"
         Height          =   285
         Index           =   0
         Left            =   1275
         TabIndex        =   16
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
         TabIndex        =   17
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
         TabIndex        =   18
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
         TabIndex        =   19
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   22
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
         TabIndex        =   72
         Top             =   375
         Width           =   540
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
         TabIndex        =   71
         Top             =   735
         Width           =   570
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
         TabIndex        =   69
         Top             =   1455
         Width           =   855
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
         TabIndex        =   68
         Top             =   2175
         Width           =   420
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
         TabIndex        =   67
         Top             =   1815
         Width           =   660
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
         TabIndex        =   66
         Top             =   2535
         Width           =   630
      End
   End
   Begin VB.Frame frWeek2 
      Caption         =   "Week 2"
      Enabled         =   0   'False
      Height          =   3015
      Left            =   2880
      TabIndex        =   56
      Top             =   4080
      Width           =   2535
      Begin MSMask.MaskEdBox medHours2 
         DataField       =   "SD_SUN_HRS2"
         Height          =   285
         Index           =   0
         Left            =   1275
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   15
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
         TabIndex        =   65
         Top             =   375
         Width           =   540
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
         TabIndex        =   64
         Top             =   735
         Width           =   570
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
         TabIndex        =   63
         Top             =   1095
         Width           =   615
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
         TabIndex        =   62
         Top             =   1455
         Width           =   855
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
         TabIndex        =   61
         Top             =   2175
         Width           =   420
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
         TabIndex        =   60
         Top             =   1815
         Width           =   660
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
         TabIndex        =   59
         Top             =   2535
         Width           =   630
      End
   End
   Begin VB.Frame frWeek1 
      Caption         =   "Week 1"
      Enabled         =   0   'False
      Height          =   3015
      Left            =   240
      TabIndex        =   48
      Top             =   4080
      Width           =   2535
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "SD_MON_HRS"
         Height          =   285
         Index           =   1
         Left            =   1275
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   8
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
         TabIndex        =   2
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
         TabIndex        =   55
         Top             =   375
         Width           =   540
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
         TabIndex        =   54
         Top             =   735
         Width           =   570
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
         TabIndex        =   53
         Top             =   1095
         Width           =   615
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
         TabIndex        =   52
         Top             =   1455
         Width           =   855
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
         TabIndex        =   51
         Top             =   2175
         Width           =   420
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
         TabIndex        =   50
         Top             =   1815
         Width           =   660
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
         TabIndex        =   49
         Top             =   2535
         Width           =   630
      End
   End
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      DataField       =   "SD_COMMENTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   1755
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   30
      Tag             =   "00-Comments"
      Top             =   7350
      Width           =   8775
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7080
      Top             =   9120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
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
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   40
      Top             =   10230
      Width           =   11670
      _Version        =   65536
      _ExtentX        =   20585
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
      Begin VB.CommandButton cmdNoCondRebuild 
         Appearance      =   0  'Flat
         Caption         =   "Rebuild Details with No Checking"
         Height          =   375
         Left            =   8160
         TabIndex        =   84
         Tag             =   "Rebuild Work Schedule Details with no checking"
         Top             =   0
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.CommandButton cmdRebuild 
         Appearance      =   0  'Flat
         Caption         =   "Rebuild Details"
         Height          =   375
         Left            =   6000
         TabIndex        =   47
         Tag             =   "Rebuild Work Schedule Details"
         Top             =   0
         Width           =   2000
      End
      Begin VB.CommandButton cmdViewRpt 
         Appearance      =   0  'Flat
         Caption         =   "Unapproved or Rejected Vacation/Time Off Requests Report"
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Tag             =   "View/Print Unapproved Vacation/Time Off Requests"
         Top             =   0
         Width           =   5565
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   9090
         Top             =   90
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   1
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         GridSource      =   "vbxTrueGrid"
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin Crystal.CrystalReport vbxCrystal1 
         Left            =   10440
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
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "SD_LDATE"
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
      Index           =   0
      Left            =   3120
      MaxLength       =   25
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   9120
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "SD_LTIME"
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
      Index           =   1
      Left            =   4920
      MaxLength       =   25
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   9120
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "SD_LUSER"
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
      Index           =   2
      Left            =   6600
      MaxLength       =   25
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   9120
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   11670
      _Version        =   65536
      _ExtentX        =   20585
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BevelInner      =   2
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin VB.Label lblEEProdLine 
         AutoSize        =   -1  'True
         Caption         =   "Product Line"
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
         Left            =   6840
         TabIndex        =   43
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   155
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
         Left            =   1440
         TabIndex        =   36
         Top             =   132
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Left            =   3100
         TabIndex        =   35
         Top             =   132
         Width           =   1740
      End
   End
   Begin INFOHR_Controls.DateLookup dlpEffectiveDate 
      DataField       =   "SD_EDATE"
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Tag             =   "41-Work Schedule From Date"
      Top             =   3120
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   503
      MultiSelect     =   -1  'True
      TextBoxWidth    =   1215
      Enabled         =   0   'False
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feScheduler.frx":0000
      Height          =   2085
      Left            =   90
      OleObjectBlob   =   "feScheduler.frx":0014
      TabIndex        =   44
      Tag             =   "Listing of Associations"
      Top             =   630
      Width           =   10875
   End
   Begin INFOHR_Controls.DateLookup dlpChangeDate 
      Height          =   285
      Left            =   1440
      TabIndex        =   80
      Tag             =   "41-Work Schedule To Date"
      Top             =   3600
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   503
      MultiSelect     =   -1  'True
      TextBoxWidth    =   1215
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.DateLookup dlpToDate 
      DataField       =   "SD_TDATE"
      Height          =   285
      Left            =   5520
      TabIndex        =   1
      Tag             =   "41-Work Schedule To Date"
      Top             =   3120
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   503
      MultiSelect     =   -1  'True
      TextBoxWidth    =   1215
      Enabled         =   0   'False
   End
   Begin VB.Label lblRotWeeks 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "rotationweeks"
      DataField       =   "SD_ROTWKS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   9360
      TabIndex        =   83
      Top             =   3240
      Visible         =   0   'False
      Width           =   990
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
      TabIndex        =   82
      Top             =   8640
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
      Left            =   240
      TabIndex        =   81
      Top             =   3645
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
      Left            =   4560
      TabIndex        =   45
      Top             =   3135
      Width           =   705
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   42
      Top             =   3135
      Width           =   885
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   41
      Top             =   7350
      Width           =   735
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "SD_EMPNBR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2160
      TabIndex        =   38
      Top             =   9120
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "SD_COMPNO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   480
      TabIndex        =   39
      Top             =   9120
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fGLBNew As Boolean
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim RSDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim Ctrl As Control 'Sam add July 2002 * Remove ADO
Dim OToDate As Date
Dim xWeek1Changed As Boolean
Dim xWeek2Changed As Boolean
Dim xWeek3Changed As Boolean
Dim xWeek4Changed As Boolean
Dim xWK1HrsChanged As String
Dim xWK2HrsChanged As String
Dim xWK3HrsChanged As String
Dim xWK4HrsChanged As String


Private Function chkEScheduler()
Dim oCode As String, OCodeD As String
Dim x As Integer
Dim resp As Integer

chkEScheduler = False

On Error GoTo chkEScheduler_Err

If Len(dlpEffectiveDate.Text) < 1 Then
    MsgBox "From Date is required."
    dlpEffectiveDate.SetFocus
    Exit Function
End If

If Not IsDate(dlpEffectiveDate.Text) Then
    MsgBox "From Date is not a valid date."
    dlpEffectiveDate.SetFocus
    Exit Function
End If

'Ticket #24485 - To Date cannot be greater than Effective Date. This is to freeze any new WS entry when another
'rotation weeks will be coming into play in near future.
If fGLBNew And IsDate(gsWS_ROTATIONWEEKSEFFDATE) Then
    If CVDate(dlpEffectiveDate.Text) < CVDate(gsWS_ROTATIONWEEKSEFFDATE) Then
        MsgBox "From Date cannot be prior to 'Effective Date' of the new '# of Work Schedule Rotation Weeks." & vbCrLf & vbCrLf & "- Number of Work Schedule Rotation Weeks = " & gsWS_ROTATIONWEEKS & vbCrLf & "- Effective Date = " & gsWS_ROTATIONWEEKSEFFDATE
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
    MsgBox "From Date cannot be greater than To Date"
    dlpEffectiveDate.SetFocus
    Exit Function
End If

'Ticket #24485 - To Date cannot be greater than Effective Date. This is to freeze any new WS entry when another
'rotation weeks will be coming into play in near future.
If Not fGLBNew And IsDate(gsWS_ROTATIONWEEKSEFFDATE) Then
    If CVDate(dlpEffectiveDate.Text) < CVDate(gsWS_ROTATIONWEEKSEFFDATE) Then  'Not already in the new WS Rotation Weeks
        If CVDate(dlpToDate.Text) >= CVDate(gsWS_ROTATIONWEEKSEFFDATE) Then
            MsgBox "To Date cannot be greater or equal to 'Effective Date' of the new '# of Work Schedule Rotation Weeks." & vbCrLf & vbCrLf & "- Number of Work Schedule Rotation Weeks = " & gsWS_ROTATIONWEEKS & vbCrLf & "- Effective Date = " & gsWS_ROTATIONWEEKSEFFDATE
            dlpToDate.SetFocus
            Exit Function
        End If
    End If
End If


'Ticket #24485 - Work Schedule Rotation Weeks
If Not fGLBNew Then
    If Len(dlpChangeDate.Text) < 1 Then
        MsgBox "Change Date is required."
        dlpChangeDate.SetFocus
        Exit Function
    End If
Else
    dlpChangeDate.Text = dlpEffectiveDate.Text
End If

If Not IsDate(dlpChangeDate.Text) Then
    MsgBox "Change Date is not a valid date."
    dlpChangeDate.SetFocus
    Exit Function
End If

If CVDate(dlpChangeDate.Text) < CVDate(dlpEffectiveDate.Text) Then
    MsgBox "Change Date cannot be prior to From Date"
    dlpChangeDate.SetFocus
    Exit Function
End If

If CVDate(dlpChangeDate.Text) > CVDate(dlpToDate.Text) Then
    MsgBox "Change Date cannot be greater than To Date"
    dlpChangeDate.SetFocus
    Exit Function
End If


'Ticket #24485 - Work Schedule Rotation Weeks
If frWeek1.Enabled Then
    For x = 0 To 6
        If Len(medHours(x)) = 0 Then medHours(x).Text = 0
        
        If Not IsNumeric(medHours(x)) Then
            MsgBox lblTitle(x + 2).Caption & " Hours is invalid"
            medHours(x).SetFocus
            Exit Function
        End If
        If Val(medHours(x)) > 99999.9999 Then
            MsgBox lblTitle(x + 2).Caption & " Hours is Invalid"
            medHours(x).SetFocus
            Exit Function
        End If
    Next x
End If

If frWeek2.Enabled Then
    For x = 0 To 6
        If Len(medHours2(x)) = 0 Then medHours2(x).Text = 0
        
        If Not IsNumeric(medHours2(x)) Then
            MsgBox lblTitle2(x).Caption & " Hours is invalid"
            medHours2(x).SetFocus
            Exit Function
        End If
        If Val(medHours2(x)) > 99999.9999 Then
            MsgBox lblTitle2(x).Caption & " Hours is Invalid"
            medHours2(x).SetFocus
            Exit Function
        End If
    Next x
End If

If frWeek3.Enabled Then
    For x = 0 To 6
        If Len(medHours3(x)) = 0 Then medHours3(x).Text = 0
        
        If Not IsNumeric(medHours3(x)) Then
            MsgBox lblTitle3(x).Caption & " Hours is invalid"
            medHours3(x).SetFocus
            Exit Function
        End If
        If Val(medHours3(x)) > 99999.9999 Then
            MsgBox lblTitle3(x).Caption & " Hours is Invalid"
            medHours3(x).SetFocus
            Exit Function
        End If
    Next x
End If

If frWeek4.Enabled Then
    For x = 0 To 6
        If Len(medHours4(x)) = 0 Then medHours4(x).Text = 0
        
        If Not IsNumeric(medHours4(x)) Then
            MsgBox lblTitle4(x).Caption & " Hours is invalid"
            medHours4(x).SetFocus
            Exit Function
        End If
        If Val(medHours4(x)) > 99999.9999 Then
            MsgBox lblTitle4(x).Caption & " Hours is Invalid"
            medHours4(x).SetFocus
            Exit Function
        End If
    Next x
End If

'Validations on Add New
If fGLBNew Then
    'Check if the same From/To Date schedule already exists. If it does then do not allow to save this.
    If ScheduleAlreadyExists(glbLEE_ID, dlpEffectiveDate.Text) Then
        MsgBox "Work Schedule for this From Date already exists. Cannot save this Work Schedule.", vbInformation, "Work Schedule already exists"
        dlpEffectiveDate.SetFocus
        Exit Function
    End If
    If ScheduleAlreadyExistsFromToDt(glbLEE_ID, dlpEffectiveDate.Text) Then
        MsgBox "Work Schedule for this From Date already exists. Cannot save this Work Schedule.", vbInformation, "Work Schedule already exists"
        dlpEffectiveDate.SetFocus
        Exit Function
    End If
    
    'Check if later Effective Date schedule already exists. If it does then do not allow to save this.
    If LaterScheduleExists(glbLEE_ID, dlpEffectiveDate.Text) Then
        MsgBox "A more recent Work Schedule already exists. Cannot save this Work Schedule.", vbInformation, "Recent Work Schedule already exists"
        dlpEffectiveDate.SetFocus
        Exit Function
    'ElseIf LaterScheduleExistsFromToDt(glbLEE_ID, dlpToDate.Text) Then
    '    MsgBox "A more recent Work Schedule already exists. Cannot save this Work Schedule.", vbInformation, "Recent Work Schedule already exists"
    '    dlpToDate.SetFocus
    '    Exit Function
    End If
    
    'Check if any Unapproved Requests (ESS) exists. If it does then do not allow to save this.
    'If UnapprovedRequestExists(glbLEE_ID) Then
    If UnapprovedRequestExistsFromToDt(glbLEE_ID, dlpEffectiveDate.Text) Then
        'MsgBox "There are Vacation or Time Off Requests pending Approval. Cannot save this Work Schedule.", vbInformation, "Vacation/Time Off Requests pending Approval"
        resp% = MsgBox("Unapproved or Rejected Vacation/Time Requests are outstanding for this employee. These requests may be invalid based on the new Work Schedule. Please have the employee verify their open requests. " & vbCrLf & vbCrLf & "Do you wish to proceed?", vbQuestion + vbYesNo, "Vacation/Time Off Requests pending Approval")
        If resp% <> 6 Then
            'dlpEffectiveDate.SetFocus
            Exit Function
        End If
    ElseIf UnapprovedRequestExistsFromToDt(glbLEE_ID, dlpToDate.Text) Then
        'MsgBox "There are Vacation or Time Off Requests pending Approval. Cannot save this Work Schedule.", vbInformation, "Vacation/Time Off Requests pending Approval"
        resp% = MsgBox("Unapproved or Rejected Vacation/Time Requests are outstanding for this employee. These requests may be invalid based on the new Work Schedule. Please have the employee verify their open requests. " & vbCrLf & vbCrLf & "Do you wish to proceed?", vbQuestion + vbYesNo, "Vacation/Time Off Requests pending Approval")
        If resp% <> 6 Then
            'dlpEffectiveDate.SetFocus
            Exit Function
        End If
    ElseIf UnapprovedRequestExistsFromToDt(glbLEE_ID, dlpEffectiveDate.Text, dlpToDate.Text) Then
        'MsgBox "There are Vacation or Time Off Requests pending Approval. Cannot save this Work Schedule.", vbInformation, "Vacation/Time Off Requests pending Approval"
        resp% = MsgBox("Unapproved or Rejected Vacation/Time Requests are outstanding for this employee. These requests may be invalid based on the new Work Schedule. Please have the employee verify their open requests. " & vbCrLf & vbCrLf & "Do you wish to proceed?", vbQuestion + vbYesNo, "Vacation/Time Off Requests pending Approval")
        If resp% <> 6 Then
            'dlpEffectiveDate.SetFocus
            Exit Function
        End If
    End If
End If

'Validations on changing existing Schedule
If Not fGLBNew Then
    If isWorkSchedule(glbLEE_ID) Then
        'Check if the same To Date schedule already exists. If it does then do not allow to save this.
        If ScheduleAlreadyExistsFromToDt(glbLEE_ID, dlpToDate.Text, RSDATA("SD_ID")) Then
            MsgBox "Work Schedule for this To Date already exists. Cannot save this Work Schedule.", vbInformation, "Work Schedule already exists"
            dlpToDate.SetFocus
            Exit Function
        End If
        
        'Check if To Date overlaps with other existing schedule. If it does then do not allow to save this.
        If OverlapScheduleExists(glbLEE_ID, dlpEffectiveDate.Text, dlpToDate.Text, RSDATA("SD_ID")) Then
            MsgBox "This Work Schedule is overlapping with other Work Schedule(s) of this employee. Cannot save this Work Schedule.", vbInformation, "Work Schedule Overlapping"
            dlpToDate.SetFocus
            Exit Function
        End If
    
        'Check if any Unapproved Requests (ESS) exists. If it does then do not allow to save this.
        'If UnapprovedRequestExists(glbLEE_ID) Then
        If UnapprovedRequestExistsFromToDt(glbLEE_ID, dlpToDate.Text) Then
            'MsgBox "There are Vacation or Time Off Requests pending Approval. Cannot save this Work Schedule.", vbInformation, "Vacation/Time Off Requests pending Approval"
            resp% = MsgBox("Unapproved or Rejected Vacation/Time Requests are outstanding for this employee. These requests may be invalid based on the revised Work Schedule. Please have the employee verify their open requests. " & vbCrLf & vbCrLf & "Do you wish to proceed?", vbQuestion + vbYesNo, "Vacation/Time Off Requests pending Approval")
            If resp% <> 6 Then
                'medHours(0).SetFocus
                Exit Function
            End If
        ElseIf UnapprovedRequestExistsFromToDt(glbLEE_ID, dlpEffectiveDate.Text, dlpToDate.Text) Then
            'MsgBox "There are Vacation or Time Off Requests pending Approval. Cannot save this Work Schedule.", vbInformation, "Vacation/Time Off Requests pending Approval"
            resp% = MsgBox("Unapproved or Rejected Vacation/Time Requests are outstanding for this employee. These requests may be invalid based on the revised Work Schedule. Please have the employee verify their open requests. " & vbCrLf & vbCrLf & "Do you wish to proceed?", vbQuestion + vbYesNo, "Vacation/Time Off Requests pending Approval")
            If resp% <> 6 Then
                'medHours(0).SetFocus
                Exit Function
            End If
        End If
        
        'Any Requests exists from ChangedDate?
        If UnapprovedRequestExistsChangeDate(glbLEE_ID, dlpChangeDate.Text) Then
            resp% = MsgBox("Unapproved or Rejected Vacation/Time Requests are outstanding for this employee from the 'Change Date'. These requests may be invalid based on the revised Work Schedule. Please have the employee verify their open requests. " & vbCrLf & vbCrLf & "Do you wish to proceed?", vbQuestion + vbYesNo, "Vacation/Time Off Requests pending Approval")
            If resp% <> 6 Then
                'medHours(0).SetFocus
                Exit Function
            End If
        End If
    Else
        MsgBox "This employee's Employment Status indicates this employee should not have Work Schedule.", vbInformation, "Work Schedule not required"
        Call cmdCancel_Click
        Exit Function
    End If
End If

'Ticket #24485 - Check if any hours has changed
If Not fGLBNew Then
    If Has_Hours_Changed(glbLEE_ID) Then
        resp% = MsgBox("Any changes to this Employee's Work Schedule hours will not be reflected in any submitted or approved Vacation/Time Requests. Prior to making this change, future-dated Vacation/Time Requests should be deleted." & vbCrLf & vbCrLf & "Are you sure you want to proceed?", vbQuestion + vbYesNo, "Submitted/Approved Vacation/Time Off Requests")
        If resp% <> 6 Then
            Exit Function
        End If
        
    End If
End If

chkEScheduler = True

Exit Function

chkEScheduler_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEScheduler", "HR_SCHEDULER", "edit/Add")
Call RollBack '23July99 js

End Function

Sub cmdCancel_Click()
Dim x
On Error GoTo Can_Err

RSDATA.CancelUpdate

'Ticket #24485 - Work Schedule Rotation Weeks
dlpChangeDate.Text = ""

Call Display_Value


fGLBNew = False

Call SET_UP_MODE
'Call ST_UPD_MODE(True)  ' reset screen's attributes

fGLBNew = False

Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Cancel", "HR_SCHEDULER", "Cancel")
Call RollBack '23July99 js

End Sub

Sub cmdClose_Click()
'Call NextForm
Unload Me
If glbOnTop = "FRMESCHEDULER" Then glbOnTop = ""

End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, x
Dim flgWorkSchExist As Boolean

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

'Validation before allowing to delete
'Check if any Unapproved Requests (ESS) exists. If it does then do not allow to save this.
'If UnapprovedRequestExists(glbLEE_ID) Then
If UnapprovedRequestExistsFromToDt(glbLEE_ID, dlpEffectiveDate.Text) Then
    MsgBox "There are Vacation or Time Off Requests pending Approval. Cannot delete this Work Schedule.", vbInformation, "Vacation/Time Off Requests pending Approval"
    Exit Sub
ElseIf UnapprovedRequestExistsFromToDt(glbLEE_ID, dlpToDate.Text) Then
    MsgBox "There are Vacation or Time Off Requests pending Approval. Cannot delete this Work Schedule.", vbInformation, "Vacation/Time Off Requests pending Approval"
    Exit Sub
End If

'Ticket #22221 - Check if Employment Status says Work Schedule. If yes, then at least one Work Schedule has to exists.
flgWorkSchExist = isWorkSchedule(glbLEE_ID)
If flgWorkSchExist Then
    'Check if more than 1 work schedule for this employee exists
    If Not MoreThanOne_WorkSchedule(glbLEE_ID) Then
        'Cannot delete this Work Schedule
        MsgBox "There must be at least one current Work Schedule for this employee.", vbInformation, "Work Schedule must exist"
        Exit Sub
    End If
End If

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub


If glbtermopen Then
  gdbAdoIhr001X.BeginTrans
  RSDATA.Delete
  gdbAdoIhr001X.CommitTrans
  Data1.Refresh
Else
  gdbAdoIhr001.BeginTrans
  RSDATA.Delete
  gdbAdoIhr001.CommitTrans
  Data1.Refresh
End If

'Ticket #24485 - Work Schedule Rotation Weeks - Delete the WS Details since the template is deleted now
Call Delete_Workschedule_Detail(glbLEE_ID)

'Ticket #24805 & Ticket #24485 - Update Log
Call HRLog(glbLEE_ID, dlpEffectiveDate, dlpToDate, "Delete", "Deleted Work Schedule", "INFOHR")

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

fGLBNew = False

Call SET_UP_MODE
'Call ST_UPD_MODE(True)

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDelete", "HR_SCHEDULER", "Delete")
Call RollBack '23July99 js

End Sub

Sub cmdNew_Click()
Dim SQLQ As String
Dim x As Integer

On Error GoTo AddN_Err

'Ticket #24485 - Do not allow to enter a WS if the Company Pref. is not set.
If lblMsg.Visible Then
    MsgBox "Cannot create Work Schedule because the '# of Work Schedule Rotation Weeks' not setup on Company Preference screen under the Setup menu.", vbInformation, "'# of Work Schedule Rotation Weeks' not setup"
    Exit Sub
End If

'Ticket #22221 - Check first if Employee's Status allows Work Schedule
fGLBNew = isWorkSchedule(glbLEE_ID)
'fglbNew = True

If fGLBNew = False Then
    MsgBox "This employee's Employment Status indicates this employee should not have Work Schedule.", vbInformation, "Work Schedule not required"
    Exit Sub
End If

'Ticket #24485 - Enable/Disable WS Rotation Weeks entry based on the global setting
Call Enable_Disable_RotationWeeks(gsWS_ROTATIONWEEKS)

Call SET_UP_MODE

Call Set_Control("B", Me)

dlpEffectiveDate.Enabled = True

'Ticket #24485 - Work Schedule Rotation Weeks - Change Date do not apply on New record.
dlpChangeDate.Enabled = False

'Ticket #28298
cmdRebuild.Enabled = False

'Ticket #28298 - WDGPHU - To correct some employee's WS Details table
If glbCompSerial = "S/N - 2411W" Then
    cmdNoCondRebuild.Enabled = False
End If

RSDATA.AddNew

If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"

'Ticket #24485 - Save the # of Rotation Weeks
lblRotWeeks = gsWS_ROTATIONWEEKS

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err


Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_SCHEDULER", "Add")
Call RollBack '23July99 js

End Sub

Sub cmdOK_Click()
Dim x
On Error GoTo Add_Err

If Not chkEScheduler() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)

'Ticket #24805 & Ticket #24485 - Update Log
If Not fGLBNew Then
    'If To Date changed, then log the change
    If CVDate(OToDate) <> CVDate(dlpToDate.Text) Then
        Call HRLog(glbLEE_ID, dlpEffectiveDate, dlpToDate, "Modify", "Work Schedule To Date changed from " & Format(OToDate, "mmm dd, yyyy") & " to " & Format(dlpToDate.Text, "mmm dd, yyyy") & " with Change Date as " & Format(dlpChangeDate.Text, "mmm dd, yyyy"), "INFOHR")
    End If
    'Check if Hours has changed
    If Has_Hours_Changed_For_Log(glbLEE_ID) Then
        'Hours has changed updated Log with which week hours changed
        If xWeek1Changed Then
            Call HRLog(glbLEE_ID, dlpEffectiveDate, dlpToDate, "Modify", "Work Schedule Hours changed for Week 1: " & xWK1HrsChanged, "INFOHR")
        End If
        If xWeek2Changed Then
            Call HRLog(glbLEE_ID, dlpEffectiveDate, dlpToDate, "Modify", "Work Schedule Hours changed for Week 2: " & xWK2HrsChanged, "INFOHR")
        End If
        If xWeek3Changed Then
            Call HRLog(glbLEE_ID, dlpEffectiveDate, dlpToDate, "Modify", "Work Schedule Hours changed for Week 3: " & xWK3HrsChanged, "INFOHR")
        End If
        If xWeek4Changed Then
            Call HRLog(glbLEE_ID, dlpEffectiveDate, dlpToDate, "Modify", "Work Schedule Hours changed for Week 4: " & xWK4HrsChanged, "INFOHR")
        End If
    End If
End If

Call Set_Control("U", Me, RSDATA)
'If glbtermopen Then
'    rsDATA!TERM_SEQ = glbTERM_Seq
'    gdbAdoIhr001X.BeginTrans
'    rsDATA.Update
'    gdbAdoIhr001X.CommitTrans
'Else
    gdbAdoIhr001.BeginTrans
    RSDATA.Update
    gdbAdoIhr001.CommitTrans
'End If

'Ticket #24485 - Create Work Schedule Detail based on the Rotation Weeks
Call AddUpdate_Workschedule_Detail(glbLEE_ID)

'Ticket #24805 & Ticket #24485 - Update Log
If fGLBNew Then
    Call HRLog(glbLEE_ID, dlpEffectiveDate, dlpToDate, "Add", "Added a new Work Schedule for " & lblRotWeeks & " Rotation " & IIf(Val(lblRotWeeks) > 1, "Weeks", "Week"), "INFOHR")
End If

Data1.Refresh

'Call ST_UPD_MODE(True)

fGLBNew = False

Call SET_UP_MODE

'Ticket #24485 - Work Schedule Rotation Weeks
dlpChangeDate.Text = ""

Me.vbxTrueGrid.SetFocus

'If NextFormIF("Association") Then
'    Call cmdNew_Click
'End If

Exit Sub

Add_Err:
If Err = 3022 Then
    'Data1.UpdateControls  ' no dups
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_SCHEDULER", "Update")
Call RollBack '23July99 js

End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Work Hours Schedule"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
MDIMain.Timer1.Enabled = False
Me.vbxCrystal.Action = 1
'vbxCrystal.Reset
MDIMain.Timer1.Enabled = True
End Sub

Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = lblEEName & "'s Work Hours Schedule"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
MDIMain.Timer1.Enabled = False
Me.vbxCrystal.Action = 1
'vbxCrystal.Reset
MDIMain.Timer1.Enabled = True
End Sub

Function EERetrieve()
Dim SQLQ As String

Screen.MousePointer = HOURGLASS

EERetrieve = False

On Error GoTo EERError

If Not glbtermopen Then
    'SQLQ = "Select * from Term_USERDEFINE_TABLE"
    'SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    'SQLQ = SQLQ & " ORDER BY UD_CODE1"
'Else
    SQLQ = "SELECT * FROM HR_SCHEDULER"
    SQLQ = SQLQ & " WHERE SD_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY SD_EDATE DESC"
End If

Data1.RecordSource = SQLQ
Data1.Refresh
EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HR_SCHEDULER", "SELECT")
Call RollBack '23July99 js

Exit Function
End Function

Private Sub cmdRebuild_Click()
    'Ticket #28298 - Rebuild the WS Details for the days where there are no Unapproved or Rejected
    'Vacation/Time Off Requests
    Dim strMsg As String
    Dim Response%
    Dim xLastDate
    
    'Make sure the Change Date is entered for the Rebuild. This applies more to the case if they ask for Rebuild with
    'hours changed.
    
    If Not fGLBNew Then
        If Len(dlpChangeDate.Text) < 1 Then
            MsgBox "Change Date is required."
            dlpChangeDate.SetFocus
            Exit Sub
        End If
    Else
        dlpChangeDate.Text = dlpEffectiveDate.Text
    End If
    
    If Not IsDate(dlpChangeDate.Text) Then
        MsgBox "Change Date is not a valid date."
        dlpChangeDate.SetFocus
        Exit Sub
    End If
    
    If CVDate(dlpChangeDate.Text) < CVDate(dlpEffectiveDate.Text) Then
        MsgBox "Change Date cannot be prior to From Date"
        dlpChangeDate.SetFocus
        Exit Sub
    End If
    
    If CVDate(dlpChangeDate.Text) > CVDate(dlpToDate.Text) Then
        MsgBox "Change Date cannot be greater than To Date"
        dlpChangeDate.SetFocus
        Exit Sub
    End If
    
    strMsg = "The Work Schedule Details file will only be rebuilt from the time period where there are no "
    strMsg = strMsg & "Approved, Unapproved or Rejected Vacation/Time Off Requests." & vbCrLf & vbCrLf
    strMsg = strMsg & "Do you wish to proceed?"
    
    Response% = MsgBox(strMsg, vbYesNo + vbQuestion, "info:HR - Work Schedule Details Rebuild")
    
    If Response% = IDNO Then
        Exit Sub
    End If
    
    'Get the Last Date of the Approved, Unapproved or Rejected Vacation/Time Off Requests.
    xLastDate = LastDateUnapprovedRequest(glbLEE_ID, dlpEffectiveDate.Text, dlpToDate.Text)
    If IsDate(xLastDate) Then
        'If the Last Date is greater than or equal to To Date then you cannot rebuild the WS Details as
        'unapproved requests exists else you can rebuild
        If CVDate(xLastDate) >= CVDate(dlpToDate.Text) Then
            'Cannot rebuild the Work Schedule Detail table
            MsgBox "Cannot rebuild the Work Schedule Details for this employee as Approved, Unapproved or Rejected Vacation/Time Off Requests exists for the entire period of the Work Hour Schedule.", vbInformation, "info:HR - Work Schedule Details Cannot Be Rebuilt"
        Else
            'Can rebuild the Work Schedule Details table from the Last Date
            Response% = MsgBox("Rebuilding the Work Schedule Details for this employee after " & Format(xLastDate, "mmm dd, yyyy") & " only. There are Approved, Unapproved or Rejected Vacation/Time Off Requests that exists until " & Format(xLastDate, "mmm dd, yyyy") & "." & vbCrLf & vbCrLf & "Do you wish to proceed?", vbYesNo + vbQuestion, "info:HR - Work Schedule Details Rebuild")
            If Response% = IDNO Then
                Exit Sub
            End If
            
            MDIMain.panHelp(1).Caption = " Rebuilding, Please Wait..."
            Screen.MousePointer = HOURGLASS
            
            Call RebuildWorkScheduleDetail(glbLEE_ID, dlpEffectiveDate.Text, dlpToDate.Text, dlpChangeDate.Text, xLastDate)
            
            'Update Log
            Call HRLog(glbLEE_ID, dlpEffectiveDate, dlpToDate, "Rebuild", "Rebuilt Work Schedule Details after " & Format(xLastDate, "mmm dd, yyyy") & " only. There are Approved, Unapproved or Rejected Vacation/Time Off Requests that exists until " & Format(xLastDate, "mmm dd, yyyy") & ".", "INFOHR")

            MDIMain.panHelp(1).Caption = " Rebuilding Complete"
            
            MsgBox "Rebuild Completed.", vbInformation, "info:HR - Work Schedule Details Rebuilt"
        End If
    Else
        'No any Requests or Unapproved or Rejected Vacation/Time Off Requests found.
        'Can rebuild the entire WS Details for the WS Date Range.
        
        MDIMain.panHelp(1).Caption = " Rebuilding, Please Wait..."
        Screen.MousePointer = HOURGLASS
        
        Call RebuildWorkScheduleDetail(glbLEE_ID, dlpEffectiveDate.Text, dlpToDate.Text, dlpChangeDate.Text)
        
        'Update Log
        Call HRLog(glbLEE_ID, dlpEffectiveDate, dlpToDate, "Rebuild", "Rebuilt Work Schedule Details for the entire period.", "INFOHR")
        
        MDIMain.panHelp(1).Caption = " Rebuilding Complete"
        
        MsgBox "Rebuild Completed.", vbInformation, "info:HR - Work Schedule Details Rebuilt"
    End If
    
    MDIMain.panHelp(1).Caption = ""
    Screen.MousePointer = DEFAULT
End Sub

Private Sub cmdNoCondRebuild_Click()
    'NEW - Rebuild with No Condition from Change Date
    
    'Ticket #28298 - Rebuild the WS Details for the days where there are no Unapproved or Rejected
    'Vacation/Time Off Requests
    Dim strMsg As String
    Dim Response%
    Dim xLastDate
    
    'Make sure the Change Date is entered for the Rebuild. This applies more to the case if they ask for Rebuild with
    'hours changed.
    
    If Not fGLBNew Then
        If Len(dlpChangeDate.Text) < 1 Then
            MsgBox "Change Date is required."
            dlpChangeDate.SetFocus
            Exit Sub
        End If
    Else
        dlpChangeDate.Text = dlpEffectiveDate.Text
    End If
    
    If Not IsDate(dlpChangeDate.Text) Then
        MsgBox "Change Date is not a valid date."
        dlpChangeDate.SetFocus
        Exit Sub
    End If
    
    If CVDate(dlpChangeDate.Text) < CVDate(dlpEffectiveDate.Text) Then
        MsgBox "Change Date cannot be prior to From Date"
        dlpChangeDate.SetFocus
        Exit Sub
    End If
    
    If CVDate(dlpChangeDate.Text) > CVDate(dlpToDate.Text) Then
        MsgBox "Change Date cannot be greater than To Date"
        dlpChangeDate.SetFocus
        Exit Sub
    End If
    
    strMsg = "The Work Schedule Details file will be rebuilt from Change Date"
    strMsg = strMsg & " up to To Date without checking for Approved, Unapproved or Rejected Vacation/Time Off Requests that may exists in this period." & vbCrLf & vbCrLf
    strMsg = strMsg & "Do you wish to proceed?"
    
    Response% = MsgBox(strMsg, vbYesNo + vbQuestion, "info:HR - Work Schedule Details Rebuild with No Checking")
    
    If Response% = IDNO Then
        Exit Sub
    End If
    
    'Rebuilding the WS Details from Change Date onwards to To Date
            
    MDIMain.panHelp(1).Caption = " Rebuilding, Please Wait..."
    Screen.MousePointer = HOURGLASS
    
    Call RebuildWorkScheduleDetail(glbLEE_ID, dlpEffectiveDate.Text, dlpToDate.Text, dlpChangeDate.Text, DateAdd("d", -1, dlpChangeDate.Text))
    
    'Update Log
    Call HRLog(glbLEE_ID, dlpEffectiveDate, dlpToDate, "Rebuild", "Rebuilt Work Schedule Details from Change Date, " & Format(dlpChangeDate.Text, "mmm dd, yyyy") & ", to To Date.", "INFOHR")
    
    MDIMain.panHelp(1).Caption = " Rebuilding with Complete"
    
    MsgBox "Rebuild Completed.", vbInformation, "info:HR - Work Schedule Details Rebuilt with No Checking"
    
    MDIMain.panHelp(1).Caption = ""
    Screen.MousePointer = DEFAULT

End Sub

Private Sub cmdViewRpt_Click()
    Screen.MousePointer = HOURGLASS
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal1.WindowShowPrintSetupBtn = True
    
    'report filename
    Me.vbxCrystal1.ReportFileName = glbIHRREPORTS & "RZUnApprvdLst.rpt"
    
    'set location for database tables
    Me.vbxCrystal1.Connect = RptODBC_SQL
    
    'window title if appropriate
    Me.vbxCrystal1.WindowTitle = "Unapproved or Rejected Vacation/Time Off Requests"
    
    Me.vbxCrystal1.SelectionFormula = "{HREMP.ED_EMPNBR} = " & glbLEE_ID
    Me.vbxCrystal1.Formulas(0) = ""
    Me.vbxCrystal1.Formulas(0) = "rpttitle = 'Unapproved or Rejected Vacation/Time Off Requests'"

    
    Me.vbxCrystal1.Destination = 0
    MDIMain.Timer1.Enabled = False

    Screen.MousePointer = DEFAULT
    Me.vbxCrystal1.Action = 1
    vbxCrystal1.Reset
    MDIMain.Timer1.Enabled = True

    Screen.MousePointer = DEFAULT
End Sub

Private Sub dlpEffectiveDate_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub dlpToDate_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
    glbOnTop = "FRMESCHEDULER"
    Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMESCHEDULER"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "FRMESCHEDULER"
'If glbtermopen Then         'Lucy July 5, 2000
'    Data1.ConnectionString = glbAdoIHRAUDIT
'Else
    Data1.ConnectionString = glbAdoIHRDB
'End If

Screen.MousePointer = DEFAULT

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
'Else
'    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
'    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

'Ticket #28298 - WDGPHU - To correct some employee's WS Details table
If glbCompSerial = "S/N - 2411W" Then
    cmdNoCondRebuild.Visible = True
Else
    cmdNoCondRebuild.Visible = False
End If

'Ticket #24485 - Enable/Disable the Works Schedule Rotation Weeks based on the Company Preference selection
Call Enable_Disable_RotationWeeks(gsWS_ROTATIONWEEKS)

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    'If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    'If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

If Len(glbLEE_SName) < 1 Then Exit Sub

Screen.MousePointer = HOURGLASS
Me.vbxTrueGrid.SetFocus

lblEENum.Caption = ShowEmpnbr(lblEEID)

Call Display_Value
Call ST_UPD_MODE(True)             '
Call INI_Controls(Me)

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
'    Call NextForm
End Sub

Private Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer
Dim x As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

fUPMode = TF    ' update mode

dlpEffectiveDate.Enabled = False
dlpToDate.Enabled = TF
txtComments.Enabled = TF

'Ticket #28298
cmdRebuild.Enabled = TF
'Ticket #28298 - WDGPHU - To correct some employee's WS Details table
If glbCompSerial = "S/N - 2411W" Then
    cmdNoCondRebuild.Enabled = TF
End If

'Ticket #24485 - Work Schedule Rotation Weeks
'Changed the logic on how it gets enabled/disable - only when the respective Week is enabled
If frWeek1.Enabled Then
    For x = 0 To 6
        medHours(x).Enabled = TF
    Next
End If
If frWeek2.Enabled Then
    For x = 0 To 6
        medHours2(x).Enabled = TF
    Next
End If
If frWeek3.Enabled Then
    For x = 0 To 6
        medHours3(x).Enabled = TF
    Next
End If
If frWeek4.Enabled Then
    For x = 0 To 6
        medHours4(x).Enabled = TF
    Next
End If

'Ticket #24485 - Work Schedule Rotation Weeks
dlpChangeDate.Enabled = TF

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
Else
    If IsDate(dlpToDate) Then
        OToDate = CVDate(dlpToDate.Text)
    End If
End If

End Sub

Private Sub medHours_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medHours_LostFocus(Index As Integer)
    If Len(medHours(Index)) = 0 Then
        medHours(Index).Text = 0
    End If
End Sub

Private Sub txtComments_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    If Not glbtermopen Then         'Lucy July 5, 2000
        SQLQ = "SELECT * FROM HR_SCHEDULER"
        SQLQ = SQLQ & " WHERE SD_EMPNBR = " & glbLEE_ID
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
       
    Data1.RecordSource = SQLQ
    Data1.Refresh

End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$
Dim SQLQ As String

On Error GoTo Tab1_Err

Call Display_Value

Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_SCHEDULER", "Add")
Call RollBack '23July99 js

End Sub

Private Function RollBack()
On Error GoTo rr
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
rr:
End Function


''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
Dim SQLQ
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    If glbtermopen Then
        RSDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        RSDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    Call SET_UP_MODE
    Exit Sub
End If
      
If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    
If Not glbtermopen Then
    'SQLQ = "SELECT * FROM Term_USERDEFINE_TABLE"
    'SQLQ = SQLQ & " WHERE TERM_SEQ = " & Data1.Recordset!TERM_SEQ
    'SQLQ = SQLQ & " ORDER BY UD_CODE1"
    'rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
'Else
    SQLQ = "SELECT * FROM HR_SCHEDULER"
    SQLQ = SQLQ & " WHERE SD_ID = " & Data1.Recordset!SD_ID
    SQLQ = SQLQ & " ORDER BY SD_EDATE DESC, SD_TDATE DESC"
    RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

If RSDATA.EOF Or RSDATA.BOF Then Exit Sub
Call Set_Control("R", Me, RSDATA)

'Ticket #24485 - If the To Date is less than the Effective Date on the Company Pref. then set the
'Work Schedule Rotation Weeks to saved # and enable the WS Weeks accordingly
If Not fGLBNew And IsDate(gsWS_ROTATIONWEEKSEFFDATE) Then
    'Not already in the new WS Rotation Weeks
    If IsDate(dlpToDate) Then
        If CVDate(dlpEffectiveDate.Text) < CVDate(gsWS_ROTATIONWEEKSEFFDATE) And CVDate(dlpToDate.Text) < CVDate(gsWS_ROTATIONWEEKSEFFDATE) Then
            Call Enable_Disable_RotationWeeks(lblRotWeeks)
        End If
    End If
End If

Call SET_UP_MODE

End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fGLBNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fGLBNew = True
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Work_Schedule
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
Printable = True
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
If fGLBNew Then
    UpdateState = NewRecord
    TF = True
ElseIf RSDATA.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If

Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
Call ST_UPD_MODE(TF)

End Sub

Private Sub lblEEID_Change()

    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
        frmEScheduler.Caption = lStr("Work Hours Schedule") & " - " & Left$(glbLEE_SName, 5)
        frmEScheduler.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID

    lblEENum = ShowEmpnbr(lblEEID)
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = glbLEE_ProdLine
    Else
        lblEEProdLine = ""
    End If
End Sub

'Private Function AlreadyExists()
'    Dim rsHrScheduler As New ADODB.Recordset
'    Dim SQLQ As String
'
'    AlreadyExists = True
'
'    SQLQ = "SELECT * FROM HR_SCHEDULER WHERE SD_EMPNBR = " & glbLEE_ID
'    SQLQ = SQLQ & " AND SD_EDATE = " & Date_SQL(dlpEffectiveDate.Text)
'    rsHrScheduler.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If Not rsHrScheduler.EOF Then
'        'Schedule already exists
'        AlreadyExists = True
'    Else
'        'Schedule does not exists
'        AlreadyExists = False
'    End If
'    rsHrScheduler.Close
'    Set rsHrScheduler = Nothing
'
'End Function
'
'Private Function LaterScheduleExists()
'    Dim rsHrScheduler As New ADODB.Recordset
'    Dim SQLQ As String
'
'    LaterScheduleExists = True
'
'    SQLQ = "SELECT * FROM HR_SCHEDULER WHERE SD_EMPNBR = " & glbLEE_ID
'    SQLQ = SQLQ & " AND SD_EDATE > " & Date_SQL(dlpEffectiveDate.Text)
'    rsHrScheduler.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If Not rsHrScheduler.EOF Then
'        'Later Schedule already exists
'        LaterScheduleExists = True
'    Else
'        'Later Schedule does not exists
'        LaterScheduleExists = False
'    End If
'    rsHrScheduler.Close
'    Set rsHrScheduler = Nothing
'
'End Function
'
'Private Function UnapprovedRequestExists()
'    Dim rsVacReq As New ADODB.Recordset
'    Dim SQLQ As String
'
'    UnapprovedRequestExists = True
'
'    'Create a query to give a list of unapproved requests and with date range greater than 1 day
'    SQLQ = "SELECT * FROM HR_VACTIMEOFF_REQ WHERE VT_EMPNBR = " & glbLEE_ID
'    SQLQ = SQLQ & " AND VT_DELFLAG=0"
'    SQLQ = SQLQ & " AND VT_FROM <> VT_TO"   'date range has more than 1 day
'    SQLQ = SQLQ & " AND (VT_APPROVED IS NULL OR VT_APPROVED = '' OR VT_APPROVED IN ('RESUBMITTED','APP/FWD'))"
'    SQLQ = SQLQ & " AND VT_VACTIME=1 "  'vacation time only
'    rsVacReq.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If Not rsVacReq.EOF Then
'        UnapprovedRequestExists = True
'    Else
'        UnapprovedRequestExists = False
'    End If
'    rsVacReq.Close
'    Set rsVacReq = Nothing
'
'    'If a list of unapproved requests are required then show the following fields:
'    'Number, Name,From Date, To Date, Requested Date, Processed Date, Supervisor Name, Status, Hours Requested
'    'VT_EMPNBR,VT_EMPNAME,VT_FROM,VT_TO,VT_REQDATE,VT_PROCDATE,VT_SUPERNAME,VT_APPROVED(Status),VT_HRS(Hours Requested)?
'End Function

Private Function MoreThanOne_WorkSchedule(xEmpNo) As Boolean
    Dim rsWorkSch As New ADODB.Recordset
    Dim SQLQ As String
    
    SQLQ = "SELECT SD_EMPNBR FROM HR_SCHEDULER WHERE SD_EMPNBR = " & xEmpNo
    rsWorkSch.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsWorkSch.EOF Then
        If rsWorkSch.RecordCount > 1 Then
            MoreThanOne_WorkSchedule = True
        Else
            MoreThanOne_WorkSchedule = False
        End If
    End If
    rsWorkSch.Close
    Set rsWorkSch = Nothing
End Function

Private Function AddUpdate_Workschedule_Detail(xEmpNo)
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
    
            
    'For details found, look for what has changed.
        '- Verify when the Change should be effective.
        '- Check if any Request exists from the Changed Date.
            '- if so then do not allow to save the change
            '- if no then save the change but from the Changed Date onwards for the respective weeks only. Also update
            '  the Changed Date as well for that record entry
    
    SQLQ = "SELECT * FROM HR_SCHEDULER_DETAIL WHERE WS_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " AND WS_FDATE = " & Date_SQL(dlpEffectiveDate.Text)
    'If To Date changed, then retrieve records by old To Date
    If CVDate(OToDate) <> CVDate(dlpToDate.Text) Then
        SQLQ = SQLQ & " AND WS_TDATE = " & Date_SQL(OToDate)
    Else
        SQLQ = SQLQ & " AND WS_TDATE = " & Date_SQL(dlpToDate.Text)
    End If
    rsSchDetail.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsSchDetail.EOF Then
        'No Work Schedule detail found. Create details
        
        'Set the Start Date of the Period so for each date upto the To Date the Work Schedule can be prepared
        xCurrDate = CVDate(dlpEffectiveDate.Text) - 1   'Starting with -1 because when adding Details, I am +1 day for WS_DATE
        
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
                    rsSchDetail("WS_DATE") = DateAdd("d", 1, xCurrDate) 'Because started with -1 day as Start Day
                    rsSchDetail("WS_HRS") = medHours(xDay)
                
                    rsSchDetail("WS_LDATE") = Date
                    rsSchDetail("WS_LTIME") = Time$
                    rsSchDetail("WS_LUSER") = glbUserID
                    rsSchDetail.Update
                    
                    'ReSet the CurrDate
                    xCurrDate = CVDate(rsSchDetail("WS_DATE"))
                    
                    'Changed from > to >= it was adding an extra day after To Date
                    'Exit the Loop if xCurrDate exceeds the ToDate.
                    If CVDate(xCurrDate) >= CVDate(dlpToDate.Text) Then Exit Do
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
                    rsSchDetail("WS_DATE") = DateAdd("d", 1, xCurrDate) 'Because started with -1 day as Start Day
                    rsSchDetail("WS_HRS") = medHours2(xDay)
                
                    rsSchDetail("WS_LDATE") = Date
                    rsSchDetail("WS_LTIME") = Time$
                    rsSchDetail("WS_LUSER") = glbUserID
                    rsSchDetail.Update
                    
                    'ReSet the CurrDate
                    xCurrDate = CVDate(rsSchDetail("WS_DATE"))
                
                    'Changed from > to >= it was adding an extra day after To Date
                    'Exit the Loop if xCurrDate exceeds the ToDate.
                    If CVDate(xCurrDate) >= CVDate(dlpToDate.Text) Then Exit Do
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
                    rsSchDetail("WS_DATE") = DateAdd("d", 1, xCurrDate) 'Because started with -1 day as Start Day
                    rsSchDetail("WS_HRS") = medHours3(xDay)
                
                    rsSchDetail("WS_LDATE") = Date
                    rsSchDetail("WS_LTIME") = Time$
                    rsSchDetail("WS_LUSER") = glbUserID
                    rsSchDetail.Update
                    
                    'ReSet the CurrDate
                    xCurrDate = CVDate(rsSchDetail("WS_DATE"))
                    
                    'Changed from > to >= it was adding an extra day after To Date
                    'Exit the Loop if xCurrDate exceeds the ToDate.
                    If CVDate(xCurrDate) >= CVDate(dlpToDate.Text) Then Exit Do
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
                    rsSchDetail("WS_DATE") = DateAdd("d", 1, xCurrDate) 'Because started with -1 day as Start Day
                    rsSchDetail("WS_HRS") = medHours4(xDay)
                
                    rsSchDetail("WS_LDATE") = Date
                    rsSchDetail("WS_LTIME") = Time$
                    rsSchDetail("WS_LUSER") = glbUserID
                    rsSchDetail.Update
                    
                    'ReSet the CurrDate
                    xCurrDate = CVDate(rsSchDetail("WS_DATE"))
                    
                    'Changed from > to >= it was adding an extra day after To Date
                    'Exit the Loop if xCurrDate exceeds the ToDate.
                    If CVDate(xCurrDate) >= CVDate(dlpToDate.Text) Then Exit Do
                Next
                
                'Reset the start day for the next Week
                xStartDay = 0
            End If
            
NextWorkSchSet:
        Loop
    Else
        'Work Schedule details found, update the existing details based on the Week's Hours changed
        'Also, check if the To Date has changed then change existing record's To Date. May need to add more/delete
        'existing work schedule records based on To Date
        
        'First change all old To Dates to new To Dates
        rsSchDetail.MoveFirst
        Do While Not rsSchDetail.EOF
            rsSchDetail("WS_TDATE") = dlpToDate.Text
                        
            rsSchDetail("WS_LDATE") = Date
            rsSchDetail("WS_LTIME") = Time$
            rsSchDetail("WS_LUSER") = glbUserID
            
            rsSchDetail.Update
        
            rsSchDetail.MoveNext
        Loop
        
        'Now if the old To Date is greater than the new To Date then delete extra old To Dates detail records
        If CVDate(OToDate) > CVDate(dlpToDate.Text) Then
            SQLQ = "DELETE FROM HR_SCHEDULER_DETAIL WHERE WS_EMPNBR = " & xEmpNo
            SQLQ = SQLQ & " AND WS_FDATE = " & Date_SQL(dlpEffectiveDate.Text)
            SQLQ = SQLQ & " AND WS_TDATE = " & Date_SQL(dlpToDate.Text)
            SQLQ = SQLQ & " AND WS_DATE > " & Date_SQL(dlpToDate.Text)
            gdbAdoIhr001.Execute SQLQ
        End If
        
        'Now if the old To Date is less than the new To Date then add the extra WS details records
        If CVDate(OToDate) < CVDate(dlpToDate.Text) Then
            xWeekNo = 0
            xWeekDayCount = 0
            xLstWSWkDay = 0
            
            SQLQ = "SELECT * FROM HR_SCHEDULER_DETAIL WHERE WS_EMPNBR = " & xEmpNo
            SQLQ = SQLQ & " AND WS_FDATE = " & Date_SQL(dlpEffectiveDate.Text)
            SQLQ = SQLQ & " AND WS_TDATE = " & Date_SQL(dlpToDate.Text)
            rsSchDetail2.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsSchDetail2.EOF Then
                'Get the last record's Week No and then count the # of days details records exists for that Week #
                rsSchDetail2.MoveLast
                xWeekNo = rsSchDetail2("WS_WEEKNO")
                xWeekDayCount = xWeekDayCount + 1
                
                'Last record's WS Date and Day of the Week.
                xLstWSDate = CVDate(rsSchDetail2("WS_DATE"))
                xLstWSWkDay = Weekday(rsSchDetail2("WS_DATE"))
                
                rsSchDetail2.MovePrevious
                Do While Not rsSchDetail2.BOF
                    'Count the # of records exists for the same Week # so we know from what day to add more details
                    If xWeekNo = rsSchDetail2("WS_WEEKNO") Then
                        xWeekDayCount = xWeekDayCount + 1
                    Else
                        'Move back to same Week #
                        rsSchDetail2.MoveNext
                        
                        'Moved up where the last day is read first (MoveLast), as we want to start from that day
                        'Last WS Date
                        'xLstWSDate = CVDate(rsSchDetail2("WS_DATE"))
                        'xLstWSWkDay = Weekday(rsSchDetail2("WS_DATE"))
                        Exit Do
                    End If
                    
                    rsSchDetail2.MovePrevious
                Loop
                
                'Add the rest of the week's detail records upto the new To Date
                'Week # to start entering the detail records from
                If xWeekDayCount = 7 Then
                    If xWeekNo = 4 Then
                        xWeekNo = 1
                    Else
                        xWeekNo = xWeekNo + 1
                    End If
                End If
                
                'Day to start adding the WS details
                If xLstWSWkDay < 7 Then
                    xStartDay = xLstWSWkDay
                Else
                    xStartDay = 0
                End If
                
                'Start adding additional day's WS Details records
                Do While CVDate(xLstWSDate) < CVDate(dlpToDate.Text)
                    'Start from the Week the last day WS detail record exists
                    If xWeekNo = 1 Then xWeekNo = 0: GoTo Week1     'Reset the WeekNo for next round
                    If xWeekNo = 2 Then xWeekNo = 0: GoTo Week2
                    If xWeekNo = 3 Then xWeekNo = 0: GoTo Week3
                    If xWeekNo = 4 Then xWeekNo = 0: GoTo Week4
Week1:
                    If frWeek1.Enabled Then
                        For xDay = xStartDay To 6
                            'Add Work Schedule for the day in a week
                            rsSchDetail2.AddNew
                            rsSchDetail2("WS_EMPNBR") = xEmpNo
                            rsSchDetail2("WS_FDATE") = dlpEffectiveDate.Text
                            rsSchDetail2("WS_TDATE") = dlpToDate.Text
                            rsSchDetail2("WS_CHGDATE") = dlpChangeDate.Text   'Change from Date will be same as From Date the first time
                        
                            rsSchDetail2("WS_WEEKNO") = 1
                            rsSchDetail2("WS_DATE") = DateAdd("d", 1, xLstWSDate)   'Because started with -1 day as Start Day
                            rsSchDetail2("WS_HRS") = medHours(xDay)
                        
                            rsSchDetail2("WS_LDATE") = Date
                            rsSchDetail2("WS_LTIME") = Time$
                            rsSchDetail2("WS_LUSER") = glbUserID
                            rsSchDetail2.Update
                            
                            'ReSet the xLstWSDate
                            xLstWSDate = CVDate(rsSchDetail2("WS_DATE"))
                            
                            'Changed from > to >= it was adding an extra day after To Date
                            'Exit the Loop if xLstWSDate exceeds the ToDate.
                            If CVDate(xLstWSDate) >= CVDate(dlpToDate.Text) Then Exit Do
                        Next
                        
                        'Reset the start day for the next Week
                        xStartDay = 0
                    End If
                    
Week2:
                    If frWeek2.Enabled Then
                        For xDay = xStartDay To 6
                            'Add Work Schedule for the day in a week
                            rsSchDetail2.AddNew
                            rsSchDetail2("WS_EMPNBR") = xEmpNo
                            rsSchDetail2("WS_FDATE") = dlpEffectiveDate.Text
                            rsSchDetail2("WS_TDATE") = dlpToDate.Text
                            rsSchDetail2("WS_CHGDATE") = dlpChangeDate.Text   'Change from Date will be same as From Date the first time
                        
                            rsSchDetail2("WS_WEEKNO") = 2
                            rsSchDetail2("WS_DATE") = DateAdd("d", 1, xLstWSDate)   'Because started with -1 day as Start Day
                            rsSchDetail2("WS_HRS") = medHours2(xDay)
                        
                            rsSchDetail2("WS_LDATE") = Date
                            rsSchDetail2("WS_LTIME") = Time$
                            rsSchDetail2("WS_LUSER") = glbUserID
                            rsSchDetail2.Update
                            
                            'ReSet the xLstWSDate
                            xLstWSDate = CVDate(rsSchDetail2("WS_DATE"))
                        
                            'Changed from > to >= it was adding an extra day after To Date
                            'Exit the Loop if xLstWSDate exceeds the ToDate.
                            If CVDate(xLstWSDate) >= CVDate(dlpToDate.Text) Then Exit Do
                        Next
                        'Reset the start day for the next Week
                        xStartDay = 0
                    Else
                        'If this week is not visible so the rest of the Weeks will not be visible too
                        GoTo NextWorkSchSet1
                    End If
                    
Week3:
                    If frWeek3.Enabled Then
                        For xDay = xStartDay To 6
                            'Add Work Schedule for the day in a week
                            rsSchDetail2.AddNew
                            rsSchDetail2("WS_EMPNBR") = xEmpNo
                            rsSchDetail2("WS_FDATE") = dlpEffectiveDate.Text
                            rsSchDetail2("WS_TDATE") = dlpToDate.Text
                            rsSchDetail2("WS_CHGDATE") = dlpChangeDate.Text   'Change from Date will be same as From Date the first time
                        
                            rsSchDetail2("WS_WEEKNO") = 3
                            rsSchDetail2("WS_DATE") = DateAdd("d", 1, xLstWSDate)   'Because started with -1 day as Start Day
                            rsSchDetail2("WS_HRS") = medHours3(xDay)
                        
                            rsSchDetail2("WS_LDATE") = Date
                            rsSchDetail2("WS_LTIME") = Time$
                            rsSchDetail2("WS_LUSER") = glbUserID
                            rsSchDetail2.Update
                            
                            'ReSet the xLstWSDate
                            xLstWSDate = CVDate(rsSchDetail2("WS_DATE"))
                            
                            'Changed from > to >= it was adding an extra day after To Date
                            'Exit the Loop if xLstWSDate exceeds the ToDate.
                            If CVDate(xLstWSDate) >= CVDate(dlpToDate.Text) Then Exit Do
                        Next
                        'Reset the start day for the next Week
                        xStartDay = 0
                    Else
                        'If this week is not visible so the rest of the Weeks will not be visible too
                        GoTo NextWorkSchSet1
                    End If
                    
Week4:
                    If frWeek4.Enabled Then
                        For xDay = xStartDay To 6
                            'Add Work Schedule for the day in a week
                            rsSchDetail2.AddNew
                            rsSchDetail2("WS_EMPNBR") = xEmpNo
                            rsSchDetail2("WS_FDATE") = dlpEffectiveDate.Text
                            rsSchDetail2("WS_TDATE") = dlpToDate.Text
                            rsSchDetail2("WS_CHGDATE") = dlpChangeDate.Text   'Change from Date will be same as From Date the first time
                        
                            rsSchDetail2("WS_WEEKNO") = 4
                            rsSchDetail2("WS_DATE") = DateAdd("d", 1, xLstWSDate)   'Because started with -1 day as Start Day
                            rsSchDetail2("WS_HRS") = medHours4(xDay)
                        
                            rsSchDetail2("WS_LDATE") = Date
                            rsSchDetail2("WS_LTIME") = Time$
                            rsSchDetail2("WS_LUSER") = glbUserID
                            rsSchDetail2.Update
                            
                            'ReSet the xLstWSDate
                            xLstWSDate = CVDate(rsSchDetail2("WS_DATE"))
                            
                            'Changed from > to >= it was adding an extra day after To Date
                            'Exit the Loop if xLstWSDate exceeds the ToDate.
                            If CVDate(xLstWSDate) >= CVDate(dlpToDate.Text) Then Exit Do
                        Next
                        'Reset the start day for the next Week
                        xStartDay = 0
                    End If
NextWorkSchSet1:
                Loop
            End If
            rsSchDetail2.Close
            Set rsSchDetail2 = Nothing
        End If
        'rsSchDetail2.Close
        'Set rsSchDetail2 = Nothing
        
        'Now change the hours if hours have changed.
        
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

Private Function Delete_Workschedule_Detail(xEmpNo)
    Dim SQLQ As String
    
    SQLQ = "DELETE FROM HR_SCHEDULER_DETAIL WHERE WS_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " AND WS_FDATE = " & Date_SQL(dlpEffectiveDate.Text)
    If IsDate(dlpToDate) Then
        SQLQ = SQLQ & " AND WS_TDATE = " & Date_SQL(dlpToDate.Text)
    Else
        SQLQ = SQLQ & " AND WS_TDATE IS NULL"
    End If
    gdbAdoIhr001.Execute SQLQ
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
    
    xWK1HrsChanged = ""
    xWK2HrsChanged = ""
    xWK3HrsChanged = ""
    xWK4HrsChanged = ""
    
    Has_Hours_Changed = False
    
    'Check if the hours has changed in the WS Master table. Check with old To Date in case the To Date has changed
    SQLQ = "SELECT * FROM HR_SCHEDULER WHERE SD_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " AND SD_EDATE = " & Date_SQL(dlpEffectiveDate.Text)
    'If To Date changed, then retrieve records by old To Date
    If CVDate(OToDate) <> CVDate(dlpToDate.Text) Then
        SQLQ = SQLQ & " AND SD_TDATE = " & Date_SQL(OToDate)
    Else
        SQLQ = SQLQ & " AND SD_TDATE = " & Date_SQL(dlpToDate.Text)
    End If
    rsSch.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsSch.EOF Then
        'Compare Week 1
        For xDay = 0 To 6
            If rsSch("SD_SUN_HRS") <> medHours(xDay) Then
                xWK1HrsChanged = xWK1HrsChanged & "Sunday from " & rsSch("SD_SUN_HRS") & " to " & medHours(xDay) & ". "
                xWeek1Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_MON_HRS") <> medHours(xDay) Then
                xWK1HrsChanged = xWK1HrsChanged & "Monday from " & rsSch("SD_MON_HRS") & " to " & medHours(xDay) & ". "
                xWeek1Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_TUE_HRS") <> medHours(xDay) Then
                xWK1HrsChanged = xWK1HrsChanged & "Tuesday from " & rsSch("SD_TUE_HRS") & " to " & medHours(xDay) & ". "
                xWeek1Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_WED_HRS") <> medHours(xDay) Then
                xWK1HrsChanged = xWK1HrsChanged & "Wednesday from " & rsSch("SD_WED_HRS") & " to " & medHours(xDay) & ". "
                xWeek1Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_THU_HRS") <> medHours(xDay) Then
                xWK1HrsChanged = xWK1HrsChanged & "Thursday from " & rsSch("SD_THU_HRS") & " to " & medHours(xDay) & ". "
                xWeek1Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_FRI_HRS") <> medHours(xDay) Then
                xWK1HrsChanged = xWK1HrsChanged & "Friday from " & rsSch("SD_FRI_HRS") & " to " & medHours(xDay) & ". "
                xWeek1Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_SAT_HRS") <> medHours(xDay) Then
                xWK1HrsChanged = xWK1HrsChanged & "Saturday from " & rsSch("SD_SAT_HRS") & " to " & medHours(xDay) & ". "
                xWeek1Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
        Next
        
        'Compare Week 2
        For xDay = 0 To 6
            If rsSch("SD_SUN_HRS2") <> medHours2(xDay) Then
                xWK2HrsChanged = xWK2HrsChanged & "Sunday from " & rsSch("SD_SUN_HRS2") & " to " & medHours2(xDay) & ". "
                xWeek2Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_MON_HRS2") <> medHours2(xDay) Then
                xWK2HrsChanged = xWK2HrsChanged & "Monday from " & rsSch("SD_MON_HRS2") & " to " & medHours2(xDay) & ". "
                xWeek2Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_TUE_HRS2") <> medHours2(xDay) Then
                xWK2HrsChanged = xWK2HrsChanged & "Tuesday from " & rsSch("SD_TUE_HRS2") & " to " & medHours2(xDay) & ". "
                xWeek2Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_WED_HRS2") <> medHours2(xDay) Then
                xWK2HrsChanged = xWK2HrsChanged & "Wednesday from " & rsSch("SD_WED_HRS2") & " to " & medHours2(xDay) & ". "
                xWeek2Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_THU_HRS2") <> medHours2(xDay) Then
                xWK2HrsChanged = xWK2HrsChanged & "Thursday from " & rsSch("SD_THU_HRS2") & " to " & medHours2(xDay) & ". "
                xWeek2Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_FRI_HRS2") <> medHours2(xDay) Then
                xWK2HrsChanged = xWK2HrsChanged & "Friday from " & rsSch("SD_FRI_HRS2") & " to " & medHours2(xDay) & ". "
                xWeek2Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_SAT_HRS2") <> medHours2(xDay) Then
                xWK2HrsChanged = xWK2HrsChanged & "Saturday from " & rsSch("SD_SAT_HRS2") & " to " & medHours2(xDay) & ". "
                xWeek2Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
        Next
        
        'Compare Week 3
        For xDay = 0 To 6
            If rsSch("SD_SUN_HRS3") <> medHours3(xDay) Then
                xWK3HrsChanged = xWK3HrsChanged & "Sunday from " & rsSch("SD_SUN_HRS3") & " to " & medHours3(xDay) & ". "
                xWeek3Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_MON_HRS3") <> medHours3(xDay) Then
                xWK3HrsChanged = xWK3HrsChanged & "Monday from " & rsSch("SD_MON_HRS3") & " to " & medHours3(xDay) & ". "
                xWeek3Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_TUE_HRS3") <> medHours3(xDay) Then
                xWK3HrsChanged = xWK3HrsChanged & "Tuesday from " & rsSch("SD_TUE_HRS3") & " to " & medHours3(xDay) & ". "
                xWeek3Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_WED_HRS3") <> medHours3(xDay) Then
                xWK3HrsChanged = xWK3HrsChanged & "Wednesday from " & rsSch("SD_WED_HRS3") & " to " & medHours3(xDay) & ". "
                xWeek3Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_THU_HRS3") <> medHours3(xDay) Then
                xWK3HrsChanged = xWK3HrsChanged & "Thursday from " & rsSch("SD_THU_HRS3") & " to " & medHours3(xDay) & ". "
                xWeek3Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_FRI_HRS3") <> medHours3(xDay) Then
                xWK3HrsChanged = xWK3HrsChanged & "Friday from " & rsSch("SD_FRI_HRS3") & " to " & medHours3(xDay) & ". "
                xWeek3Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_SAT_HRS3") <> medHours3(xDay) Then
                xWK3HrsChanged = xWK3HrsChanged & "Saturday from " & rsSch("SD_SAT_HRS3") & " to " & medHours3(xDay) & ". "
                xWeek3Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
        Next
        
        'Compare Week 4
        For xDay = 0 To 6
            If rsSch("SD_SUN_HRS4") <> medHours4(xDay) Then
                xWK4HrsChanged = xWK4HrsChanged & "Sunday from " & rsSch("SD_SUN_HRS4") & " to " & medHours4(xDay) & ". "
                xWeek4Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_MON_HRS4") <> medHours4(xDay) Then
                xWK4HrsChanged = xWK4HrsChanged & "Monday from " & rsSch("SD_MON_HRS4") & " to " & medHours4(xDay) & ". "
                xWeek4Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_TUE_HRS4") <> medHours4(xDay) Then
                xWK4HrsChanged = xWK4HrsChanged & "Tuesday from " & rsSch("SD_TUE_HRS4") & " to " & medHours4(xDay) & ". "
                xWeek4Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_WED_HRS4") <> medHours4(xDay) Then
                xWK4HrsChanged = xWK4HrsChanged & "Wednesday from " & rsSch("SD_WED_HRS4") & " to " & medHours4(xDay) & ". "
                xWeek4Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_THU_HRS4") <> medHours4(xDay) Then
                xWK4HrsChanged = xWK4HrsChanged & "Thursday from " & rsSch("SD_THU_HRS4") & " to " & medHours4(xDay) & ". "
                xWeek4Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_FRI_HRS4") <> medHours4(xDay) Then
                xWK4HrsChanged = xWK4HrsChanged & "Friday from " & rsSch("SD_FRI_HRS4") & " to " & medHours4(xDay) & ". "
                xWeek4Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
            If rsSch("SD_SAT_HRS4") <> medHours4(xDay) Then
                xWK4HrsChanged = xWK4HrsChanged & "Saturday from " & rsSch("SD_SAT_HRS4") & " to " & medHours4(xDay) & ". "
                xWeek4Changed = True
                Has_Hours_Changed = True
                Exit For
            End If
        Next
    End If
    rsSch.Close
    Set rsSch = Nothing
    
End Function

Private Function Has_Hours_Changed_For_Log(xEmpNo) As Boolean
    Dim rsSch As New ADODB.Recordset
    Dim SQLQ As String
    Dim xDay As Integer
    
    'Initialise
    xWeek1Changed = False
    xWeek2Changed = False
    xWeek3Changed = False
    xWeek4Changed = False
    
    xWK1HrsChanged = ""
    xWK2HrsChanged = ""
    xWK3HrsChanged = ""
    xWK4HrsChanged = ""
    
    Has_Hours_Changed_For_Log = False
    
    'Check if the hours has changed in the WS Master table. Check with old To Date in case the To Date has changed
    SQLQ = "SELECT * FROM HR_SCHEDULER WHERE SD_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " AND SD_EDATE = " & Date_SQL(dlpEffectiveDate.Text)
    'If To Date changed, then retrieve records by old To Date
    If CVDate(OToDate) <> CVDate(dlpToDate.Text) Then
        SQLQ = SQLQ & " AND SD_TDATE = " & Date_SQL(OToDate)
    Else
        SQLQ = SQLQ & " AND SD_TDATE = " & Date_SQL(dlpToDate.Text)
    End If
    rsSch.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsSch.EOF Then
        'Compare Week 1
        'For xDay = 0 To 6
            If rsSch("SD_SUN_HRS") <> medHours(0) Then
                xWK1HrsChanged = xWK1HrsChanged & "Sunday from " & rsSch("SD_SUN_HRS") & " to " & medHours(0) & ". "
                xWeek1Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK1
            End If
            If rsSch("SD_MON_HRS") <> medHours(1) Then
                xWK1HrsChanged = xWK1HrsChanged & "Monday from " & rsSch("SD_MON_HRS") & " to " & medHours(1) & ". "
                xWeek1Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK1
            End If
            If rsSch("SD_TUE_HRS") <> medHours(2) Then
                xWK1HrsChanged = xWK1HrsChanged & "Tuesday from " & rsSch("SD_TUE_HRS") & " to " & medHours(2) & ". "
                xWeek1Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK1
            End If
            If rsSch("SD_WED_HRS") <> medHours(3) Then
                xWK1HrsChanged = xWK1HrsChanged & "Wednesday from " & rsSch("SD_WED_HRS") & " to " & medHours(3) & ". "
                xWeek1Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK1
            End If
            If rsSch("SD_THU_HRS") <> medHours(4) Then
                xWK1HrsChanged = xWK1HrsChanged & "Thursday from " & rsSch("SD_THU_HRS") & " to " & medHours(4) & ". "
                xWeek1Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK1
            End If
            If rsSch("SD_FRI_HRS") <> medHours(5) Then
                xWK1HrsChanged = xWK1HrsChanged & "Friday from " & rsSch("SD_FRI_HRS") & " to " & medHours(5) & ". "
                xWeek1Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK1
            End If
            If rsSch("SD_SAT_HRS") <> medHours(6) Then
                xWK1HrsChanged = xWK1HrsChanged & "Saturday from " & rsSch("SD_SAT_HRS") & " to " & medHours(6) & ". "
                xWeek1Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK1
            End If
'NextDayWK1:
        'Next
        
        'Compare Week 2
        'For xDay = 0 To 6
            If rsSch("SD_SUN_HRS2") <> medHours2(0) Then
                xWK2HrsChanged = xWK2HrsChanged & "Sunday from " & rsSch("SD_SUN_HRS2") & " to " & medHours2(0) & ". "
                xWeek2Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK2
            End If
            If rsSch("SD_MON_HRS2") <> medHours2(1) Then
                xWK2HrsChanged = xWK2HrsChanged & "Monday from " & rsSch("SD_MON_HRS2") & " to " & medHours2(1) & ". "
                xWeek2Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK2
            End If
            If rsSch("SD_TUE_HRS2") <> medHours2(2) Then
                xWK2HrsChanged = xWK2HrsChanged & "Tuesday from " & rsSch("SD_TUE_HRS2") & " to " & medHours2(2) & ". "
                xWeek2Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK2
            End If
            If rsSch("SD_WED_HRS2") <> medHours2(3) Then
                xWK2HrsChanged = xWK2HrsChanged & "Wednesday from " & rsSch("SD_WED_HRS2") & " to " & medHours2(3) & ". "
                xWeek2Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK2
            End If
            If rsSch("SD_THU_HRS2") <> medHours2(4) Then
                xWK2HrsChanged = xWK2HrsChanged & "Thursday from " & rsSch("SD_THU_HRS2") & " to " & medHours2(4) & ". "
                xWeek2Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK2
            End If
            If rsSch("SD_FRI_HRS2") <> medHours2(5) Then
                xWK2HrsChanged = xWK2HrsChanged & "Friday from " & rsSch("SD_FRI_HRS2") & " to " & medHours2(5) & ". "
                xWeek2Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK2
            End If
            If rsSch("SD_SAT_HRS2") <> medHours2(6) Then
                xWK2HrsChanged = xWK2HrsChanged & "Saturday from " & rsSch("SD_SAT_HRS2") & " to " & medHours2(6) & ". "
                xWeek2Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK2
            End If
'NextDayWK2:
        'Next
        
        'Compare Week 3
        'For xDay = 0 To 6
            If rsSch("SD_SUN_HRS3") <> medHours3(0) Then
                xWK3HrsChanged = xWK3HrsChanged & "Sunday from " & rsSch("SD_SUN_HRS3") & " to " & medHours3(0) & ". "
                xWeek3Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK3
            End If
            If rsSch("SD_MON_HRS3") <> medHours3(1) Then
                xWK3HrsChanged = xWK3HrsChanged & "Monday from " & rsSch("SD_MON_HRS3") & " to " & medHours3(1) & ". "
                xWeek3Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK3
            End If
            If rsSch("SD_TUE_HRS3") <> medHours3(2) Then
                xWK3HrsChanged = xWK3HrsChanged & "Tuesday from " & rsSch("SD_TUE_HRS3") & " to " & medHours3(2) & ". "
                xWeek3Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK3
            End If
            If rsSch("SD_WED_HRS3") <> medHours3(3) Then
                xWK3HrsChanged = xWK3HrsChanged & "Wednesday from " & rsSch("SD_WED_HRS3") & " to " & medHours3(3) & ". "
                xWeek3Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK3
            End If
            If rsSch("SD_THU_HRS3") <> medHours3(4) Then
                xWK3HrsChanged = xWK3HrsChanged & "Thursday from " & rsSch("SD_THU_HRS3") & " to " & medHours3(4) & ". "
                xWeek3Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK3
            End If
            If rsSch("SD_FRI_HRS3") <> medHours3(5) Then
                xWK3HrsChanged = xWK3HrsChanged & "Friday from " & rsSch("SD_FRI_HRS3") & " to " & medHours3(5) & ". "
                xWeek3Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK3
            End If
            If rsSch("SD_SAT_HRS3") <> medHours3(6) Then
                xWK3HrsChanged = xWK3HrsChanged & "Saturday from " & rsSch("SD_SAT_HRS3") & " to " & medHours3(6) & ". "
                xWeek3Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK3
            End If
'NextDayWK3:
        'Next
        
        'Compare Week 4
        'For xDay = 0 To 6
            If rsSch("SD_SUN_HRS4") <> medHours4(0) Then
                xWK4HrsChanged = xWK4HrsChanged & "Sunday from " & rsSch("SD_SUN_HRS4") & " to " & medHours4(0) & ". "
                xWeek4Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK4
            End If
            If rsSch("SD_MON_HRS4") <> medHours4(1) Then
                xWK4HrsChanged = xWK4HrsChanged & "Monday from " & rsSch("SD_MON_HRS4") & " to " & medHours4(1) & ". "
                xWeek4Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK4
            End If
            If rsSch("SD_TUE_HRS4") <> medHours4(2) Then
                xWK4HrsChanged = xWK4HrsChanged & "Tuesday from " & rsSch("SD_TUE_HRS4") & " to " & medHours4(2) & ". "
                xWeek4Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK4
            End If
            If rsSch("SD_WED_HRS4") <> medHours4(3) Then
                xWK4HrsChanged = xWK4HrsChanged & "Wednesday from " & rsSch("SD_WED_HRS4") & " to " & medHours4(3) & ". "
                xWeek4Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK4
            End If
            If rsSch("SD_THU_HRS4") <> medHours4(4) Then
                xWK4HrsChanged = xWK4HrsChanged & "Thursday from " & rsSch("SD_THU_HRS4") & " to " & medHours4(4) & ". "
                xWeek4Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK4
            End If
            If rsSch("SD_FRI_HRS4") <> medHours4(5) Then
                xWK4HrsChanged = xWK4HrsChanged & "Friday from " & rsSch("SD_FRI_HRS4") & " to " & medHours4(5) & ". "
                xWeek4Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK4
            End If
            If rsSch("SD_SAT_HRS4") <> medHours4(6) Then
                xWK4HrsChanged = xWK4HrsChanged & "Saturday from " & rsSch("SD_SAT_HRS4") & " to " & medHours4(6) & ". "
                xWeek4Changed = True
                Has_Hours_Changed_For_Log = True
                'GoTo NextDayWK4
            End If
'NextDayWK4:
        'Next
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

Private Sub RebuildWorkScheduleDetail(xEmpNo, xFromDate, xToDate, xChangeDate, Optional xLastDate)
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
    
    'Last Date is the last date the Unapproved/Rejected Requests records found in the table. So any changes to the
    'WS Detail tables can be done after this Last Date.
    
    'Last Date passed then
        'Delete the WS Details from the Last Date onward upto To Date for the From Date to To Date Period
    'If No Last Date then
        'Delete the existing WS Details for the From Date to To Date period only
    'Rebuild the WS Details from
        'Last Date, if Last Date passed, upto To Date
        'From Date to To Date, if Last Date not passed.
            
    'For details found, look for what has changed.
        '- Verify when the Change should be effective.
        '- Check if any Request exists from the Changed Date.
            '- if so then do not allow to save the change
            '- if no then save the change but from the Changed Date onwards for the respective weeks only. Also update
            '  the Changed Date as well for that record entry
    
    'Delete the existing records for the From / To Date Period and based on the Last Date if passed
    SQLQ = "DELETE FROM HR_SCHEDULER_DETAIL WHERE WS_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " AND WS_FDATE = " & Date_SQL(xFromDate)
    SQLQ = SQLQ & " AND WS_TDATE = " & Date_SQL(xToDate)
    If Not IsMissing(xLastDate) Then
        SQLQ = SQLQ & " AND WS_DATE > " & Date_SQL(xLastDate)
    End If
    gdbAdoIhr001.Execute SQLQ
    
    'If Last Date NOT missing
    If Not IsMissing(xLastDate) Then
        'Check if the Last Date is less than the To Date then add the remaining WS details records
        If CVDate(xLastDate) < CVDate(xToDate) Then
            xWeekNo = 0
            xWeekDayCount = 0
            xLstWSWkDay = 0
            
            SQLQ = "SELECT * FROM HR_SCHEDULER_DETAIL WHERE WS_EMPNBR = " & xEmpNo
            SQLQ = SQLQ & " AND WS_FDATE = " & Date_SQL(xFromDate)
            SQLQ = SQLQ & " AND WS_TDATE = " & Date_SQL(xToDate)
            rsSchDetail2.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsSchDetail2.EOF Then
                'Get the last record's Week No and then count the # of days details records exists for that Week #
                rsSchDetail2.MoveLast
                xWeekNo = rsSchDetail2("WS_WEEKNO")
                xWeekDayCount = xWeekDayCount + 1
                
                'Last record's WS Date and Day of the Week.
                xLstWSDate = CVDate(rsSchDetail2("WS_DATE"))
                xLstWSWkDay = Weekday(rsSchDetail2("WS_DATE"))
                
                rsSchDetail2.MovePrevious
                Do While Not rsSchDetail2.BOF
                    'Count the # of records exists for the same Week # so we know from what day to add more details
                    If xWeekNo = rsSchDetail2("WS_WEEKNO") Then
                        xWeekDayCount = xWeekDayCount + 1
                    Else
                        'Move back to same Week #
                        rsSchDetail2.MoveNext
                        
                        'Moved up where the last day is read first (MoveLast), as we want to start from that day
                        'Last WS Date
                        'xLstWSDate = CVDate(rsSchDetail2("WS_DATE"))
                        'xLstWSWkDay = Weekday(rsSchDetail2("WS_DATE"))
                        Exit Do
                    End If
                    
                    rsSchDetail2.MovePrevious
                Loop
                
                'Add the rest of the week's detail records upto the new To Date
                'Week # to start entering the detail records from
                If xWeekDayCount = 7 Then
                    If xWeekNo = 4 Then
                        xWeekNo = 1
                    Else
                        xWeekNo = xWeekNo + 1
                    End If
                End If
                
                'Day to start adding the WS details
                If xLstWSWkDay < 7 Then
                    xStartDay = xLstWSWkDay
                Else
                    xStartDay = 0
                End If
                
                'Start adding additional day's WS Details records
                Do While CVDate(xLstWSDate) < CVDate(xToDate)
                    'Start from the Week the last day WS detail record exists
                    If xWeekNo = 1 Then xWeekNo = 0: GoTo Week1     'Reset the WeekNo for next round
                    If xWeekNo = 2 Then xWeekNo = 0: GoTo Week2
                    If xWeekNo = 3 Then xWeekNo = 0: GoTo Week3
                    If xWeekNo = 4 Then xWeekNo = 0: GoTo Week4
Week1:
                    If frWeek1.Enabled Then
                        For xDay = xStartDay To 6
                            'Add Work Schedule for the day in a week
                            rsSchDetail2.AddNew
                            rsSchDetail2("WS_EMPNBR") = xEmpNo
                            rsSchDetail2("WS_FDATE") = xFromDate
                            rsSchDetail2("WS_TDATE") = xToDate
                            rsSchDetail2("WS_CHGDATE") = xChangeDate   'Change from Date will be same as From Date the first time
                        
                            rsSchDetail2("WS_WEEKNO") = 1
                            rsSchDetail2("WS_DATE") = DateAdd("d", 1, xLstWSDate)   'Because started with -1 day as Start Day
                            rsSchDetail2("WS_HRS") = medHours(xDay)
                        
                            rsSchDetail2("WS_LDATE") = Date
                            rsSchDetail2("WS_LTIME") = Time$
                            rsSchDetail2("WS_LUSER") = glbUserID
                            rsSchDetail2.Update
                            
                            'ReSet the xLstWSDate
                            xLstWSDate = CVDate(rsSchDetail2("WS_DATE"))
                            
                            'Changed from > to >= it was adding an extra day after To Date
                            'Exit the Loop if xLstWSDate exceeds the ToDate.
                            If CVDate(xLstWSDate) >= CVDate(xToDate) Then Exit Do
                        Next
                        
                        'Reset the start day for the next Week
                        xStartDay = 0
                    End If
                    
Week2:
                    If frWeek2.Enabled Then
                        For xDay = xStartDay To 6
                            'Add Work Schedule for the day in a week
                            rsSchDetail2.AddNew
                            rsSchDetail2("WS_EMPNBR") = xEmpNo
                            rsSchDetail2("WS_FDATE") = xFromDate
                            rsSchDetail2("WS_TDATE") = xToDate
                            rsSchDetail2("WS_CHGDATE") = xChangeDate   'Change from Date will be same as From Date the first time
                        
                            rsSchDetail2("WS_WEEKNO") = 2
                            rsSchDetail2("WS_DATE") = DateAdd("d", 1, xLstWSDate)   'Because started with -1 day as Start Day
                            rsSchDetail2("WS_HRS") = medHours2(xDay)
                        
                            rsSchDetail2("WS_LDATE") = Date
                            rsSchDetail2("WS_LTIME") = Time$
                            rsSchDetail2("WS_LUSER") = glbUserID
                            rsSchDetail2.Update
                            
                            'ReSet the xLstWSDate
                            xLstWSDate = CVDate(rsSchDetail2("WS_DATE"))
                        
                            'Changed from > to >= it was adding an extra day after To Date
                            'Exit the Loop if xLstWSDate exceeds the ToDate.
                            If CVDate(xLstWSDate) >= CVDate(xToDate) Then Exit Do
                        Next
                        'Reset the start day for the next Week
                        xStartDay = 0
                    Else
                        'If this week is not visible so the rest of the Weeks will not be visible too
                        GoTo NextWorkSchSet1
                    End If
                    
Week3:
                    If frWeek3.Enabled Then
                        For xDay = xStartDay To 6
                            'Add Work Schedule for the day in a week
                            rsSchDetail2.AddNew
                            rsSchDetail2("WS_EMPNBR") = xEmpNo
                            rsSchDetail2("WS_FDATE") = xFromDate
                            rsSchDetail2("WS_TDATE") = xToDate
                            rsSchDetail2("WS_CHGDATE") = xChangeDate   'Change from Date will be same as From Date the first time
                        
                            rsSchDetail2("WS_WEEKNO") = 3
                            rsSchDetail2("WS_DATE") = DateAdd("d", 1, xLstWSDate)   'Because started with -1 day as Start Day
                            rsSchDetail2("WS_HRS") = medHours3(xDay)
                        
                            rsSchDetail2("WS_LDATE") = Date
                            rsSchDetail2("WS_LTIME") = Time$
                            rsSchDetail2("WS_LUSER") = glbUserID
                            rsSchDetail2.Update
                            
                            'ReSet the xLstWSDate
                            xLstWSDate = CVDate(rsSchDetail2("WS_DATE"))
                            
                            'Changed from > to >= it was adding an extra day after To Date
                            'Exit the Loop if xLstWSDate exceeds the ToDate.
                            If CVDate(xLstWSDate) >= CVDate(xToDate) Then Exit Do
                        Next
                        'Reset the start day for the next Week
                        xStartDay = 0
                    Else
                        'If this week is not visible so the rest of the Weeks will not be visible too
                        GoTo NextWorkSchSet1
                    End If
                    
Week4:
                    If frWeek4.Enabled Then
                        For xDay = xStartDay To 6
                            'Add Work Schedule for the day in a week
                            rsSchDetail2.AddNew
                            rsSchDetail2("WS_EMPNBR") = xEmpNo
                            rsSchDetail2("WS_FDATE") = xFromDate
                            rsSchDetail2("WS_TDATE") = xToDate
                            rsSchDetail2("WS_CHGDATE") = xChangeDate   'Change from Date will be same as From Date the first time
                        
                            rsSchDetail2("WS_WEEKNO") = 4
                            rsSchDetail2("WS_DATE") = DateAdd("d", 1, xLstWSDate)   'Because started with -1 day as Start Day
                            rsSchDetail2("WS_HRS") = medHours4(xDay)
                        
                            rsSchDetail2("WS_LDATE") = Date
                            rsSchDetail2("WS_LTIME") = Time$
                            rsSchDetail2("WS_LUSER") = glbUserID
                            rsSchDetail2.Update
                            
                            'ReSet the xLstWSDate
                            xLstWSDate = CVDate(rsSchDetail2("WS_DATE"))
                            
                            'Changed from > to >= it was adding an extra day after To Date
                            'Exit the Loop if xLstWSDate exceeds the ToDate.
                            If CVDate(xLstWSDate) >= CVDate(xToDate) Then Exit Do
                        Next
                        'Reset the start day for the next Week
                        xStartDay = 0
                    End If
NextWorkSchSet1:
                Loop
            End If
            rsSchDetail2.Close
            Set rsSchDetail2 = Nothing
        End If
    Else
        'Last Date is missing, add all the WS Details starting From Date to To Date
        SQLQ = "SELECT * FROM HR_SCHEDULER_DETAIL WHERE WS_EMPNBR = " & xEmpNo
        SQLQ = SQLQ & " AND WS_FDATE = " & Date_SQL(xFromDate)
        SQLQ = SQLQ & " AND WS_TDATE = " & Date_SQL(xToDate)
        rsSchDetail.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
        If rsSchDetail.EOF Then
            'No Work Schedule detail found. Create details
            
            'Set the Start Date of the Period so for each date upto the To Date the Work Schedule can be prepared
            xCurrDate = CVDate(xFromDate) - 1   'Starting with -1 because when adding Details, I am +1 day for WS_DATE
            
            'Get the day of the Week to start from
            xLstWSWkDay = 0
            xLstWSWkDay = Weekday(xCurrDate)
            
            'Compute the Day to start adding the WS details
            If xLstWSWkDay < 7 Then
                xStartDay = xLstWSWkDay
            Else
                xStartDay = 0
            End If
            
            
            Do While CVDate(xCurrDate) < CVDate(xToDate)
                
                If frWeek1.Enabled Then
                    For xDay = xStartDay To 6
                        'Add Work Schedule for the day in a week
                        rsSchDetail.AddNew
                        rsSchDetail("WS_EMPNBR") = xEmpNo
                        rsSchDetail("WS_FDATE") = xFromDate
                        rsSchDetail("WS_TDATE") = xToDate
                        rsSchDetail("WS_CHGDATE") = xChangeDate   'Change from Date will be same as From Date the first time
                    
                        rsSchDetail("WS_WEEKNO") = 1
                        rsSchDetail("WS_DATE") = DateAdd("d", 1, xCurrDate) 'Because started with -1 day as Start Day
                        rsSchDetail("WS_HRS") = medHours(xDay)
                    
                        rsSchDetail("WS_LDATE") = Date
                        rsSchDetail("WS_LTIME") = Time$
                        rsSchDetail("WS_LUSER") = glbUserID
                        rsSchDetail.Update
                        
                        'ReSet the CurrDate
                        xCurrDate = CVDate(rsSchDetail("WS_DATE"))
                        
                        'Changed from > to >= it was adding an extra day after To Date
                        'Exit the Loop if xCurrDate exceeds the ToDate.
                        If CVDate(xCurrDate) >= CVDate(xToDate) Then Exit Do
                    Next
                    
                    'Reset the start day for the next Week
                    xStartDay = 0
                End If
                
                If frWeek2.Enabled Then
                    For xDay = 0 To 6
                        'Add Work Schedule for the day in a week
                        rsSchDetail.AddNew
                        rsSchDetail("WS_EMPNBR") = xEmpNo
                        rsSchDetail("WS_FDATE") = xFromDate
                        rsSchDetail("WS_TDATE") = xToDate
                        rsSchDetail("WS_CHGDATE") = xChangeDate   'Change from Date will be same as From Date the first time
                    
                        rsSchDetail("WS_WEEKNO") = 2
                        rsSchDetail("WS_DATE") = DateAdd("d", 1, xCurrDate) 'Because started with -1 day as Start Day
                        rsSchDetail("WS_HRS") = medHours2(xDay)
                    
                        rsSchDetail("WS_LDATE") = Date
                        rsSchDetail("WS_LTIME") = Time$
                        rsSchDetail("WS_LUSER") = glbUserID
                        rsSchDetail.Update
                        
                        'ReSet the CurrDate
                        xCurrDate = CVDate(rsSchDetail("WS_DATE"))
                    
                        'Changed from > to >= it was adding an extra day after To Date
                        'Exit the Loop if xCurrDate exceeds the ToDate.
                        If CVDate(xCurrDate) >= CVDate(xToDate) Then Exit Do
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
                        rsSchDetail("WS_FDATE") = xFromDate
                        rsSchDetail("WS_TDATE") = xToDate
                        rsSchDetail("WS_CHGDATE") = xChangeDate   'Change from Date will be same as From Date the first time
                    
                        rsSchDetail("WS_WEEKNO") = 3
                        rsSchDetail("WS_DATE") = DateAdd("d", 1, xCurrDate) 'Because started with -1 day as Start Day
                        rsSchDetail("WS_HRS") = medHours3(xDay)
                    
                        rsSchDetail("WS_LDATE") = Date
                        rsSchDetail("WS_LTIME") = Time$
                        rsSchDetail("WS_LUSER") = glbUserID
                        rsSchDetail.Update
                        
                        'ReSet the CurrDate
                        xCurrDate = CVDate(rsSchDetail("WS_DATE"))
                        
                        'Changed from > to >= it was adding an extra day after To Date
                        'Exit the Loop if xCurrDate exceeds the ToDate.
                        If CVDate(xCurrDate) >= CVDate(xToDate) Then Exit Do
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
                        rsSchDetail("WS_FDATE") = xFromDate
                        rsSchDetail("WS_TDATE") = xToDate
                        rsSchDetail("WS_CHGDATE") = xChangeDate   'Change from Date will be same as From Date the first time
                    
                        rsSchDetail("WS_WEEKNO") = 4
                        rsSchDetail("WS_DATE") = DateAdd("d", 1, xCurrDate) 'Because started with -1 day as Start Day
                        rsSchDetail("WS_HRS") = medHours4(xDay)
                    
                        rsSchDetail("WS_LDATE") = Date
                        rsSchDetail("WS_LTIME") = Time$
                        rsSchDetail("WS_LUSER") = glbUserID
                        rsSchDetail.Update
                        
                        'ReSet the CurrDate
                        xCurrDate = CVDate(rsSchDetail("WS_DATE"))
                        
                        'Changed from > to >= it was adding an extra day after To Date
                        'Exit the Loop if xCurrDate exceeds the ToDate.
                        If CVDate(xCurrDate) >= CVDate(xToDate) Then Exit Do
                    Next
                    
                    'Reset the start day for the next Week
                    xStartDay = 0
                End If
                
NextWorkSchSet:
            Loop
        End If
        rsSchDetail.Close
        Set rsSchDetail = Nothing
    End If
    
End Sub
